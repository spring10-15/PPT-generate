import { z } from "zod";
import {
  SUPPORTED_LAYOUT_TYPES,
  type GeneratorInput,
  type OutlineDocument,
  type ParsedDocument,
  type ParsedSection,
  type SlideLayoutType
} from "@/lib/types";
import { normalizeOutlineDocument, stripTitlePrefix } from "@/lib/outline-format";
import {
  buildDirectoryBrief,
  buildPageTitle,
  buildStructuredContent,
  composeSlideSummary
} from "@/lib/slide-content";

const API_URL =
  process.env.SILICONFLOW_API_URL ?? "https://api.siliconflow.cn/v1/chat/completions";
const MODEL = process.env.SILICONFLOW_MODEL ?? "Pro/zai-org/GLM-5";
const LAYOUT_OPTIONS_TEXT = SUPPORTED_LAYOUT_TYPES.join(" / ");
const TARGET_DIRECTORY_LIMIT = 3;
const TEXT_VARIETY_SEQUENCE: SlideLayoutType[] = ["overview", "vertical", "three-column", "split-grid"];
const PROGRESS_CUE_RE = /(步骤|阶段|推进|路径|计划|里程碑|排期|节奏|实施|动作|举措|落地)/;
const HIERARCHY_CUE_RE = /(体系|架构|分层|能力|层级|模块|框架|分类|结构|版图|矩阵)/;
const GENERIC_OVERVIEW_TITLE_RE = /^(概览|方案思路)$/u;
const WEAK_TITLE_END_RE = /(与|及|和|或|的|等|及其|以及|并|并且|在|对|向|为)$/u;
const TITLE_SEGMENT_SPLIT_RE = /[，,：:；;（）()【】/]|(?:\s*[-—－]\s*)|(?:与|及|和|或|以及|并且|并)/u;
const TITLE_KEYWORD_RE =
  /[\u4e00-\u9fa5A-Za-z0-9]{1,10}?(?:场景|应用|方案|协作|价值|流程|绑定|内容|能力|体系|架构|矩阵|模块|数据|开发|运营|管理|办公|创作|服务|平台|合规|治理|安全|知识|任务|团队|文件|业务|项目)/gu;
const SEMANTIC_SUFFIXES = [
  "知识库",
  "场景",
  "应用",
  "方案",
  "协作",
  "价值",
  "流程",
  "绑定",
  "内容",
  "能力",
  "体系",
  "架构",
  "矩阵",
  "模块",
  "数据",
  "开发",
  "运营",
  "管理",
  "办公",
  "创作",
  "服务",
  "平台",
  "合规",
  "治理",
  "安全",
  "知识",
  "任务",
  "团队",
  "文件",
  "业务",
  "项目"
];

const outlineSchema = z.object({
  cover: z.object({
    title: z.string().min(1),
    subtitle: z.string().min(1)
  }),
  directory: z.array(
    z.object({
      title: z.string().min(1),
      description: z.string().min(1),
      slideCount: z.number().int().min(1)
    })
  ),
  slides: z.array(
    z.object({
      directoryTitle: z.string().min(1),
      title: z.string().min(1),
      type: z.enum(SUPPORTED_LAYOUT_TYPES),
      content: z.object({
        intro: z.string().min(1),
        regularTitle: z.string().min(1),
        description: z.string().min(1)
      }),
      imageSuggestion: z.string().min(1)
    })
  )
});

type AIOutline = z.infer<typeof outlineSchema>;
type PlanningSection = ParsedSection & {
  childSections: ParsedSection[];
  sourceSections: ParsedSection[];
};
type PlanningGroup = {
  title: string;
  sections: ParsedSection[];
};
type PagePlan = {
  titleHint: string;
  regularTitleHint: string;
  introHint: string;
  summary: string;
  isOverviewLead: boolean;
};

function isPlanningSection(section: ParsedSection | PlanningSection): section is PlanningSection {
  return Array.isArray((section as PlanningSection).childSections);
}

function normalizeTitle(title: string) {
  return title.replace(/\s+/g, " ").trim();
}

function cleanSectionTitle(title: string) {
  return stripTitlePrefix(normalizeTitle(title)).replace(/（[^）]*）|\([^)]*\)/gu, "").trim();
}

function trimWeakTitleEnding(text: string) {
  let current = text.trim();

  while (current && WEAK_TITLE_END_RE.test(current)) {
    current = current.replace(WEAK_TITLE_END_RE, "").trim();
  }

  return current;
}

function buildSuffixFragments(text: string, maxLength: number) {
  const candidates: string[] = [];

  SEMANTIC_SUFFIXES.forEach((suffix) => {
    const index = text.lastIndexOf(suffix);
    if (index < 0) {
      return;
    }

    const before = text.slice(0, index).trim();
    const splitParts = before
      .split(/(?:与|及|和|并|以及|级|层|流|端|型|类)/u)
      .map((part) => trimWeakTitleEnding(part.trim()))
      .filter(Boolean);
    const baseCandidates = [
      splitParts[0],
      splitParts[splitParts.length - 1],
      before.slice(-2),
      before.slice(-3),
      before.slice(0, Math.max(2, maxLength - suffix.length))
    ]
      .map((item) => trimWeakTitleEnding(item ?? ""))
      .filter(Boolean);

    baseCandidates.forEach((base) => {
      const candidate = trimWeakTitleEnding(`${base}${suffix}`);
      if (candidate && candidate.length <= maxLength) {
        candidates.push(candidate);
      }
    });
  });

  return candidates.filter((item, index) => item && candidates.indexOf(item) === index);
}

function collectSemanticFragments(title: string, maxLength: number) {
  const cleaned = cleanSectionTitle(title);
  if (!cleaned) {
    return [];
  }

  const suffixFragments = buildSuffixFragments(cleaned, maxLength);
  const keywordMatches = Array.from(cleaned.matchAll(TITLE_KEYWORD_RE))
    .map((match) => trimWeakTitleEnding(match[0].trim()))
    .filter((item) => item.length <= maxLength);
  const segments = cleaned
    .split(TITLE_SEGMENT_SPLIT_RE)
    .map((segment) => trimWeakTitleEnding(segment.trim()))
    .filter((item) => item && item.length <= maxLength);
  const candidates = [
    ...suffixFragments,
    ...keywordMatches,
    ...segments,
    trimWeakTitleEnding(cleaned.slice(0, maxLength).trim())
  ];

  return candidates.filter((item, index) => item && candidates.indexOf(item) === index);
}

function shortenTitleFragment(title: string, maxLength = 6) {
  const cleaned = cleanSectionTitle(title);
  if (!cleaned) {
    return "";
  }

  if (cleaned.length <= maxLength) {
    return trimWeakTitleEnding(cleaned);
  }

  return collectSemanticFragments(cleaned, maxLength)[0] ?? trimWeakTitleEnding(cleaned.slice(0, maxLength).trim());
}

function combineFragments(fragments: string[], maxLength: number) {
  let combined = "";

  for (const fragment of fragments) {
    const next = `${combined}${fragment}`;
    if (next.length > maxLength) {
      break;
    }
    combined = next;
  }

  return combined;
}

function buildPlanningDirectoryBrief(section: PlanningSection) {
  const title = cleanSectionTitle(section.title);
  const titleFragments = section.childSections
    .filter((child) => !GENERIC_OVERVIEW_TITLE_RE.test(cleanSectionTitle(child.title)))
    .map((child) => shortenTitleFragment(child.title, 5))
    .filter(Boolean)
    .reduce<string[]>((result, item) => {
      if (!result.includes(item)) {
        result.push(item);
      }
      return result;
    }, []);
  const combined = combineFragments(titleFragments, 10);

  return (
    combined ||
    buildDirectoryBrief({
      title,
      summary: section.childSections
        .slice(0, 3)
        .map((child) => child.paragraphs[0] || child.title)
        .filter(Boolean)
        .join("；") || section.paragraphs.join("；")
    })
  );
}

function buildOverviewIntro(section: PlanningSection) {
  const leadParagraphs = [
    ...(section.childSections[0]?.paragraphs ?? []),
    ...(section.paragraphs ?? [])
  ].filter(Boolean);
  const leadText =
    leadParagraphs.sort((left, right) => right.length - left.length)[0] ??
    "";
  const normalized = normalizeTitle(leadText)
    .replace(/下面.+$/u, "")
    .replace(/适合.+$/u, "")
    .trim();
  const structured = buildStructuredContent({
    title: cleanSectionTitle(section.title),
    intro: normalized,
    summary: normalized
  });
  const afterContrast = structured.intro.split(/而是/u)[1]?.trim() || structured.intro;

  return afterContrast.slice(0, 30);
}

function buildOverviewRegularTitle(section: PlanningSection) {
  const title = cleanSectionTitle(section.title);
  const fragments = section.childSections
    .filter((child) => !GENERIC_OVERVIEW_TITLE_RE.test(cleanSectionTitle(child.title)))
    .map((child) => shortenTitleFragment(child.title, 5))
    .filter(Boolean)
    .reduce<string[]>((result, item) => {
      if (!result.includes(item)) {
        result.push(item);
      }
      return result;
    }, []);
  const combined = combineFragments(fragments.slice(0, 2), 15);

  return combined || `${shortenTitleFragment(title, 10) || title}要点`.slice(0, 15);
}

function buildOverviewDescription(section: PlanningSection) {
  const actualChildren = section.childSections.filter(
    (child, index) => index > 0 || !GENERIC_OVERVIEW_TITLE_RE.test(cleanSectionTitle(child.title))
  );

  return actualChildren
    .slice(0, 2)
    .map((child) => {
      const title = cleanSectionTitle(child.title);
      const summary = splitIntoUnits(child.paragraphs.join("；")).slice(0, 1).join("；") || child.paragraphs[0] || title;
      return `${title}：${summary}`;
    })
    .join("；");
}

function toMatchKey(value: string | undefined) {
  return normalizeTitle(value ?? "").replace(/[：:；;。！？,.，、】【（）()[\]\s-]/g, "").toLowerCase();
}

function matchSectionByTitle<T extends { title: string }>(title: string, sections: T[]) {
  const target = toMatchKey(title);
  if (!target) {
    return undefined;
  }

  return (
    sections.find((section) => toMatchKey(section.title) === target) ??
    sections.find((section) => {
      const key = toMatchKey(section.title);
      return key.includes(target) || target.includes(key);
    })
  );
}

function hasSectionContent(section: ParsedSection) {
  return section.paragraphs.length > 0 || section.tables.length > 0 || section.imageIds.length > 0;
}

function buildPlanningSection(title: string, sections: ParsedSection[], level = 1): PlanningSection {
  const contentSections = sections.filter(hasSectionContent);
  const primary = contentSections[0] ?? sections[0];

  return {
    id: primary?.id ?? `planning-${normalizeTitle(title)}`,
    title: normalizeTitle(title),
    level,
    paragraphs: contentSections.flatMap((section) => section.paragraphs).filter(Boolean),
    tables: contentSections.flatMap((section) => section.tables),
    imageIds: Array.from(new Set(contentSections.flatMap((section) => section.imageIds))),
    isHeadingOnly: false,
    childSections: contentSections.length > 0 ? contentSections : sections,
    sourceSections: sections
  };
}

function buildPlanningGroups(section: PlanningSection): PlanningGroup[] {
  const groups: PlanningGroup[] = [];
  let currentGroup: PlanningGroup | null = null;
  const leadingSections: ParsedSection[] = [];

  section.sourceSections.forEach((sourceSection) => {
    if (!hasSectionContent(sourceSection) && !normalizeTitle(sourceSection.title)) {
      return;
    }

    if (sourceSection.level === 2) {
      if (currentGroup) {
        groups.push(currentGroup);
      }

      currentGroup = {
        title: cleanSectionTitle(sourceSection.title) || cleanSectionTitle(section.title),
        sections: hasSectionContent(sourceSection) ? [sourceSection] : []
      };
      return;
    }

    if (sourceSection.level > 2) {
      if (!currentGroup) {
        currentGroup = {
          title: cleanSectionTitle(section.title),
          sections: []
        };
      }

      if (hasSectionContent(sourceSection)) {
        currentGroup.sections.push(sourceSection);
      }
      return;
    }

    if (hasSectionContent(sourceSection)) {
      if (currentGroup) {
        currentGroup.sections.push(sourceSection);
      } else {
        leadingSections.push(sourceSection);
      }
    }
  });

  if (currentGroup) {
    groups.push(currentGroup);
  }

  if (groups.length === 0) {
    return [
      {
        title: cleanSectionTitle(section.title),
        sections: section.childSections
      }
    ];
  }

  if (leadingSections.length > 0) {
    groups[0] = {
      ...groups[0],
      sections: [...leadingSections, ...groups[0].sections]
    };
  }

  return groups.filter((group) => group.sections.length > 0);
}

function estimateGroupWeight(group: PlanningGroup) {
  const paragraphWeight = group.sections.reduce((sum, section) => sum + Math.max(1, section.paragraphs.length), 0);
  return Math.max(1, group.sections.length * 2 + paragraphWeight);
}

function packGroupsIntoPages(groups: PlanningGroup[], pageCount: number) {
  if (groups.length === 0) {
    return Array.from({ length: pageCount }, () => [] as PlanningGroup[]);
  }

  const pages: PlanningGroup[][] = [];
  const weights = groups.map(estimateGroupWeight);
  let startIndex = 0;

  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const pagesLeft = pageCount - pageIndex;
    const groupsLeft = groups.length - startIndex;

    if (groupsLeft <= 0) {
      pages.push([]);
      continue;
    }

    if (pagesLeft === 1) {
      pages.push(groups.slice(startIndex));
      break;
    }

    const remainingWeight = weights.slice(startIndex).reduce((sum, weight) => sum + weight, 0);
    const targetWeight = remainingWeight / pagesLeft;
    const current: PlanningGroup[] = [];
    let currentWeight = 0;

    while (startIndex < groups.length) {
      const remainingGroupsAfterCurrent = groups.length - (startIndex + 1);
      const mustLeave = pagesLeft - 1;

      current.push(groups[startIndex]);
      currentWeight += weights[startIndex];
      startIndex += 1;

      if (remainingGroupsAfterCurrent < mustLeave) {
        continue;
      }

      if (currentWeight >= targetWeight) {
        break;
      }
    }

    pages.push(current);
  }

  return pages;
}

function buildGroupSummary(group: PlanningGroup) {
  return group.sections
    .slice(0, 4)
    .map((child) => {
      const title = cleanSectionTitle(child.title);
      const summary = splitIntoUnits(child.paragraphs.join("；")).slice(0, 1).join("；") || child.paragraphs[0] || title;
      return `${title}：${summary}`;
    })
    .join("；");
}

function buildChunkTitleHint(groupChunk: PlanningGroup[], fallbackTitle: string) {
  const fragments = groupChunk
    .map((group) => shortenTitleFragment(group.title))
    .filter(Boolean)
    .reduce<string[]>((result, item) => {
      if (!result.includes(item)) {
        result.push(item);
      }
      return result;
    }, []);
  const combined = combineFragments(fragments.slice(0, 2), 10);

  return combined || shortenTitleFragment(fallbackTitle);
}

function buildPagePlans(section: PlanningSection, pageCount: number): PagePlan[] {
  const groups = buildPlanningGroups(section);
  const packedGroups = packGroupsIntoPages(groups, Math.max(1, pageCount));

  return packedGroups.map((groupChunk, pageIndex) => {
    const firstGroup = groupChunk[0];
    const firstSection = firstGroup?.sections[0];
    const sectionTitle = firstGroup ? shortenTitleFragment(firstGroup.title) : shortenTitleFragment(section.title);
    const summary = groupChunk.map(buildGroupSummary).filter(Boolean).join("；") || `${cleanSectionTitle(section.title)}相关内容`;
    const isOverviewLead =
      pageIndex === 0 &&
      firstSection &&
      GENERIC_OVERVIEW_TITLE_RE.test(cleanSectionTitle(firstSection.title));
    const chunkTitle = buildChunkTitleHint(groupChunk, section.title);

    return {
      titleHint: groupChunk.length > 1 ? chunkTitle : sectionTitle || cleanSectionTitle(section.title),
      regularTitleHint: groupChunk.length > 1 ? chunkTitle : sectionTitle || cleanSectionTitle(section.title),
      introHint:
        firstSection?.paragraphs[0] ??
        groupChunk.flatMap((group) => group.sections).flatMap((child) => child.paragraphs)[0] ??
        section.paragraphs[0] ??
        "",
      summary,
      isOverviewLead
    };
  });
}

function mergePlanningSections(groups: PlanningSection[], maxDirectories: number) {
  if (groups.length <= maxDirectories) {
    return groups;
  }

  const groupSize = Math.ceil(groups.length / maxDirectories);
  const merged: PlanningSection[] = [];

  for (let index = 0; index < groups.length; index += groupSize) {
    const chunk = groups.slice(index, index + groupSize);
    const first = chunk[0];
    const last = chunk[chunk.length - 1];
    const combinedTitle =
      chunk.length === 1
        ? first.title
        : normalizeTitle(first.title) || normalizeTitle(last.title) || `目录${merged.length + 1}`;
    const childSections = chunk.flatMap((group) => group.childSections);
    merged.push(buildPlanningSection(combinedTitle, childSections, first.level));
  }

  return merged;
}

function groupSectionsForPlanning(sections: ParsedSection[], maxDirectories: number): PlanningSection[] {
  const normalizedSections = sections.filter(
    (section) => normalizeTitle(section.title) || hasSectionContent(section)
  );
  const levelOneAnchors = normalizedSections.filter(
    (section, index) =>
      section.level === 1 &&
      !(index === 0 && normalizeTitle(section.title) === "概览")
  );

  if (levelOneAnchors.length > 0) {
    const groups: PlanningSection[] = [];
    const leadingSections = normalizedSections.filter(
      (section) =>
        hasSectionContent(section) &&
        normalizedSections.indexOf(section) < normalizedSections.indexOf(levelOneAnchors[0])
    );

    levelOneAnchors.forEach((anchor, anchorIndex) => {
      const anchorPosition = normalizedSections.indexOf(anchor);
      const nextAnchorPosition =
        anchorIndex < levelOneAnchors.length - 1
          ? normalizedSections.indexOf(levelOneAnchors[anchorIndex + 1])
          : normalizedSections.length;
      const contentSections = normalizedSections.slice(anchorPosition + 1, nextAnchorPosition).filter(hasSectionContent);
      const groupSections = anchorIndex === 0 ? [...leadingSections, ...contentSections] : contentSections;

      if (groupSections.length > 0) {
        groups.push(buildPlanningSection(anchor.title, groupSections, anchor.level));
      }
    });

    if (groups.length > 0) {
      return mergePlanningSections(groups, maxDirectories);
    }
  }

  const contentSections = normalizedSections.filter(hasSectionContent);
  if (contentSections.length <= maxDirectories) {
    return contentSections.map((section) => buildPlanningSection(section.title, [section], section.level));
  }

  const groupSize = Math.ceil(contentSections.length / maxDirectories);
  const merged: PlanningSection[] = [];

  for (let index = 0; index < contentSections.length; index += groupSize) {
    const group = contentSections.slice(index, index + groupSize);
    const first = group[0];
    const last = group[group.length - 1];
    const mergedTitle =
      group.length === 1
        ? normalizeTitle(first.title)
        : `${normalizeTitle(first.title)}-${normalizeTitle(last.title)}`;

    merged.push(buildPlanningSection(mergedTitle, group, first.level));
  }

  return merged;
}

function trimSections(sections: PlanningSection[]) {
  return sections.slice(0, 16).map((section) => ({
    title: section.title,
    paragraphs:
      section.childSections.length > 0
        ? section.childSections
            .slice(0, 4)
            .map((child) => [child.title, ...child.paragraphs.slice(0, 1)].filter(Boolean).join("："))
        : section.paragraphs.slice(0, 4),
    hasTable: section.tables.length > 0,
    hasImage: section.imageIds.length > 0
  }));
}

function buildPrompt(doc: ParsedDocument, input: GeneratorInput, detailPages: number): string {
  const planningSections = groupSectionsForPlanning(
    doc.sections,
    Math.min(TARGET_DIRECTORY_LIMIT, detailPages)
  );

  return [
    "你是一个严格遵循企业汇报材料的 PPT 架构规划器。",
    "只能基于用户文档内容和少量通用连接表述进行重组，不能引入新事实。",
    "输出必须是 JSON，禁止输出 Markdown 代码块。",
    "目录标题和页标题只返回纯标题文本，不要自带任何数字编号、括号编号或层级编号，系统会统一编号。",
    "目录简要说明、内容简介、常规标题、说明性文字都只返回纯文本，不要自带编号。",
    "目录和页内容必须严格按原文顺序组织，不能打乱先后顺序。",
    "同一页内的页标题、内容简介、常规标题、说明性文字不要重复表达同一句内容。",
    "",
    "任务约束：",
    `- 总页数: ${input.totalPages}，其中封面 1 页，目录 1 页，详情页 ${detailPages} 页`,
    "- 目录项可以跨多页，但不能把多个目录项合并在同一页",
    `- 目录项尽可能不超过 ${Math.min(TARGET_DIRECTORY_LIMIT, detailPages)} 个，优先聚合相近内容`,
    "- 如果信息过多，优先压缩摘要表达，不要删除目录结构",
    `- 页型只能从 ${LAYOUT_OPTIONS_TEXT} 里选，同一种页型可以重复使用`,
    "- 页型要与内容量匹配：简短概述优先 overview / two-column，中量内容优先 three-column / split-grid，高密内容优先 four-column / progress / hierarchy",
    "- image-left / image-right / image 页只有文档里确有图片时才优先使用，否则用 overview 或 vertical",
    "- table 页只有文档里确有表格时才优先使用",
    "- 不要因为字数限制主动裁短标题或摘要，优先保留原始语义完整性",
    "- 不要输出强调标题字段",
    "- 目标受众只用于帮助理解汇报语气，不要直接写进封面标题或封面副标题",
    "",
    "请返回 JSON：",
    JSON.stringify(
      {
        cover: { title: "封面标题", subtitle: "封面副标题" },
        directory: [{ title: "目录项", description: "目录简述", slideCount: 1 }],
        slides: [
          {
            directoryTitle: "目录项",
            title: "页标题",
            type: "overview",
            content: {
              intro: "内容简介",
              regularTitle: "常规标题",
              description: "说明性文字"
            },
            imageSuggestion: "配图建议或占位说明"
          }
        ]
      },
      null,
      2
    ),
    "",
    "用户信息：",
    JSON.stringify(
      {
        userName: input.userName,
        targetAudience: input.targetAudience,
        estimatedMinutes: input.estimatedMinutes,
        titleGuess: doc.titleGuess,
        subtitleGuess: doc.subtitleGuess,
        sections: trimSections(planningSections)
      },
      null,
      2
    )
  ].join("\n");
}

async function requestOutline(doc: ParsedDocument, input: GeneratorInput, detailPages: number) {
  const apiKey = process.env.SILICONFLOW_API_KEY;

  if (!apiKey) {
    return null;
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 20000);

  const response = await fetch(API_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: MODEL,
      temperature: 0.2,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content: "你输出的唯一内容必须是合法 JSON。"
        },
        {
          role: "user",
          content: buildPrompt(doc, input, detailPages)
        }
      ]
    }),
    signal: controller.signal
  }).finally(() => {
    clearTimeout(timeout);
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`SiliconFlow 请求失败: ${response.status} ${body}`);
  }

  const payload = (await response.json()) as {
    choices?: Array<{
      message?: {
        content?: string;
      };
    }>;
  };

  const content = payload.choices?.[0]?.message?.content;
  if (!content) {
    throw new Error("SiliconFlow 未返回可解析内容。");
  }

  return outlineSchema.parse(JSON.parse(content));
}

function distributePages(sections: PlanningSection[], detailPages: number): number[] {
  const sectionCount = sections.length;
  const allocation = new Array(sectionCount).fill(1);
  const weights = sections.map((section) =>
    buildPlanningGroups(section).reduce((sum, group) => sum + estimateGroupWeight(group), 0)
  );
  let remaining = detailPages - sectionCount;

  while (remaining > 0 && sectionCount > 0) {
    let bestIndex = 0;
    let bestScore = -Infinity;

    weights.forEach((weight, index) => {
      const score = weight / allocation[index];
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    allocation[bestIndex] += 1;
    remaining -= 1;
  }

  return allocation;
}

function isUsableOutline(
  result: AIOutline,
  doc: ParsedDocument,
  detailPages: number
): boolean {
  const sumPages = result.directory.reduce((sum, item) => sum + item.slideCount, 0);
  const directoryTitles = new Set(result.directory.map((item) => item.title));
  const maximumDirectoryCount = Math.min(TARGET_DIRECTORY_LIMIT, detailPages);
  const planningSections = groupSectionsForPlanning(doc.sections, maximumDirectoryCount);

  if (result.directory.length === 0 || result.slides.length === 0) {
    return false;
  }

  if (sumPages !== result.slides.length || result.slides.length !== detailPages) {
    return false;
  }

  if (directoryTitles.size !== result.directory.length) {
    return false;
  }

  if (result.directory.length > maximumDirectoryCount) {
    return false;
  }

  if (planningSections.length > 0 && result.directory.length !== planningSections.length) {
    return false;
  }

  return result.slides.every((slide) => directoryTitles.has(slide.directoryTitle));
}

function splitIntoUnits(text: string): string[] {
  const units = text
    .split(/[。！？；;\n]/)
    .flatMap((sentence) => sentence.split(/[，,、]/))
    .map((item) => item.trim())
    .filter((item) => item.length >= 6);

  if (units.length > 0) {
    return units;
  }

  const normalized = text.trim();
  return normalized ? [normalized] : [];
}

function buildSummaryBuckets(section: PlanningSection, pageCount: number): string[] {
  const contentSections = section.childSections.filter(hasSectionContent);
  if (contentSections.length > 0) {
    const sectionChunks: ParsedSection[][] = [];
    let cursor = 0;
    let remainingItems = contentSections.length;
    let remainingPages = Math.max(1, pageCount);

    while (remainingPages > 0) {
      const size = remainingItems > 0 ? Math.ceil(remainingItems / remainingPages) : 0;
      sectionChunks.push(contentSections.slice(cursor, cursor + size));
      cursor += size;
      remainingItems -= size;
      remainingPages -= 1;
    }

    return sectionChunks.map((chunk) =>
      chunk.length > 0
        ? chunk
            .map((child) => {
              const summaryUnits = splitIntoUnits(child.paragraphs.join("；"));
              const summary = summaryUnits.slice(0, 2).join("；") || child.paragraphs[0] || child.title;
              return `${child.title}：${summary}`;
            })
            .join("；")
        : `${section.title}相关内容`
    );
  }

  const baseText = section.paragraphs.join("；");
  const paragraphUnits = section.paragraphs.flatMap(splitIntoUnits);
  const units = Array.from(new Set(paragraphUnits.length > 0 ? paragraphUnits : [baseText]));
  const buckets = Array.from({ length: pageCount }, () => [] as string[]);

  units.forEach((unit, index) => {
    buckets[index % pageCount].push(unit);
  });

  return buckets.map((bucket, index) => {
    if (bucket.length > 0) {
      return bucket.join("；");
    }

    const fallbackSource =
      splitIntoUnits(baseText)[index] ??
      section.paragraphs[index] ??
      section.paragraphs[0] ??
      `${section.title}相关内容`;

    return fallbackSource.trim();
  });
}

type SummaryProfile = {
  textLength: number;
  unitCount: number;
  avgUnitLength: number;
  hasProgressCue: boolean;
  hasHierarchyCue: boolean;
  hasMixedStageCue: boolean;
};

function buildSummaryProfile(summary: string): SummaryProfile {
  const normalized = summary.trim();
  const units = splitIntoUnits(summary);
  const unitCount = Math.max(1, units.length);
  const totalLength = units.reduce((sum, unit) => sum + unit.length, 0);

  return {
    textLength: normalized.length,
    unitCount,
    avgUnitLength: totalLength > 0 ? totalLength / unitCount : normalized.length,
    hasProgressCue: PROGRESS_CUE_RE.test(normalized),
    hasHierarchyCue: HIERARCHY_CUE_RE.test(normalized),
    hasMixedStageCue: /(前置准备|绑定前提|核心价值)/u.test(normalized)
  };
}

function scoreTextLayout(
  type: SlideLayoutType,
  profile: SummaryProfile,
  options: {
    pageIndex: number;
    pageCount: number;
    hasTable: boolean;
    hasImage: boolean;
  }
) {
  const { unitCount, textLength, avgUnitLength, hasProgressCue, hasHierarchyCue, hasMixedStageCue } = profile;
  let score = 0;

  switch (type) {
    case "table":
      score = options.hasTable && options.pageIndex === 0 ? 180 : -999;
      break;
    case "two-column":
      score = options.hasImage && options.pageIndex === 0 ? 150 : -999;
      break;
    case "overview":
      score = 76;
      if (unitCount <= 2) score += 36;
      if (unitCount === 3) score += 18;
      if (avgUnitLength >= 16) score += 8;
      if (unitCount >= 5) score -= 24;
      if (textLength >= 140) score -= 8;
      break;
    case "vertical":
      score = 62;
      if (unitCount >= 3 && unitCount <= 4) score += 22;
      if (avgUnitLength >= 14) score += 8;
      if (unitCount >= 6) score -= 10;
      break;
    case "three-column":
      score = 42;
      if (unitCount >= 4 && unitCount <= 6) score += 28;
      if (avgUnitLength <= 20) score += 6;
      if (unitCount <= 3) score -= 28;
      if (unitCount >= 7) score -= 6;
      if (avgUnitLength >= 24) score -= 6;
      break;
    case "split-grid":
      score = 40;
      if (unitCount >= 4 && unitCount <= 6) score += 18;
      if (textLength >= 120) score += 6;
      if (unitCount <= 3) score -= 18;
      break;
    case "four-column":
      score = 12;
      if (unitCount >= 7) score += 34;
      if (avgUnitLength <= 16) score += 8;
      if (unitCount <= 5) score -= 36;
      break;
    case "progress":
      score = hasProgressCue && unitCount >= 5 ? 92 : 8;
      if (unitCount >= 5) score += 14;
      if (unitCount <= 4) score -= 72;
      if (hasMixedStageCue) score -= 80;
      break;
    case "hierarchy":
      score = hasHierarchyCue ? 82 : 16;
      if (unitCount >= 4 && unitCount <= 6) score += 12;
      if (unitCount <= 3) score -= 52;
      break;
    case "image-right":
    case "image":
    case "image-left":
      score = options.hasImage && options.pageIndex === 0 ? 160 : -999;
      break;
    default:
      score = 0;
  }

  return score;
}

function chooseLayoutSequence(
  section: PlanningSection,
  summaries: string[]
): SlideLayoutType[] {
  return summaries.map((summary, index) => {
    const profile = buildSummaryProfile(summary);
    const ranked = SUPPORTED_LAYOUT_TYPES.map((type) => ({
      type,
      score: scoreTextLayout(type, profile, {
        pageIndex: index,
        pageCount: summaries.length,
        hasTable: section.tables.length > 0 && index === 0,
        hasImage: section.imageIds.length > 0 && index === 0
      })
    }))
      .filter((item) => item.score > -999)
      .sort((left, right) => right.score - left.score);

    if (profile.hasMixedStageCue && ranked[0]?.type === "progress") {
      return profile.unitCount >= 4 ? "split-grid" : "vertical";
    }

    return (
      ranked[0]?.type ??
      (section.tables.length > 0
        ? ("table" as const)
        : section.imageIds.length > 0
          ? ("image-right" as const)
          : TEXT_VARIETY_SEQUENCE[index % TEXT_VARIETY_SEQUENCE.length])
    );
  });
}

function createFallbackOutline(
  doc: ParsedDocument,
  input: GeneratorInput,
  detailPages: number
): AIOutline {
  const sections = groupSectionsForPlanning(
    doc.sections,
    Math.min(TARGET_DIRECTORY_LIMIT, detailPages)
  ).slice(0, detailPages);
  const pageDistribution = distributePages(sections, detailPages);

  const directory = sections.map((section, index) => ({
    title: section.title,
    description: buildPlanningDirectoryBrief(section),
    slideCount: pageDistribution[index]
  }));

  const slides = sections.flatMap((section, index) =>
    (() => {
      const pagePlans = buildPagePlans(section, pageDistribution[index]);
      const layoutSequence = chooseLayoutSequence(
        section,
        pagePlans.map((plan) => plan.summary)
      );

      return new Array(pageDistribution[index]).fill(0).map((_, slideIndex) => {
      const pagePlan = pagePlans[slideIndex] ?? pagePlans[0];
      let type = layoutSequence[slideIndex % layoutSequence.length];
      if (
        type === "progress" &&
        /(前置准备|绑定前提|核心价值)/u.test(pagePlan?.summary ?? "")
      ) {
        type = /核心价值/u.test(pagePlan?.summary ?? "") ? "vertical" : "split-grid";
      }
      const rawSectionTitle = pagePlan?.titleHint || cleanSectionTitle(section.title);
      const isOverviewLead = Boolean(pagePlan?.isOverviewLead);
      const pageTitle = isOverviewLead
        ? buildPageTitle(
            {
              sectionTitle: `${cleanSectionTitle(section.title)}概览`,
              summary: pagePlan?.summary ?? section.paragraphs[0]
            },
            slideIndex + 1
          )
        : buildPageTitle(
            {
              sectionTitle: rawSectionTitle,
              summary: pagePlan?.summary ?? section.paragraphs[0]
            },
            slideIndex + 1
          );
      const overviewDescription = buildOverviewDescription(section);
      const structuredContent = isOverviewLead
        ? buildStructuredContent({
            title: pageTitle,
            intro: buildOverviewIntro(section),
            regularTitle: buildOverviewRegularTitle(section),
            description: overviewDescription || pagePlan?.summary || section.paragraphs[0]
          })
        : buildStructuredContent({
            title: pageTitle,
            intro: pagePlan?.introHint,
            regularTitle: pagePlan?.regularTitleHint,
            description: pagePlan?.summary ?? section.paragraphs[0]
          });

      return {
        directoryTitle: section.title,
        title: pageTitle,
        type,
        content: structuredContent,
        imageSuggestion:
          section.imageIds.length > 0
            ? "优先使用文档原图；若版面不足，保留图片占位。"
            : "文档无原图，保留模板占位框。"
      };
      });
    })()
  );

  return {
    cover: {
      title: doc.titleGuess,
      subtitle: doc.subtitleGuess
    },
    directory,
    slides
  };
}

export async function generateOutlineWithAI(
  doc: ParsedDocument,
  input: GeneratorInput
): Promise<AIOutline> {
  const detailPages = input.totalPages - 2;

  try {
    const result = await requestOutline(doc, input, detailPages);
    if (result && isUsableOutline(result, doc, detailPages)) {
      return result;
    }
  } catch (error) {
    console.warn("SiliconFlow outline generation failed, using fallback.", error);
  }

  return createFallbackOutline(doc, input, detailPages);
}

export function mergeOutlineWithDocument(
  generated: AIOutline,
  doc: ParsedDocument,
  input: GeneratorInput
): OutlineDocument {
  const planningSections = groupSectionsForPlanning(
    doc.sections,
    Math.min(TARGET_DIRECTORY_LIMIT, input.totalPages - 2)
  );
  const directory: OutlineDocument["directory"] = [];
  let pagePointer = 3;

  for (let index = 0; index < generated.directory.length; index += 1) {
    const item = generated.directory[index];
    directory.push({
      id: `dir-${index + 1}`,
      title: item.title,
      description: buildDirectoryBrief({
        title: item.title,
        summary: item.description
      }),
      pageStart: pagePointer,
      pageCount: item.slideCount
    });
    pagePointer += item.slideCount;
  }

  const directoryMap = new Map(directory.map((item) => [toMatchKey(item.title), item]));
  const rawSlides = generated.slides.map((slide, index) => {
    const matchedSection =
      matchSectionByTitle(slide.directoryTitle, planningSections) ??
      matchSectionByTitle(slide.directoryTitle, doc.sections) ??
      planningSections[index % Math.max(1, planningSections.length)] ??
      doc.sections[index % Math.max(1, doc.sections.length)];
    const directoryItem =
      directoryMap.get(toMatchKey(slide.directoryTitle)) ?? directory[index % directory.length];

    return {
      id: `slide-${index + 1}`,
      directoryId: directoryItem.id,
      directoryTitle: directoryItem.title,
      title: slide.title,
      type: slide.type,
      content: buildStructuredContent({
        title: slide.title,
        intro: slide.content.intro,
        regularTitle: slide.content.regularTitle,
        description: slide.content.description
      }),
      summary: composeSlideSummary(
        buildStructuredContent({
          title: slide.title,
          intro: slide.content.intro,
          regularTitle: slide.content.regularTitle,
          description: slide.content.description
        })
      ),
      imageSuggestion: slide.imageSuggestion,
      imageIds: matchedSection?.imageIds ?? [],
      table: matchedSection?.tables.length ? matchedSection.tables : undefined
    };
  });

  const slides = rawSlides.map((slide) => ({ ...slide }));

  directory.forEach((directoryItem) => {
    const matchedSection =
      matchSectionByTitle(directoryItem.title, planningSections) ??
      matchSectionByTitle(directoryItem.title, doc.sections) ??
      matchSectionByTitle(
        rawSlides.find((slide) => slide.directoryId === directoryItem.id)?.directoryTitle ?? "",
        planningSections
      ) ??
      matchSectionByTitle(
        rawSlides.find((slide) => slide.directoryId === directoryItem.id)?.directoryTitle ?? "",
        doc.sections
      );
    const relatedIndexes = slides
      .map((slide, index) => ({ slide, index }))
      .filter(({ slide }) => slide.directoryId === directoryItem.id);

    if (!matchedSection || relatedIndexes.length === 0) {
      return;
    }

    const planningSection: PlanningSection =
      isPlanningSection(matchedSection)
        ? matchedSection
        : buildPlanningSection(matchedSection.title, [matchedSection], matchedSection.level);

    const summaries = relatedIndexes.map(({ slide }) => slide.summary.trim());
    const hasDuplicateSummaries = new Set(summaries).size !== summaries.length;
    const fallbackSummaries = buildSummaryBuckets(planningSection, relatedIndexes.length);
    const effectiveSummaries = hasDuplicateSummaries ? fallbackSummaries : summaries;
    const layoutSequence = chooseLayoutSequence(planningSection, effectiveSummaries);

    relatedIndexes.forEach(({ index }, order) => {
      slides[index] = {
        ...slides[index],
        type: layoutSequence[order % layoutSequence.length],
        ...(hasDuplicateSummaries
          ? (() => {
              const nextContent = buildStructuredContent({
                title: slides[index].title,
                summary: fallbackSummaries[order]
              });

              return {
                content: nextContent,
                summary: composeSlideSummary(nextContent)
              };
            })()
          : {})
      };
    });
  });

  return normalizeOutlineDocument({
    cover: {
      title: generated.cover.title,
      subtitle: generated.cover.subtitle,
      userName: input.userName,
      dateLabel: new Intl.DateTimeFormat("zh-CN", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
      })
        .format(new Date())
        .replace(/\//g, "-")
    },
    directoryTitle: "目录",
    directory,
    slides,
    pageSummary: {
      totalPages: input.totalPages,
      coverPages: 1,
      directoryPages: 1,
      detailPages: input.totalPages - 2
    },
    extractedImages: doc.images
  });
}
