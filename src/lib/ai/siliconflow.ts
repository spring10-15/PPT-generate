import { z } from "zod";
import {
  SUPPORTED_LAYOUT_TYPES,
  type GeneratorInput,
  type OutlineDocument,
  type ParsedDocument,
  type ParsedSection,
  type SlideLayoutType
} from "@/lib/types";
import { normalizeOutlineDocument } from "@/lib/outline-format";
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

function normalizeTitle(title: string) {
  return title.replace(/\s+/g, " ").trim();
}

function groupSectionsForPlanning(sections: ParsedSection[], maxDirectories: number) {
  if (sections.length <= maxDirectories) {
    return sections;
  }

  const groupSize = Math.ceil(sections.length / maxDirectories);
  const merged: ParsedSection[] = [];

  for (let index = 0; index < sections.length; index += groupSize) {
    const group = sections.slice(index, index + groupSize);
    const first = group[0];
    const last = group[group.length - 1];
    const mergedTitle =
      group.length === 1
        ? normalizeTitle(first.title)
        : `${normalizeTitle(first.title)}-${normalizeTitle(last.title)}`;

    merged.push({
      id: first.id,
      title: mergedTitle,
      level: first.level,
      paragraphs: group.flatMap((section) => section.paragraphs).filter(Boolean),
      tables: group.flatMap((section) => section.tables),
      imageIds: Array.from(new Set(group.flatMap((section) => section.imageIds)))
    });
  }

  return merged;
}

function trimSections(sections: ParsedSection[]) {
  return sections.slice(0, 16).map((section) => ({
    title: section.title,
    paragraphs: section.paragraphs.slice(0, 4),
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

function distributePages(sectionCount: number, detailPages: number): number[] {
  const base = new Array(sectionCount).fill(1);
  let remaining = detailPages - sectionCount;
  let index = 0;

  while (remaining > 0 && sectionCount > 0) {
    base[index % sectionCount] += 1;
    remaining -= 1;
    index += 1;
  }

  return base;
}

function isUsableOutline(
  result: AIOutline,
  doc: ParsedDocument,
  detailPages: number
): boolean {
  const sumPages = result.directory.reduce((sum, item) => sum + item.slideCount, 0);
  const directoryTitles = new Set(result.directory.map((item) => item.title));
  const maximumDirectoryCount = Math.min(TARGET_DIRECTORY_LIMIT, detailPages);

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

function buildSummaryBuckets(section: ParsedDocument["sections"][number], pageCount: number): string[] {
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
    hasHierarchyCue: HIERARCHY_CUE_RE.test(normalized)
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
  const { unitCount, textLength, avgUnitLength, hasProgressCue, hasHierarchyCue } = profile;
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
      score = hasProgressCue ? 86 : 18;
      if (unitCount >= 4) score += 14;
      if (unitCount <= 3) score -= 22;
      break;
    case "hierarchy":
      score = hasHierarchyCue ? 82 : 16;
      if (unitCount >= 4 && unitCount <= 6) score += 12;
      if (unitCount <= 3) score -= 18;
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
  section: ParsedDocument["sections"][number],
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
  const pageDistribution = distributePages(sections.length, detailPages);

  const directory = sections.map((section, index) => ({
    title: section.title,
    description: buildDirectoryBrief({
      title: section.title,
      summary: section.paragraphs.join("；")
    }),
    slideCount: pageDistribution[index]
  }));

  const slides = sections.flatMap((section, index) =>
    new Array(pageDistribution[index]).fill(0).map((_, slideIndex) => {
      const summaryBuckets = buildSummaryBuckets(section, pageDistribution[index]);
      const layoutSequence = chooseLayoutSequence(section, summaryBuckets);
      const type = layoutSequence[slideIndex % layoutSequence.length];
      const pageTitle = buildPageTitle(
        {
          sectionTitle: section.title,
          summary: summaryBuckets[slideIndex] ?? section.paragraphs[0]
        },
        slideIndex + 1
      );

      return {
        directoryTitle: section.title,
        title: pageTitle,
        type,
        content: buildStructuredContent({
          title: pageTitle,
          summary: summaryBuckets[slideIndex] ?? section.paragraphs[0]
        }),
        imageSuggestion:
          section.imageIds.length > 0
            ? "优先使用文档原图；若版面不足，保留图片占位。"
            : "文档无原图，保留模板占位框。"
      };
    })
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

  const directoryMap = new Map(directory.map((item) => [item.title, item]));
  const rawSlides = generated.slides.map((slide, index) => {
    const matchedSection =
      planningSections.find((section) => section.title === slide.directoryTitle) ??
      doc.sections.find((section) => section.title === slide.directoryTitle) ??
      planningSections[index % Math.max(1, planningSections.length)] ??
      doc.sections[index % Math.max(1, doc.sections.length)];
    const directoryItem = directoryMap.get(slide.directoryTitle) ?? directory[index % directory.length];

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
      planningSections.find((section) => section.title === directoryItem.title) ??
      doc.sections.find((section) => section.title === directoryItem.title) ??
      planningSections.find(
        (section) =>
          section.title === rawSlides.find((slide) => slide.directoryId === directoryItem.id)?.directoryTitle
      ) ??
      doc.sections.find(
        (section) =>
          section.title === rawSlides.find((slide) => slide.directoryId === directoryItem.id)?.directoryTitle
      );
    const relatedIndexes = slides
      .map((slide, index) => ({ slide, index }))
      .filter(({ slide }) => slide.directoryId === directoryItem.id);

    if (!matchedSection || relatedIndexes.length === 0) {
      return;
    }

    const summaries = relatedIndexes.map(({ slide }) => slide.summary.trim());
    const hasDuplicateSummaries = new Set(summaries).size !== summaries.length;
    const fallbackSummaries = buildSummaryBuckets(matchedSection, relatedIndexes.length);
    const effectiveSummaries = hasDuplicateSummaries ? fallbackSummaries : summaries;
    const layoutSequence = chooseLayoutSequence(matchedSection, effectiveSummaries);

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
