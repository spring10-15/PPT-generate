import type { DirectoryItem, OutlineDocument, OutlineSlide } from "@/lib/types";
import { buildDirectoryBrief, composeSlideSummary } from "@/lib/slide-content";

const CHINESE_NUMERALS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const TITLE_PREFIX_PATTERNS = [
  /^\s*第?[一二三四五六七八九十百千]+[章节部分篇节项]\s*/,
  /^\s*[（(【[]\s*[一二三四五六七八九十百千]+\s*[】）)\]]\s*[、.．]?\s*/,
  /^\s*[一二三四五六七八九十百千]+[、.．]\s*/,
  /^\s*[（(【[]\s*\d+(?:\.\d+)*\s*[】）)\]]\s*[、.．]?\s*/,
  /^\s*\d+(?:\.\d+)*[、.．]\s*/,
  /^\s*[【[]?[一二三四五六七八九十0-9]+[】\]]\s*/
];
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

function pickKeywordTitle(text: string, maxLength: number) {
  const base = stripTitlePrefix(text).replace(/（[^）]*）|\([^)]*\)/gu, "").trim();
  if (!base) {
    return "";
  }

  if (base.length <= maxLength) {
    return trimWeakTitleEnding(base);
  }

  const matches = Array.from(base.matchAll(TITLE_KEYWORD_RE))
    .map((match) => trimWeakTitleEnding(match[0].trim()))
    .filter((candidate) => candidate && candidate.length <= maxLength)
    .sort((left, right) => right.length - left.length);

  return matches[0] ?? "";
}

function buildSuffixTitle(text: string, maxLength: number) {
  for (const suffix of SEMANTIC_SUFFIXES) {
    const index = text.lastIndexOf(suffix);
    if (index < 0) {
      continue;
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

    for (const base of baseCandidates) {
      const candidate = trimWeakTitleEnding(`${base}${suffix}`);
      if (candidate && candidate.length <= maxLength) {
        return candidate;
      }
    }
  }

  return "";
}

function truncateCoreText(text: string, maxLength: number) {
  const normalized = stripTitlePrefix(text).replace(/\s+/g, " ").trim();
  if (!normalized || normalized.length <= maxLength) {
    return trimWeakTitleEnding(normalized);
  }

  const suffixTitle = buildSuffixTitle(normalized, maxLength);
  if (suffixTitle) {
    return suffixTitle;
  }

  const keywordTitle = pickKeywordTitle(normalized, maxLength);
  if (keywordTitle) {
    return keywordTitle;
  }

  const segments = normalized
    .split(TITLE_SEGMENT_SPLIT_RE)
    .map((segment) => trimWeakTitleEnding(segment.trim()))
    .filter(Boolean);
  const bestSegment = segments
    .filter((segment) => segment.length <= maxLength)
    .sort((left, right) => right.length - left.length)[0];

  if (bestSegment) {
    return bestSegment;
  }

  const clipped = trimWeakTitleEnding(normalized.slice(0, maxLength).trim());
  if (clipped) {
    return clipped;
  }

  return normalized.slice(0, maxLength).trim();
}

function toChineseNumeral(value: number): string {
  if (value <= 10) {
    return value === 10 ? "十" : CHINESE_NUMERALS[value];
  }

  if (value < 20) {
    return `十${CHINESE_NUMERALS[value - 10]}`;
  }

  if (value < 100) {
    const tens = Math.floor(value / 10);
    const ones = value % 10;
    return `${CHINESE_NUMERALS[tens]}十${ones === 0 ? "" : CHINESE_NUMERALS[ones]}`;
  }

  return String(value);
}

export function stripTitlePrefix(text: string) {
  let current = text.replace(/\s+/g, " ").trim();

  while (current) {
    const next = TITLE_PREFIX_PATTERNS.reduce((result, pattern) => {
      if (result !== current) {
        return result;
      }

      return current.replace(pattern, "").trim();
    }, current);

    if (next === current) {
      break;
    }

    current = next;
  }

  current = current.replace(/^\d+(?=[\u4e00-\u9fa5])/u, "").trim();
  current = current.replace(/[-_.。．、:：]+\s*$/u, "").trim();
  current = current.replace(/[-_.。．、:：]\s*\d+[.。．、:：]*$/u, "").trim();

  return current;
}

function trimWeakTitleEnding(text: string) {
  let current = text.trim();

  while (current && WEAK_TITLE_END_RE.test(current)) {
    current = current.replace(WEAK_TITLE_END_RE, "").trim();
  }

  return current;
}

export function formatPrimaryLabel(order: number, text: string) {
  const cleaned = stripTitlePrefix(text) || "未命名标题";
  return `【${toChineseNumeral(order)}】${cleaned}`;
}

export function formatSecondaryLabel(order: number, text: string) {
  const cleaned = stripTitlePrefix(text) || "未命名要点";
  return `【${order}】${cleaned}`;
}

function buildDirectoryMap(directory: DirectoryItem[], slides: OutlineSlide[]) {
  const orderedIds = directory.map((item) => item.id);
  const slideOnlyIds = slides
    .map((slide) => slide.directoryId)
    .filter((directoryId, index, list) => list.indexOf(directoryId) === index && !orderedIds.includes(directoryId));

  return [...orderedIds, ...slideOnlyIds]
    .map((directoryId) => ({
      directory:
        directory.find((item) => item.id === directoryId) ??
        ({
          id: directoryId,
          title: slides.find((slide) => slide.directoryId === directoryId)?.directoryTitle ?? "未命名目录",
          description: buildDirectoryBrief({
            title: slides.find((slide) => slide.directoryId === directoryId)?.directoryTitle,
            description: slides.find((slide) => slide.directoryId === directoryId)?.content.description
          }),
          pageStart: 0,
          pageCount: 0
        } satisfies DirectoryItem),
      slides: slides.filter((slide) => slide.directoryId === directoryId)
    }))
    .filter((group) => group.slides.length > 0);
}

export function normalizeOutlineDocument(outline: OutlineDocument): OutlineDocument {
  const grouped = buildDirectoryMap(outline.directory, outline.slides);
  let pageStart = 3;

  const directory = grouped.map((group, index) => {
    const baseTitle =
      truncateCoreText(group.directory.title, 5) ||
      truncateCoreText(group.slides[0]?.directoryTitle ?? "", 5) ||
      truncateCoreText(group.slides[0]?.title ?? "", 5) ||
      `目录${index + 1}`;

    const pageCount = group.slides.length;
    const item: DirectoryItem = {
      id: group.directory.id,
      title: baseTitle,
      description:
        buildDirectoryBrief({
          title: baseTitle,
          description: group.directory.description || group.slides[0]?.content.description
        }) || "要点概述",
      pageStart,
      pageCount
    };

    pageStart += pageCount;
    return item;
  });

  const directoryIndexMap = new Map(directory.map((item, index) => [item.id, index]));
  const slides = grouped.flatMap((group) => {
    const directoryIndex = directoryIndexMap.get(group.directory.id) ?? 0;
    const directoryTitle = directory[directoryIndex]?.title ?? truncateCoreText(group.directory.title, 5);
    const directoryBase = stripTitlePrefix(directoryTitle);

    return group.slides.map((slide) => {
      const baseTitle =
        truncateCoreText(slide.title, 10) ||
        truncateCoreText(slide.directoryTitle, 10) ||
        truncateCoreText(directoryBase, 10);

      return {
        ...slide,
        directoryId: group.directory.id,
        directoryTitle,
        title: baseTitle,
        summary: composeSlideSummary(slide.content)
      };
    });
  });

  return {
    ...outline,
    directory,
    slides,
    pageSummary: {
      totalPages: slides.length + 2,
      coverPages: 1,
      directoryPages: 1,
      detailPages: slides.length
    }
  };
}
