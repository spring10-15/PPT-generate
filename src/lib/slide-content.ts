import type { SlideContentDraft } from "@/lib/types";

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
const GENERIC_TITLE_RE =
  /(方案思路|整体思路|具体操作流程|前置准备|重点推荐|全程自动化|自动化|能力说明|核心说明|详细说明|业务说明|内容概述|整体概览|项目概述)/u;
const TITLE_KEYWORD_RE =
  /[\u4e00-\u9fa5A-Za-z0-9]{1,10}?(?:场景|应用|方案|协作|价值|流程|绑定|内容|能力|体系|架构|矩阵|模块|数据|开发|运营|管理|办公|创作|服务|平台|合规|治理|安全|知识|任务|团队|文件|业务|项目)/gu;
const GENERIC_HEADING_RE = /^(概览|方案思路|前置准备|具体操作流程|核心价值|绑定前提)$/u;
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

function normalizeText(value: string | undefined) {
  return (value ?? "").replace(/\s+/g, " ").trim();
}

function toComparableText(value: string | undefined) {
  return stripTitlePrefix(normalizeText(value ?? "")).replace(/[：:；;。！？,.，、】【（）()[\]\s-]/g, "");
}

function stripDecorators(text: string) {
  return normalizeText(text)
    .replace(/[“”"「」]/gu, "")
    .replace(/[‘’']/gu, "")
    .replace(/^[：:，,；;\-—\s]+/u, "")
    .replace(/[：:，,；;\-—\s]+$/u, "");
}

function pickKeywordTitle(text: string, maxLength: number) {
  const withoutParentheses = stripDecorators(text).replace(/（[^）]*）|\([^)]*\)/gu, "").trim();
  if (!withoutParentheses) {
    return "";
  }

  if (withoutParentheses.length <= maxLength) {
    return withoutParentheses;
  }

  const matches = Array.from(withoutParentheses.matchAll(TITLE_KEYWORD_RE))
    .map((match) => match[0])
    .filter((candidate) => candidate.length <= maxLength)
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

function stripTitlePrefix(text: string) {
  let current = normalizeText(text);

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

function compactTitle(text: string, maxLength: number) {
  const normalized = stripDecorators(stripTitlePrefix(text));
  if (!normalized) {
    return "";
  }

  const keywordTitle = pickKeywordTitle(normalized, maxLength);
  if (keywordTitle) {
    return keywordTitle;
  }

  if (normalized.length <= maxLength) {
    return trimWeakTitleEnding(normalized);
  }

  const suffixTitle = buildSuffixTitle(normalized, maxLength);
  if (suffixTitle) {
    return suffixTitle;
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
  return clipped || normalized.slice(0, maxLength).trim();
}

function scoreTitleCandidate(candidate: string, sourceIndex: number, maxLength: number) {
  const cleaned = compactTitle(candidate, maxLength);
  if (!cleaned) {
    return -999;
  }

  let score = 0;
  score += 80 - sourceIndex * 12;
  score += Math.min(cleaned.length, maxLength) * 3;

  if (GENERIC_TITLE_RE.test(cleaned)) {
    score -= 40;
  }

  if (/[“”"「」]/u.test(candidate)) {
    score -= 18;
  }

  if (/[：:]/u.test(candidate)) {
    score += 8;
  }

  return score;
}

function pickSemanticTitle(candidates: string[], maxLength: number, fallback: string) {
  const ranked = candidates
    .map((candidate, index) => ({
      value: compactTitle(candidate, maxLength),
      score: scoreTitleCandidate(candidate, index, maxLength)
    }))
    .filter((candidate) => candidate.value)
    .sort((left, right) => right.score - left.score);

  return ranked[0]?.value || compactTitle(fallback, maxLength) || fallback;
}

function isGenericHeading(text: string | undefined) {
  return GENERIC_HEADING_RE.test(stripDecorators(stripTitlePrefix(text ?? "")));
}

function truncatePhrase(text: string, maxLength: number) {
  const normalized = normalizeText(text);
  if (!normalized || normalized.length <= maxLength) {
    return normalized;
  }

  const units = splitUnits(normalized);
  if (units.length > 1) {
    let result = "";

    units.forEach((unit) => {
      const next = result ? `${result}；${unit}` : unit;
      if (next.length <= maxLength) {
        result = next;
      }
    });

    if (result) {
      return result;
    }
  }

  return normalized.slice(0, maxLength).trim();
}

function dedupeUnits(units: string[], blocked: string[] = []) {
  const blockedSet = new Set(
    blocked
      .map((item) => stripTitlePrefix(normalizeText(item)))
      .filter(Boolean)
  );
  const seen = new Set<string>();

  return units.filter((unit) => {
    const normalized = stripTitlePrefix(normalizeText(unit));
    if (!normalized || blockedSet.has(normalized) || seen.has(normalized)) {
      return false;
    }
    seen.add(normalized);
    return true;
  });
}

function splitUnits(text: string) {
  return normalizeText(text)
    .split(/[。！？；;\n]/)
    .flatMap((part) => part.split(/[，,、]/))
    .map((part) => normalizeText(part))
    .filter(Boolean);
}

export function splitUnitTitleBody(unit: string) {
  const normalized = stripDecorators(unit);
  if (!normalized) {
    return {
      title: "",
      body: ""
    };
  }

  const colonMatch = normalized.match(/^([^：:]{2,36})[：:]\s*(.+)$/u);
  if (colonMatch) {
    return {
      title: compactTitle(colonMatch[1], 15),
      body: stripDecorators(colonMatch[2])
    };
  }

  const quoteMatch = normalized.match(/^([^“”"「」]{2,36})[“”"「」](.+)[”"」]?$/u);
  if (quoteMatch) {
    return {
      title: compactTitle(quoteMatch[1], 15),
      body: stripDecorators(quoteMatch[2])
    };
  }

  return {
    title: "",
    body: normalized
  };
}

function buildIntro(text: string) {
  const sentence = normalizeText(text).split(/[。！？;\n]/)[0] ?? "";
  return truncatePhrase(sentence, 30);
}

function buildRegularTitle(title: string, description: string) {
  const titleSource = stripTitlePrefix(normalizeText(title));
  const units = dedupeUnits(splitUnits(description), [titleSource]);
  const labeled = units
    .map((unit) => splitUnitTitleBody(unit).title)
    .filter(Boolean);

  if (labeled.length > 0) {
    return pickSemanticTitle(labeled, 15, "核心内容");
  }

  if (units.length > 0) {
    return pickSemanticTitle([units[0], titleSource], 15, "核心内容");
  }

  if (titleSource) {
    return truncatePhrase(titleSource, 15);
  }

  const firstUnit = splitUnits(description)[0] ?? description;
  return truncatePhrase(firstUnit, 15) || "核心内容";
}

export function composeSlideSummary(content: SlideContentDraft) {
  return [content.intro, content.regularTitle, content.description]
    .map((item) => normalizeText(item))
    .filter(Boolean)
    .join("；");
}

export function buildDirectoryBrief(input: { title?: string; summary?: string; description?: string }) {
  const source = input.description || input.summary || input.title || "";
  const blocked = [input.title || ""];
  const units = dedupeUnits(splitUnits(source), blocked);
  const labeledTitles = units
    .map((unit) => splitUnitTitleBody(unit).title)
    .filter(Boolean);
  const keywordCandidate = pickSemanticTitle(
    [
      ...labeledTitles,
      ...units.map((unit) => splitUnitTitleBody(unit).body || unit),
      normalizeText(source)
    ],
    10,
    stripDecorators(stripTitlePrefix(input.title || "")) || "要点概述"
  );

  return compactTitle(keywordCandidate, 10) || truncatePhrase(keywordCandidate, 10) || "要点概述";
}

export function buildPageTitle(input: { sectionTitle?: string; summary?: string }, fallbackIndex = 1) {
  const sectionTitle = stripTitlePrefix(normalizeText(input.sectionTitle || ""));
  const units = dedupeUnits(splitUnits(input.summary || ""), [sectionTitle]);
  const labeledTitles = units
    .map((unit) => splitUnitTitleBody(unit).title)
    .filter((title) => title && !isGenericHeading(title));
  const candidate = pickSemanticTitle(
    [sectionTitle, ...labeledTitles, ...units],
    10,
    sectionTitle || `要点${fallbackIndex}`
  );
  return compactTitle(candidate, 10) || `要点${fallbackIndex}`;
}

export function buildStructuredContent(input: {
  title?: string;
  summary?: string;
  intro?: string;
  regularTitle?: string;
  description?: string;
}): SlideContentDraft {
  const title = normalizeText(input.title || "");
  const rawSource = input.description || input.summary || "";
  const sourceUnits = dedupeUnits(splitUnits(rawSource), [title]);
  const labeledFirstUnit = sourceUnits
    .map((unit) => splitUnitTitleBody(unit))
    .find((item) => (item.title && !isGenericHeading(item.title)) || item.body);
  const regularTitle =
    truncatePhrase(
      normalizeText(input.regularTitle) ||
        labeledFirstUnit?.title ||
        sourceUnits[0] ||
        buildRegularTitle(title, rawSource),
      15
    ) || "核心内容";
  const explicitIntroUnits = dedupeUnits(splitUnits(normalizeText(input.intro)), [title, regularTitle]);
  const remainingUnits = dedupeUnits(sourceUnits, [regularTitle]);
  const selectedIntroUnits = (explicitIntroUnits.length > 0 ? explicitIntroUnits : remainingUnits).slice(0, 2);
  const introCandidate =
    selectedIntroUnits.length > 0
      ? selectedIntroUnits
          .map((unit) => {
            const labeled = splitUnitTitleBody(unit);
            return labeled.body || unit;
          })
          .join("；")
      : normalizeText(input.intro) || buildIntro(input.summary || rawSource);
  const intro = truncatePhrase(introCandidate, 30);
  const descriptionUnits = dedupeUnits(splitUnits(input.description || rawSource), [
    title,
    regularTitle,
    ...selectedIntroUnits
  ]);
  const description =
    descriptionUnits.length > 0
      ? descriptionUnits
          .slice(0, 6)
          .map((unit) => {
            const labeled = splitUnitTitleBody(unit);
            if (toComparableText(labeled.title) === toComparableText(regularTitle)) {
              return labeled.body || unit;
            }
            return unit;
          })
          .filter(Boolean)
          .join("；")
      : normalizeText(input.description || "");

  return {
    intro: intro || truncatePhrase(description, 30),
    regularTitle: regularTitle || "核心内容",
    description
  };
}
