import type { SlideContentDraft } from "@/lib/types";

const TITLE_PREFIX_PATTERNS = [
  /^\s*第?[一二三四五六七八九十百千]+[章节部分篇节项]\s*/,
  /^\s*[（(【[]\s*[一二三四五六七八九十百千]+\s*[】）)\]]\s*[、.．]?\s*/,
  /^\s*[一二三四五六七八九十百千]+[、.．]\s*/,
  /^\s*[（(【[]\s*\d+(?:\.\d+)*\s*[】）)\]]\s*[、.．]?\s*/,
  /^\s*\d+(?:\.\d+)*[、.．]\s*/,
  /^\s*[【[]?[一二三四五六七八九十0-9]+[】\]]\s*/
];

function normalizeText(value: string | undefined) {
  return (value ?? "").replace(/\s+/g, " ").trim();
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

  return current;
}

function truncatePhrase(text: string, maxLength: number) {
  const normalized = normalizeText(text);
  void maxLength;
  return normalized;
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

function buildIntro(text: string) {
  const sentence = normalizeText(text).split(/[。！？;\n]/)[0] ?? "";
  return truncatePhrase(sentence, 30);
}

function buildRegularTitle(title: string, description: string) {
  const titleSource = stripTitlePrefix(normalizeText(title));
  const units = dedupeUnits(splitUnits(description), [titleSource]);

  if (units.length > 0) {
    return truncatePhrase(units[0], 15) || "核心内容";
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
  const candidate = units[0] ?? normalizeText(source);
  return truncatePhrase(candidate, 10) || "要点概述";
}

export function buildPageTitle(input: { sectionTitle?: string; summary?: string }, fallbackIndex = 1) {
  const sectionTitle = stripTitlePrefix(normalizeText(input.sectionTitle || ""));
  const units = dedupeUnits(splitUnits(input.summary || ""), [sectionTitle]);
  const candidate = units[0] ?? (sectionTitle || `要点${fallbackIndex}`);
  return truncatePhrase(candidate, 7) || `要点${fallbackIndex}`;
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
  const regularTitle =
    truncatePhrase(
      normalizeText(input.regularTitle) || sourceUnits[0] || buildRegularTitle(title, rawSource),
      15
    ) || "核心内容";
  const explicitIntroUnits = dedupeUnits(splitUnits(normalizeText(input.intro)), [title, regularTitle]);
  const remainingUnits = dedupeUnits(sourceUnits, [regularTitle]);
  const selectedIntroUnits = (explicitIntroUnits.length > 0 ? explicitIntroUnits : remainingUnits).slice(0, 2);
  const introCandidate =
    selectedIntroUnits.length > 0
      ? selectedIntroUnits.join("；")
      : normalizeText(input.intro) || buildIntro(input.summary || rawSource);
  const intro = truncatePhrase(introCandidate, 30);
  const descriptionUnits = dedupeUnits(splitUnits(input.description || rawSource), [
    title,
    regularTitle,
    ...selectedIntroUnits
  ]);
  const description =
    descriptionUnits.length > 0
      ? descriptionUnits.slice(0, 6).join("；")
      : normalizeText(input.description || "");

  return {
    intro: intro || truncatePhrase(description, 30),
    regularTitle: regularTitle || "核心内容",
    description
  };
}
