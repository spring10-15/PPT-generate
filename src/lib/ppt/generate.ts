import { readFile } from "node:fs/promises";
import { load as loadXml } from "cheerio";
import type { CheerioAPI, Cheerio } from "cheerio";
import type { Element } from "domhandler";
import JSZip from "jszip";
import type { ExtractedImage, OutlineDocument, OutlineSlide, SlideLayoutType } from "@/lib/types";
import {
  normalizeOutlineDocument,
  stripTitlePrefix
} from "@/lib/outline-format";
import { getTemplatePath } from "@/lib/template";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const BLANK_IMAGE_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9GNfQAAAAASUVORK5CYII=";
const TEMPLATE_PLACEHOLDER_TEXT = [
  "参考版式",
  "内容简介",
  "常规标题",
  "强调标题",
  "添加标题",
  "放置示意图",
  "这是个小标题",
  "这是说明性文字",
  "这是个标题",
  "这是个示例",
  "总体说明性文字",
  "请在此处添加"
];

const COVER_SOURCE_SLIDE = 1;
const SLIDE_RELATIONSHIP_BASE_ID = 100;

const SLOT_LIMITS = {
  coverTitle: 24,
  coverSubtitle: 38,
  coverMeta: 48,
  intro: 30,
  pageTitle: 10,
  badge: 12,
  smallTitle: 15,
  smallBody: 60,
  mediumBody: 82,
  largeBody: 110,
  directoryLine: 5,
  directoryBrief: 10,
  compactLine: 18,
  regularLine: 24,
  tableHeader: 10,
  tableRow: 10,
  tableCell: 18
} as const;

type TemplatePlan = {
  sourceSlide: number;
  fill: (xml: CheerioAPI, outline: OutlineDocument, slide?: OutlineSlide) => void;
  withImage?: boolean;
};

type TableDraft = {
  columnHeaders: string[];
  rowHeaders: string[];
  cellBodies: string[][];
};

type CardDraft = {
  heading: string;
  bodyLines: string[];
};

type SectionDraft = {
  label: string;
  cards: CardDraft[];
};

type SlotSpec = {
  maxLines: number;
  maxLength: number;
  chunkLength?: number;
};

type TextStyleSpec = {
  fontSize?: number;
  colorHex?: string;
  bold?: boolean;
};

const DETAIL_TEMPLATE_MAP: Record<SlideLayoutType, TemplatePlan> = {
  overview: { sourceSlide: 13, fill: fillOverviewSlide },
  "two-column": { sourceSlide: 4, fill: fillTwoColumnSlide, withImage: true },
  "three-column": { sourceSlide: 5, fill: fillThreeColumnSlide },
  "four-column": { sourceSlide: 9, fill: fillFourColumnSlide },
  progress: { sourceSlide: 10, fill: fillProgressSlide },
  vertical: { sourceSlide: 11, fill: fillVerticalSplitSlide },
  "split-grid": { sourceSlide: 12, fill: fillSplitGridSlide },
  hierarchy: { sourceSlide: 15, fill: fillHierarchySlide },
  table: { sourceSlide: 16, fill: fillTableSlide },
  image: { sourceSlide: 4, fill: fillImageSlide, withImage: true },
  "image-left": { sourceSlide: 4, fill: fillImageSlide, withImage: true },
  "image-right": { sourceSlide: 4, fill: fillImageSlide, withImage: true }
};

const REFERENCE_TITLE_STYLE: TextStyleSpec = { fontSize: 28, colorHex: "0070C0", bold: false };
const INTRO_STYLE: TextStyleSpec = { fontSize: 16, colorHex: "0070C0", bold: false };
const REGULAR_TITLE_STYLE: TextStyleSpec = { fontSize: 14, colorHex: "EFEFEF", bold: false };
const SMALL_TITLE_STYLE: TextStyleSpec = { fontSize: 14, colorHex: "000000", bold: false };
const DESCRIPTION_STYLE: TextStyleSpec = { fontSize: 16, colorHex: "C7C713", bold: false };
const DIRECTORY_BRIEF_STYLE: TextStyleSpec = { fontSize: 12, colorHex: "C7C713", bold: false };

function normalizeText(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function truncateText(text: string, maxLength: number) {
  const normalized = normalizeText(text);
  void maxLength;
  return normalized;
}

function splitSummary(summary: string) {
  const pieces = summary
    .split(/[。！？；;\n]/)
    .flatMap((part) => part.split(/[，,、]/))
    .map((part) => normalizeText(part))
    .filter(Boolean);

  if (pieces.length > 0) {
    return Array.from(new Set(pieces));
  }

  const fallback = normalizeText(summary);
  return fallback ? [fallback] : ["待补充"];
}

function toComparableText(text: string) {
  return stripTitlePrefix(normalizeText(text)).replace(/[：:；;。！？,.，、】【（）()[\]\s]/g, "");
}

function dedupeTextItems(items: string[], blocked: string[] = []) {
  const blockedSet = new Set(blocked.map((item) => toComparableText(item)).filter(Boolean));
  const seen = new Set<string>();

  return items.filter((item) => {
    const comparable = toComparableText(item);
    if (!comparable || blockedSet.has(comparable) || seen.has(comparable)) {
      return false;
    }

    seen.add(comparable);
    return true;
  });
}

function buildHeading(text: string, fallback: string) {
  const source = normalizeText(text)
    .replace(/[：:；;。！？,.，、]/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .join("");
  const candidate = source;
  return candidate || fallback;
}

function buildBuckets(items: string[], count: number) {
  const buckets = Array.from({ length: count }, () => [] as string[]);
  items.forEach((item, index) => {
    buckets[index % count].push(item);
  });
  return buckets;
}

function buildPlainSlot(text: string, fallback: string, maxLength: number = SLOT_LIMITS.smallTitle) {
  return truncateText(stripTitlePrefix(text) || fallback, maxLength);
}

function joinBucket(items: string[], maxLength: number) {
  return truncateText(items.join("；"), maxLength);
}

function buildParagraphLines(items: string[], maxLines: number, maxLength: number) {
  return items
    .slice(0, maxLines)
    .map((item) => truncateText(item, maxLength))
    .filter(Boolean);
}

function splitLongTextForSlots(text: string, chunkLength: number) {
  const normalized = normalizeText(text);
  void chunkLength;
  return normalized ? [normalized] : [];
}

function expandUnitsForTarget(items: string[], targetCount: number, chunkLength: number) {
  const expanded = items.map((item) => normalizeText(item)).filter(Boolean);

  while (expanded.length < targetCount) {
    let splitIndex = -1;
    let longest = 0;

    expanded.forEach((item, index) => {
      const pieces = splitLongTextForSlots(item, chunkLength);
      if (pieces.length > 1 && item.length > longest) {
        splitIndex = index;
        longest = item.length;
      }
    });

    if (splitIndex < 0) {
      break;
    }

    const pieces = splitLongTextForSlots(expanded[splitIndex], chunkLength);
    expanded.splice(splitIndex, 1, ...pieces);
  }

  return expanded;
}

function buildBodyLines(text: string, maxLines: number, maxLength: number) {
  if (!normalizeText(text)) {
    return [" "];
  }

  const units = expandUnitsForTarget(splitSummary(text), maxLines, maxLength);
  return buildParagraphLines(units.length > 0 ? units : [text], maxLines, maxLength);
}

function buildCardDraft(
  text: string,
  fallback: string,
  maxLines: number,
  maxLength: number
): CardDraft {
  const source = normalizeText(text);
  if (!source) {
    return {
      heading: " ",
      bodyLines: [" "]
    };
  }

  const heading = buildPlainSlot(buildHeading(source, fallback), fallback);
  const strippedBody = normalizeText(
    source
      .replace(heading, "")
      .replace(/^[：:，,；;\-—\s]+/, "")
  );
  const bodySource =
    strippedBody && toComparableText(strippedBody) !== toComparableText(heading) ? strippedBody : "";

  return {
    heading,
    bodyLines: bodySource
      ? buildBodyLines(truncateText(bodySource, SLOT_LIMITS.smallBody), maxLines, maxLength)
      : [" "]
  };
}

function buildCardsBySlotSpecs(items: string[], slotSpecs: SlotSpec[], fallbackPrefix: string) {
  const targetCount = slotSpecs.length;
  const defaultChunkLength = Math.max(...slotSpecs.map((slot) => slot.chunkLength ?? slot.maxLength));
  const expanded = expandUnitsForTarget(items, targetCount, defaultChunkLength);

  return slotSpecs.map((slot, index) =>
    buildCardDraft(
      expanded[index] ?? "",
      `${fallbackPrefix}${index + 1}`,
      slot.maxLines,
      slot.maxLength
    )
  );
}

function buildSectionDrafts(
  items: string[],
  sectionSizes: number[],
  options: {
    sectionPrefix: string;
    preferredLabels?: string[];
    cardPrefix: string;
    bodyLines: number;
    bodyLength: number;
    chunkLength: number;
  }
): SectionDraft[] {
  const totalSlots =
    sectionSizes.reduce((sum, size) => sum + size, 0) +
    sectionSizes.reduce((sum, _size, sectionIndex) => {
      const preferredLabel = options.preferredLabels?.[sectionIndex] ?? "";
      return sum + (normalizeText(preferredLabel) ? 0 : 1);
    }, 0);
  const expanded = expandUnitsForTarget(items, totalSlots, options.chunkLength);
  let cursor = 0;
  let absoluteOrder = 1;

  return sectionSizes.map((size, sectionIndex) => {
    const preferredLabel = options.preferredLabels?.[sectionIndex] ?? "";
    const hasPreferredLabel = Boolean(normalizeText(preferredLabel));
    const labelSource =
      preferredLabel ||
      expanded[cursor] ||
      "";
    const sectionItems = hasPreferredLabel
      ? expanded.slice(cursor, cursor + size)
      : expanded.slice(cursor + 1, cursor + 1 + size);
    cursor += size + (hasPreferredLabel ? 0 : 1);
    const cards = Array.from({ length: size }, (_, slotIndex) =>
      buildCardDraft(
        sectionItems[slotIndex] ?? "",
        `${options.cardPrefix}${absoluteOrder + slotIndex}`,
        options.bodyLines,
        options.bodyLength
      )
    );
    absoluteOrder += size;

    return {
      label: labelSource ? buildPlainSlot(labelSource, labelSource) : " ",
      cards
    };
  });
}

function pickImage(outline: OutlineDocument, slide: OutlineSlide | undefined): ExtractedImage | undefined {
  if (!slide || slide.imageIds.length === 0) {
    return undefined;
  }

  return outline.extractedImages.find((image) => image.id === slide.imageIds[0]);
}

function parseDataUri(dataUri: string) {
  const match = dataUri.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    throw new Error("图片数据格式无效。");
  }

  return {
    mimeType: match[1],
    buffer: Buffer.from(match[2], "base64")
  };
}

function extensionFromMime(mimeType: string) {
  switch (mimeType) {
    case "image/jpeg":
      return "jpg";
    case "image/png":
      return "png";
    case "image/gif":
      return "gif";
    case "image/tiff":
      return "tiff";
    case "image/svg+xml":
      return "svg";
    case "image/webp":
      return "webp";
    case "image/bmp":
      return "bmp";
    default:
      return "png";
  }
}

function mimeFromExtension(extension: string) {
  switch (extension) {
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "webp":
      return "image/webp";
    case "bmp":
      return "image/bmp";
    default:
      return "image/png";
  }
}

function findShapeByName(xml: CheerioAPI, name: string, occurrence = 0) {
  return xml("p\\:sp")
    .filter((_, element) => xml("p\\:cNvPr", element).attr("name") === name)
    .eq(occurrence);
}

function ensureTextRun(paragraph: Cheerio<Element>, templateRunXml?: string) {
  const endParagraph = paragraph.children("a\\:endParaRPr").first();
  const runXml = templateRunXml || "<a:r><a:t/></a:r>";
  if (endParagraph.length > 0) {
    endParagraph.before(runXml);
  } else {
    paragraph.append(runXml);
  }

  return paragraph.children("a\\:r").first();
}

function writeParagraphText(paragraph: Cheerio<Element>, text: string) {
  const templateRunXml = paragraph.children("a\\:r").first().toString();
  paragraph.children("a\\:r, a\\:fld, a\\:br").remove();
  const run = ensureTextRun(paragraph, templateRunXml);
  run.children("a\\:t, a\\:fld, a\\:br").remove();
  let textNode = run.children("a\\:t").first();
  if (textNode.length === 0) {
    run.append("<a:t/>");
    textNode = run.children("a\\:t").first();
  }

  const value = text || " ";
  textNode.text(value);
  if (/^\s|\s$/.test(value)) {
    textNode.attr("xml:space", "preserve");
  } else {
    textNode.removeAttr("xml:space");
  }
}

function setShapeParagraphs(
  xml: CheerioAPI,
  shape: Cheerio<Element>,
  paragraphs: string[],
  fallbackText = " ",
  style?: TextStyleSpec
) {
  if (shape.length === 0) {
    return;
  }

  const textBody = shape.children("p\\:txBody").first();
  if (textBody.length === 0) {
    return;
  }

  const existingParagraphs = textBody.children("a\\:p");
  const templateParagraph = existingParagraphs.first().clone();
  const values = paragraphs.length > 0 ? paragraphs : [fallbackText];

  existingParagraphs.remove();

  values.forEach((value) => {
    const nextParagraph = templateParagraph.clone();
    writeParagraphText(nextParagraph, value);
    applyParagraphStyle(xml, nextParagraph, style);
    textBody.append(nextParagraph);
  });
}

function setShapeTextByName(
  xml: CheerioAPI,
  name: string,
  text: string,
  options?: { occurrence?: number; style?: TextStyleSpec }
) {
  const shape = findShapeByName(xml, name, options?.occurrence ?? 0);
  setShapeParagraphs(xml, shape, [text], " ", options?.style);
}

function setShapeLinesByName(
  xml: CheerioAPI,
  name: string,
  lines: string[],
  options?: { occurrence?: number; style?: TextStyleSpec }
) {
  const shape = findShapeByName(xml, name, options?.occurrence ?? 0);
  setShapeParagraphs(xml, shape, lines, " ", options?.style);
}

function applyParagraphStyle(xml: CheerioAPI, paragraph: Cheerio<Element>, style?: TextStyleSpec) {
  if (!style) {
    return;
  }

  paragraph.children("a\\:r").each((_, element) => {
    const run = xml(element);
    let runProperties = run.children("a\\:rPr").first();

    if (runProperties.length === 0) {
      run.prepend("<a:rPr/>");
      runProperties = run.children("a\\:rPr").first();
    }

    if (style.fontSize) {
      runProperties.attr("sz", String(style.fontSize * 100));
    }

    if (style.bold === true) {
      runProperties.attr("b", "1");
    } else if (style.bold === false) {
      runProperties.removeAttr("b");
    }

    if (style.colorHex) {
      runProperties.children("a\\:solidFill").remove();
      runProperties.append(`<a:solidFill><a:srgbClr val="${style.colorHex}"/></a:solidFill>`);
    }
  });
}

function clearTemplatePlaceholderText(xml: CheerioAPI) {
  xml("p\\:sp p\\:txBody a\\:p").each((_, element) => {
    const paragraph = xml(element);
    const text = paragraph
      .find("a\\:t")
      .map((__, node) => xml(node).text())
      .get()
      .join("");

    if (TEMPLATE_PLACEHOLDER_TEXT.some((keyword) => text.includes(keyword))) {
      writeParagraphText(paragraph, " ");
    }
  });
}

function buildSlideDraft(slide: OutlineSlide) {
  const description = normalizeText(slide.content.description ?? "");
  const units = description
    ? dedupeTextItems(splitSummary(description), [
        slide.title,
        slide.content.intro,
        slide.content.regularTitle
      ])
    : [];
  const intro = truncateText(slide.content.intro || slide.summary, SLOT_LIMITS.intro);
  const unitPool = units.length > 0 ? units : [];
  return {
    intro,
    title: buildPlainSlot(slide.title, "要点", SLOT_LIMITS.pageTitle),
    regularTitle: buildPlainSlot(slide.content.regularTitle, "核心内容"),
    description,
    units: unitPool,
    buckets: buildBuckets(unitPool, 6)
  };
}

function buildTableDraft(slide: OutlineSlide) {
  const summaryUnits = splitSummary(slide.summary);
  const rows = slide.table ?? [];
  const columnHeaders =
    rows[0]?.slice(0, 4).map((cell) => truncateText(cell, SLOT_LIMITS.tableHeader)) ??
    new Array(4).fill(0).map((_, index) => `列${index + 1}`);
  const rowHeaders =
    rows
      .slice(1, 5)
      .map((row, index) => truncateText(row[0] || `项${index + 1}`, SLOT_LIMITS.tableRow)) ??
    [];
  const flattenedCells = rows
    .slice(1)
    .flatMap((row) => row.slice(1))
    .map((cell) => normalizeText(cell))
    .filter(Boolean);
  const bodySource = flattenedCells.length > 0 ? flattenedCells : summaryUnits;
  const cellBodies = buildBuckets(bodySource, 6).map((bucket) => buildParagraphLines(bucket, 3, SLOT_LIMITS.tableCell));

  return {
    columnHeaders: [
      columnHeaders[0] ?? "列1",
      columnHeaders[1] ?? "列2",
      columnHeaders[2] ?? "列3",
      columnHeaders[3] ?? "列4"
    ],
    rowHeaders: [
      rowHeaders[0] ?? "项1",
      rowHeaders[1] ?? "项2",
      rowHeaders[2] ?? "项3",
      rowHeaders[3] ?? "项4"
    ],
    cellBodies
  } satisfies TableDraft;
}

function fillCoverSlide(xml: CheerioAPI, outline: OutlineDocument) {
  setShapeTextByName(xml, "标题 1", truncateText(outline.cover.title, SLOT_LIMITS.coverTitle));
  setShapeTextByName(xml, "文本占位符 2", truncateText(outline.cover.subtitle, SLOT_LIMITS.coverSubtitle));
  setShapeTextByName(xml, "文本占位符 3", truncateText(outline.cover.userName, SLOT_LIMITS.coverMeta));
  setShapeTextByName(xml, "文本占位符 4", outline.cover.dateLabel);
}

function fillDirectorySlide(xml: CheerioAPI, outline: OutlineDocument) {
  setShapeTextByName(xml, "Title 1", "目录");
  outline.directory.slice(0, 3).forEach((item, index) => {
    setShapeTextByName(
      xml,
      "TextBox 11",
      truncateText(stripTitlePrefix(item.title), SLOT_LIMITS.directoryLine) || " ",
      { occurrence: index }
    );
    setShapeTextByName(
      xml,
      "TextBox 12",
      truncateText(item.description, SLOT_LIMITS.directoryBrief) || " ",
      { occurrence: index, style: DIRECTORY_BRIEF_STYLE }
    );
  });

  for (let index = outline.directory.length; index < 3; index += 1) {
    setShapeTextByName(xml, "TextBox 11", " ", { occurrence: index });
    setShapeTextByName(xml, "TextBox 12", " ", { occurrence: index });
  }
}

function fillImageSlide(xml: CheerioAPI, outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const cards = buildCardsBySlotSpecs(
    draft.units,
    new Array(4).fill(0).map(() => ({
      maxLines: 2,
      maxLength: SLOT_LIMITS.smallBody,
      chunkLength: SLOT_LIMITS.smallBody
    })),
    "要点"
  );

  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "Rectangle 2", draft.regularTitle || "核心内容", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 5", " ", { style: REGULAR_TITLE_STYLE });
  setShapeTextByName(xml, "TextBox 8", cards[0]?.heading ?? "核心信息", { style: SMALL_TITLE_STYLE });
  setShapeLinesByName(xml, "TextBox 9", cards[0]?.bodyLines ?? [" "], { style: DESCRIPTION_STYLE });

  [
    { title: "TextBox 11", body: "TextBox 12", index: 1 },
    { title: "TextBox 13", body: "TextBox 14", index: 2 },
    { title: "TextBox 15", body: "TextBox 16", index: 3 }
  ].forEach((slot) => {
    setShapeTextByName(xml, slot.title, cards[slot.index]?.heading ?? " ", { style: SMALL_TITLE_STYLE });
    setShapeLinesByName(xml, slot.body, cards[slot.index]?.bodyLines ?? [" "], {
      style: DESCRIPTION_STYLE
    });
  });
}

function fillTwoColumnSlide(xml: CheerioAPI, outline: OutlineDocument, slide?: OutlineSlide) {
  fillImageSlide(xml, outline, slide);
}

function fillThreeColumnSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const sections = buildSectionDrafts(draft.units, [2, 2, 2], {
    sectionPrefix: "模块",
    preferredLabels: [draft.regularTitle, "", ""],
    cardPrefix: "要点",
    bodyLines: 2,
    bodyLength: SLOT_LIMITS.smallBody,
    chunkLength: SLOT_LIMITS.smallBody
  });
  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "Rectangle 2", (sections[0]?.label ?? draft.regularTitle) || "模块一", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 18", sections[1]?.label ?? "模块二", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 22", sections[2]?.label ?? "模块三", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "TextBox 11", sections[0]?.cards[0]?.heading ?? " ", {
    occurrence: 0,
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 12", sections[0]?.cards[0]?.bodyLines ?? [" "], {
    occurrence: 0,
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 26", sections[0]?.cards[1]?.heading ?? " ", {
    occurrence: 0,
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 27", sections[0]?.cards[1]?.bodyLines ?? [" "], {
    occurrence: 0,
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 11", sections[1]?.cards[0]?.heading ?? " ", {
    occurrence: 1,
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 12", sections[1]?.cards[0]?.bodyLines ?? [" "], {
    occurrence: 1,
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 26", sections[1]?.cards[1]?.heading ?? " ", {
    occurrence: 1,
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 27", sections[1]?.cards[1]?.bodyLines ?? [" "], {
    occurrence: 1,
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 23", sections[2]?.cards[0]?.heading ?? " ", {
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 24", sections[2]?.cards[0]?.bodyLines ?? [" "], {
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 30", sections[2]?.cards[1]?.heading ?? " ", {
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 31", sections[2]?.cards[1]?.bodyLines ?? [" "], {
    style: DESCRIPTION_STYLE
  });
}

function fillFourColumnSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const sections = buildSectionDrafts(draft.units, [3, 3, 3, 3], {
    sectionPrefix: "分区",
    preferredLabels: [draft.regularTitle, "", "", ""],
    cardPrefix: "要点",
    bodyLines: 3,
    bodyLength: SLOT_LIMITS.smallBody,
    chunkLength: SLOT_LIMITS.smallBody
  });
  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });

  [
    { label: "Rectangle 18", heading: "TextBox 19", body: "TextBox 20", sectionIndex: 0, cardIndex: 0 },
    { label: "Rectangle 18", heading: "TextBox 51", body: "TextBox 52", sectionIndex: 0, cardIndex: 1 },
    { label: "Rectangle 18", heading: "TextBox 53", body: "TextBox 54", sectionIndex: 0, cardIndex: 2 },
    { label: "Rectangle 56", heading: "TextBox 57", body: "TextBox 58", sectionIndex: 1, cardIndex: 0 },
    { label: "Rectangle 56", heading: "TextBox 59", body: "TextBox 60", sectionIndex: 1, cardIndex: 1 },
    { label: "Rectangle 56", heading: "TextBox 61", body: "TextBox 62", sectionIndex: 1, cardIndex: 2 },
    { label: "Rectangle 64", heading: "TextBox 65", body: "TextBox 66", sectionIndex: 2, cardIndex: 0 },
    { label: "Rectangle 64", heading: "TextBox 67", body: "TextBox 68", sectionIndex: 2, cardIndex: 1 },
    { label: "Rectangle 64", heading: "TextBox 69", body: "TextBox 70", sectionIndex: 2, cardIndex: 2 },
    { label: "Rectangle 72", heading: "TextBox 73", body: "TextBox 74", sectionIndex: 3, cardIndex: 0 },
    { label: "Rectangle 72", heading: "TextBox 75", body: "TextBox 76", sectionIndex: 3, cardIndex: 1 },
    { label: "Rectangle 72", heading: "TextBox 77", body: "TextBox 78", sectionIndex: 3, cardIndex: 2 }
  ].forEach((slot) => {
    const card = sections[slot.sectionIndex]?.cards[slot.cardIndex];
    setShapeTextByName(xml, slot.label, sections[slot.sectionIndex]?.label ?? " ", {
      style: REGULAR_TITLE_STYLE
    });
    setShapeTextByName(xml, slot.heading, card?.heading ?? " ", { style: SMALL_TITLE_STYLE });
    setShapeLinesByName(xml, slot.body, card?.bodyLines ?? [" "], { style: DESCRIPTION_STYLE });
  });
}

function fillProgressSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const sections = buildSectionDrafts(draft.units, [3, 2, 3, 2], {
    sectionPrefix: "阶段",
    preferredLabels: [draft.regularTitle, "", "", ""],
    cardPrefix: "步骤",
    bodyLines: 2,
    bodyLength: SLOT_LIMITS.smallBody,
    chunkLength: SLOT_LIMITS.smallBody
  });
  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "Pentagon 18", (sections[0]?.label ?? draft.regularTitle) || "阶段一", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Chevron 56", sections[1]?.label ?? "阶段二", { style: REGULAR_TITLE_STYLE });
  setShapeTextByName(xml, "Chevron 36", sections[2]?.label ?? "阶段三", { style: REGULAR_TITLE_STYLE });
  setShapeTextByName(xml, "Chevron 37", sections[3]?.label ?? "阶段四", { style: REGULAR_TITLE_STYLE });

  [
    { heading: "TextBox 19", body: "TextBox 20", sectionIndex: 0, cardIndex: 0 },
    { heading: "TextBox 51", body: "TextBox 52", sectionIndex: 0, cardIndex: 1 },
    { heading: "TextBox 53", body: "TextBox 54", sectionIndex: 0, cardIndex: 2 },
    { heading: "TextBox 40", body: "TextBox 41", sectionIndex: 1, cardIndex: 0 },
    { heading: "TextBox 48", body: "TextBox 49", sectionIndex: 1, cardIndex: 1 },
    { heading: "TextBox 82", body: "TextBox 83", sectionIndex: 2, cardIndex: 0 },
    { heading: "TextBox 84", body: "TextBox 85", sectionIndex: 2, cardIndex: 1 },
    { heading: "TextBox 86", body: "TextBox 87", sectionIndex: 2, cardIndex: 2 },
    { heading: "TextBox 88", body: "TextBox 89", sectionIndex: 3, cardIndex: 0 },
    { heading: "TextBox 91", body: "TextBox 92", sectionIndex: 3, cardIndex: 1 }
  ].forEach((slot) => {
    const card = sections[slot.sectionIndex]?.cards[slot.cardIndex];
    setShapeTextByName(xml, slot.heading, card?.heading ?? " ", { style: SMALL_TITLE_STYLE });
    setShapeLinesByName(xml, slot.body, card?.bodyLines ?? [" "], { style: DESCRIPTION_STYLE });
  });
}

function fillVerticalSplitSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const sections = buildSectionDrafts(draft.units, [2, 2], {
    sectionPrefix: "板块",
    preferredLabels: [draft.regularTitle, ""],
    cardPrefix: "要点",
    bodyLines: 3,
    bodyLength: SLOT_LIMITS.smallBody,
    chunkLength: SLOT_LIMITS.smallBody
  });

  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "Rectangle 18", draft.regularTitle || "上部模块", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 38", sections[1]?.label ?? "下部模块", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "TextBox 19", sections[0]?.cards[0]?.heading ?? " ", {
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 20", sections[0]?.cards[0]?.bodyLines ?? [" "], {
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 35", sections[0]?.cards[1]?.heading ?? " ", {
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 36", sections[0]?.cards[1]?.bodyLines ?? [" "], {
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 39", sections[1]?.cards[0]?.heading ?? " ", {
    occurrence: 0,
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 40", sections[1]?.cards[0]?.bodyLines ?? [" "], {
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "TextBox 41", sections[1]?.cards[1]?.heading ?? " ", {
    style: SMALL_TITLE_STYLE
  });
  setShapeLinesByName(xml, "TextBox 42", sections[1]?.cards[1]?.bodyLines ?? [" "], {
    style: DESCRIPTION_STYLE
  });
}

function fillSplitGridSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const sections = buildSectionDrafts(draft.units, [3, 3], {
    sectionPrefix: "板块",
    preferredLabels: [draft.regularTitle, ""],
    cardPrefix: "要点",
    bodyLines: 4,
    bodyLength: SLOT_LIMITS.smallBody,
    chunkLength: SLOT_LIMITS.smallBody
  });
  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "Rectangle 18", (sections[0]?.label ?? draft.regularTitle) || "上部板块", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 38", sections[1]?.label ?? "下部板块", {
    style: REGULAR_TITLE_STYLE
  });

  [
    { heading: "TextBox 19", body: "TextBox 20", sectionIndex: 0, cardIndex: 0 },
    { heading: "TextBox 15", body: "TextBox 16", sectionIndex: 0, cardIndex: 1 },
    { heading: "TextBox 21", body: "TextBox 22", sectionIndex: 0, cardIndex: 2 },
    { heading: "TextBox 23", body: "TextBox 24", sectionIndex: 1, cardIndex: 0 },
    { heading: "TextBox 25", body: "TextBox 26", sectionIndex: 1, cardIndex: 1 },
    { heading: "TextBox 27", body: "TextBox 28", sectionIndex: 1, cardIndex: 2 }
  ].forEach((slot) => {
    const card = sections[slot.sectionIndex]?.cards[slot.cardIndex];
    setShapeTextByName(xml, slot.heading, card?.heading ?? " ", { style: SMALL_TITLE_STYLE });
    setShapeLinesByName(xml, slot.body, card?.bodyLines ?? [" "], { style: DESCRIPTION_STYLE });
  });
}

function fillOverviewSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const topUnits = draft.units.slice(0, 2);
  const lowerSections = buildSectionDrafts(draft.units.slice(2), [1, 1, 1], {
    sectionPrefix: "分项",
    preferredLabels: ["", "", ""],
    cardPrefix: "要点",
    bodyLines: 2,
    bodyLength: SLOT_LIMITS.regularLine,
    chunkLength: SLOT_LIMITS.regularLine
  });

  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "Rectangle 18", draft.regularTitle || "总体概述", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeLinesByName(
    xml,
    "TextBox 20",
    buildBodyLines(
      joinBucket(topUnits.length > 0 ? topUnits : [draft.description], SLOT_LIMITS.largeBody),
      3,
      SLOT_LIMITS.regularLine
    ),
    { style: DESCRIPTION_STYLE }
  );

  [
    { badge: "Rectangle 38", heading: "TextBox 39", body: "TextBox 40", index: 0 },
    { badge: "Rectangle 24", heading: "TextBox 25", body: "TextBox 26", index: 1 },
    { badge: "Rectangle 28", heading: "TextBox 29", body: "TextBox 30", index: 2 }
  ].forEach((slot) => {
    const section = lowerSections[slot.index];
    setShapeTextByName(xml, slot.badge, section?.label ?? " ", { style: REGULAR_TITLE_STYLE });
    setShapeTextByName(xml, slot.heading, section?.cards[0]?.heading ?? " ", {
      style: SMALL_TITLE_STYLE
    });
    setShapeLinesByName(xml, slot.body, section?.cards[0]?.bodyLines ?? [" "], {
      style: DESCRIPTION_STYLE
    });
  });
}

function fillHierarchySlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const cards = buildCardsBySlotSpecs(
    draft.units,
    [
      { maxLines: 2, maxLength: SLOT_LIMITS.regularLine, chunkLength: SLOT_LIMITS.regularLine },
      { maxLines: 2, maxLength: SLOT_LIMITS.regularLine, chunkLength: SLOT_LIMITS.regularLine },
      { maxLines: 3, maxLength: SLOT_LIMITS.compactLine, chunkLength: SLOT_LIMITS.compactLine },
      { maxLines: 3, maxLength: SLOT_LIMITS.compactLine, chunkLength: SLOT_LIMITS.compactLine },
      { maxLines: 3, maxLength: SLOT_LIMITS.compactLine, chunkLength: SLOT_LIMITS.compactLine }
    ],
    "层级"
  );

  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });
  setShapeTextByName(xml, "TextBox 2", truncateText(draft.description, SLOT_LIMITS.largeBody), {
    style: DESCRIPTION_STYLE
  });
  setShapeTextByName(xml, "Down Arrow Callout 18", draft.regularTitle || cards[0]?.heading || "模块一", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 31", cards[1]?.heading || "主题总览", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Down Arrow Callout 32", cards[2]?.heading ?? "模块三", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Down Arrow Callout 33", cards[3]?.heading ?? "模块四", {
    style: REGULAR_TITLE_STYLE
  });
  setShapeTextByName(xml, "Rectangle 5", " ", { style: REGULAR_TITLE_STYLE });
  setShapeTextByName(xml, "Rectangle 55", " ", { style: REGULAR_TITLE_STYLE });
  setShapeTextByName(xml, "Rectangle 56", " ", { style: REGULAR_TITLE_STYLE });
  setShapeTextByName(xml, "Rectangle 57", " ", { style: REGULAR_TITLE_STYLE });

  [
    { heading: "TextBox 19", body: "TextBox 20", index: 0 },
    { heading: "TextBox 36", body: "TextBox 53", index: 1 },
    { heading: "TextBox 51", body: "TextBox 52", index: 2 },
    { heading: "TextBox 59", body: "TextBox 60", index: 3 },
    { heading: "TextBox 61", body: "TextBox 62", index: 4 }
  ].forEach((slot) => {
    const card = cards[slot.index];
    setShapeTextByName(xml, slot.heading, card?.heading ?? " ", { style: SMALL_TITLE_STYLE });
    setShapeLinesByName(xml, slot.body, card?.bodyLines ?? [" "], { style: DESCRIPTION_STYLE });
  });
}

function fillTableSlide(xml: CheerioAPI, _outline: OutlineDocument, slide?: OutlineSlide) {
  if (!slide) {
    return;
  }

  const draft = buildSlideDraft(slide);
  const tableDraft = buildTableDraft(slide);

  setShapeTextByName(xml, "Title 1", draft.title, { style: REFERENCE_TITLE_STYLE });
  setShapeTextByName(xml, "文本框 39", draft.intro || " ", { style: INTRO_STYLE });

  [
    { name: "Rectangle 5", value: tableDraft.columnHeaders[0] },
    { name: "Rectangle 55", value: tableDraft.columnHeaders[1] },
    { name: "Rectangle 56", value: tableDraft.columnHeaders[2] },
    { name: "Rectangle 57", value: tableDraft.columnHeaders[3] }
  ].forEach((item) => {
    setShapeTextByName(xml, item.name, item.value);
  });

  [
    { name: "Pentagon 18", value: tableDraft.rowHeaders[0] },
    { name: "Pentagon 32", value: tableDraft.rowHeaders[1] },
    { name: "Pentagon 33", value: tableDraft.rowHeaders[2] },
    { name: "Pentagon 31", value: tableDraft.rowHeaders[3] }
  ].forEach((item) => {
    setShapeTextByName(xml, item.name, item.value);
  });

  [
    { name: "TextBox 30", lines: tableDraft.cellBodies[0] },
    { name: "TextBox 38", lines: tableDraft.cellBodies[1] },
    { name: "TextBox 34", lines: tableDraft.cellBodies[2] },
    { name: "TextBox 40", lines: tableDraft.cellBodies[3] },
    { name: "TextBox 39", lines: tableDraft.cellBodies[4] },
    { name: "TextBox 37", lines: tableDraft.cellBodies[5] }
  ].forEach((item) => {
    setShapeLinesByName(xml, item.name, item.lines.length > 0 ? item.lines : [" "]);
  });
}

function serializeXml(xml: CheerioAPI) {
  return `${XML_HEADER}${xml.root().children().toString()}`;
}

function updatePresentationXml(presentationXml: string, slideCount: number) {
  const xml = loadXml(presentationXml, { xmlMode: true });
  const slideList = xml("p\\:sldIdLst").first();
  slideList.empty();

  Array.from({ length: slideCount }, (_, index) => index + 1).forEach((slideNumber, index) => {
    slideList.append(
      `<p:sldId id="${256 + index}" r:id="rId${SLIDE_RELATIONSHIP_BASE_ID + slideNumber - 1}"/>`
    );
  });

  return serializeXml(xml);
}

function updatePresentationRelationships(presentationRelsXml: string, slideCount: number) {
  const xml = loadXml(presentationRelsXml, { xmlMode: true });
  const root = xml("Relationships").first();

  xml("Relationship")
    .filter((_, element) => Boolean(xml(element).attr("Type")?.endsWith("/slide")))
    .remove();

  Array.from({ length: slideCount }, (_, index) => index + 1).forEach((slideNumber, index) => {
    root.append(
      `<Relationship Id="rId${SLIDE_RELATIONSHIP_BASE_ID + index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNumber}.xml"/>`
    );
  });

  return serializeXml(xml);
}

function updateAppXml(appXml: string, slideCount: number) {
  const xml = loadXml(appXml, { xmlMode: true });
  xml("Slides").text(String(slideCount));
  return serializeXml(xml);
}

function updateContentTypeDefaults(contentTypesXml: string, extensions: Set<string>) {
  const xml = loadXml(contentTypesXml, { xmlMode: true });
  const root = xml("Types").first();
  const existing = new Set(
    xml("Default")
      .map((_, element) => xml(element).attr("Extension"))
      .get()
      .filter(Boolean)
  );

  extensions.forEach((extension) => {
    if (existing.has(extension)) {
      return;
    }

    root.append(
      `<Default Extension="${extension}" ContentType="${mimeFromExtension(extension)}"/>`
    );
  });

  return serializeXml(xml);
}

function updateSlideOverrides(contentTypesXml: string, slideCount: number, highestExistingSlide: number) {
  const xml = loadXml(contentTypesXml, { xmlMode: true });
  const root = xml("Types").first();
  const overrides = xml("Override");

  overrides
    .filter((_, element) => {
      const partName = xml(element).attr("PartName") ?? "";
      return /^\/ppt\/slides\/slide\d+\.xml$/.test(partName);
    })
    .remove();

  Array.from({ length: Math.max(slideCount, highestExistingSlide) }, (_, index) => index + 1).forEach(
    (slideNumber) => {
      root.append(
        `<Override PartName="/ppt/slides/slide${slideNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
      );
    }
  );

  return serializeXml(xml);
}

function retargetImageRelationship(relsXml: string, target: string) {
  const xml = loadXml(relsXml, { xmlMode: true });
  const imageRelationship = xml("Relationship")
    .filter((_, element) => Boolean(xml(element).attr("Type")?.endsWith("/image")))
    .first();

  if (imageRelationship.length > 0) {
    imageRelationship.attr("Target", target);
  }

  return serializeXml(xml);
}

function buildDetailPlan(slide: OutlineSlide) {
  return DETAIL_TEMPLATE_MAP[slide.type] ?? DETAIL_TEMPLATE_MAP.overview;
}

async function loadSourceSlides(zip: JSZip, slideNumbers: number[]) {
  const unique = Array.from(new Set(slideNumbers));
  const cache = new Map<number, { xml: string; rels: string }>();

  for (const slideNumber of unique) {
    const xml = await zip.file(`ppt/slides/slide${slideNumber}.xml`)?.async("string");
    const rels =
      (await zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)?.async("string")) ??
      `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

    if (!xml) {
      throw new Error(`模板缺少第 ${slideNumber} 页，无法生成导出文件。`);
    }

    cache.set(slideNumber, { xml, rels });
  }

  return cache;
}

function getHighestExistingSlide(zip: JSZip) {
  const slideNumbers = Object.keys(zip.files)
    .map((name) => name.match(/^ppt\/slides\/slide(\d+)\.xml$/)?.[1])
    .filter(Boolean)
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));

  return slideNumbers.length > 0 ? Math.max(...slideNumbers) : 0;
}

export async function generatePptBuffer(outline: OutlineDocument): Promise<Buffer> {
  const normalizedOutline = normalizeOutlineDocument(outline);
  const templateBuffer = await readFile(getTemplatePath());
  const zip = await JSZip.loadAsync(templateBuffer);
  const highestExistingSlide = getHighestExistingSlide(zip);
  const directorySourceSlide = highestExistingSlide;
  const slideCount = normalizedOutline.slides.length + 2;
  const detailSources = normalizedOutline.slides.map((slide) => buildDetailPlan(slide).sourceSlide);
  const sourceSlides = await loadSourceSlides(zip, [COVER_SOURCE_SLIDE, directorySourceSlide, ...detailSources]);
  const extraExtensions = new Set<string>(["png"]);

  zip.file("ppt/slides/media/generated-blank.png", Buffer.from(BLANK_IMAGE_BASE64, "base64"));

  {
    const coverSource = sourceSlides.get(COVER_SOURCE_SLIDE);
    if (!coverSource) {
      throw new Error("模板缺少封面页。");
    }
    const xml = loadXml(coverSource.xml, { xmlMode: true });
    fillCoverSlide(xml, normalizedOutline);
    clearTemplatePlaceholderText(xml);
    zip.file("ppt/slides/slide1.xml", serializeXml(xml));
    zip.file("ppt/slides/_rels/slide1.xml.rels", coverSource.rels);
  }

  {
    const directorySource = sourceSlides.get(directorySourceSlide);
    if (!directorySource) {
      throw new Error("模板缺少目录参考页。");
    }
    const xml = loadXml(directorySource.xml, { xmlMode: true });
    fillDirectorySlide(xml, normalizedOutline);
    clearTemplatePlaceholderText(xml);
    zip.file("ppt/slides/slide2.xml", serializeXml(xml));
    zip.file("ppt/slides/_rels/slide2.xml.rels", directorySource.rels);
  }

  for (const [index, slide] of normalizedOutline.slides.entries()) {
    const outputSlideNumber = index + 3;
    const plan = buildDetailPlan(slide);
    const source = sourceSlides.get(plan.sourceSlide);
    if (!source) {
      throw new Error(`模板缺少第 ${plan.sourceSlide} 页参考版式。`);
    }

    const xml = loadXml(source.xml, { xmlMode: true });
    plan.fill(xml, normalizedOutline, slide);
    clearTemplatePlaceholderText(xml);

    let rels = source.rels;
    if (plan.withImage) {
      const image = pickImage(normalizedOutline, slide);
      const fileName = image
        ? (() => {
            const { mimeType, buffer } = parseDataUri(image.dataUri);
            const extension = extensionFromMime(mimeType);
            const mediaName = `generated-slide-${outputSlideNumber}.${extension}`;
            zip.file(`ppt/slides/media/${mediaName}`, buffer);
            extraExtensions.add(extension);
            return mediaName;
          })()
        : "generated-blank.png";
      rels = retargetImageRelationship(rels, `media/${fileName}`);
    }

    zip.file(`ppt/slides/slide${outputSlideNumber}.xml`, serializeXml(xml));
    zip.file(`ppt/slides/_rels/slide${outputSlideNumber}.xml.rels`, rels);
  }

  const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
  const presentationRelsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string");
  const appXml = await zip.file("docProps/app.xml")?.async("string");
  const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");

  if (!presentationXml || !presentationRelsXml || !appXml || !contentTypesXml) {
    throw new Error("模板文件结构不完整，缺少导出所需的元数据。");
  }

  zip.file("ppt/presentation.xml", updatePresentationXml(presentationXml, slideCount));
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    updatePresentationRelationships(presentationRelsXml, slideCount)
  );
  zip.file("docProps/app.xml", updateAppXml(appXml, slideCount));
  zip.file(
    "[Content_Types].xml",
    updateSlideOverrides(
      updateContentTypeDefaults(contentTypesXml, extraExtensions),
      slideCount,
      highestExistingSlide
    )
  );

  return await zip.generateAsync({ type: "nodebuffer" });
}
