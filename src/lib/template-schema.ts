import { execFile } from "node:child_process";
import { join } from "node:path";
import { promisify } from "node:util";
import { getTemplatePath, loadTemplateTokens } from "@/lib/template";
import type {
  SlideLayoutType,
  TemplateLayoutSchema,
  TemplatePlaceholderKind,
  TemplatePlaceholderSchema,
  TemplateSchemaLibrary
} from "@/lib/types";

const execFileAsync = promisify(execFile);

export const FIELD_LIMITS = {
  coverTitle: 24,
  coverSubtitle: 38,
  coverMeta: 48,
  directoryTitle: 5,
  directoryDescription: 10,
  pageTitle: 10,
  intro: 30,
  regularTitle: 15,
  smallTitle: 15,
  detailBody: 60,
  overviewLeadBody: 120,
  hierarchyLeadBody: 120,
  tableHeader: 10,
  tableRow: 10,
  tableCell: 18
} as const;

type PlaceholderRefConfig = {
  name: string;
  occurrence?: number;
  role: string;
  fallbackMaxChars: number;
  kind?: TemplatePlaceholderKind;
  maxItems?: number;
  description?: string;
};

type LayoutConfig = {
  id: string;
  name: string;
  layoutType: SlideLayoutType;
  sourceSlide: number;
  summary: string;
  detailItems: number;
  refs: PlaceholderRefConfig[];
};

type ScannedPlaceholder = {
  id: string;
  name: string;
  occurrence: number;
  kind: TemplatePlaceholderKind;
  sampleText?: string;
  placeholderType?: string | null;
  placeholderIndex?: string | null;
  xPt?: number;
  yPt?: number;
  widthPt?: number;
  heightPt?: number;
  fontSizePt?: number;
  maxChars?: number;
  maxLines?: number;
};

type ScannedSlide = {
  sourceSlide: number;
  placeholderCount: number;
  capacities: {
    textCount: number;
    pictureSlots: number;
    tableSlots: number;
    chartSlots: number;
  };
  placeholders: ScannedPlaceholder[];
};

type TemplateScanPayload = {
  templatePath: string;
  slideCount: number;
  slides: ScannedSlide[];
};

function ref(
  name: string,
  role: string,
  fallbackMaxChars: number,
  options?: {
    occurrence?: number;
    kind?: TemplatePlaceholderKind;
    maxItems?: number;
    description?: string;
  }
): PlaceholderRefConfig {
  return {
    name,
    role,
    fallbackMaxChars,
    occurrence: options?.occurrence,
    kind: options?.kind,
    maxItems: options?.maxItems,
    description: options?.description
  };
}

function textPlaceholder(
  id: string,
  name: string,
  role: string,
  maxChars: number,
  occurrence = 1
): TemplatePlaceholderSchema {
  return {
    id,
    name,
    role,
    kind: "TEXT",
    occurrence,
    maxChars
  };
}

const coverRefs: PlaceholderRefConfig[] = [
  ref("标题 1", "coverTitle", FIELD_LIMITS.coverTitle),
  ref("文本占位符 2", "coverSubtitle", FIELD_LIMITS.coverSubtitle),
  ref("文本占位符 3", "coverAuthor", FIELD_LIMITS.coverMeta),
  ref("文本占位符 4", "coverDate", FIELD_LIMITS.coverMeta)
];

const directoryRefs: PlaceholderRefConfig[] = [
  ref("TextBox 11", "directoryTitle", FIELD_LIMITS.directoryTitle, { occurrence: 1 }),
  ref("TextBox 11", "directoryTitle", FIELD_LIMITS.directoryTitle, { occurrence: 2 }),
  ref("TextBox 11", "directoryTitle", FIELD_LIMITS.directoryTitle, { occurrence: 3 }),
  ref("TextBox 12", "directoryDescription", FIELD_LIMITS.directoryDescription, { occurrence: 1 }),
  ref("TextBox 12", "directoryDescription", FIELD_LIMITS.directoryDescription, { occurrence: 2 }),
  ref("TextBox 12", "directoryDescription", FIELD_LIMITS.directoryDescription, { occurrence: 3 })
];

const layoutConfigs: LayoutConfig[] = [
  {
    id: "layout-13-overview",
    name: "总体结构",
    layoutType: "overview",
    sourceSlide: 13,
    summary: "概述型版式，适合内容较少或需要总览+三点拆解的页面。",
    detailItems: 5,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 20", "detailBody", FIELD_LIMITS.overviewLeadBody, { maxItems: 2 }),
      ref("Rectangle 38", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 39", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 40", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("Rectangle 24", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 25", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 26", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("Rectangle 28", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 29", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 30", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-04-two-column",
    name: "左右结构 1:1",
    layoutType: "two-column",
    sourceSlide: 4,
    summary: "带图片区的左右结构，适合有原图且要点不多的页面。",
    detailItems: 4,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 2", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 8", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 9", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 11", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 12", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 13", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 14", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 15", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 16", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-05-three-column",
    name: "左右结构 1:1:1",
    layoutType: "three-column",
    sourceSlide: 5,
    summary: "三列均分版式，适合 4-6 个并列要点。",
    detailItems: 6,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 2", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 22", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 11", "smallTitle", FIELD_LIMITS.smallTitle, { occurrence: 1 }),
      ref("TextBox 12", "detailBody", FIELD_LIMITS.detailBody, { occurrence: 1, maxItems: 1 }),
      ref("TextBox 26", "smallTitle", FIELD_LIMITS.smallTitle, { occurrence: 1 }),
      ref("TextBox 27", "detailBody", FIELD_LIMITS.detailBody, { occurrence: 1, maxItems: 1 }),
      ref("TextBox 11", "smallTitle", FIELD_LIMITS.smallTitle, { occurrence: 2 }),
      ref("TextBox 12", "detailBody", FIELD_LIMITS.detailBody, { occurrence: 2, maxItems: 1 }),
      ref("TextBox 26", "smallTitle", FIELD_LIMITS.smallTitle, { occurrence: 2 }),
      ref("TextBox 27", "detailBody", FIELD_LIMITS.detailBody, { occurrence: 2, maxItems: 1 }),
      ref("TextBox 23", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 24", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 30", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 31", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-09-four-column",
    name: "左右结构 1:1:1:1",
    layoutType: "four-column",
    sourceSlide: 9,
    summary: "四栏高密结构，适合大量并列短要点。",
    detailItems: 12,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 56", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 64", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 72", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 19", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 20", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 51", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 52", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 53", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 54", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 57", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 58", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 59", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 60", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 61", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 62", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 65", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 66", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 67", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 68", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 69", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 70", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 73", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 74", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 75", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 76", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 77", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 78", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-10-progress",
    name: "推进图",
    layoutType: "progress",
    sourceSlide: 10,
    summary: "推进/路径/阶段型版式，适合步骤推进和里程碑内容。",
    detailItems: 10,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Pentagon 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Chevron 56", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Chevron 36", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Chevron 37", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 19", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 20", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 51", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 52", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 53", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 54", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 40", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 41", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 48", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 49", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 82", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 83", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 84", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 85", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 86", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 87", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 88", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 89", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 91", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 92", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-11-vertical",
    name: "上下结构",
    layoutType: "vertical",
    sourceSlide: 11,
    summary: "上下双层版式，适合 2 个板块各自展开。",
    detailItems: 4,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 38", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 19", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 20", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 35", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 36", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 39", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 40", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 41", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 42", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-12-split-grid",
    name: "上下分区结构",
    layoutType: "split-grid",
    sourceSlide: 12,
    summary: "上下两大块、每块多要点的密集版式。",
    detailItems: 6,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 38", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 19", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 20", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 15", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 16", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 21", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 22", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 23", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 24", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 25", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 26", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 27", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 28", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-15-hierarchy",
    name: "层级结构",
    layoutType: "hierarchy",
    sourceSlide: 15,
    summary: "架构/体系/能力分层内容优先选用的层级页。",
    detailItems: 5,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Down Arrow Callout 18", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Rectangle 31", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Down Arrow Callout 32", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("Down Arrow Callout 33", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 2", "detailSummary", FIELD_LIMITS.hierarchyLeadBody),
      ref("TextBox 19", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 20", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 36", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 53", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 51", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 52", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 59", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 60", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 61", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 62", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-16-table",
    name: "表格结构",
    layoutType: "table",
    sourceSlide: 16,
    summary: "原文包含表格时优先匹配的表格式页面。",
    detailItems: 6,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 5", "tableHeader", FIELD_LIMITS.tableHeader),
      ref("Rectangle 55", "tableHeader", FIELD_LIMITS.tableHeader),
      ref("Rectangle 56", "tableHeader", FIELD_LIMITS.tableHeader),
      ref("Rectangle 57", "tableHeader", FIELD_LIMITS.tableHeader),
      ref("Pentagon 18", "tableRow", FIELD_LIMITS.tableRow),
      ref("Pentagon 32", "tableRow", FIELD_LIMITS.tableRow),
      ref("Pentagon 33", "tableRow", FIELD_LIMITS.tableRow),
      ref("Pentagon 31", "tableRow", FIELD_LIMITS.tableRow),
      ref("TextBox 30", "tableCell", FIELD_LIMITS.tableCell),
      ref("TextBox 38", "tableCell", FIELD_LIMITS.tableCell),
      ref("TextBox 34", "tableCell", FIELD_LIMITS.tableCell),
      ref("TextBox 40", "tableCell", FIELD_LIMITS.tableCell),
      ref("TextBox 39", "tableCell", FIELD_LIMITS.tableCell),
      ref("TextBox 37", "tableCell", FIELD_LIMITS.tableCell)
    ]
  },
  {
    id: "layout-04-image",
    name: "图片页",
    layoutType: "image",
    sourceSlide: 4,
    summary: "图片优先版式。",
    detailItems: 4,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 2", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 8", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 9", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 11", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 12", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 13", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 14", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 15", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 16", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-04-image-left",
    name: "图片左置",
    layoutType: "image-left",
    sourceSlide: 4,
    summary: "图片在左、文字在右的图片版式。",
    detailItems: 4,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 2", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 8", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 9", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 11", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 12", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 13", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 14", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 15", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 16", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  },
  {
    id: "layout-04-image-right",
    name: "图片右置",
    layoutType: "image-right",
    sourceSlide: 4,
    summary: "图片在右、文字在左的图片版式。",
    detailItems: 4,
    refs: [
      ref("Title 1", "pageTitle", FIELD_LIMITS.pageTitle),
      ref("文本框 39", "intro", FIELD_LIMITS.intro),
      ref("Rectangle 2", "regularTitle", FIELD_LIMITS.regularTitle),
      ref("TextBox 8", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 9", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 11", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 12", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 13", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 14", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 }),
      ref("TextBox 15", "smallTitle", FIELD_LIMITS.smallTitle),
      ref("TextBox 16", "detailBody", FIELD_LIMITS.detailBody, { maxItems: 1 })
    ]
  }
];

let cachedScanPromise: Promise<TemplateScanPayload | null> | null = null;

function placeholderId(slideNumber: number, name: string, occurrence = 1) {
  return `slide-${slideNumber}:${name}:${occurrence}`;
}

async function loadPythonTemplateScan(): Promise<TemplateScanPayload | null> {
  if (!cachedScanPromise) {
    cachedScanPromise = (async () => {
      try {
        const scriptPath = join(process.cwd(), "scripts", "scan_template_schema.py");
        const { stdout } = await execFileAsync("python3", [scriptPath, getTemplatePath()], {
          maxBuffer: 20 * 1024 * 1024
        });
        return JSON.parse(stdout) as TemplateScanPayload;
      } catch (error) {
        console.warn("Python template scanner failed, falling back to static schema.", error);
        return null;
      }
    })();
  }

  return cachedScanPromise;
}

function normalizeMatchText(value: string | undefined) {
  return (value ?? "").replace(/\s+/g, "").toLowerCase();
}

function scannedPlaceholderKey(placeholder: ScannedPlaceholder) {
  return `${placeholder.name}#${placeholder.occurrence}`;
}

function scoreScannedPlaceholder(item: PlaceholderRefConfig, placeholder: ScannedPlaceholder) {
  if (item.kind && placeholder.kind !== item.kind) {
    return Number.NEGATIVE_INFINITY;
  }

  const text = normalizeMatchText(placeholder.sampleText);
  const name = normalizeMatchText(placeholder.name);
  const placeholderType = normalizeMatchText(placeholder.placeholderType ?? undefined);
  const font = placeholder.fontSizePt ?? 0;
  const y = placeholder.yPt ?? Number.MAX_SAFE_INTEGER;
  const height = placeholder.heightPt ?? 0;
  const width = placeholder.widthPt ?? 0;
  const maxChars = placeholder.maxChars ?? 0;
  let score = 0;

  if (normalizeMatchText(item.name) === name) {
    score += 260;
    if ((item.occurrence ?? 1) === placeholder.occurrence) {
      score += 80;
    }
  }

  switch (item.role) {
    case "coverTitle":
      if (placeholderType === "ctrtitle" || name.includes("标题")) score += 180;
      if (font >= 24) score += 80;
      if (y < 120) score += 40;
      break;
    case "coverSubtitle":
      if (placeholderType === "subTitle".toLowerCase() || name.includes("文本占位符2")) score += 160;
      if (font >= 16 && font < 24) score += 50;
      if (y > 100 && y < 220) score += 30;
      break;
    case "coverAuthor":
      if (name.includes("文本占位符3")) score += 180;
      if (y > 320) score += 30;
      break;
    case "coverDate":
      if (name.includes("文本占位符4")) score += 180;
      if (y > 360) score += 30;
      break;
    case "directoryTitle":
      if (text.includes("这是个示例")) score += 160;
      if (font >= 13 && font <= 15) score += 30;
      if (y > 200 && y < 280) score += 30;
      break;
    case "directoryDescription":
      if (text.includes("5g的性能目标") || text.includes("说明性文字")) score += 150;
      if (font <= 12.5) score += 30;
      if (y > 260) score += 30;
      break;
    case "pageTitle":
      if (text.includes("参考版式") || text === "目录") score += 220;
      if (placeholderType === "title" || name.includes("title")) score += 100;
      if (font >= 24) score += 80;
      if (y < 60) score += 40;
      break;
    case "intro":
      if (text.includes("内容简介")) score += 220;
      if (font >= 16) score += 40;
      if (y > 60 && y < 130) score += 40;
      if (width > 700) score += 10;
      break;
    case "regularTitle":
      if (text.includes("常规标题")) score += 220;
      if (font >= 13 && font <= 15) score += 30;
      if (maxChars > 0 && maxChars <= 20) score += 20;
      if (/rectangle|chevron|pentagon|callout/.test(name)) score += 20;
      break;
    case "smallTitle":
      if (text.includes("小标题") || text.includes("这是个示例") || text.includes("概述")) score += 180;
      if (font >= 13 && font <= 15) score += 30;
      if (height > 0 && height <= 45) score += 20;
      if (maxChars > 0 && maxChars <= 30) score += 20;
      break;
    case "detailSummary":
      if (text.includes("总体说明性文字")) score += 260;
      if (font >= 12 && font <= 16) score += 20;
      break;
    case "detailBody":
      if (
        text.includes("说明性文字") ||
        text.includes("5g的性能目标") ||
        text.includes("峰值速率") ||
        text.includes("虚拟现实")
      ) {
        score += 180;
      }
      if (font <= 12.5) score += 40;
      if (height >= 35) score += 20;
      if (maxChars >= 25) score += 20;
      break;
    case "tableHeader":
      if (text.includes("添加标题")) score += 220;
      if (font <= 12.5) score += 20;
      if (y < 220) score += 20;
      break;
    case "tableRow":
      if (/pentagon/.test(name) && text.includes("常规标题")) score += 260;
      break;
    case "tableCell":
      if (text.includes("峰值速率") || text.includes("5g的性能目标")) score += 180;
      if (font <= 12.5) score += 20;
      if (width > 140 && width < 190) score += 20;
      break;
    default:
      break;
  }

  return score;
}

function resolveScannedPlaceholder(
  slide: ScannedSlide | undefined,
  item: PlaceholderRefConfig,
  usedKeys: Set<string>
): ScannedPlaceholder | undefined {
  if (!slide) {
    return undefined;
  }

  const occurrence = item.occurrence ?? 1;
  const exact = slide.placeholders.find(
    (placeholder) =>
      placeholder.name === item.name &&
      placeholder.occurrence === occurrence &&
      (!item.kind || placeholder.kind === item.kind) &&
      !usedKeys.has(scannedPlaceholderKey(placeholder))
  );

  if (exact) {
    return exact;
  }

  return slide.placeholders
    .filter((placeholder) => !usedKeys.has(scannedPlaceholderKey(placeholder)))
    .map((placeholder) => ({
      placeholder,
      score: scoreScannedPlaceholder(item, placeholder)
    }))
    .filter((candidate) => Number.isFinite(candidate.score) && candidate.score > 0)
    .sort((left, right) => {
      if (right.score !== left.score) {
        return right.score - left.score;
      }

      const leftY = left.placeholder.yPt ?? Number.MAX_SAFE_INTEGER;
      const rightY = right.placeholder.yPt ?? Number.MAX_SAFE_INTEGER;
      if (leftY !== rightY) {
        return leftY - rightY;
      }

      return left.placeholder.occurrence - right.placeholder.occurrence;
    })[0]?.placeholder;
}

function toPlaceholderSchema(
  slideNumber: number,
  item: PlaceholderRefConfig,
  scanned: ScannedPlaceholder | undefined
): TemplatePlaceholderSchema {
  const scannedMaxChars =
    scanned?.kind === "TEXT" || !scanned?.kind ? scanned?.maxChars : undefined;
  const maxChars =
    scannedMaxChars && item.fallbackMaxChars > 0
      ? Math.min(scannedMaxChars, item.fallbackMaxChars)
      : scannedMaxChars ?? item.fallbackMaxChars;

  return {
    id: placeholderId(slideNumber, item.name, item.occurrence ?? 1),
    name: item.name,
    occurrence: item.occurrence ?? 1,
    role: item.role,
    kind: scanned?.kind ?? item.kind ?? "TEXT",
    maxChars: scanned?.kind === "TEXT" || !scanned?.kind ? maxChars : undefined,
    maxLines: scanned?.maxLines,
    widthPt: scanned?.widthPt,
    heightPt: scanned?.heightPt,
    fontSizePt: scanned?.fontSizePt,
    maxItems: item.maxItems,
    description: item.description
  };
}

function collectNonTextPlaceholders(
  slide: ScannedSlide | undefined,
  usedKeys: Set<string>,
  sourceSlide: number
): TemplatePlaceholderSchema[] {
  if (!slide) {
    return [];
  }

  return slide.placeholders
    .filter((item) => item.kind !== "TEXT")
    .filter((item) => !usedKeys.has(`${item.name}#${item.occurrence}`))
    .map((item) => ({
      id: placeholderId(sourceSlide, item.name, item.occurrence),
      name: item.name,
      occurrence: item.occurrence,
      role:
        item.kind === "PICTURE" ? "picture" : item.kind === "TABLE" ? "table" : "chart",
      kind: item.kind,
      widthPt: item.widthPt,
      heightPt: item.heightPt,
      fontSizePt: item.fontSizePt,
      description: item.sampleText || undefined
    }));
}

function buildLayoutSchema(config: LayoutConfig, slide: ScannedSlide | undefined): TemplateLayoutSchema {
  const usedKeys = new Set<string>();
  const placeholders = config.refs.map((item) => {
    const resolved = resolveScannedPlaceholder(slide, item, usedKeys);
    if (resolved) {
      usedKeys.add(scannedPlaceholderKey(resolved));
    } else {
      usedKeys.add(`${item.name}#${item.occurrence ?? 1}`);
    }
    return toPlaceholderSchema(config.sourceSlide, item, resolved);
  });

  const inferredDetailItems = placeholders
    .filter((item) => item.role === "detailBody" || item.role === "detailSummary")
    .reduce((sum, item) => sum + Math.max(1, item.maxItems ?? 1), 0);

  placeholders.push(...collectNonTextPlaceholders(slide, usedKeys, config.sourceSlide));

  return {
    id: config.id,
    name: config.name,
    layoutType: config.layoutType,
    sourceSlide: config.sourceSlide,
    summary: config.summary,
    capacities: {
      detailItems: inferredDetailItems > 0 ? inferredDetailItems : config.detailItems,
      pictureSlots:
        slide?.capacities.pictureSlots ?? placeholders.filter((item) => item.kind === "PICTURE").length,
      tableSlots:
        slide?.capacities.tableSlots ?? placeholders.filter((item) => item.kind === "TABLE").length,
      chartSlots:
        slide?.capacities.chartSlots ?? placeholders.filter((item) => item.kind === "CHART").length
    },
    placeholders
  };
}

function buildFallbackSchema(colors: Awaited<ReturnType<typeof loadTemplateTokens>>["colors"], metrics: Awaited<ReturnType<typeof loadTemplateTokens>>["metrics"]): TemplateSchemaLibrary {
  return {
    cover: {
      sourceSlide: 1,
      placeholders: coverRefs.map((item) =>
        textPlaceholder(
          placeholderId(1, item.name, item.occurrence ?? 1),
          item.name,
          item.role,
          item.fallbackMaxChars,
          item.occurrence ?? 1
        )
      )
    },
    directory: {
      sourceSlide: 22,
      maxItems: 3,
      placeholders: directoryRefs.map((item) =>
        textPlaceholder(
          placeholderId(22, item.name, item.occurrence ?? 1),
          item.name,
          item.role,
          item.fallbackMaxChars,
          item.occurrence ?? 1
        )
      )
    },
    detailLayouts: layoutConfigs.map((config) => buildLayoutSchema(config, undefined)),
    colors,
    metrics,
    scanner: {
      engine: "static-fallback",
      templatePath: getTemplatePath(),
      slideCount: layoutConfigs.length + 2
    }
  };
}

export function getDetailLayoutSchema(
  schema: TemplateSchemaLibrary,
  layoutType: SlideLayoutType
): TemplateLayoutSchema | undefined {
  return schema.detailLayouts.find((item) => item.layoutType === layoutType);
}

export async function loadTemplateSchema(): Promise<TemplateSchemaLibrary> {
  const tokens = await loadTemplateTokens();
  const scan = await loadPythonTemplateScan();

  if (!scan) {
    return buildFallbackSchema(tokens.colors, tokens.metrics);
  }

  const slideMap = new Map(scan.slides.map((slide) => [slide.sourceSlide, slide]));
  const directorySourceSlide = scan.slides[scan.slides.length - 1]?.sourceSlide ?? scan.slideCount;

  return {
    cover: {
      sourceSlide: 1,
      placeholders: (() => {
        const usedKeys = new Set<string>();
        return coverRefs.map((item) => {
          const resolved = resolveScannedPlaceholder(slideMap.get(1), item, usedKeys);
          if (resolved) {
            usedKeys.add(scannedPlaceholderKey(resolved));
          }
          return toPlaceholderSchema(1, item, resolved);
        });
      })()
    },
    directory: {
      sourceSlide: directorySourceSlide,
      maxItems: 3,
      placeholders: (() => {
        const usedKeys = new Set<string>();
        return directoryRefs.map((item) => {
          const resolved = resolveScannedPlaceholder(slideMap.get(directorySourceSlide), item, usedKeys);
          if (resolved) {
            usedKeys.add(scannedPlaceholderKey(resolved));
          }
          return toPlaceholderSchema(directorySourceSlide, item, resolved);
        });
      })()
    },
    detailLayouts: layoutConfigs.map((config) => buildLayoutSchema(config, slideMap.get(config.sourceSlide))),
    colors: tokens.colors,
    metrics: tokens.metrics,
    scanner: {
      engine: "python-ooxml-scan",
      templatePath: scan.templatePath,
      slideCount: scan.slideCount
    }
  };
}
