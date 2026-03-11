import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type { TemplateColorTokens, TemplateMetrics } from "@/lib/types";

const EMU_PER_INCH = 914400;
const DEFAULT_TEMPLATE_PATH =
  process.env.PPT_TEMPLATE_PATH ?? "/Users/springwater/Desktop/中国移动PPT模板.pptx";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: ""
});

export type TemplateTokens = {
  colors: TemplateColorTokens;
  metrics: TemplateMetrics;
};

let cachedTokens: Promise<TemplateTokens> | null = null;

function deepFindFirst(node: unknown, key: string): unknown {
  if (!node || typeof node !== "object") {
    return undefined;
  }

  if (Array.isArray(node)) {
    for (const item of node) {
      const match = deepFindFirst(item, key);
      if (match !== undefined) {
        return match;
      }
    }

    return undefined;
  }

  if (key in (node as Record<string, unknown>)) {
    return (node as Record<string, unknown>)[key];
  }

  for (const value of Object.values(node as Record<string, unknown>)) {
    const match = deepFindFirst(value, key);
    if (match !== undefined) {
      return match;
    }
  }

  return undefined;
}

function extractThemeColor(
  colorScheme: Record<string, unknown>,
  key: string,
  fallback: string
): string {
  const entry = colorScheme[key];

  if (entry && typeof entry === "object" && "a:srgbClr" in entry) {
    const clr = (entry as Record<string, unknown>)["a:srgbClr"];

    if (clr && typeof clr === "object" && "val" in clr) {
      return `#${String((clr as Record<string, unknown>).val)}`;
    }
  }

  return fallback;
}

function collectChartColors(slideRoot: Record<string, unknown>): string[] {
  const colors = new Set<string>();

  const walk = (node: unknown) => {
    if (!node || typeof node !== "object") {
      return;
    }

    if (Array.isArray(node)) {
      node.forEach(walk);
      return;
    }

    const objectNode = node as Record<string, unknown>;

    if ("a:srgbClr" in objectNode) {
      const value = objectNode["a:srgbClr"];
      if (value && typeof value === "object" && "val" in value) {
        colors.add(`#${String((value as Record<string, unknown>).val)}`);
      }
    }

    Object.values(objectNode).forEach(walk);
  };

  walk(slideRoot);

  return Array.from(colors);
}

async function readTemplateTokensFromFile(path: string): Promise<TemplateTokens> {
  const buffer = await readFile(path);
  const zip = await JSZip.loadAsync(buffer);
  const themeXml = await zip.file("ppt/theme/theme1.xml")?.async("string");
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
  const colorSlideXml = await zip.file("ppt/slides/slide2.xml")?.async("string");

  if (!themeXml || !presentationXml || !colorSlideXml) {
    throw new Error("PPT 模板文件缺少必要的主题或页面数据。");
  }

  const theme = parser.parse(themeXml) as Record<string, unknown>;
  const presentation = parser.parse(presentationXml) as Record<string, unknown>;
  const colorSlide = parser.parse(colorSlideXml) as Record<string, unknown>;

  const colorScheme = deepFindFirst(theme, "a:clrScheme") as Record<string, unknown>;
  const slideSize = deepFindFirst(presentation, "p:sldSz") as Record<string, unknown>;
  const chartColors = collectChartColors(colorSlide).filter(
    (color) =>
      !["#0070BF", "#8CC121", "#C6C7C6", "#EEEFEE", "#C00000", "#FFFE01"].includes(color)
  );

  return {
    colors: {
      primary: extractThemeColor(colorScheme, "a:accent1", "#0070BF"),
      secondary1: extractThemeColor(colorScheme, "a:accent2", "#8CC121"),
      secondary2: extractThemeColor(colorScheme, "a:accent3", "#C6C7C6"),
      secondary3: extractThemeColor(colorScheme, "a:accent4", "#EEEFEE"),
      accent: extractThemeColor(colorScheme, "a:accent5", "#C00000"),
      highlight: extractThemeColor(colorScheme, "a:accent6", "#FFFE01"),
      darkText: extractThemeColor(colorScheme, "a:dk1", "#000000"),
      lightText: extractThemeColor(colorScheme, "a:lt1", "#FFFFFF"),
      chart1: chartColors[0] ?? "#009DD9",
      chart2: chartColors[1] ?? "#F6B961",
      chart3: chartColors[2] ?? "#61A3B5"
    },
    metrics: {
      widthInches: Number(slideSize.cx) / EMU_PER_INCH,
      heightInches: Number(slideSize.cy) / EMU_PER_INCH
    }
  };
}

export function getTemplatePath(): string {
  return DEFAULT_TEMPLATE_PATH;
}

export async function loadTemplateTokens(): Promise<TemplateTokens> {
  if (!cachedTokens) {
    cachedTokens = readTemplateTokensFromFile(getTemplatePath());
  }

  return cachedTokens;
}
