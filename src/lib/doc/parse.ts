import { access } from "node:fs/promises";
import { constants } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import mammoth from "mammoth";
import { load as loadHtml } from "cheerio";
import WordExtractor from "word-extractor";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import type { ParsedDocument, ParsedSection } from "@/lib/types";

const execFileAsync = promisify(execFile);

function makeId(prefix: string, index: number): string {
  return `${prefix}-${index + 1}`;
}

function normalizeText(text: string): string {
  return text.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function buildFallbackSectionTitle(index: number) {
  return `章节 ${index + 1}`;
}

function detectStructuredHeading(text: string) {
  const normalized = normalizeText(text);
  if (!normalized || normalized.length > 48) {
    return null;
  }

  const patterns: Array<{ regex: RegExp; level: number }> = [
    { regex: /^\d+\.\d+\.\d+\.?\s*(.+)$/u, level: 3 },
    { regex: /^\d+\.\d+\s*(.+)$/u, level: 2 },
    { regex: /^(第[一二三四五六七八九十百]+[章节部分篇]|[一二三四五六七八九十]+[、.．])\s*(.+)$/u, level: 1 },
    { regex: /^\d+\s*[、.．]\s*(.+)$/u, level: 1 }
  ];

  for (const pattern of patterns) {
    const match = normalized.match(pattern.regex);
    if (!match) {
      continue;
    }

    const title = normalizeText(match[2] ?? match[1] ?? "");
    if (!title) {
      return null;
    }

    return {
      title,
      level: pattern.level
    };
  }

  return null;
}

function guessTitleFromSections(sections: ParsedSection[], rawText: string): string {
  if (sections.length > 0 && sections[0].title) {
    return sections[0].title;
  }

  return rawText
    .split(/\n+/)
    .map((line) => normalizeText(line))
    .find((line) => line.length >= 8) ?? "项目汇报";
}

function buildSubtitle(rawText: string): string {
  const sentence = rawText
    .split(/[。！？\n]/)
    .map((line) => normalizeText(line))
    .find((line) => line.length >= 10);

  return sentence ?? "基于上传材料自动生成的汇报摘要";
}

function inferLayoutType(section: ParsedSection) {
  if (section.tables.length > 0) {
    return "table" as const;
  }

  if (section.imageIds.length > 0) {
    return "image" as const;
  }

  const paragraphLength = section.paragraphs.join("").length;

  if (paragraphLength > 360) {
    return "vertical" as const;
  }

  if (paragraphLength > 180) {
    return "two-column" as const;
  }

  return "overview" as const;
}

function sectionFromParagraphs(
  title: string,
  paragraphs: string[],
  level: number,
  index: number
): ParsedSection {
  return {
    id: makeId("section", index),
    title,
    level,
    paragraphs: paragraphs.filter(Boolean),
    tables: [],
    imageIds: []
  };
}

function parseMarkdownOrText(fileName: string, rawText: string, fileType: "md" | "txt" | "text"): ParsedDocument {
  const sections: ParsedSection[] = [];
  const lines = rawText.split(/\r?\n/);
  let currentTitle = "概览";
  let currentLevel = 1;
  let currentParagraphs: string[] = [];
  let currentTables: string[][] = [];
  let currentHasExplicitHeading = false;
  let sectionIndex = 0;

  const flush = (forceHeadingOnly = false) => {
    const normalizedParagraphs = currentParagraphs.map((text) => normalizeText(text)).filter(Boolean);
    const hasContent = normalizedParagraphs.length > 0 || currentTables.length > 0;
    if (!hasContent && !forceHeadingOnly && sections.length > 0) {
      return;
    }

    sections.push({
      id: makeId("section", sectionIndex),
      title: currentTitle,
      level: currentLevel,
      paragraphs: normalizedParagraphs,
      tables: currentTables,
      imageIds: [],
      isHeadingOnly: !hasContent && currentHasExplicitHeading
    });
    sectionIndex += 1;
    currentTables = [];
    currentParagraphs = [];
    currentHasExplicitHeading = false;
  };

  const pushParagraph = (value: string) => {
    const normalized = normalizeText(value);
    if (normalized) {
      currentParagraphs.push(normalized);
    }
  };

  lines.forEach((rawLine) => {
    const line = rawLine.replace(/\t/g, "    ");
    const trimmed = normalizeText(line);

    if (!trimmed) {
      return;
    }

    const markdownHeading = line.match(/^(#{1,6})\s+(.+)$/);
    if (markdownHeading) {
      if (currentParagraphs.length > 0 || currentTables.length > 0) {
        flush();
      } else if (currentHasExplicitHeading) {
        flush(true);
      }
      currentTitle = normalizeText(markdownHeading[2]) || buildFallbackSectionTitle(sectionIndex);
      currentLevel = markdownHeading[1].length;
      currentHasExplicitHeading = true;
      return;
    }

    const textHeading = detectStructuredHeading(trimmed);
    if (textHeading) {
      if (currentParagraphs.length > 0 || currentTables.length > 0) {
        flush();
      } else if (currentHasExplicitHeading) {
        flush(true);
      }
      currentTitle = textHeading.title || buildFallbackSectionTitle(sectionIndex);
      currentLevel = textHeading.level;
      currentHasExplicitHeading = true;
      return;
    }

    if (/^\|.+\|$/.test(trimmed)) {
      const cells = trimmed
        .split("|")
        .map((item) => normalizeText(item))
        .filter(Boolean);
      if (cells.length > 0) {
        currentTables.push(cells);
      }
      return;
    }

    if (/^[-*]\s+/.test(trimmed)) {
      pushParagraph(trimmed.replace(/^[-*]\s+/, ""));
      return;
    }

    pushParagraph(trimmed);
  });

  if (currentParagraphs.length > 0 || currentTables.length > 0 || currentHasExplicitHeading || sections.length === 0) {
    flush(currentHasExplicitHeading);
  }

  return {
    fileName,
    fileType,
    rawText,
    titleGuess: guessTitleFromSections(sections, rawText),
    subtitleGuess: buildSubtitle(rawText),
    sections: sections.map((section, index) => ({
      ...section,
      title: section.title || buildFallbackSectionTitle(index)
    })),
    images: []
  };
}

function parseDocxHtml(
  fileName: string,
  html: string,
  rawText: string,
  imageMap: Map<string, string>
): ParsedDocument {
  const $ = loadHtml(html);
  const sections: ParsedSection[] = [];
  const images: ParsedDocument["images"] = [];
  let currentTitle = "概览";
  let currentLevel = 1;
  let currentParagraphs: string[] = [];
  let currentTables: string[][] = [];
  let currentImageIds: string[] = [];
  let currentHasExplicitHeading = false;
  let sectionIndex = 0;
  let imageIndex = 0;

  const flush = (forceHeadingOnly = false) => {
    const normalizedParagraphs = currentParagraphs.map((text) => normalizeText(text)).filter(Boolean);
    const hasContent =
      normalizedParagraphs.length > 0 ||
      currentTables.length > 0 ||
      currentImageIds.length > 0;
    if (
      !hasContent &&
      !forceHeadingOnly &&
      sections.length > 0
    ) {
      return;
    }

    sections.push({
      id: makeId("section", sectionIndex),
      title: currentTitle,
      level: currentLevel,
      paragraphs: normalizedParagraphs,
      tables: currentTables,
      imageIds: currentImageIds,
      isHeadingOnly: !hasContent && currentHasExplicitHeading
    });
    sectionIndex += 1;
    currentParagraphs = [];
    currentTables = [];
    currentImageIds = [];
    currentHasExplicitHeading = false;
  };

  const beginHeading = (title: string, level: number) => {
    if (currentParagraphs.length > 0 || currentTables.length > 0 || currentImageIds.length > 0) {
      flush();
    } else if (currentHasExplicitHeading) {
      flush(true);
    }

    currentTitle = normalizeText(title) || `章节 ${sectionIndex + 1}`;
    currentLevel = level;
    currentParagraphs = [];
    currentTables = [];
    currentImageIds = [];
    currentHasExplicitHeading = true;
  };

  $("body")
    .children()
    .each((_, element) => {
      const tagName = element.tagName?.toLowerCase();

      if (!tagName) {
        return;
      }

      if (/^h[1-6]$/.test(tagName)) {
        beginHeading($(element).text(), Number(tagName.slice(1)));
        return;
      }

      if (tagName === "p" || tagName === "li") {
        const text = normalizeText($(element).text());
        const inlineHeading = detectStructuredHeading(text);
        if (inlineHeading) {
          beginHeading(inlineHeading.title, inlineHeading.level);
          return;
        }

        if (text) {
          currentParagraphs.push(text);
        }

        $(element)
          .find("img")
          .each((__, img) => {
            const src = $(img).attr("src");
            if (!src) {
              return;
            }

            const dataUri = imageMap.get(src) ?? src;
            const mimeType = dataUri.slice(5, dataUri.indexOf(";")) || "image/png";
            const imageId = makeId("img", imageIndex);
            imageIndex += 1;
            images.push({
              id: imageId,
              name: `${imageId}.${mimeType.split("/")[1] ?? "png"}`,
              mimeType,
              dataUri
            });
            currentImageIds.push(imageId);
          });
        return;
      }

      if (tagName === "table") {
        const rows = $(element)
          .find("tr")
          .toArray()
          .map((row) =>
            $(row)
              .find("th,td")
              .toArray()
              .map((cell) => normalizeText($(cell).text()))
          )
          .filter((row) => row.some(Boolean));

        if (rows.length > 0) {
          if (currentParagraphs.length === 0) {
            currentParagraphs.push("表格信息");
          }
          currentTables.push(...rows);
        }
      }
    });

  if (currentParagraphs.length > 0 || currentTables.length > 0 || currentImageIds.length > 0 || currentHasExplicitHeading || sections.length === 0) {
    flush(currentHasExplicitHeading);
  }

  return {
    fileName,
    fileType: "docx",
    rawText,
    titleGuess: guessTitleFromSections(sections, rawText),
    subtitleGuess: buildSubtitle(rawText),
    sections: sections.map((section, index) => ({
      ...section,
      title: section.title || buildFallbackSectionTitle(index)
    })),
    images
  };
}

async function convertLegacyDocToDocx(buffer: Buffer): Promise<Buffer | null> {
  const tempDir = await mkdtemp(join(tmpdir(), "ppt-doc-"));
  const inputPath = join(tempDir, "source.doc");
  const outputPath = join(tempDir, "source.docx");

  try {
    await access("/Applications/LibreOffice.app/Contents/MacOS/soffice", constants.X_OK);
  } catch {
    return null;
  }

  try {
    await writeFile(inputPath, buffer);
    await execFileAsync("/Applications/LibreOffice.app/Contents/MacOS/soffice", [
      "--headless",
      "--convert-to",
      "docx",
      "--outdir",
      tempDir,
      inputPath
    ]);
    return await readFile(outputPath);
  } catch {
    return null;
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
}

async function parseDocBuffer(fileName: string, buffer: Buffer): Promise<ParsedDocument> {
  const extractor = new WordExtractor();
  const extracted = await extractor.extract(buffer);
  const rawText = normalizeText(extracted.getBody()) || "未能提取有效正文，请检查源文件内容。";
  const paragraphs = rawText
    .split(/\n+/)
    .map((line) => normalizeText(line))
    .filter(Boolean);
  const sections = (paragraphs.length > 0 ? paragraphs : ["项目概述"])
    .slice(0, 8)
    .map((paragraph, index) =>
      sectionFromParagraphs(
        index === 0 ? paragraph || "概览" : buildFallbackSectionTitle(index),
        [paragraph],
        1,
        index
      )
    );

  return {
    fileName,
    fileType: "doc",
    rawText,
    titleGuess: guessTitleFromSections(sections, rawText),
    subtitleGuess: buildSubtitle(rawText),
    sections,
    images: []
  };
}

export async function parseWordDocument(file: File): Promise<ParsedDocument> {
  const fileName = file.name;
  const extension = fileName.split(".").pop()?.toLowerCase();
  const buffer = Buffer.from(await file.arrayBuffer());

  if (extension === "docx") {
    const imageMap = new Map<string, string>();
    const htmlResult = await mammoth.convertToHtml(
      { buffer },
      {
        convertImage: mammoth.images.imgElement(async (element) => {
          const contentType = element.contentType || "image/png";
          const base64 = await element.readAsBase64String();
          const dataUri = `data:${contentType};base64,${base64}`;
          imageMap.set(dataUri, dataUri);
          return { src: dataUri };
        })
      }
    );
    const rawTextResult = await mammoth.extractRawText({ buffer });

    return parseDocxHtml(fileName, htmlResult.value, rawTextResult.value, imageMap);
  }

  if (extension === "doc") {
    const converted = await convertLegacyDocToDocx(buffer);
    if (converted) {
      const convertedFile = new File([new Uint8Array(converted)], `${fileName}.docx`, {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });
      return parseWordDocument(convertedFile);
    }

    return parseDocBuffer(fileName, buffer);
  }

  if (extension === "md" || extension === "markdown") {
    const rawText = Buffer.from(buffer).toString("utf-8");
    return parseMarkdownOrText(fileName, rawText, "md");
  }

  if (extension === "txt") {
    const rawText = Buffer.from(buffer).toString("utf-8");
    return parseMarkdownOrText(fileName, rawText, "txt");
  }

  throw new Error("仅支持 .doc、.docx、.md 或 .txt 文件。");
}

export async function parseSourceDocument(
  source: File | { name?: string; text: string; type?: "md" | "txt" | "text" }
): Promise<ParsedDocument> {
  if (source instanceof File) {
    return parseWordDocument(source);
  }

  const fileType = source.type ?? "text";
  const fileName =
    source.name?.trim() ||
    (fileType === "md" ? "聊天输入素材.md" : fileType === "txt" ? "聊天输入素材.txt" : "聊天输入素材.txt");
  return parseMarkdownOrText(fileName, source.text, fileType);
}

export function suggestLayoutForSection(section: ParsedSection) {
  return inferLayoutType(section);
}
