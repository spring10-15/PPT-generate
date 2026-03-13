import { validateAssemblyDocument } from "@/lib/assembly";
import { generateOutlineWithAI, mergeOutlineWithDocument } from "@/lib/ai/siliconflow";
import { parseSourceDocument, parseWordDocument } from "@/lib/doc/parse";
import { loadTemplateSchema } from "@/lib/template-schema";
import type {
  AssemblyInstructionDocument,
  AssemblySlideInstruction,
  GenerationPipeline,
  GeneratorInput,
  OutlineDocument,
  OutlineRelation,
  ParsedDocument,
  StructuredOutlineDocument,
  StructuredOutlineHierarchyNode,
  StructuredOutlinePage,
  TemplateLayoutSchema,
  TemplatePlaceholderSchema,
  TemplateSchemaLibrary
} from "@/lib/types";

const PROGRESS_CUE_RE = /(步骤|阶段|推进|路径|计划|里程碑|排期|节奏|实施|动作|举措|落地)/;
const HIERARCHY_CUE_RE = /(体系|架构|分层|能力|层级|模块|框架|分类|结构|版图|矩阵)/;

function normalizeText(text: string | undefined) {
  return (text ?? "").replace(/\s+/g, " ").trim();
}

function toComparableText(text: string | undefined) {
  return normalizeText(text).replace(/[：:；;。！？,.，、】【（）()[\]\s-]/g, "");
}

function splitPoints(text: string) {
  return normalizeText(text)
    .split(/[。！？；;\n]/)
    .flatMap((part) => part.split(/[，,、]/))
    .map((part) => normalizeText(part))
    .filter(Boolean);
}

function dedupePoints(points: string[], blocked: string[] = []) {
  const blockedSet = new Set(
    blocked
      .flatMap((item) => [item, ...splitPoints(item)])
      .map((item) => toComparableText(item))
      .filter(Boolean)
  );
  const seen = new Set<string>();

  return points.filter((point) => {
    const comparable = toComparableText(point);
    if (!comparable || blockedSet.has(comparable) || seen.has(comparable)) {
      return false;
    }

    seen.add(comparable);
    return true;
  });
}

function fitWithinLimit(text: string, maxChars: number) {
  const normalized = normalizeText(text);
  if (!normalized) {
    return "";
  }

  if (normalized.length <= maxChars) {
    return normalized;
  }

  const units = splitPoints(normalized);
  if (units.length > 1) {
    let result = "";

    units.forEach((unit) => {
      const next = result ? `${result}；${unit}` : unit;
      if (next.length <= maxChars) {
        result = next;
      }
    });

    if (result) {
      return result;
    }
  }

  return normalized.slice(0, maxChars).trim();
}

function chunkItems<T>(items: T[], size: number) {
  const chunkSize = Math.max(1, size);
  const chunks: T[][] = [];

  for (let index = 0; index < items.length; index += chunkSize) {
    chunks.push(items.slice(index, index + chunkSize));
  }

  return chunks.length > 0 ? chunks : [[]];
}

function inferRelation(title: string, summary: string, pointCount: number): OutlineRelation {
  const source = `${title} ${summary}`;
  if (PROGRESS_CUE_RE.test(source)) {
    return "sequential";
  }

  if (HIERARCHY_CUE_RE.test(source)) {
    return "supporting";
  }

  if (pointCount > 1) {
    return "parallel";
  }

  return "supporting";
}

function buildHierarchy(document: StructuredOutlineDocument): StructuredOutlineHierarchyNode[] {
  const directoryNodes = document.directory.map((item) => {
    const pages = document.pages.filter((page) => page.directoryId === item.id);

    return {
      id: `node-${item.id}`,
      level: "H2" as const,
      title: item.title,
      summary: item.description,
      relation: "sequential" as const,
      pointCount: pages.length,
      children: pages.map((page) => ({
        id: `node-${page.id}`,
        level: "H3" as const,
        title: page.title,
        summary: page.summary,
        relation: page.relation,
        pointCount: page.keyPoints.length,
        children: page.keyPoints.map((point, pointIndex) => ({
          id: `node-${page.id}-point-${pointIndex + 1}`,
          level: "H4" as const,
          title: `要点 ${pointIndex + 1}`,
          summary: point,
          relation: "parallel" as const,
          pointCount: 0,
          children: []
        }))
      }))
    };
  });

  return [
    {
      id: "node-cover",
      level: "H1",
      title: document.cover.title,
      summary: document.cover.subtitle,
      relation: "root",
      pointCount: directoryNodes.length,
      children: directoryNodes
    }
  ];
}

function syncStructuredOutlineWithAssembly(
  structuredOutline: StructuredOutlineDocument,
  assembly: AssemblyInstructionDocument
): StructuredOutlineDocument {
  const activeDirectoryIds = new Set(assembly.directory.map((item) => item.id));
  const directoryMap = new Map(
    assembly.directory.map((item) => [
      item.id,
      {
        id: item.id,
        order: item.order,
        title: item.fields.title.value,
        description: item.fields.description.value
      }
    ])
  );

  const nextDocument: StructuredOutlineDocument = {
    ...structuredOutline,
    directory: assembly.directory.map((item) => ({
      id: item.id,
      order: item.order,
      title: item.fields.title.value,
      description: item.fields.description.value
    })),
    pages: structuredOutline.pages
      .filter((page) => activeDirectoryIds.has(page.directoryId))
      .map((page) => ({
        ...page,
        directoryTitle: directoryMap.get(page.directoryId)?.title ?? page.directoryTitle
      }))
  };

  nextDocument.hierarchy = buildHierarchy(nextDocument);
  return nextDocument;
}

function buildStructuredOutlineDocument(
  parsed: ParsedDocument,
  outline: OutlineDocument,
  input: GeneratorInput
): StructuredOutlineDocument {
  const pages: StructuredOutlinePage[] = outline.slides.map((slide, index) => {
    const keyPoints = dedupePoints(
      [
        ...splitPoints(slide.content.description || ""),
        ...splitPoints(slide.content.intro || ""),
        ...splitPoints(slide.summary || "")
      ],
      [slide.title, slide.content.regularTitle, slide.content.intro]
    );
    const summary = normalizeText(slide.summary || slide.content.description || slide.content.intro);

    return {
      id: `page-${index + 1}`,
      directoryId: slide.directoryId,
      directoryTitle: slide.directoryTitle,
      title: slide.title,
      intro: slide.content.intro,
      regularTitle: slide.content.regularTitle,
      description: slide.content.description,
      summary,
      keyPoints,
      relation: inferRelation(slide.title, summary, keyPoints.length),
      suggestedLayoutType: slide.type,
      imageIds: slide.imageIds,
      table: slide.table,
      hasImage: slide.imageIds.length > 0,
      hasTable: Boolean(slide.table?.length)
    };
  });

  const structured: StructuredOutlineDocument = {
    cover: {
      title: outline.cover.title,
      subtitle: outline.cover.subtitle,
      userName: outline.cover.userName,
      targetAudience: input.targetAudience,
      estimatedMinutes: input.estimatedMinutes,
      dateLabel: outline.cover.dateLabel
    },
    directory: outline.directory.map((item, index) => ({
      id: item.id,
      order: index + 1,
      title: item.title,
      description: item.description
    })),
    pages,
    hierarchy: [],
    extractedImages: parsed.images,
    source: {
      fileName: parsed.fileName,
      totalPages: input.totalPages,
      targetAudience: input.targetAudience,
      estimatedMinutes: input.estimatedMinutes
    }
  };

  structured.hierarchy = buildHierarchy(structured);
  return structured;
}

function scoreLayout(page: StructuredOutlinePage, layout: TemplateLayoutSchema) {
  let score = 0;
  const pointCount = Math.max(1, page.keyPoints.length);
  const capacityDelta = Math.abs(layout.capacities.detailItems - pointCount);

  score -= capacityDelta * 10;
  score -= Math.max(0, layout.capacities.detailItems - pointCount - 2) * 4;

  if (layout.layoutType === page.suggestedLayoutType) {
    score += 20;
  }

  if (page.hasTable) {
    score += layout.layoutType === "table" ? 200 : -400;
  }

  if (page.hasImage) {
    score += ["image", "image-left", "image-right", "two-column"].includes(layout.layoutType) ? 80 : -30;
  } else if (layout.capacities.pictureSlots > 0 || ["image", "image-left", "image-right"].includes(layout.layoutType)) {
    score -= layout.layoutType === "two-column" ? 80 : 120;
  }

  if (page.relation === "sequential") {
    score += layout.layoutType === "progress" ? 90 : 0;
  }

  if (page.relation === "supporting") {
    score += layout.layoutType === "hierarchy" ? 60 : 0;
  }

  if (pointCount <= 2) {
    score += ["overview", "vertical", "two-column"].includes(layout.layoutType) ? 30 : -15;
  }

  if (pointCount >= 7) {
    score += ["four-column", "progress", "split-grid"].includes(layout.layoutType) ? 30 : 0;
  }

  return score;
}

function chooseBestLayout(page: StructuredOutlinePage, schema: TemplateSchemaLibrary) {
  const ranked = schema.detailLayouts
    .map((layout) => ({
      layout,
      score: scoreLayout(page, layout)
    }))
    .sort((left, right) => right.score - left.score);

  return ranked[0]?.layout ?? schema.detailLayouts[0];
}

function getRolePlaceholders(
  placeholders: TemplatePlaceholderSchema[],
  roles: string | string[]
): TemplatePlaceholderSchema[] {
  const roleSet = new Set(Array.isArray(roles) ? roles : [roles]);
  return placeholders.filter((placeholder) => placeholder.kind === "TEXT" && roleSet.has(placeholder.role));
}

function getRoleMaxChars(
  placeholders: TemplatePlaceholderSchema[],
  role: string,
  fallback: number
) {
  const match = getRolePlaceholders(placeholders, role)[0];
  return match?.maxChars ?? fallback;
}

function getRolePlaceholderName(
  placeholders: TemplatePlaceholderSchema[],
  role: string,
  fallback: string
) {
  const match = getRolePlaceholders(placeholders, role)[0];
  return match?.name ?? fallback;
}

function getDetailSlotLimits(layout: TemplateLayoutSchema) {
  const detailPlaceholders = getRolePlaceholders(layout.placeholders, ["detailBody", "detailSummary"]);
  const slotLimits: number[] = [];

  detailPlaceholders.forEach((placeholder) => {
    const itemCount = Math.max(1, placeholder.maxItems ?? 1);
    const totalChars = placeholder.maxChars ?? itemCount * 60;
    const perItemChars = Math.max(1, Math.floor(totalChars / itemCount));

    for (let index = 0; index < itemCount; index += 1) {
      slotLimits.push(perItemChars);
    }
  });

  if (slotLimits.length > 0) {
    return slotLimits;
  }

  return Array.from({ length: Math.max(1, layout.capacities.detailItems) }, () => 60);
}

function buildAssemblyInstructions(
  structuredOutline: StructuredOutlineDocument,
  templateSchema: TemplateSchemaLibrary
): AssemblyInstructionDocument {
  const slides: AssemblySlideInstruction[] = [];
  const directoryTitleMaxChars = getRoleMaxChars(
    templateSchema.directory.placeholders,
    "directoryTitle",
    5
  );
  const directoryDescriptionMaxChars = getRoleMaxChars(
    templateSchema.directory.placeholders,
    "directoryDescription",
    10
  );
  const directoryTitlePlaceholderName = getRolePlaceholderName(
    templateSchema.directory.placeholders,
    "directoryTitle",
    "TextBox 11"
  );
  const directoryDescriptionPlaceholderName = getRolePlaceholderName(
    templateSchema.directory.placeholders,
    "directoryDescription",
    "TextBox 12"
  );

  structuredOutline.pages.forEach((page) => {
    const layout = chooseBestLayout(page, templateSchema);
    const pageTitleMaxChars = getRoleMaxChars(layout.placeholders, "pageTitle", 10);
    const introMaxChars = getRoleMaxChars(layout.placeholders, "intro", 30);
    const regularTitleMaxChars = getRoleMaxChars(layout.placeholders, "regularTitle", 15);
    const pageTitlePlaceholderName = getRolePlaceholderName(layout.placeholders, "pageTitle", "Title 1");
    const introPlaceholderName = getRolePlaceholderName(layout.placeholders, "intro", "文本框 39");
    const regularTitlePlaceholderName = getRolePlaceholderName(
      layout.placeholders,
      "regularTitle",
      "Regular Title"
    );
    const detailPlaceholderName = getRolePlaceholderName(
      layout.placeholders,
      "detailBody",
      getRolePlaceholderName(layout.placeholders, "detailSummary", "Detail Text Group")
    );
    const detailSlotLimits = getDetailSlotLimits(layout);
    const detailCapacity = Math.max(1, detailSlotLimits.length);
    const sourcePoints = dedupePoints(page.keyPoints, [page.title, page.intro, page.regularTitle]);
    const fallbackPoints = dedupePoints(
      splitPoints(page.description || page.summary),
      [page.title, page.regularTitle]
    );
    const normalizedPoints = (sourcePoints.length > 0
      ? sourcePoints
      : fallbackPoints.length > 0
        ? fallbackPoints
        : [page.description || page.summary])
      .map((point, index) => fitWithinLimit(point, detailSlotLimits[index % detailCapacity] ?? 60))
      .filter(Boolean);
    const pointChunks = chunkItems(normalizedPoints, detailCapacity);

    pointChunks.forEach((chunk, chunkIndex) => {
      const chunkSlotLimits = detailSlotLimits.slice(0, chunk.length);
      const descriptionMaxChars =
        chunkSlotLimits.reduce((sum, limit) => sum + limit, 0) ||
        detailSlotLimits.reduce((sum, limit) => sum + limit, 0);
      const fittedChunk = chunk.map((point, pointIndex) =>
        fitWithinLimit(point, chunkSlotLimits[pointIndex] ?? detailSlotLimits[pointIndex] ?? 60)
      );
      const description = fitWithinLimit(fittedChunk.join("；"), descriptionMaxChars);
      const notes =
        pointChunks.length > 1
          ? [`当前页要点 ${normalizedPoints.length} 个，超过版式承载 ${detailCapacity} 个，已自动拆为 ${pointChunks.length} 页。`]
          : [];

      slides.push({
        id: `${page.id}-slide-${chunkIndex + 1}`,
        sourcePageId: page.id,
        directoryId: page.directoryId,
        directoryTitle: page.directoryTitle,
        pageNumber: slides.length + 1,
        layoutId: layout.id,
        layoutName: layout.name,
        layoutType: layout.layoutType,
        sourceSlide: layout.sourceSlide,
        detailPointCount: chunk.length,
        detailCapacity,
        fields: {
          pageTitle: {
            id: `${page.id}-field-page-title`,
            label: "页标题",
            placeholderId: `${layout.id}:pageTitle`,
            placeholderName: pageTitlePlaceholderName,
            maxChars: pageTitleMaxChars,
            value: fitWithinLimit(page.title, pageTitleMaxChars),
            overflow: false
          },
          intro: {
            id: `${page.id}-field-intro`,
            label: "内容简介",
            placeholderId: `${layout.id}:intro`,
            placeholderName: introPlaceholderName,
            maxChars: introMaxChars,
            value: fitWithinLimit(page.intro || page.summary, introMaxChars),
            overflow: false
          },
          regularTitle: {
            id: `${page.id}-field-regular-title`,
            label: "常规标题",
            placeholderId: `${layout.id}:regularTitle`,
            placeholderName: regularTitlePlaceholderName,
            maxChars: regularTitleMaxChars,
            value: fitWithinLimit(page.regularTitle || page.title, regularTitleMaxChars),
            overflow: false
          },
          description: {
            id: `${page.id}-field-description-${chunkIndex + 1}`,
            label: "说明性文字",
            placeholderId: `${layout.id}:detailBody`,
            placeholderName: detailPlaceholderName,
            maxChars: descriptionMaxChars,
            value: description,
            overflow: false,
            helper: `当前版式可承载 ${detailCapacity} 个要点，字符上限按节点二 Python 扫描出的文本坑位自动汇总。`
          }
        },
        detailPoints: fittedChunk,
        slotMode: "derived",
        slotContent: {
          sections: [],
          unassigned: []
        },
        notes,
        imageIds: page.imageIds,
        table: page.table,
        overflow: false
      });
    });

    if (pointChunks.length === 0) {
      slides.push({
        id: `${page.id}-slide-1`,
        sourcePageId: page.id,
        directoryId: page.directoryId,
        directoryTitle: page.directoryTitle,
        pageNumber: slides.length + 1,
        layoutId: layout.id,
        layoutName: layout.name,
        layoutType: layout.layoutType,
        sourceSlide: layout.sourceSlide,
        detailPointCount: 0,
        detailCapacity,
        fields: {
          pageTitle: {
            id: `${page.id}-field-page-title`,
            label: "页标题",
            placeholderId: `${layout.id}:pageTitle`,
            placeholderName: pageTitlePlaceholderName,
            maxChars: pageTitleMaxChars,
            value: fitWithinLimit(page.title, pageTitleMaxChars),
            overflow: false
          },
          intro: {
            id: `${page.id}-field-intro`,
            label: "内容简介",
            placeholderId: `${layout.id}:intro`,
            placeholderName: introPlaceholderName,
            maxChars: introMaxChars,
            value: fitWithinLimit(page.intro || page.summary, introMaxChars),
            overflow: false
          },
          regularTitle: {
            id: `${page.id}-field-regular-title`,
            label: "常规标题",
            placeholderId: `${layout.id}:regularTitle`,
            placeholderName: regularTitlePlaceholderName,
            maxChars: regularTitleMaxChars,
            value: fitWithinLimit(page.regularTitle || page.title, regularTitleMaxChars),
            overflow: false
          },
          description: {
            id: `${page.id}-field-description`,
            label: "说明性文字",
            placeholderId: `${layout.id}:detailBody`,
            placeholderName: detailPlaceholderName,
            maxChars: detailSlotLimits.reduce((sum, limit) => sum + limit, 0),
            value: fitWithinLimit(
              page.description || page.summary,
              detailSlotLimits.reduce((sum, limit) => sum + limit, 0)
            ),
            overflow: false,
            helper: `当前版式可承载 ${detailCapacity} 个要点，字符上限按节点二 Python 扫描出的文本坑位自动汇总。`
          }
        },
        detailPoints: [],
        slotMode: "derived",
        slotContent: {
          sections: [],
          unassigned: []
        },
        notes: [],
        imageIds: page.imageIds,
        table: page.table,
        overflow: false
      });
    }
  });

  const slideCountByDirectory = new Map<string, number>();
  slides.forEach((slide) => {
    slideCountByDirectory.set(slide.directoryId, (slideCountByDirectory.get(slide.directoryId) ?? 0) + 1);
  });

  let pageStart = 3;
  const directory = structuredOutline.directory.map((item) => {
    const pageCount = slideCountByDirectory.get(item.id) ?? 0;
    const next = {
      id: item.id,
      order: item.order,
      pageStart,
      pageCount,
      fields: {
        title: {
          id: `${item.id}-title`,
          label: "目录标题",
          placeholderId: "directory:title",
          placeholderName: directoryTitlePlaceholderName,
          maxChars: directoryTitleMaxChars,
          value: fitWithinLimit(item.title, directoryTitleMaxChars),
          overflow: false
        },
        description: {
          id: `${item.id}-description`,
          label: "简要说明",
          placeholderId: "directory:description",
          placeholderName: directoryDescriptionPlaceholderName,
          maxChars: directoryDescriptionMaxChars,
          value: fitWithinLimit(item.description, directoryDescriptionMaxChars),
          overflow: false
        }
      }
    };
    pageStart += pageCount;
    return next;
  });

  return validateAssemblyDocument({
    cover: structuredOutline.cover,
    directoryTitle: "目录",
    directory,
    slides,
    extractedImages: structuredOutline.extractedImages,
    templateSchema,
    structuredOutline,
    validation: {
      hasOverflow: false,
      issues: []
    }
  });
}

export function rebuildPipelineFromConfirmedDirectory(
  pipeline: GenerationPipeline
): GenerationPipeline {
  const node1 = syncStructuredOutlineWithAssembly(pipeline.node1, pipeline.node3);
  const node3 = buildAssemblyInstructions(node1, pipeline.node2);

  return {
    ...pipeline,
    node1,
    node3
  };
}

export async function buildGenerationPipeline(
  file: File,
  input: GeneratorInput
): Promise<GenerationPipeline> {
  return buildGenerationPipelineFromSource(file, input);
}

export async function buildGenerationPipelineFromParsed(
  parsed: ParsedDocument,
  input: GeneratorInput
): Promise<GenerationPipeline> {
  if (input.totalPages < 3) {
    throw new Error("总页数至少需要 3 页，才能包含封面、目录和正文。");
  }

  const aiOutline = await generateOutlineWithAI(parsed, input);
  const mergedOutline = mergeOutlineWithDocument(aiOutline, parsed, input);
  const node1 = buildStructuredOutlineDocument(parsed, mergedOutline, input);
  const node2 = await loadTemplateSchema();
  const node3 = buildAssemblyInstructions(node1, node2);

  return {
    input,
    node1,
    node2,
    node3
  };
}

export async function buildGenerationPipelineFromSource(
  source: File | { name?: string; text: string; type?: "md" | "txt" | "text" },
  input: GeneratorInput
): Promise<GenerationPipeline> {
  const parsed = source instanceof File ? await parseWordDocument(source) : await parseSourceDocument(source);
  return buildGenerationPipelineFromParsed(parsed, input);
}
