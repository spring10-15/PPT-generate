import { normalizeOutlineDocument } from "@/lib/outline-format";
import { composeSlideSummary, splitUnitTitleBody } from "@/lib/slide-content";
import type {
  AssemblyInstructionDocument,
  AssemblySlideInstruction,
  OutlineDocument,
  OutlineSlide,
  SlideLayoutType,
  SlideSlotCard,
  SlideSlotContent,
  SlideSlotSection,
  TemplateLayoutSchema,
  TemplatePlaceholderSchema
} from "@/lib/types";

const LAYOUT_SECTION_SIZES: Partial<Record<SlideLayoutType, number[]>> = {
  overview: [2, 1, 1, 1],
  "two-column": [4],
  "three-column": [2, 2, 2],
  "four-column": [3, 3, 3, 3],
  progress: [3, 2, 3, 2],
  vertical: [2, 2],
  "split-grid": [3, 3],
  hierarchy: [1, 1, 1, 1, 1],
  image: [4],
  "image-left": [4],
  "image-right": [4]
};

function countChars(value: string) {
  return value.trim().length;
}

function normalizeFieldValue(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeText(value: string | undefined) {
  return normalizeFieldValue(value ?? "");
}

function toComparableText(value: string | undefined) {
  return normalizeText(value).replace(/[：:；;。！？,.，、】【（）()[\]\s-]/g, "");
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

function truncateText(value: string, maxChars: number) {
  const normalized = normalizeText(value);
  if (!normalized || normalized.length <= maxChars) {
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

function splitLongText(text: string, chunkLength: number) {
  const normalized = normalizeText(text);
  if (!normalized) {
    return [];
  }

  const units = splitPoints(normalized);
  if (units.length > 1) {
    const chunks: string[] = [];
    let current = "";

    units.forEach((unit) => {
      const next = current ? `${current}；${unit}` : unit;
      if (next.length <= chunkLength) {
        current = next;
        return;
      }

      if (current) {
        chunks.push(current);
      }

      current = unit.length <= chunkLength ? unit : unit.slice(0, chunkLength).trim();
    });

    if (current) {
      chunks.push(current);
    }

    return chunks.filter(Boolean);
  }

  if (normalized.length <= chunkLength) {
    return [normalized];
  }

  const chunks: string[] = [];
  for (let index = 0; index < normalized.length; index += chunkLength) {
    chunks.push(normalized.slice(index, index + chunkLength).trim());
  }
  return chunks.filter(Boolean);
}

function expandPointsForTarget(points: string[], targetCount: number, chunkLength: number) {
  const expanded = points.map((point) => normalizeText(point)).filter(Boolean);

  while (expanded.length < targetCount) {
    let splitIndex = -1;
    let longest = 0;

    expanded.forEach((point, index) => {
      const pieces = splitLongText(point, chunkLength);
      if (pieces.length > 1 && point.length > longest) {
        splitIndex = index;
        longest = point.length;
      }
    });

    if (splitIndex < 0) {
      break;
    }

    const pieces = splitLongText(expanded[splitIndex], chunkLength);
    expanded.splice(splitIndex, 1, ...pieces);
  }

  return expanded;
}

function getRolePlaceholders(
  layout: TemplateLayoutSchema,
  roles: string | string[]
) {
  const roleSet = new Set(Array.isArray(roles) ? roles : [roles]);
  return layout.placeholders.filter((placeholder) => placeholder.kind === "TEXT" && roleSet.has(placeholder.role));
}

type SlotBlueprintSection = {
  labelPlaceholder?: TemplatePlaceholderSchema;
  detailPlaceholders: Array<{
    titlePlaceholder?: TemplatePlaceholderSchema;
    bodyPlaceholder?: TemplatePlaceholderSchema;
  }>;
};

function buildSlotBlueprint(layout: TemplateLayoutSchema): SlotBlueprintSection[] {
  const sectionSizes = LAYOUT_SECTION_SIZES[layout.layoutType];
  const regularPlaceholders = getRolePlaceholders(layout, "regularTitle");
  const smallTitlePlaceholders = getRolePlaceholders(layout, "smallTitle");
  const bodyPlaceholders = getRolePlaceholders(layout, ["detailBody", "detailSummary"]);
  const expandedBodyPlaceholders = bodyPlaceholders.flatMap((placeholder) =>
    Array.from({ length: Math.max(1, placeholder.maxItems ?? 1) }, () => placeholder)
  );
  const totalDetailSlots = expandedBodyPlaceholders.length;

  if (!sectionSizes || totalDetailSlots === 0) {
    return [];
  }

  const leadingBodyOnlySlots = Math.max(0, totalDetailSlots - smallTitlePlaceholders.length);
  let detailIndex = 0;

  return sectionSizes.map((size, sectionIndex) => {
    const detailPlaceholders = Array.from({ length: size }, () => {
      const titleIndex = detailIndex - leadingBodyOnlySlots;
      const next = {
        titlePlaceholder: titleIndex >= 0 ? smallTitlePlaceholders[titleIndex] : undefined,
        bodyPlaceholder: expandedBodyPlaceholders[detailIndex]
      };
      detailIndex += 1;
      return next;
    }).filter((slot) => slot.titlePlaceholder || slot.bodyPlaceholder);

    return {
      labelPlaceholder: regularPlaceholders[sectionIndex],
      detailPlaceholders
    };
  });
}

function buildHeading(text: string, fallback: string, maxChars: number) {
  const labeled = splitUnitTitleBody(text);
  if (labeled.title) {
    return truncateText(labeled.title, maxChars) || fallback;
  }

  const candidate = normalizeText(text)
    .replace(/[：:；;。！？,.，、]/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .join("");

  return truncateText(candidate || fallback, maxChars) || fallback;
}

function buildBodyLines(text: string, maxChars: number, maxLines: number) {
  const units = splitLongText(text, Math.max(8, Math.min(maxChars, Math.floor(maxChars / Math.max(1, maxLines)))));
  if (units.length === 0) {
    return [" "];
  }

  return units.slice(0, maxLines).map((unit) => truncateText(unit, maxChars)).filter(Boolean);
}

function buildSlotCard(
  text: string,
  fallback: string,
  order: number,
  titlePlaceholder: TemplatePlaceholderSchema | undefined,
  bodyPlaceholder: TemplatePlaceholderSchema | undefined,
  slideId: string
): SlideSlotCard {
  const titleMaxChars = titlePlaceholder?.maxChars ?? 15;
  const bodyMaxChars = bodyPlaceholder?.maxChars ?? 60;
  const maxLines = Math.max(1, bodyPlaceholder?.maxLines ?? 3);
  const normalized = normalizeText(text);

  if (!normalized) {
    return {
      id: `${slideId}-slot-card-${order}`,
      order,
      title: " ",
      body: " ",
      bodyLines: [" "],
      titleMaxChars,
      bodyMaxChars,
      titleOverflow: false,
      bodyOverflow: false
    };
  }

  const units = splitPoints(normalized);
  const labeled = splitUnitTitleBody(normalized);
  const title = buildHeading(labeled.title || units[0] || normalized, fallback, titleMaxChars);
  const bodySource = labeled.body || dedupePoints(units.length > 1 ? units.slice(1) : [normalized], [title]).join("；");
  const bodyCandidate = dedupePoints(splitPoints(bodySource || normalized), [title]).join("；");
  const body = truncateText(bodyCandidate || normalized, bodyMaxChars) || " ";

  return {
    id: `${slideId}-slot-card-${order}`,
    order,
    title,
    body,
    bodyLines: buildBodyLines(body, bodyMaxChars, maxLines),
    titleMaxChars,
    bodyMaxChars,
    titleOverflow: countChars(title) > titleMaxChars,
    bodyOverflow: countChars(body) > bodyMaxChars
  };
}

function buildEmptySlotCard(
  order: number,
  titlePlaceholder: TemplatePlaceholderSchema | undefined,
  bodyPlaceholder: TemplatePlaceholderSchema | undefined,
  slideId: string
) {
  return buildSlotCard("", `要点${order}`, order, titlePlaceholder, bodyPlaceholder, slideId);
}

function normalizeSlotCard(
  card: SlideSlotCard | undefined,
  fallback: string,
  order: number,
  titlePlaceholder: TemplatePlaceholderSchema | undefined,
  bodyPlaceholder: TemplatePlaceholderSchema | undefined,
  slideId: string
): SlideSlotCard {
  const emptyCard = buildEmptySlotCard(order, titlePlaceholder, bodyPlaceholder, slideId);
  const title = normalizeText(card?.title);
  const body = normalizeText(card?.body);
  const bodyMaxChars = bodyPlaceholder?.maxChars ?? card?.bodyMaxChars ?? emptyCard.bodyMaxChars;
  const titleMaxChars = titlePlaceholder?.maxChars ?? card?.titleMaxChars ?? emptyCard.titleMaxChars;
  const maxLines = Math.max(1, bodyPlaceholder?.maxLines ?? 3);

  return {
    ...emptyCard,
    title: title || emptyCard.title,
    body: body || emptyCard.body,
    bodyLines: buildBodyLines(body || emptyCard.body, bodyMaxChars, maxLines),
    titleMaxChars,
    bodyMaxChars,
    titleOverflow: countChars(title) > titleMaxChars,
    bodyOverflow: countChars(body) > bodyMaxChars
  };
}

function buildStructuredSlotContent(
  slide: AssemblySlideInstruction,
  layout: TemplateLayoutSchema | undefined
): SlideSlotContent {
  if (!layout || slide.layoutType === "table") {
    return {
      sections: [],
      unassigned: []
    };
  }

  const blueprint = buildSlotBlueprint(layout);
  if (blueprint.length === 0) {
    return {
      sections: [],
      unassigned: []
    };
  }

  const totalDetailSlots = blueprint.reduce((sum, section) => sum + section.detailPlaceholders.length, 0);
  const maxChunkLength =
    Math.max(
      24,
      ...blueprint.flatMap((section) =>
        section.detailPlaceholders.map((slot) => slot.bodyPlaceholder?.maxChars ?? 60)
      )
    ) || 60;
  const sourcePoints = dedupePoints(
    [...slide.detailPoints, ...splitPoints(slide.fields.description.value)],
    [slide.fields.pageTitle.value, slide.fields.intro.value, slide.fields.regularTitle.value]
  );
  const expandedPoints = expandPointsForTarget(
    sourcePoints.length > 0 ? sourcePoints : [slide.fields.description.value],
    totalDetailSlots,
    maxChunkLength
  );
  const sections: SlideSlotSection[] = [];
  let pointIndex = 0;

  blueprint.forEach((section, sectionIndex) => {
    const cards = section.detailPlaceholders.map((slot, detailIndex) => {
      const point = expandedPoints[pointIndex] ?? "";
      pointIndex += 1;
      return buildSlotCard(
        point,
        `要点${pointIndex}`,
        detailIndex + 1,
        slot.titlePlaceholder,
        slot.bodyPlaceholder,
        `${slide.id}-section-${sectionIndex + 1}`
      );
    });

    const labelMaxChars = section.labelPlaceholder?.maxChars ?? slide.fields.regularTitle.maxChars;
    const labelSource =
      sectionIndex === 0
        ? slide.fields.regularTitle.value || cards[0]?.title || slide.fields.pageTitle.value
        : cards[0]?.title || slide.fields.regularTitle.value || `模块${sectionIndex + 1}`;

    sections.push({
      id: `${slide.id}-section-${sectionIndex + 1}`,
      order: sectionIndex + 1,
      label: truncateText(labelSource, labelMaxChars) || " ",
      labelMaxChars,
      labelOverflow: false,
      cards
    });
  });

  return {
    sections,
    unassigned: expandedPoints.slice(pointIndex)
  };
}

function normalizeManualSlotContent(
  slide: AssemblySlideInstruction,
  layout: TemplateLayoutSchema | undefined
): SlideSlotContent {
  if (!layout || slide.layoutType === "table") {
    return slide.slotContent;
  }

  const blueprint = buildSlotBlueprint(layout);
  if (blueprint.length === 0) {
    return slide.slotContent;
  }

  const sections = blueprint.map((section, sectionIndex) => {
    const sourceSection = slide.slotContent.sections[sectionIndex];
    const labelMaxChars = section.labelPlaceholder?.maxChars ?? slide.fields.regularTitle.maxChars;
    const cards = section.detailPlaceholders.map((slot, detailIndex) =>
      normalizeSlotCard(
        sourceSection?.cards[detailIndex],
        `要点${detailIndex + 1}`,
        detailIndex + 1,
        slot.titlePlaceholder,
        slot.bodyPlaceholder,
        `${slide.id}-section-${sectionIndex + 1}`
      )
    );
    const label = normalizeText(sourceSection?.label) || (sectionIndex === 0 ? slide.fields.regularTitle.value : " ");

    return {
      id: sourceSection?.id ?? `${slide.id}-section-${sectionIndex + 1}`,
      order: sectionIndex + 1,
      label: label || " ",
      labelMaxChars,
      labelOverflow: countChars(label) > labelMaxChars,
      cards
    };
  });

  return {
    sections,
    unassigned: (slide.slotContent.unassigned ?? []).map((item) => normalizeText(item)).filter(Boolean)
  };
}

function summarizeSlotContent(slotContent: SlideSlotContent) {
  const points = dedupePoints(
    slotContent.sections.flatMap((section) =>
      section.cards.flatMap((card) => [card.title, card.body, ...card.bodyLines])
    )
  );
  return {
    detailPoints: points,
    regularTitle: normalizeText(slotContent.sections[0]?.label),
    description: truncateText(points.join("；"), Math.max(60, points.join("；").length))
  };
}

function slotContentHasOverflow(slotContent: SlideSlotContent) {
  return slotContent.sections.some(
    (section) =>
      section.labelOverflow ||
      section.cards.some((card) => card.titleOverflow || card.bodyOverflow)
  );
}

function withOverflow<T extends { value: string; maxChars: number; overflow: boolean }>(field: T): T {
  const nextValue = normalizeFieldValue(field.value);
  return {
    ...field,
    value: nextValue,
    overflow: countChars(nextValue) > field.maxChars
  };
}

function recomputeSlide(
  slide: AssemblySlideInstruction,
  document: AssemblyInstructionDocument
): AssemblySlideInstruction {
  const layout =
    document.templateSchema.detailLayouts.find((item) => item.id === slide.layoutId) ??
    document.templateSchema.detailLayouts.find((item) => item.sourceSlide === slide.sourceSlide);
  const slotContent =
    slide.slotMode === "manual"
      ? normalizeManualSlotContent(slide, layout)
      : buildStructuredSlotContent(slide, layout);
  const slotSummary = summarizeSlotContent(slotContent);
  const pageTitle = withOverflow(slide.fields.pageTitle);
  const intro = withOverflow(slide.fields.intro);
  const regularTitle = withOverflow({
    ...slide.fields.regularTitle,
    value: slotSummary.regularTitle || slide.fields.regularTitle.value
  });
  const description = withOverflow({
    ...slide.fields.description,
    value: truncateText(slotSummary.description || slide.fields.description.value, slide.fields.description.maxChars)
  });
  const overflow =
    pageTitle.overflow ||
    intro.overflow ||
    regularTitle.overflow ||
    description.overflow ||
    slotContentHasOverflow(slotContent);

  return {
    ...slide,
    fields: {
      pageTitle,
      intro,
      regularTitle,
      description
    },
    detailPoints: slotSummary.detailPoints.length > 0 ? slotSummary.detailPoints : slide.detailPoints,
    slotMode: slide.slotMode,
    slotContent,
    overflow
  };
}

export function validateAssemblyDocument(document: AssemblyInstructionDocument): AssemblyInstructionDocument {
  const slideCountByDirectory = new Map<string, number>();
  document.slides.forEach((slide) => {
    slideCountByDirectory.set(slide.directoryId, (slideCountByDirectory.get(slide.directoryId) ?? 0) + 1);
  });

  let pageStart = 3;
  const directory = document.directory
    .map((item) => {
      const title = withOverflow(item.fields.title);
      const description = withOverflow(item.fields.description);
      const pageCount = slideCountByDirectory.get(item.id) ?? 0;
      const next = {
        ...item,
        pageStart,
        pageCount,
        fields: {
          title,
          description
        }
      };
      pageStart += pageCount;
      return next;
    })
    .filter((item) => item.pageCount > 0);

  const slides = document.slides.map((slide, index) => ({
    ...recomputeSlide(slide, document),
    pageNumber: index + 1
  }));

  const issues = [
    ...directory
      .flatMap((item) => [
        item.fields.title.overflow
          ? `目录“${item.fields.title.value || item.id}”标题超出 ${item.fields.title.maxChars} 字`
          : "",
        item.fields.description.overflow
          ? `目录“${item.fields.title.value || item.id}”简述超出 ${item.fields.description.maxChars} 字`
          : ""
      ])
      .filter(Boolean),
    ...slides
      .flatMap((slide) => [
        slide.fields.pageTitle.overflow
          ? `第 ${slide.pageNumber} 页标题超出 ${slide.fields.pageTitle.maxChars} 字`
          : "",
        slide.fields.intro.overflow
          ? `第 ${slide.pageNumber} 页内容简介超出 ${slide.fields.intro.maxChars} 字`
          : "",
        slide.fields.regularTitle.overflow
          ? `第 ${slide.pageNumber} 页常规标题超出 ${slide.fields.regularTitle.maxChars} 字`
          : "",
        slide.fields.description.overflow
          ? `第 ${slide.pageNumber} 页说明性文字超出 ${slide.fields.description.maxChars} 字`
          : ""
      ])
      .filter(Boolean),
    ...slides
      .flatMap((slide) =>
        slide.slotContent.sections.flatMap((section) => [
          section.labelOverflow
            ? `第 ${slide.pageNumber} 页第 ${section.order} 分区标题超出 ${section.labelMaxChars} 字`
            : "",
          ...section.cards.flatMap((card) => [
            card.titleOverflow
              ? `第 ${slide.pageNumber} 页第 ${section.order} 分区第 ${card.order} 条小标题超出 ${card.titleMaxChars} 字`
              : "",
            card.bodyOverflow
              ? `第 ${slide.pageNumber} 页第 ${section.order} 分区第 ${card.order} 条说明性文字超出 ${card.bodyMaxChars} 字`
              : ""
          ])
        ])
      )
      .filter(Boolean)
  ];

  return {
    ...document,
    directory,
    slides,
    validation: {
      hasOverflow: issues.length > 0,
      issues
    }
  };
}

export function assemblyToOutlineDocument(document: AssemblyInstructionDocument): OutlineDocument {
  const validated = validateAssemblyDocument(document);
  const directory = validated.directory.map((item) => ({
    id: item.id,
    title: item.fields.title.value,
    description: item.fields.description.value,
    pageStart: item.pageStart,
    pageCount: item.pageCount
  }));

  const slides: OutlineSlide[] = validated.slides.map((slide) => {
    const content = {
      intro: slide.fields.intro.value,
      regularTitle: slide.fields.regularTitle.value,
      description: slide.fields.description.value
    };

    return {
      id: slide.id,
      directoryId: slide.directoryId,
      directoryTitle: slide.directoryTitle,
      title: slide.fields.pageTitle.value,
      type: slide.layoutType,
      summary: composeSlideSummary(content),
      content,
      imageSuggestion:
        slide.imageIds.length > 0
          ? "优先使用文档原图；若版面不足，保留图片占位。"
          : "文档无原图，保留模板占位框。",
      imageIds: slide.imageIds,
      table: slide.table,
      detailPoints: slide.detailPoints,
      slotContent: slide.slotContent
    };
  });

  return normalizeOutlineDocument({
    cover: {
      title: validated.cover.title,
      subtitle: validated.cover.subtitle,
      userName: validated.cover.userName,
      dateLabel: validated.cover.dateLabel
    },
    directoryTitle: validated.directoryTitle,
    directory,
    slides,
    pageSummary: {
      totalPages: slides.length + 2,
      coverPages: 1,
      directoryPages: 1,
      detailPages: slides.length
    },
    extractedImages: validated.extractedImages
  });
}
