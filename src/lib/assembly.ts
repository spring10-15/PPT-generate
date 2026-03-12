import { normalizeOutlineDocument } from "@/lib/outline-format";
import { composeSlideSummary } from "@/lib/slide-content";
import type {
  AssemblyInstructionDocument,
  AssemblySlideInstruction,
  OutlineDocument,
  OutlineSlide
} from "@/lib/types";

function countChars(value: string) {
  return value.trim().length;
}

function normalizeFieldValue(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function withOverflow<T extends { value: string; maxChars: number; overflow: boolean }>(field: T): T {
  const nextValue = normalizeFieldValue(field.value);
  return {
    ...field,
    value: nextValue,
    overflow: countChars(nextValue) > field.maxChars
  };
}

function recomputeSlide(slide: AssemblySlideInstruction): AssemblySlideInstruction {
  const pageTitle = withOverflow(slide.fields.pageTitle);
  const intro = withOverflow(slide.fields.intro);
  const regularTitle = withOverflow(slide.fields.regularTitle);
  const description = withOverflow(slide.fields.description);
  const overflow =
    pageTitle.overflow || intro.overflow || regularTitle.overflow || description.overflow;

  return {
    ...slide,
    fields: {
      pageTitle,
      intro,
      regularTitle,
      description
    },
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
    ...recomputeSlide(slide),
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
      table: slide.table
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
