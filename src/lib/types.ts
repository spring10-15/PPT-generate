export const SUPPORTED_LAYOUT_TYPES = [
  "overview",
  "two-column",
  "three-column",
  "four-column",
  "progress",
  "vertical",
  "split-grid",
  "hierarchy",
  "table",
  "image",
  "image-left",
  "image-right"
] as const;

export type SlideLayoutType = (typeof SUPPORTED_LAYOUT_TYPES)[number];

export type ExtractedImage = {
  id: string;
  name: string;
  mimeType: string;
  dataUri: string;
};

export type ParsedSection = {
  id: string;
  title: string;
  level: number;
  paragraphs: string[];
  tables: string[][];
  imageIds: string[];
};

export type ParsedDocument = {
  fileName: string;
  fileType: "docx" | "doc";
  rawText: string;
  titleGuess: string;
  subtitleGuess: string;
  sections: ParsedSection[];
  images: ExtractedImage[];
};

export type DirectoryItem = {
  id: string;
  title: string;
  description: string;
  pageStart: number;
  pageCount: number;
};

export type SlideContentDraft = {
  intro: string;
  regularTitle: string;
  description: string;
};

export type OutlineSlide = {
  id: string;
  directoryId: string;
  directoryTitle: string;
  title: string;
  type: SlideLayoutType;
  summary: string;
  content: SlideContentDraft;
  imageSuggestion: string;
  imageIds: string[];
  table?: string[][];
};

export type OutlineDocument = {
  cover: {
    title: string;
    subtitle: string;
    userName: string;
    dateLabel: string;
  };
  directoryTitle: string;
  directory: DirectoryItem[];
  slides: OutlineSlide[];
  pageSummary: {
    totalPages: number;
    coverPages: number;
    directoryPages: number;
    detailPages: number;
  };
  extractedImages: ExtractedImage[];
};

export type GeneratorInput = {
  userName: string;
  totalPages: number;
};

export type TemplateColorTokens = {
  primary: string;
  accent: string;
  highlight: string;
  secondary1: string;
  secondary2: string;
  secondary3: string;
  chart1: string;
  chart2: string;
  chart3: string;
  darkText: string;
  lightText: string;
};

export type TemplateMetrics = {
  widthInches: number;
  heightInches: number;
};
