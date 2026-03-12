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
  fileType: "docx" | "doc" | "md" | "txt" | "text";
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
  targetAudience: string;
  estimatedMinutes: number | null;
  totalPages: number;
};

export type OutlineRelation = "root" | "parallel" | "sequential" | "supporting";

export type OutlineHierarchyLevel = "H1" | "H2" | "H3" | "H4";

export type StructuredOutlineHierarchyNode = {
  id: string;
  level: OutlineHierarchyLevel;
  title: string;
  summary: string;
  relation: OutlineRelation;
  pointCount: number;
  children: StructuredOutlineHierarchyNode[];
};

export type StructuredOutlineDirectoryItem = {
  id: string;
  order: number;
  title: string;
  description: string;
};

export type StructuredOutlinePage = {
  id: string;
  directoryId: string;
  directoryTitle: string;
  title: string;
  intro: string;
  regularTitle: string;
  description: string;
  summary: string;
  keyPoints: string[];
  relation: OutlineRelation;
  suggestedLayoutType: SlideLayoutType;
  imageIds: string[];
  table?: string[][];
  hasImage: boolean;
  hasTable: boolean;
};

export type StructuredOutlineDocument = {
  cover: {
    title: string;
    subtitle: string;
    userName: string;
    targetAudience: string;
    estimatedMinutes: number | null;
    dateLabel: string;
  };
  directory: StructuredOutlineDirectoryItem[];
  pages: StructuredOutlinePage[];
  hierarchy: StructuredOutlineHierarchyNode[];
  extractedImages: ExtractedImage[];
  source: {
    fileName: string;
    totalPages: number;
    targetAudience: string;
    estimatedMinutes: number | null;
  };
};

export type TemplatePlaceholderKind = "TEXT" | "PICTURE" | "TABLE" | "CHART";

export type TemplatePlaceholderSchema = {
  id: string;
  name: string;
  kind: TemplatePlaceholderKind;
  role: string;
  occurrence?: number;
  maxChars?: number;
  maxItems?: number;
  maxLines?: number;
  widthPt?: number;
  heightPt?: number;
  fontSizePt?: number;
  required?: boolean;
  description?: string;
};

export type TemplateLayoutSchema = {
  id: string;
  name: string;
  layoutType: SlideLayoutType;
  sourceSlide: number;
  summary: string;
  capacities: {
    detailItems: number;
    pictureSlots: number;
    tableSlots: number;
    chartSlots: number;
  };
  placeholders: TemplatePlaceholderSchema[];
};

export type TemplateSchemaLibrary = {
  cover: {
    sourceSlide: number;
    placeholders: TemplatePlaceholderSchema[];
  };
  directory: {
    sourceSlide: number;
    maxItems: number;
    placeholders: TemplatePlaceholderSchema[];
  };
  detailLayouts: TemplateLayoutSchema[];
  colors: TemplateColorTokens;
  metrics: TemplateMetrics;
  scanner: {
    engine: "python-ooxml-scan" | "static-fallback";
    templatePath: string;
    slideCount: number;
  };
};

export type AssemblyTextField = {
  id: string;
  label: string;
  placeholderId: string;
  placeholderName: string;
  maxChars: number;
  value: string;
  overflow: boolean;
  helper?: string;
};

export type AssemblyDirectoryItem = {
  id: string;
  order: number;
  pageStart: number;
  pageCount: number;
  fields: {
    title: AssemblyTextField;
    description: AssemblyTextField;
  };
};

export type AssemblySlideInstruction = {
  id: string;
  sourcePageId: string;
  directoryId: string;
  directoryTitle: string;
  pageNumber: number;
  layoutId: string;
  layoutName: string;
  layoutType: SlideLayoutType;
  sourceSlide: number;
  detailPointCount: number;
  detailCapacity: number;
  fields: {
    pageTitle: AssemblyTextField;
    intro: AssemblyTextField;
    regularTitle: AssemblyTextField;
    description: AssemblyTextField;
  };
  notes: string[];
  imageIds: string[];
  table?: string[][];
  overflow: boolean;
};

export type AssemblyInstructionDocument = {
  cover: {
    title: string;
    subtitle: string;
    userName: string;
    targetAudience: string;
    estimatedMinutes: number | null;
    dateLabel: string;
  };
  directoryTitle: string;
  directory: AssemblyDirectoryItem[];
  slides: AssemblySlideInstruction[];
  extractedImages: ExtractedImage[];
  templateSchema: TemplateSchemaLibrary;
  structuredOutline: StructuredOutlineDocument;
  validation: {
    hasOverflow: boolean;
    issues: string[];
  };
};

export type GenerationPipeline = {
  input: GeneratorInput;
  node1: StructuredOutlineDocument;
  node2: TemplateSchemaLibrary;
  node3: AssemblyInstructionDocument;
};

export type SourceMaterialDraft = {
  kind: "file" | "text";
  name: string;
  preview: string;
  textContent?: string;
};

export type ChatSessionInputDraft = {
  userName: string;
  targetAudience: string;
  estimatedMinutes: number | null;
  totalPages: number | null;
};

export type ChatSessionStage = "collect" | "directory" | "summary" | "ready";

export type ChatSessionState = {
  stage: ChatSessionStage;
  confidence: number;
  input: ChatSessionInputDraft;
  sourceMaterial: SourceMaterialDraft | null;
  pipeline: GenerationPipeline | null;
};

export type ChatRouteResponse = {
  state: ChatSessionState;
  assistantMessages: string[];
  action?: "none" | "export";
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
