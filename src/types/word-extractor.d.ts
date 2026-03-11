declare module "word-extractor" {
  export interface ExtractedDocument {
    getBody(): string;
  }

  export default class WordExtractor {
    extract(source: string | Buffer): Promise<ExtractedDocument>;
  }
}
