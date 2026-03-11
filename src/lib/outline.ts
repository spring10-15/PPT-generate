import { generateOutlineWithAI, mergeOutlineWithDocument } from "@/lib/ai/siliconflow";
import { parseWordDocument } from "@/lib/doc/parse";
import type { GeneratorInput, OutlineDocument } from "@/lib/types";

export async function buildOutline(file: File, input: GeneratorInput): Promise<OutlineDocument> {
  if (input.totalPages < 3) {
    throw new Error("总页数至少需要 3 页，才能包含封面、目录和正文。");
  }

  const parsed = await parseWordDocument(file);
  const aiOutline = await generateOutlineWithAI(parsed, input);
  return mergeOutlineWithDocument(aiOutline, parsed, input);
}
