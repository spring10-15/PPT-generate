import { NextResponse } from "next/server";
import { z } from "zod";
import { generatePptBuffer } from "@/lib/ppt/generate";
import { SUPPORTED_LAYOUT_TYPES } from "@/lib/types";

const slideSchema = z.object({
  id: z.string(),
  directoryId: z.string(),
  directoryTitle: z.string(),
  title: z.string(),
  type: z.enum(SUPPORTED_LAYOUT_TYPES),
  summary: z.string(),
  content: z.object({
    intro: z.string(),
    regularTitle: z.string(),
    description: z.string()
  }),
  imageSuggestion: z.string(),
  imageIds: z.array(z.string()),
  table: z.array(z.array(z.string())).optional()
});

const outlineSchema = z.object({
  cover: z.object({
    title: z.string(),
    subtitle: z.string(),
    userName: z.string(),
    dateLabel: z.string()
  }),
  directoryTitle: z.string(),
  directory: z.array(
    z.object({
      id: z.string(),
      title: z.string(),
      description: z.string(),
      pageStart: z.number().int(),
      pageCount: z.number().int()
    })
  ),
  slides: z.array(slideSchema),
  pageSummary: z.object({
    totalPages: z.number().int(),
    coverPages: z.number().int(),
    directoryPages: z.number().int(),
    detailPages: z.number().int()
  }),
  extractedImages: z.array(
    z.object({
      id: z.string(),
      name: z.string(),
      mimeType: z.string(),
      dataUri: z.string()
    })
  )
});

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const outline = outlineSchema.parse(body);
    const buffer = await generatePptBuffer(outline);
    const fileName = `${outline.cover.title.replace(/[\\/:*?"<>|]/g, "_") || "汇报材料"}.pptx`;

    return new NextResponse(new Uint8Array(buffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`
      }
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "导出 PPT 失败。";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
