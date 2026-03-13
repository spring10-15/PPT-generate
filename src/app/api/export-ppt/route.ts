import { NextResponse } from "next/server";
import { assemblyToOutlineDocument, validateAssemblyDocument } from "@/lib/assembly";
import { generatePptBuffer } from "@/lib/ppt/generate";
import type { AssemblyInstructionDocument } from "@/lib/types";

function assertAssemblyDocument(body: unknown): AssemblyInstructionDocument {
  if (!body || typeof body !== "object") {
    throw new Error("导出参数无效。");
  }

  const candidate = body as Partial<AssemblyInstructionDocument>;
  if (
    !candidate.cover ||
    !candidate.directory ||
    !candidate.slides ||
    !candidate.extractedImages ||
    !candidate.templateSchema
  ) {
    throw new Error("导出参数缺少必要的装配指令。");
  }

  return candidate as AssemblyInstructionDocument;
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const assembly = validateAssemblyDocument(assertAssemblyDocument(body));

    if (assembly.validation.hasOverflow) {
      throw new Error("当前仍有超出模板物理边界的字段，请先修正后再导出。");
    }

    const outline = assemblyToOutlineDocument(assembly);
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
