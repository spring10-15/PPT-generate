import { NextResponse } from "next/server";
import { z } from "zod";
import { buildGenerationPipeline } from "@/lib/pipeline";

const schema = z.object({
  userName: z.string().min(1, "请填写用户名称"),
  targetAudience: z.string().min(1, "请填写目标受众"),
  estimatedMinutes: z.coerce.number().int().min(1, "预估演讲时长至少为 1 分钟").nullable(),
  totalPages: z.coerce.number().int().min(3, "总页数至少为 3")
});

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const parsedInput = schema.parse({
      userName: formData.get("userName"),
      targetAudience: formData.get("targetAudience"),
      estimatedMinutes: formData.get("estimatedMinutes")
        ? Number(formData.get("estimatedMinutes"))
        : null,
      totalPages: formData.get("totalPages")
    });

    if (!(file instanceof File)) {
      throw new Error("请上传 doc、docx、md 或 txt 文件。");
    }

    const pipeline = await buildGenerationPipeline(file, parsedInput);
    return NextResponse.json(pipeline);
  } catch (error) {
    const message = error instanceof Error ? error.message : "生成摘要失败。";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
