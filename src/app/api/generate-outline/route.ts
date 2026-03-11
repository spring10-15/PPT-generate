import { NextResponse } from "next/server";
import { z } from "zod";
import { buildOutline } from "@/lib/outline";

const schema = z.object({
  userName: z.string().min(1, "请填写用户名称"),
  totalPages: z.coerce.number().int().min(3, "总页数至少为 3")
});

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const parsedInput = schema.parse({
      userName: formData.get("userName"),
      totalPages: formData.get("totalPages")
    });

    if (!(file instanceof File)) {
      throw new Error("请上传 doc 或 docx 文件。");
    }

    const outline = await buildOutline(file, parsedInput);
    return NextResponse.json(outline);
  } catch (error) {
    const message = error instanceof Error ? error.message : "生成摘要失败。";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
