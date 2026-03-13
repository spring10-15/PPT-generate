import { NextResponse } from "next/server";
import { z } from "zod";
import { handleChatTurn } from "@/lib/chat-orchestrator";
import type { ChatDirective, ChatSessionState } from "@/lib/types";

const stateSchema = z
  .object({
    stage: z.enum(["collect", "directory", "summary", "ready"]),
    confidence: z.number(),
    input: z.object({
      userName: z.string(),
      targetAudience: z.string(),
      estimatedMinutes: z.number().nullable(),
      totalPages: z.number().nullable()
    }),
    sourceMaterial: z
      .object({
        kind: z.enum(["file", "text"]),
        name: z.string(),
        preview: z.string(),
        textContent: z.string().optional()
      })
      .nullable(),
    pipeline: z.unknown().nullable()
  })
  .nullable();

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const message = String(formData.get("message") ?? "");
    const rawState = formData.get("state");
    const rawDirective = formData.get("directive");
    const parsedState = rawState
      ? stateSchema.parse(JSON.parse(String(rawState))) as ChatSessionState
      : null;
    const parsedDirective = rawDirective
      ? (JSON.parse(String(rawDirective)) as ChatDirective)
      : null;

    const response = await handleChatTurn({
      state: parsedState,
      message,
      file: file instanceof File ? file : null,
      directive: parsedDirective
    });

    return NextResponse.json(response);
  } catch (error) {
    const message = error instanceof Error ? error.message : "聊天处理失败。";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
