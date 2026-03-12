import { z } from "zod";
import { validateAssemblyDocument } from "@/lib/assembly";
import { buildGenerationPipelineFromSource } from "@/lib/pipeline";
import type {
  AssemblyInstructionDocument,
  ChatRouteResponse,
  ChatSessionState,
  GenerationPipeline,
  GeneratorInput,
  SourceMaterialDraft
} from "@/lib/types";

const API_URL =
  process.env.SILICONFLOW_API_URL ?? "https://api.siliconflow.cn/v1/chat/completions";
const MODEL = process.env.SILICONFLOW_MODEL ?? "Pro/zai-org/GLM-5";

const collectTurnSchema = z.object({
  reset: z.boolean().optional(),
  userName: z.string().optional(),
  targetAudience: z.string().optional(),
  estimatedMinutes: z.number().int().positive().optional(),
  totalPages: z.number().int().min(3).optional(),
  sourceText: z.string().optional()
});

const directoryTurnSchema = z.object({
  action: z.enum(["confirm", "revise", "unknown"]).default("unknown"),
  updates: z
    .array(
      z.object({
        order: z.number().int().min(1),
        title: z.string().optional(),
        description: z.string().optional(),
        remove: z.boolean().optional()
      })
    )
    .default([])
});

const summaryTurnSchema = z.object({
  action: z.enum(["confirm", "revise", "unknown"]).default("unknown"),
  updates: z
    .array(
      z.object({
        pageNumber: z.number().int().min(1),
        remove: z.boolean().optional(),
        pageTitle: z.string().optional(),
        intro: z.string().optional(),
        regularTitle: z.string().optional(),
        description: z.string().optional()
      })
    )
    .default([])
});

const initialState: ChatSessionState = {
  stage: "collect",
  confidence: 0,
  input: {
    userName: "",
    targetAudience: "",
    estimatedMinutes: null,
    totalPages: null
  },
  sourceMaterial: null,
  pipeline: null
};

function normalizeText(value: string | undefined) {
  return (value ?? "").replace(/\s+/g, " ").trim();
}

function truncateText(value: string, maxChars = 120) {
  const normalized = normalizeText(value);
  if (normalized.length <= maxChars) {
    return normalized;
  }

  return `${normalized.slice(0, Math.max(1, maxChars - 1)).trim()}…`;
}

function isResetIntent(message: string) {
  return /(重新开始|重置|清空|从头开始|reset)/i.test(message);
}

function isExportIntent(message: string) {
  return /(生成|导出|下载).{0,8}(ppt|pptx)?/i.test(message);
}

function shouldTreatAsSourceText(message: string) {
  const normalized = message.trim();
  if (normalized.length >= 120) {
    return true;
  }

  return normalized.includes("\n") && normalized.length >= 60;
}

function computeConfidence(state: ChatSessionState) {
  let score = 0;
  if (normalizeText(state.input.userName)) score += 25;
  if (normalizeText(state.input.targetAudience)) score += 20;
  if (state.input.estimatedMinutes && state.input.estimatedMinutes >= 1) score += 15;
  if (state.input.totalPages && state.input.totalPages >= 3) score += 15;
  if (state.sourceMaterial) score += 25;
  return Math.min(100, score);
}

function buildMissingPrompt(state: ChatSessionState) {
  if (!normalizeText(state.input.userName)) {
    return "先告诉我汇报人名称。";
  }

  if (!normalizeText(state.input.targetAudience)) {
    return "请继续告诉我这份 PPT 的目标受众。";
  }

  if (!state.input.estimatedMinutes) {
    return "再告诉我预估演讲时长，单位分钟。";
  }

  if (!state.input.totalPages) {
    return "请告诉我需要生成多少页 PPT，总页数至少 3 页。";
  }

  if (!state.sourceMaterial) {
    return "现在请上传素材文件，支持 `.doc / .docx / .md / .txt`，也可以直接把纯文字素材粘贴到对话框里。";
  }

  return "信息已经齐了，我准备开始生成目录和摘要。";
}

async function requestStructuredJson<T>(
  schema: z.ZodSchema<T>,
  systemPrompt: string,
  userPrompt: string
): Promise<T | null> {
  const apiKey = process.env.SILICONFLOW_API_KEY;
  if (!apiKey) {
    return null;
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 12000);

  try {
    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: MODEL,
        temperature: 0.2,
        response_format: { type: "json_object" },
        messages: [
          {
            role: "system",
            content: systemPrompt
          },
          {
            role: "user",
            content: userPrompt
          }
        ]
      }),
      signal: controller.signal
    });

    if (!response.ok) {
      return null;
    }

    const payload = (await response.json()) as {
      choices?: Array<{
        message?: {
          content?: string;
        };
      }>;
    };
    const content = payload.choices?.[0]?.message?.content;
    if (!content) {
      return null;
    }

    const parsed = schema.safeParse(JSON.parse(content));
    return parsed.success ? parsed.data : null;
  } catch {
    return null;
  } finally {
    clearTimeout(timeout);
  }
}

async function parseCollectTurn(message: string) {
  const heuristic: z.infer<typeof collectTurnSchema> = {};
  const normalized = normalizeText(message);

  if (!normalized) {
    return heuristic;
  }

  if (isResetIntent(normalized)) {
    heuristic.reset = true;
  }

  const pageMatch = normalized.match(/(\d+)\s*页/);
  if (pageMatch) {
    heuristic.totalPages = Number(pageMatch[1]);
  }

  const minuteMatch = normalized.match(/(\d+)\s*(分钟|min)/i);
  if (minuteMatch) {
    heuristic.estimatedMinutes = Number(minuteMatch[1]);
  }

  const nameMatch =
    normalized.match(/(?:汇报人|署名|我是|我叫)\s*[:：]?\s*([^\n，,。；;]{2,24})/) ??
    (/^[^\d，,。；;]{2,12}$/.test(normalized) ? [normalized, normalized] : null);
  if (nameMatch?.[1]) {
    heuristic.userName = normalizeText(nameMatch[1]);
  }

  const audienceMatch = normalized.match(/(?:目标受众|汇报对象|面向|给|对象是)\s*[:：]?\s*(.+)$/);
  if (audienceMatch?.[1]) {
    heuristic.targetAudience = normalizeText(audienceMatch[1]);
  }

  if (shouldTreatAsSourceText(message)) {
    heuristic.sourceText = message.trim();
  }

  const aiParsed = await requestStructuredJson(
    collectTurnSchema,
    [
      "你是一个对话信息抽取器。",
      "你只负责从用户最新一条消息里抽取明确表达的信息。",
      "如果用户这条消息里包含大段素材正文、Markdown、方案内容或纯文字材料，请把它放入 sourceText。",
      "不要猜测没有明确说出的信息。",
      "只输出 JSON。"
    ].join("\n"),
    JSON.stringify({ message }, null, 2)
  );

  return {
    ...heuristic,
    ...aiParsed
  };
}

async function parseDirectoryTurn(message: string, pipeline: GenerationPipeline) {
  const normalized = normalizeText(message);
  const heuristic: z.infer<typeof directoryTurnSchema> = {
    action: "unknown",
    updates: []
  };

  if (/确认目录|目录确认|目录没问题|就这样|确认/.test(normalized) && !/修改|删除/.test(normalized)) {
    heuristic.action = "confirm";
  }

  for (const match of normalized.matchAll(/删除目录\s*(\d+)/g)) {
    heuristic.action = "revise";
    heuristic.updates.push({
      order: Number(match[1]),
      remove: true
    });
  }

  for (const match of normalized.matchAll(/目录\s*(\d+).*?标题(?:改为|为)\s*[:：]?\s*([^\n]+)/g)) {
    heuristic.action = "revise";
    heuristic.updates.push({
      order: Number(match[1]),
      title: normalizeText(match[2])
    });
  }

  for (const match of normalized.matchAll(/目录\s*(\d+).*?(?:简述|说明|描述)(?:改为|为)\s*[:：]?\s*([^\n]+)/g)) {
    heuristic.action = "revise";
    heuristic.updates.push({
      order: Number(match[1]),
      description: normalizeText(match[2])
    });
  }

  const aiParsed = await requestStructuredJson(
    directoryTurnSchema,
    [
      "你负责解析用户对 PPT 目录的确认或修改意图。",
      "如果用户明确确认目录，action 返回 confirm。",
      "如果用户要求修改或删除目录项，action 返回 revise，并按目录序号输出 updates。",
      "如果无法判断，action 返回 unknown。",
      "只输出 JSON。"
    ].join("\n"),
    JSON.stringify(
      {
        message,
        directory: pipeline.node3.directory.map((item) => ({
          order: item.order,
          title: item.fields.title.value,
          description: item.fields.description.value
        }))
      },
      null,
      2
    )
  );

  return {
    action: aiParsed?.action ?? heuristic.action,
    updates: aiParsed?.updates?.length ? aiParsed.updates : heuristic.updates
  };
}

async function parseSummaryTurn(message: string, pipeline: GenerationPipeline) {
  const normalized = normalizeText(message);
  const heuristic: z.infer<typeof summaryTurnSchema> = {
    action: "unknown",
    updates: []
  };

  if (/确认摘要|摘要确认|摘要没问题|可以生成|确认/.test(normalized) && !/修改|删除/.test(normalized)) {
    heuristic.action = "confirm";
  }

  for (const match of normalized.matchAll(/删除(?:第)?\s*(\d+)\s*页/g)) {
    heuristic.action = "revise";
    heuristic.updates.push({
      pageNumber: Number(match[1]),
      remove: true
    });
  }

  const fieldPatterns: Array<{
    field: "pageTitle" | "intro" | "regularTitle" | "description";
    pattern: RegExp;
  }> = [
    { field: "pageTitle", pattern: /第\s*(\d+)\s*页.*?(?:页标题|标题)(?:改为|为)\s*[:：]?\s*([^\n]+)/g },
    { field: "intro", pattern: /第\s*(\d+)\s*页.*?内容简介(?:改为|为)\s*[:：]?\s*([^\n]+)/g },
    { field: "regularTitle", pattern: /第\s*(\d+)\s*页.*?常规标题(?:改为|为)\s*[:：]?\s*([^\n]+)/g },
    { field: "description", pattern: /第\s*(\d+)\s*页.*?说明性文字(?:改为|为)\s*[:：]?\s*([^\n]+)/g }
  ];

  fieldPatterns.forEach(({ field, pattern }) => {
    for (const match of normalized.matchAll(pattern)) {
      heuristic.action = "revise";
      heuristic.updates.push({
        pageNumber: Number(match[1]),
        [field]: normalizeText(match[2])
      });
    }
  });

  const aiParsed = await requestStructuredJson(
    summaryTurnSchema,
    [
      "你负责解析用户对 PPT 详情摘要的确认或修改意图。",
      "如果用户明确确认摘要，action 返回 confirm。",
      "如果用户要求修改或删除页面，action 返回 revise，并按页码输出 updates。",
      "页码以当前列表中的 pageNumber 为准。",
      "只输出 JSON。"
    ].join("\n"),
    JSON.stringify(
      {
        message,
        slides: pipeline.node3.slides.map((slide) => ({
          pageNumber: slide.pageNumber,
          pageTitle: slide.fields.pageTitle.value,
          intro: slide.fields.intro.value,
          regularTitle: slide.fields.regularTitle.value,
          description: truncateText(slide.fields.description.value, 90)
        }))
      },
      null,
      2
    )
  );

  return {
    action: aiParsed?.action ?? heuristic.action,
    updates: aiParsed?.updates?.length ? aiParsed.updates : heuristic.updates
  };
}

function buildGeneratorInput(state: ChatSessionState): GeneratorInput {
  if (
    !state.input.totalPages ||
    !normalizeText(state.input.userName) ||
    !normalizeText(state.input.targetAudience)
  ) {
    throw new Error("当前信息还不完整，暂时不能生成 PPT。");
  }

  return {
    userName: state.input.userName,
    targetAudience: state.input.targetAudience,
    estimatedMinutes: state.input.estimatedMinutes,
    totalPages: state.input.totalPages
  };
}

function formatDirectoryMarkdown(pipeline: GenerationPipeline) {
  const lines = [
    "### 目录架构",
    ...pipeline.node3.directory.map(
      (item) =>
        `${item.order}. **${item.fields.title.value || "未命名目录"}**\n   - 简要说明：${item.fields.description.value || "待补充"}\n   - 覆盖页数：${item.pageCount} 页`
    ),
    "",
    "如果确认，请直接回复 `确认目录`。",
    "如果要修改，可以直接说：`目录 1 标题改为……`、`目录 2 简述改为……`、`删除目录 3`。"
  ];

  return lines.join("\n");
}

function formatSummaryMarkdown(pipeline: GenerationPipeline) {
  const lines = [
    "### 详情摘要",
    ...pipeline.node3.slides.map(
      (slide) =>
        `#### 第 ${slide.pageNumber} 页 · ${slide.layoutName}\n- 页标题：${slide.fields.pageTitle.value || "待补充"}\n- 内容简介：${slide.fields.intro.value || "待补充"}\n- 常规标题：${slide.fields.regularTitle.value || "待补充"}\n- 说明性文字：${truncateText(slide.fields.description.value || "待补充", 180)}`
    ),
    "",
    "如果确认，请回复 `确认摘要`。",
    "如果要修改，可以直接说：`第 2 页 标题改为……`、`第 3 页 内容简介改为……`、`删除第 4 页`。"
  ];

  return lines.join("\n");
}

function formatValidationIssues(document: AssemblyInstructionDocument) {
  if (!document.validation.hasOverflow) {
    return "";
  }

  return [
    "",
    "当前还有超出模板红线的字段：",
    ...document.validation.issues.map((issue) => `- ${issue}`)
  ].join("\n");
}

function applyDirectoryUpdates(
  pipeline: GenerationPipeline,
  updates: z.infer<typeof directoryTurnSchema>["updates"]
): GenerationPipeline {
  let node3 = pipeline.node3;

  updates.forEach((update) => {
    const target = node3.directory.find((item) => item.order === update.order);
    if (!target) {
      return;
    }

    if (update.remove) {
      node3 = {
        ...node3,
        directory: node3.directory.filter((item) => item.id !== target.id),
        slides: node3.slides.filter((slide) => slide.directoryId !== target.id)
      };
      node3 = validateAssemblyDocument(node3);
      return;
    }

    node3 = {
      ...node3,
      directory: node3.directory.map((item) =>
        item.id === target.id
          ? {
              ...item,
              fields: {
                ...item.fields,
                title: update.title
                  ? { ...item.fields.title, value: update.title }
                  : item.fields.title,
                description: update.description
                  ? { ...item.fields.description, value: update.description }
                  : item.fields.description
              }
            }
          : item
      ),
      slides: update.title
        ? node3.slides.map((slide) =>
            slide.directoryId === target.id ? { ...slide, directoryTitle: update.title ?? slide.directoryTitle } : slide
          )
        : node3.slides
    };
    node3 = validateAssemblyDocument(node3);
  });

  return {
    ...pipeline,
    node3
  };
}

function applySummaryUpdates(
  pipeline: GenerationPipeline,
  updates: z.infer<typeof summaryTurnSchema>["updates"]
): GenerationPipeline {
  let node3 = pipeline.node3;

  updates.forEach((update) => {
    const target = node3.slides.find((slide) => slide.pageNumber === update.pageNumber);
    if (!target) {
      return;
    }

    if (update.remove) {
      node3 = {
        ...node3,
        slides: node3.slides.filter((slide) => slide.id !== target.id)
      };
      node3 = validateAssemblyDocument(node3);
      return;
    }

    node3 = {
      ...node3,
      slides: node3.slides.map((slide) =>
        slide.id === target.id
          ? {
              ...slide,
              fields: {
                ...slide.fields,
                pageTitle: update.pageTitle
                  ? { ...slide.fields.pageTitle, value: update.pageTitle }
                  : slide.fields.pageTitle,
                intro: update.intro ? { ...slide.fields.intro, value: update.intro } : slide.fields.intro,
                regularTitle: update.regularTitle
                  ? { ...slide.fields.regularTitle, value: update.regularTitle }
                  : slide.fields.regularTitle,
                description: update.description
                  ? { ...slide.fields.description, value: update.description }
                  : slide.fields.description
              }
            }
          : slide
      )
    };
    node3 = validateAssemblyDocument(node3);
  });

  return {
    ...pipeline,
    node3
  };
}

function withConfidence(state: ChatSessionState): ChatSessionState {
  return {
    ...state,
    confidence: computeConfidence(state)
  };
}

function resolveSourceMaterial(
  existing: SourceMaterialDraft | null,
  file: File | null,
  sourceText: string | undefined
): SourceMaterialDraft | null {
  if (file) {
    return {
      kind: "file",
      name: file.name,
      preview: `已接收文件 ${file.name}`
    };
  }

  if (sourceText) {
    return {
      kind: "text",
      name: "聊天输入素材.txt",
      preview: truncateText(sourceText, 80),
      textContent: sourceText
    };
  }

  return existing;
}

export function createInitialChatState(): ChatSessionState {
  return initialState;
}

export async function handleChatTurn(params: {
  state: ChatSessionState | null;
  message: string;
  file: File | null;
}): Promise<ChatRouteResponse> {
  const currentState = params.state ? { ...params.state } : createInitialChatState();
  const message = params.message.trim();

  if (isResetIntent(message)) {
    return {
      state: createInitialChatState(),
      assistantMessages: ["好的，我们重新开始。先告诉我汇报人名称。"],
      action: "none"
    };
  }

  if (currentState.stage === "collect") {
    const collect = await parseCollectTurn(message);
    if (collect.reset) {
      return {
        state: createInitialChatState(),
        assistantMessages: ["好的，我们重新开始。先告诉我汇报人名称。"],
        action: "none"
      };
    }

    const nextState = withConfidence({
      ...currentState,
      input: {
        userName: collect.userName ? normalizeText(collect.userName) : currentState.input.userName,
        targetAudience: collect.targetAudience
          ? normalizeText(collect.targetAudience)
          : currentState.input.targetAudience,
        estimatedMinutes:
          typeof collect.estimatedMinutes === "number"
            ? collect.estimatedMinutes
            : currentState.input.estimatedMinutes,
        totalPages:
          typeof collect.totalPages === "number" ? collect.totalPages : currentState.input.totalPages
      },
      sourceMaterial: resolveSourceMaterial(currentState.sourceMaterial, params.file, collect.sourceText)
    });

    if (nextState.confidence >= 95 && nextState.sourceMaterial) {
      if (nextState.sourceMaterial.kind === "file" && !params.file) {
        return {
          state: nextState,
          assistantMessages: [
            "我已经拿到足够的信息了，但还需要你重新附带一次素材文件，才能继续生成目录架构。"
          ],
          action: "none"
        };
      }

      const pipeline = await buildGenerationPipelineFromSource(
        params.file ??
          ({
            name: nextState.sourceMaterial.name,
            text: nextState.sourceMaterial.textContent ?? "",
            type: nextState.sourceMaterial.name.endsWith(".md") ? "md" : "text"
          } as const),
        buildGeneratorInput(nextState)
      );

      const stateWithPipeline: ChatSessionState = {
        ...nextState,
        stage: "directory",
        pipeline
      };

      return {
        state: stateWithPipeline,
        assistantMessages: [
          "我已经完成目录规划，下面请你直接在对话里确认或修改。",
          formatDirectoryMarkdown(pipeline)
        ],
        action: "none"
      };
    }

    return {
      state: nextState,
      assistantMessages: [
        buildMissingPrompt(nextState)
      ],
      action: "none"
    };
  }

  if (!currentState.pipeline) {
    return {
      state: createInitialChatState(),
      assistantMessages: ["会话状态丢失了，我们重新开始。先告诉我汇报人名称。"],
      action: "none"
    };
  }

  if (currentState.stage === "directory") {
    const parsed = await parseDirectoryTurn(message, currentState.pipeline);

    if (parsed.action === "confirm") {
      const nextState: ChatSessionState = {
        ...currentState,
        stage: "summary"
      };

      return {
        state: nextState,
        assistantMessages: [
          "目录已确认。下面是详情摘要，请继续在对话里确认或修改。",
          formatSummaryMarkdown(currentState.pipeline)
        ],
        action: "none"
      };
    }

    if (parsed.updates.length > 0) {
      const nextPipeline = applyDirectoryUpdates(currentState.pipeline, parsed.updates);
      const nextState: ChatSessionState = {
        ...currentState,
        pipeline: nextPipeline
      };

      return {
        state: nextState,
        assistantMessages: [
          "目录已按你的意见更新。请继续确认。",
          formatDirectoryMarkdown(nextPipeline)
        ],
        action: "none"
      };
    }

    return {
      state: currentState,
      assistantMessages: [
        "我还没有识别到明确的目录确认或修改指令。你可以直接回复 `确认目录`，或者说 `目录 1 标题改为……`。"
      ],
      action: "none"
    };
  }

  if (currentState.stage === "summary") {
    const parsed = await parseSummaryTurn(message, currentState.pipeline);

    if (parsed.action === "confirm") {
      const validated = validateAssemblyDocument(currentState.pipeline.node3);
      const nextState: ChatSessionState = {
        ...currentState,
        stage: validated.validation.hasOverflow ? "summary" : "ready",
        pipeline: {
          ...currentState.pipeline,
          node3: validated
        }
      };

      return {
        state: nextState,
        assistantMessages: validated.validation.hasOverflow
          ? [
              "摘要已经锁定，但当前仍有字段超出模板红线，请先继续修改。",
              formatSummaryMarkdown(nextState.pipeline!),
              formatValidationIssues(validated)
            ]
          : [
              "详情摘要已确认，所有字段都在模板红线内。现在你可以直接回复 `生成PPT`。"
            ],
        action: "none"
      };
    }

    if (parsed.updates.length > 0) {
      const nextPipeline = applySummaryUpdates(currentState.pipeline, parsed.updates);
      const validated = validateAssemblyDocument(nextPipeline.node3);
      const nextState: ChatSessionState = {
        ...currentState,
        pipeline: {
          ...nextPipeline,
          node3: validated
        }
      };

      return {
        state: nextState,
        assistantMessages: [
          "详情摘要已按你的意见更新。请继续确认。",
          formatSummaryMarkdown(nextState.pipeline!),
          formatValidationIssues(validated)
        ].filter(Boolean),
        action: "none"
      };
    }

    return {
      state: currentState,
      assistantMessages: [
        "我还没有识别到明确的摘要确认或修改指令。你可以回复 `确认摘要`，或者说 `第 2 页 标题改为……`。"
      ],
      action: "none"
    };
  }

  if (isExportIntent(message)) {
    return {
      state: currentState,
      assistantMessages: ["收到，我现在开始生成 PPT 终稿。"],
      action: "export"
    };
  }

  const parsed = await parseSummaryTurn(message, currentState.pipeline);
  if (parsed.updates.length > 0) {
    const nextPipeline = applySummaryUpdates(currentState.pipeline, parsed.updates);
    const validated = validateAssemblyDocument(nextPipeline.node3);
    const nextState: ChatSessionState = {
      ...currentState,
      pipeline: {
        ...nextPipeline,
        node3: validated
      }
    };

    return {
      state: nextState,
      assistantMessages: [
        "已按你的意见更新终稿前的摘要内容。确认无误后，回复 `生成PPT` 即可。",
        formatSummaryMarkdown(nextState.pipeline!),
        formatValidationIssues(validated)
      ].filter(Boolean),
      action: "none"
    };
  }

  return {
    state: currentState,
    assistantMessages: ["当前已经进入终稿阶段。你可以继续修改页摘要，或者直接回复 `生成PPT`。"],
    action: "none"
  };
}
