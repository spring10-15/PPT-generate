"use client";

import Image from "next/image";
import { useSearchParams } from "next/navigation";
import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import type { ChangeEvent, FormEvent } from "react";
import type {
  ChatDirective,
  ChatRouteResponse,
  ChatSessionState,
  GenerationPipeline
} from "@/lib/types";

type ChatRole = "assistant" | "user";

type ChatAttachment = {
  name: string;
  extension: string;
  sizeLabel: string;
};

type ChatMessage = {
  id: string;
  role: ChatRole;
  content: string;
  createdAt: string;
  attachment?: ChatAttachment;
};

type DirectoryDraftItem = {
  id: string;
  order: number;
  title: string;
  titleMaxChars: number;
  description: string;
  descriptionMaxChars: number;
};

type SlotCardDraft = {
  id: string;
  order: number;
  title: string;
  titleMaxChars: number;
  body: string;
  bodyMaxChars: number;
};

type SlotSectionDraft = {
  id: string;
  order: number;
  label: string;
  labelMaxChars: number;
  cards: SlotCardDraft[];
};

type SummaryDraftSlide = {
  id: string;
  pageNumber: number;
  layoutName: string;
  pageTitle: string;
  pageTitleMaxChars: number;
  intro: string;
  introMaxChars: number;
  sections: SlotSectionDraft[];
};

type PreviewPreset =
  | "initial"
  | "filled"
  | "textmaterial"
  | "thinking"
  | "directory"
  | "directory-editing"
  | "summary"
  | "summary-overflow"
  | "ready"
  | "success";

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

function makeMessage(role: ChatRole, content: string, attachment?: ChatAttachment): ChatMessage {
  return {
    id: `msg-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    role,
    content,
    createdAt: new Date().toISOString(),
    attachment
  };
}

function createInitialMessages(): ChatMessage[] {
  return [
    makeMessage(
      "assistant",
      "欢迎使用 PPT 生成助手。\n\n我们会通过对话完成信息收集、目录确认、槽位级摘要确认和终稿导出。\n\n先告诉我汇报人名称。"
    )
  ];
}

function extractErrorMessage(error: unknown) {
  if (error instanceof Error) {
    return error.message;
  }

  return "处理失败，请稍后重试。";
}

function formatFileSize(size: number) {
  if (size >= 1024 * 1024) {
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }

  if (size >= 1024) {
    return `${Math.round(size / 1024)} KB`;
  }

  return `${size} B`;
}

function buildAttachment(file: File | null): ChatAttachment | undefined {
  if (!file) {
    return undefined;
  }

  const extension = file.name.includes(".") ? file.name.split(".").pop()?.toUpperCase() ?? "FILE" : "FILE";
  return {
    name: file.name,
    extension,
    sizeLabel: formatFileSize(file.size)
  };
}

function shouldShowTimeDivider(previous: ChatMessage | undefined, current: ChatMessage) {
  if (!previous) {
    return true;
  }

  const previousTime = new Date(previous.createdAt).getTime();
  const currentTime = new Date(current.createdAt).getTime();
  return currentTime - previousTime >= 5 * 60 * 1000;
}

function formatTimeDivider(isoTime: string) {
  return new Intl.DateTimeFormat("zh-CN", {
    hour: "2-digit",
    minute: "2-digit"
  }).format(new Date(isoTime));
}

function countChars(value: string) {
  return value.trim().length;
}

function buildDirectoryDraft(pipeline: GenerationPipeline | null): DirectoryDraftItem[] {
  if (!pipeline) {
    return [];
  }

  return pipeline.node3.directory.map((item) => ({
    id: item.id,
    order: item.order,
    title: item.fields.title.value,
    titleMaxChars: item.fields.title.maxChars,
    description: item.fields.description.value,
    descriptionMaxChars: item.fields.description.maxChars
  }));
}

function buildSummaryDraft(pipeline: GenerationPipeline | null): SummaryDraftSlide[] {
  if (!pipeline) {
    return [];
  }

  return pipeline.node3.slides.map((slide) => ({
    id: slide.id,
    pageNumber: slide.pageNumber,
    layoutName: slide.layoutName,
    pageTitle: slide.fields.pageTitle.value,
    pageTitleMaxChars: slide.fields.pageTitle.maxChars,
    intro: slide.fields.intro.value,
    introMaxChars: slide.fields.intro.maxChars,
    sections: slide.slotContent.sections.map((section) => ({
      id: section.id,
      order: section.order,
      label: section.label,
      labelMaxChars: section.labelMaxChars,
      cards: section.cards.map((card) => ({
        id: card.id,
        order: card.order,
        title: card.title,
        titleMaxChars: card.titleMaxChars,
        body: card.body,
        bodyMaxChars: card.bodyMaxChars
      }))
    }))
  }));
}

function createPreviewDirectoryDraft(): DirectoryDraftItem[] {
  return [
    {
      id: "preview-directory-1",
      order: 1,
      title: "技术方案",
      titleMaxChars: 5,
      description: "系统总体方案",
      descriptionMaxChars: 10
    },
    {
      id: "preview-directory-2",
      order: 2,
      title: "应用场景",
      titleMaxChars: 5,
      description: "办公协同场景",
      descriptionMaxChars: 10
    },
    {
      id: "preview-directory-3",
      order: 3,
      title: "落地价值",
      titleMaxChars: 5,
      description: "项目业务价值",
      descriptionMaxChars: 10
    }
  ];
}

function createPreviewSummaryDraft(): SummaryDraftSlide[] {
  return [
    {
      id: "preview-slide-1",
      pageNumber: 1,
      layoutName: "总体",
      pageTitle: "技术方案概览",
      pageTitleMaxChars: 12,
      intro: "统一梳理文档解析、目录生成与模板装配流程。",
      introMaxChars: 30,
      sections: [
        {
          id: "preview-slide-1-section-1",
          order: 1,
          label: "核心能力",
          labelMaxChars: 10,
          cards: [
            {
              id: "preview-slide-1-card-1",
              order: 1,
              title: "文档解析",
              titleMaxChars: 15,
              body: "自动抽取标题层级、关键段落、表格与图片素材。",
              bodyMaxChars: 60
            },
            {
              id: "preview-slide-1-card-2",
              order: 2,
              title: "模板匹配",
              titleMaxChars: 15,
              body: "根据版式槽位与字数边界，为内容选择合适模板。",
              bodyMaxChars: 60
            },
            {
              id: "preview-slide-1-card-3",
              order: 3,
              title: "终稿导出",
              titleMaxChars: 15,
              body: "保留模板样式和原生元素，导出可编辑 PPTX。",
              bodyMaxChars: 60
            }
          ]
        }
      ]
    },
    {
      id: "preview-slide-2",
      pageNumber: 2,
      layoutName: "左右结构",
      pageTitle: "应用场景展开",
      pageTitleMaxChars: 12,
      intro: "围绕协同办公与业务提效场景展开页级说明。",
      introMaxChars: 30,
      sections: [
        {
          id: "preview-slide-2-section-1",
          order: 1,
          label: "协同办公",
          labelMaxChars: 10,
          cards: [
            {
              id: "preview-slide-2-card-1",
              order: 1,
              title: "资料沉淀",
              titleMaxChars: 15,
              body: "将方案、纪要和规则统一沉淀，支撑后续知识复用。",
              bodyMaxChars: 60
            },
            {
              id: "preview-slide-2-card-2",
              order: 2,
              title: "执行协同",
              titleMaxChars: 15,
              body: "围绕任务分工、过程同步和结果复盘形成闭环。",
              bodyMaxChars: 60
            }
          ]
        }
      ]
    }
  ];
}

function createPreviewSnapshot(preset: PreviewPreset): {
  messages: ChatMessage[];
  chatState: ChatSessionState;
  status: string;
  directoryDraft: DirectoryDraftItem[];
  summaryDraft: SummaryDraftSlide[];
  showThinking: boolean;
  validationIssues: string[];
} {
  const sourceAttachment: ChatAttachment = {
    name: "企业协同方案说明.docx",
    extension: "DOCX",
    sizeLabel: "428 KB"
  };

  const baseMessages: ChatMessage[] = [
    makeMessage(
      "assistant",
      "欢迎使用 PPT 生成助手。\n\n我们会通过对话完成信息收集、目录确认、槽位级摘要确认和终稿导出。\n\n先告诉我汇报人名称。"
    ),
    makeMessage("user", "汇报人是张明，这次汇报控制在 12 页，演讲时长 15 分钟。"),
    makeMessage("assistant", "好的，我已经记录汇报人、页数和演讲时长。\n\n请继续上传素材文件，或者直接粘贴文字内容。"),
    makeMessage("user", "我先上传方案材料。", sourceAttachment)
  ];

  const baseState: ChatSessionState = {
    stage: "collect",
    confidence: 75,
    input: {
      userName: "张明",
      targetAudience: "",
      estimatedMinutes: 15,
      totalPages: 12
    },
    sourceMaterial: {
      kind: "file",
      name: "企业协同方案说明.docx",
      preview: "已上传方案材料"
    },
    pipeline: null
  };

  if (preset === "initial") {
    return {
      messages: createInitialMessages(),
      chatState: initialState,
      status: "可直接输入文字，或点击左下角 + 上传素材。",
      directoryDraft: [] as DirectoryDraftItem[],
      summaryDraft: [] as SummaryDraftSlide[],
      showThinking: false,
      validationIssues: []
    };
  }

  if (preset === "filled") {
    return {
      messages: baseMessages,
      chatState: baseState,
      status: "已记录汇报人、时长、页数，并附带素材文件。",
      directoryDraft: [] as DirectoryDraftItem[],
      summaryDraft: [] as SummaryDraftSlide[],
      showThinking: false,
      validationIssues: []
    };
  }

  if (preset === "textmaterial") {
    return {
      messages: [
        makeMessage(
          "assistant",
          "欢迎使用 PPT 生成助手。\n\n我们会通过对话完成信息收集、目录确认、槽位级摘要确认和终稿导出。\n\n先告诉我汇报人名称。"
        ),
        makeMessage("user", "汇报人是张明，这次给业务团队汇报，时长 15 分钟，想做 12 页。"),
        makeMessage(
          "assistant",
          "好的，我已经记录了汇报人、目标受众、时长和页数。\n\n如果方便，可以直接把文字素材粘贴到聊天框里。"
        ),
        makeMessage(
          "user",
          "以下是方案摘要：知识平台负责沉淀业务资料与知识资产，智能助手负责问答、任务分发与结果回收，两者通过文档、任务和知识库建立统一协同闭环。"
        )
      ],
      chatState: {
        stage: "collect",
        confidence: 95,
        input: {
          userName: "张明",
          targetAudience: "业务团队",
          estimatedMinutes: 15,
          totalPages: 12
        },
        sourceMaterial: {
          kind: "text",
          name: "聊天输入素材",
          preview: "已粘贴一段方案摘要"
        },
        pipeline: null
      },
      status: "已收集到完整信息，准备生成目录。",
      directoryDraft: [] as DirectoryDraftItem[],
      summaryDraft: [] as SummaryDraftSlide[],
      showThinking: false,
      validationIssues: []
    };
  }

  if (preset === "thinking") {
    return {
      messages: baseMessages,
      chatState: {
        ...baseState,
        confidence: 95
      },
      status: "正在解析素材并生成目录。",
      directoryDraft: [] as DirectoryDraftItem[],
      summaryDraft: [] as SummaryDraftSlide[],
      showThinking: true,
      validationIssues: []
    };
  }

  if (preset === "directory") {
    return {
      messages: [
        ...baseMessages,
        makeMessage(
          "assistant",
          "我已经结合素材梳理出目录，请直接在下方卡片里调整标题和简要说明。"
        )
      ],
      chatState: {
        ...baseState,
        stage: "directory",
        confidence: 100
      },
      status: "请先确认目录结构。",
      directoryDraft: createPreviewDirectoryDraft(),
      summaryDraft: [] as SummaryDraftSlide[],
      showThinking: false,
      validationIssues: []
    };
  }

  if (preset === "directory-editing") {
    const draft = createPreviewDirectoryDraft();
    draft[1] = {
      ...draft[1],
      title: "核心场景",
      description: "覆盖办公与行业"
    };

    return {
      messages: [
        ...baseMessages,
        makeMessage(
          "assistant",
          "我已经结合素材梳理出目录，请直接在下方卡片里调整标题和简要说明。"
        ),
        makeMessage("user", "第二个目录我想换成更聚焦的表述。")
      ],
      chatState: {
        ...baseState,
        stage: "directory",
        confidence: 100
      },
      status: "目录已编辑，可继续保存或确认。",
      directoryDraft: draft,
      summaryDraft: [] as SummaryDraftSlide[],
      showThinking: false,
      validationIssues: []
    };
  }

  if (preset === "summary") {
    return {
      messages: [
        ...baseMessages,
        makeMessage("assistant", "目录已确认。我已按目录结构展开详情摘要，请继续确认每一页的槽位内容。")
      ],
      chatState: {
        ...baseState,
        stage: "summary",
        confidence: 100
      },
      status: "请确认每一页的槽位摘要。",
      directoryDraft: createPreviewDirectoryDraft(),
      summaryDraft: createPreviewSummaryDraft(),
      showThinking: false,
      validationIssues: []
    };
  }

  if (preset === "summary-overflow") {
    const draft = createPreviewSummaryDraft();
    draft[0] = {
      ...draft[0],
      intro: "统一梳理文档解析、目录生成、模板装配与多角色协同确认流程，确保最终成稿既符合固定模板也保持业务语义完整。",
      sections: draft[0].sections.map((section, sectionIndex) =>
        sectionIndex === 0
          ? {
              ...section,
              cards: section.cards.map((card, cardIndex) =>
                cardIndex === 0
                  ? {
                      ...card,
                      body: "自动抽取标题层级、关键段落、表格与图片素材，并结合语义摘要对文档中的商业逻辑、业务结构和关键数据进行稳定重组。"
                    }
                  : card
              )
            }
          : section
      )
    };

    return {
      messages: [
        ...baseMessages,
        makeMessage("assistant", "目录已确认。我已按目录结构展开详情摘要，请继续确认每一页的槽位内容。"),
        makeMessage("user", "第一页我想再补充完整一点。")
      ],
      chatState: {
        ...baseState,
        stage: "summary",
        confidence: 100
      },
      status: "当前有超出模板红线的内容，请先收敛文案。",
      directoryDraft: createPreviewDirectoryDraft(),
      summaryDraft: draft,
      showThinking: false,
      validationIssues: ["页码 1 内容简介超出模板上限", "页码 1 条目 1.1 说明性文字超出模板上限"]
    };
  }

  if (preset === "ready") {
    return {
      messages: [
        ...baseMessages,
        makeMessage("assistant", "目录和摘要都已确认，可以直接生成 PPT 终稿。")
      ],
      chatState: {
        ...baseState,
        stage: "ready",
        confidence: 100
      },
      status: "摘要已确认，可以生成 PPT。",
      directoryDraft: createPreviewDirectoryDraft(),
      summaryDraft: createPreviewSummaryDraft(),
      showThinking: false,
      validationIssues: []
    };
  }

  return {
    messages: [
      ...baseMessages,
      makeMessage("assistant", "目录和摘要都已确认，可以直接生成 PPT 终稿。"),
      makeMessage("user", "生成PPT"),
      makeMessage("assistant", "PPT 终稿已生成，下载已经开始。")
    ],
    chatState: {
      ...baseState,
      stage: "ready",
      confidence: 100
    },
    status: "PPT 终稿已生成并开始下载。",
    directoryDraft: createPreviewDirectoryDraft(),
    summaryDraft: createPreviewSummaryDraft(),
    showThinking: false,
    validationIssues: []
  };
}

function buildStatusByStage(response: ChatRouteResponse) {
  if (response.action === "export") {
    return "准备导出 PPT 终稿。";
  }

  switch (response.state.stage) {
    case "directory":
      return "请在聊天内确认或修改目录。";
    case "summary":
      return "请在聊天内确认或修改槽位摘要。";
    case "ready":
      return "摘要已就绪，可以直接生成 PPT。";
    default:
      return "继续告诉我还缺少的信息。";
  }
}

function buildThinkingLabel(state: ChatSessionState, isExporting: boolean) {
  if (isExporting || state.stage === "ready") {
    return "正在生成 PPT 终稿";
  }

  switch (state.stage) {
    case "directory":
      return "正在整理目录回复";
    case "summary":
      return "正在整理详情摘要";
    default:
      return "正在思考回复";
  }
}

function buildDirectoryDirective(
  pipeline: GenerationPipeline | null,
  draft: DirectoryDraftItem[]
): Extract<ChatDirective, { type: "directory-update" }> | null {
  if (!pipeline) {
    return null;
  }

  const updates: Array<{
    order: number;
    title?: string;
    description?: string;
    remove?: boolean;
  }> = [];

  pipeline.node3.directory.forEach((item) => {
    const target = draft.find((draftItem) => draftItem.id === item.id);
    if (!target) {
      updates.push({
        order: item.order,
        remove: true
      });
      return;
    }

    const titleChanged = target.title.trim() !== item.fields.title.value.trim();
    const descriptionChanged = target.description.trim() !== item.fields.description.value.trim();
    if (titleChanged || descriptionChanged) {
      updates.push({
        order: item.order,
        title: titleChanged ? target.title : undefined,
        description: descriptionChanged ? target.description : undefined
      });
    }
  });

  return updates.length > 0
    ? {
        type: "directory-update",
        updates
      }
    : null;
}

function buildSummaryDirective(
  pipeline: GenerationPipeline | null,
  draft: SummaryDraftSlide[]
): Extract<ChatDirective, { type: "summary-update" }> | null {
  if (!pipeline) {
    return null;
  }

  const updates: Array<{
    pageNumber: number;
    remove?: boolean;
    pageTitle?: string;
    intro?: string;
    sectionOrder?: number;
    sectionLabel?: string;
    cardOrder?: number;
    cardTitle?: string;
    cardBody?: string;
  }> = [];

  pipeline.node3.slides.forEach((slide) => {
    const target = draft.find((draftSlide) => draftSlide.id === slide.id);
    if (!target) {
      updates.push({
        pageNumber: slide.pageNumber,
        remove: true
      });
      return;
    }

    if (target.pageTitle.trim() !== slide.fields.pageTitle.value.trim()) {
      updates.push({
        pageNumber: slide.pageNumber,
        pageTitle: target.pageTitle
      });
    }

    if (target.intro.trim() !== slide.fields.intro.value.trim()) {
      updates.push({
        pageNumber: slide.pageNumber,
        intro: target.intro
      });
    }

    slide.slotContent.sections.forEach((section) => {
      const targetSection = target.sections.find((item) => item.id === section.id);
      if (!targetSection) {
        return;
      }

      if (targetSection.label.trim() !== section.label.trim()) {
        updates.push({
          pageNumber: slide.pageNumber,
          sectionOrder: section.order,
          sectionLabel: targetSection.label
        });
      }

      section.cards.forEach((card) => {
        const targetCard = targetSection.cards.find((item) => item.id === card.id);
        if (!targetCard) {
          return;
        }

        if (targetCard.title.trim() !== card.title.trim()) {
          updates.push({
            pageNumber: slide.pageNumber,
            sectionOrder: section.order,
            cardOrder: card.order,
            cardTitle: targetCard.title
          });
        }

        if (targetCard.body.trim() !== card.body.trim()) {
          updates.push({
            pageNumber: slide.pageNumber,
            sectionOrder: section.order,
            cardOrder: card.order,
            cardBody: targetCard.body
          });
        }
      });
    });
  });

  return updates.length > 0
    ? {
        type: "summary-update",
        updates
      }
    : null;
}

type DirectoryPanelProps = {
  draft: DirectoryDraftItem[];
  onChange: (updater: (current: DirectoryDraftItem[]) => DirectoryDraftItem[]) => void;
  onSave: () => void;
  onConfirm: () => void;
  disabled: boolean;
};

function DirectoryReviewPanel(props: DirectoryPanelProps) {
  return (
    <div className="review-stack">
      <div className="review-head">
        <div>
          <strong>目录确认</strong>
          <p>直接在聊天里修改标题和简要说明，然后保存。</p>
        </div>
        <div className="review-actions">
          <button type="button" className="mini-btn secondary" onClick={props.onSave} disabled={props.disabled}>
            暂存
          </button>
          <button type="button" className="mini-btn primary" onClick={props.onConfirm} disabled={props.disabled}>
            提交
          </button>
        </div>
      </div>

      <div className="review-scroll directory-review-scroll smooth-scroll">
        {props.draft.map((item) => (
          <article key={item.id} className="review-card">
            <div className="review-card-head">
              <span className="review-index">目录 {item.order}</span>
              <button
                type="button"
                className="review-remove"
                onClick={() =>
                  props.onChange((current) => current.filter((currentItem) => currentItem.id !== item.id))
                }
                disabled={props.disabled}
              >
                删除
              </button>
            </div>

            <label className={`slot-field ${countChars(item.title) > item.titleMaxChars ? "is-overflow" : ""}`}>
              <span>目录标题</span>
              <input
                value={item.title}
                onChange={(event) =>
                  props.onChange((current) =>
                    current.map((currentItem) =>
                      currentItem.id === item.id ? { ...currentItem, title: event.target.value } : currentItem
                    )
                  )
                }
                disabled={props.disabled}
              />
              <em>
                {countChars(item.title)}/{item.titleMaxChars}
              </em>
            </label>

            <label
              className={`slot-field ${countChars(item.description) > item.descriptionMaxChars ? "is-overflow" : ""}`}
            >
              <span>简要说明</span>
              <textarea
                value={item.description}
                onChange={(event) =>
                  props.onChange((current) =>
                    current.map((currentItem) =>
                      currentItem.id === item.id
                        ? { ...currentItem, description: event.target.value }
                        : currentItem
                    )
                  )
                }
                disabled={props.disabled}
              />
              <em>
                {countChars(item.description)}/{item.descriptionMaxChars}
              </em>
            </label>
          </article>
        ))}
      </div>
    </div>
  );
}

type SummaryPanelProps = {
  draft: SummaryDraftSlide[];
  validationIssues: string[];
  onChange: (updater: (current: SummaryDraftSlide[]) => SummaryDraftSlide[]) => void;
  onSave: () => void;
  onConfirm: () => void;
  onGenerate: () => void;
  disabled: boolean;
  isReady: boolean;
};

function SummaryReviewPanel(props: SummaryPanelProps) {
  return (
    <div className="review-stack">
      <div className="review-head">
        <div>
          <strong>槽位级摘要确认</strong>
          <p>页标题、内容简介、分区标题、小标题和说明性文字都可以直接编辑。</p>
        </div>
        <div className="review-actions">
          <button type="button" className="mini-btn secondary" onClick={props.onSave} disabled={props.disabled}>
            暂存
          </button>
          <button type="button" className="mini-btn secondary" onClick={props.onConfirm} disabled={props.disabled}>
            提交
          </button>
          {props.isReady ? (
            <button type="button" className="mini-btn primary" onClick={props.onGenerate} disabled={props.disabled}>
              生成 PPT
            </button>
          ) : null}
        </div>
      </div>

      {props.validationIssues.length > 0 ? (
        <div className="validation-box">
          <strong>当前还有红线问题</strong>
          <ul>
            {props.validationIssues.map((issue) => (
              <li key={issue}>{issue}</li>
            ))}
          </ul>
        </div>
      ) : null}

      <div className="review-scroll summary-review-scroll smooth-scroll">
        {props.draft.map((slide) => (
          <article key={slide.id} className="review-card summary-card">
            <div className="review-card-head">
              <div>
                <span className="review-index">页码 {slide.pageNumber}</span>
                <strong className="review-layout">{slide.layoutName}</strong>
              </div>
              <button
                type="button"
                className="review-remove"
                onClick={() =>
                  props.onChange((current) => current.filter((currentSlide) => currentSlide.id !== slide.id))
                }
                disabled={props.disabled}
              >
                删除此页
              </button>
            </div>

            <div className="slot-grid slot-grid-top">
              <label className={`slot-field ${countChars(slide.pageTitle) > slide.pageTitleMaxChars ? "is-overflow" : ""}`}>
                <span>页标题</span>
                <input
                  value={slide.pageTitle}
                  onChange={(event) =>
                    props.onChange((current) =>
                      current.map((currentSlide) =>
                        currentSlide.id === slide.id
                          ? { ...currentSlide, pageTitle: event.target.value }
                          : currentSlide
                      )
                    )
                  }
                  disabled={props.disabled}
                />
                <em>
                  {countChars(slide.pageTitle)}/{slide.pageTitleMaxChars}
                </em>
              </label>

              <label className={`slot-field ${countChars(slide.intro) > slide.introMaxChars ? "is-overflow" : ""}`}>
                <span>内容简介</span>
                <textarea
                  value={slide.intro}
                  onChange={(event) =>
                    props.onChange((current) =>
                      current.map((currentSlide) =>
                        currentSlide.id === slide.id ? { ...currentSlide, intro: event.target.value } : currentSlide
                      )
                    )
                  }
                  disabled={props.disabled}
                />
                <em>
                  {countChars(slide.intro)}/{slide.introMaxChars}
                </em>
              </label>
            </div>

            <div className="section-stack">
              {slide.sections.map((section) => (
                <section key={section.id} className="section-card">
                  <label
                    className={`slot-field ${countChars(section.label) > section.labelMaxChars ? "is-overflow" : ""}`}
                  >
                    <span>分区 {section.order} 标题</span>
                    <input
                      value={section.label}
                      onChange={(event) =>
                        props.onChange((current) =>
                          current.map((currentSlide) =>
                            currentSlide.id === slide.id
                              ? {
                                  ...currentSlide,
                                  sections: currentSlide.sections.map((currentSection) =>
                                    currentSection.id === section.id
                                      ? { ...currentSection, label: event.target.value }
                                      : currentSection
                                  )
                                }
                              : currentSlide
                          )
                        )
                      }
                      disabled={props.disabled}
                    />
                    <em>
                      {countChars(section.label)}/{section.labelMaxChars}
                    </em>
                  </label>

                  <div className="card-grid">
                    {section.cards.map((card) => (
                      <article key={card.id} className="slot-card">
                        <div className="slot-card-head">条目 {section.order}.{card.order}</div>
                        <label
                          className={`slot-field ${countChars(card.title) > card.titleMaxChars ? "is-overflow" : ""}`}
                        >
                          <span>小标题</span>
                          <input
                            value={card.title}
                            onChange={(event) =>
                              props.onChange((current) =>
                                current.map((currentSlide) =>
                                  currentSlide.id === slide.id
                                    ? {
                                        ...currentSlide,
                                        sections: currentSlide.sections.map((currentSection) =>
                                          currentSection.id === section.id
                                            ? {
                                                ...currentSection,
                                                cards: currentSection.cards.map((currentCard) =>
                                                  currentCard.id === card.id
                                                    ? { ...currentCard, title: event.target.value }
                                                    : currentCard
                                                )
                                              }
                                            : currentSection
                                        )
                                      }
                                    : currentSlide
                                )
                              )
                            }
                            disabled={props.disabled}
                          />
                          <em>
                            {countChars(card.title)}/{card.titleMaxChars}
                          </em>
                        </label>

                        <label
                          className={`slot-field ${countChars(card.body) > card.bodyMaxChars ? "is-overflow" : ""}`}
                        >
                          <span>说明性文字</span>
                          <textarea
                            value={card.body}
                            onChange={(event) =>
                              props.onChange((current) =>
                                current.map((currentSlide) =>
                                  currentSlide.id === slide.id
                                    ? {
                                        ...currentSlide,
                                        sections: currentSlide.sections.map((currentSection) =>
                                          currentSection.id === section.id
                                            ? {
                                                ...currentSection,
                                                cards: currentSection.cards.map((currentCard) =>
                                                  currentCard.id === card.id
                                                    ? { ...currentCard, body: event.target.value }
                                                    : currentCard
                                                )
                                              }
                                            : currentSection
                                        )
                                      }
                                    : currentSlide
                                )
                              )
                            }
                            disabled={props.disabled}
                          />
                          <em>
                            {countChars(card.body)}/{card.bodyMaxChars}
                          </em>
                        </label>
                      </article>
                    ))}
                  </div>
                </section>
              ))}
            </div>
          </article>
        ))}
      </div>
    </div>
  );
}

export default function HomePage() {
  const searchParams = useSearchParams();
  const previewPreset = useMemo(() => {
    const value = searchParams.get("preset");
    if (
      value === "initial" ||
      value === "filled" ||
      value === "textmaterial" ||
      value === "thinking" ||
      value === "directory" ||
      value === "directory-editing" ||
      value === "summary" ||
      value === "summary-overflow" ||
      value === "success" ||
      value === "ready"
    ) {
      return value as PreviewPreset;
    }

    return null;
  }, [searchParams]);
  const previewSnapshot = useMemo(
    () => (previewPreset ? createPreviewSnapshot(previewPreset) : null),
    [previewPreset]
  );
  const isPreviewMode = Boolean(previewSnapshot);
  const [messages, setMessages] = useState<ChatMessage[]>(() => createInitialMessages());
  const [chatState, setChatState] = useState<ChatSessionState>(initialState);
  const [composer, setComposer] = useState("");
  const [attachedFile, setAttachedFile] = useState<File | null>(null);
  const [status, setStatus] = useState("可直接输入文字，或点击左下角 + 上传素材。");
  const [isSending, setIsSending] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [directoryDraft, setDirectoryDraft] = useState<DirectoryDraftItem[]>([]);
  const [summaryDraft, setSummaryDraft] = useState<SummaryDraftSlide[]>([]);
  const [previewValidationIssues, setPreviewValidationIssues] = useState<string[]>([]);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const threadRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!previewSnapshot) {
      return;
    }

    setMessages(previewSnapshot.messages);
    setChatState(previewSnapshot.chatState);
    setComposer("");
    setAttachedFile(null);
    setStatus(previewSnapshot.status);
    setDirectoryDraft(previewSnapshot.directoryDraft);
    setSummaryDraft(previewSnapshot.summaryDraft);
    setPreviewValidationIssues(previewSnapshot.validationIssues);
  }, [previewSnapshot]);

  useEffect(() => {
    if (isPreviewMode) {
      return;
    }

    setDirectoryDraft(buildDirectoryDraft(chatState.pipeline));
    setSummaryDraft(buildSummaryDraft(chatState.pipeline));
    setPreviewValidationIssues([]);
  }, [chatState.pipeline, isPreviewMode]);

  useEffect(() => {
    const node = threadRef.current;
    if (!node) {
      return;
    }

    node.scrollTo({
      top: node.scrollHeight,
      behavior: "smooth"
    });
  }, [messages, chatState.stage, directoryDraft.length, summaryDraft.length]);

  const canSend = useMemo(() => Boolean(composer.trim()) || Boolean(attachedFile), [attachedFile, composer]);

  const stageLabel = useMemo(() => {
    switch (chatState.stage) {
      case "directory":
        return "目录确认";
      case "summary":
        return "槽位确认";
      case "ready":
        return "终稿导出";
      default:
        return "信息收集";
    }
  }, [chatState.stage]);
  const thinkingLabel = useMemo(
    () => buildThinkingLabel(chatState, isExporting),
    [chatState, isExporting]
  );

  const handleReset = () => {
    if (isPreviewMode) {
      return;
    }

    setMessages(createInitialMessages());
    setChatState(initialState);
    setComposer("");
    setAttachedFile(null);
    setDirectoryDraft([]);
    setSummaryDraft([]);
    setPreviewValidationIssues([]);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setStatus("可直接输入文字，或点击左下角 + 上传素材。");
  };

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (isPreviewMode) {
      return;
    }

    const nextFile = event.target.files?.[0] ?? null;
    setAttachedFile(nextFile);
    if (nextFile) {
      setStatus(`已附带素材：${nextFile.name}`);
    }
  };

  const triggerAttach = () => {
    fileInputRef.current?.click();
  };

  const clearAttachedFile = () => {
    setAttachedFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleExport = async (nextState: ChatSessionState) => {
    if (!nextState.pipeline) {
      return;
    }

    setIsExporting(true);
    setStatus("正在生成 PPT 终稿。");

    try {
      const response = await fetch("/api/export-ppt", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(nextState.pipeline.node3)
      });

      if (!response.ok) {
        const payload = (await response.json()) as { error?: string };
        throw new Error(payload.error ?? "导出失败");
      }

      const blob = await response.blob();
      const fileName =
        response.headers.get("Content-Disposition")?.match(/filename\*=UTF-8''(.+)$/)?.[1] ??
        "固定格式汇报.pptx";
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = decodeURIComponent(fileName);
      anchor.click();
      URL.revokeObjectURL(url);
      setStatus("PPT 终稿已生成并开始下载。");
      setMessages((current) => [...current, makeMessage("assistant", "PPT 终稿已生成，下载已经开始。")]);
    } catch (error) {
      const message = extractErrorMessage(error);
      setStatus(message);
      setMessages((current) => [...current, makeMessage("assistant", message)]);
    } finally {
      setIsExporting(false);
    }
  };

  const sendTurn = async(params: {
    message: string;
    file?: File | null;
    directive?: ChatDirective | null;
    userContent?: string;
    attachment?: ChatAttachment;
    clearComposer?: boolean;
    clearAttachment?: boolean;
  }) => {
    if (isSending || isExporting) {
      return;
    }

    if (params.userContent || params.attachment) {
      setMessages((current) => [
        ...current,
        makeMessage("user", params.userContent ?? "", params.attachment)
      ]);
    }

    if (params.clearComposer) {
      setComposer("");
    }
    if (params.clearAttachment) {
      setAttachedFile(null);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }

    setIsSending(true);
    setStatus("正在处理中。");

    try {
      const formData = new FormData();
      formData.set("message", params.message);
      formData.set("state", JSON.stringify(chatState));

      if (params.file) {
        formData.set("file", params.file);
      }

      if (params.directive) {
        formData.set("directive", JSON.stringify(params.directive));
      }

      const response = await fetch("/api/chat", {
        method: "POST",
        body: formData
      });
      const payload = (await response.json()) as ChatRouteResponse | { error: string };

      if (!response.ok || "error" in payload) {
        throw new Error("error" in payload ? payload.error : "对话处理失败");
      }

      setChatState(payload.state);
      setMessages((current) => [
        ...current,
        ...payload.assistantMessages.map((content) => makeMessage("assistant", content))
      ]);
      setStatus(buildStatusByStage(payload));

      if (payload.action === "export") {
        await handleExport(payload.state);
      }
    } catch (error) {
      const message = extractErrorMessage(error);
      setStatus(message);
      setMessages((current) => [...current, makeMessage("assistant", message)]);
    } finally {
      setIsSending(false);
    }
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!canSend || isPreviewMode) {
      return;
    }

    const outgoingFile = attachedFile;
    const outgoingMessage = composer.trim();

    await sendTurn({
      message: outgoingMessage,
      file: outgoingFile,
      userContent: outgoingMessage,
      attachment: buildAttachment(outgoingFile),
      clearComposer: true,
      clearAttachment: true
    });
  };

  const handleDirectorySave = async () => {
    const directive = buildDirectoryDirective(chatState.pipeline, directoryDraft);
    if (!directive) {
      setStatus("目录还没有新的修改。");
      return;
    }

    await sendTurn({
      message: "请按我刚刚在目录卡片里的修改更新目录。",
      directive,
      userContent: `我已调整 ${directive.updates.length} 处目录内容，请更新。`
    });
  };

  const handleDirectoryConfirm = async () => {
    await sendTurn({
      message: "确认目录",
      directive: { type: "confirm-directory" },
      userContent: "确认目录"
    });
  };

  const handleSummarySave = async () => {
    const directive = buildSummaryDirective(chatState.pipeline, summaryDraft);
    if (!directive) {
      setStatus("摘要槽位还没有新的修改。");
      return;
    }

    await sendTurn({
      message: "请按我刚刚在槽位卡片里的修改更新摘要。",
      directive,
      userContent: `我已调整 ${directive.updates.length} 处摘要槽位，请更新。`
    });
  };

  const handleSummaryConfirm = async () => {
    await sendTurn({
      message: "确认摘要",
      directive: { type: "confirm-summary" },
      userContent: "确认摘要"
    });
  };

  const handleGenerate = async () => {
    await sendTurn({
      message: "生成PPT",
      userContent: "生成PPT"
    });
  };

  const renderLiveReview = () => {
    if ((chatState.stage === "directory" && directoryDraft.length > 0) || (isPreviewMode && previewPreset === "directory")) {
      return (
        <div className="chat-row is-assistant">
          <div className="chat-bubble assistant review-bubble">
            <DirectoryReviewPanel
              draft={directoryDraft}
              onChange={setDirectoryDraft}
              onSave={handleDirectorySave}
              onConfirm={handleDirectoryConfirm}
              disabled={isPreviewMode || isSending || isExporting}
            />
          </div>
        </div>
      );
    }

    if ((chatState.stage === "summary" || chatState.stage === "ready") && summaryDraft.length > 0) {
      return (
        <div className="chat-row is-assistant">
          <div className="chat-bubble assistant review-bubble">
            <SummaryReviewPanel
              draft={summaryDraft}
              validationIssues={isPreviewMode ? previewValidationIssues : chatState.pipeline?.node3.validation.issues ?? []}
              onChange={setSummaryDraft}
              onSave={handleSummarySave}
              onConfirm={handleSummaryConfirm}
              onGenerate={handleGenerate}
              disabled={isPreviewMode || isSending || isExporting}
              isReady={
                chatState.stage === "ready" &&
                !(chatState.pipeline?.node3.validation.hasOverflow ?? false)
              }
            />
          </div>
        </div>
      );
    }

    return null;
  };

  return (
    <main className="chat-app-shell">
      <section className="cover-mark">
        <Image
          src="/logo.png"
          alt="抱谷科技"
          className="brand-logo"
          width={760}
          height={365}
          priority
          unoptimized
        />
      </section>

      <section className="chat-only-card">
        <header className="chat-only-header">
          <div className="chat-only-headline" aria-live="polite">
            <span className="chat-only-headline-dot" />
            <span>PPT 生成助手 · {stageLabel}</span>
          </div>

          <div className="chat-only-meta">
            <span className="meta-chip">置信度：{chatState.confidence}%</span>
            <button
              type="button"
              className="btn btn-secondary btn-small"
              onClick={handleReset}
              disabled={isPreviewMode}
            >
              重新开始
            </button>
          </div>
        </header>

        <div ref={threadRef} className="chat-only-thread smooth-scroll">
          {messages.map((message, index) => (
            <Fragment key={message.id}>
              {shouldShowTimeDivider(messages[index - 1], message) ? (
                <div className="chat-time-divider">
                  <span>{formatTimeDivider(message.createdAt)}</span>
                </div>
              ) : null}

              <div className={`chat-row ${message.role === "assistant" ? "is-assistant" : "is-user"}`}>
                <div className={`chat-bubble ${message.role === "assistant" ? "assistant" : "user"}`}>
                  {message.attachment ? (
                    <div className="chat-file-card">
                      <div className="chat-file-icon">{message.attachment.extension}</div>
                      <div className="chat-file-meta">
                        <strong>{message.attachment.name}</strong>
                        <span>{message.attachment.sizeLabel}</span>
                      </div>
                    </div>
                  ) : null}

                  {message.content ? <div className="chat-bubble-text">{message.content}</div> : null}
                </div>
              </div>
            </Fragment>
          ))}

          {isSending || isExporting || previewSnapshot?.showThinking ? (
            <div className="chat-row is-assistant">
              <div className="chat-bubble assistant thinking-bubble">
                <div className="thinking-dots" aria-hidden="true">
                  <span />
                  <span />
                  <span />
                </div>
                <div className="thinking-text">
                  {previewSnapshot?.showThinking ? "正在解析素材并整理目录" : thinkingLabel}
                </div>
              </div>
            </div>
          ) : null}

          {renderLiveReview()}
        </div>

        <footer className="chat-only-compose">
          <input
            ref={fileInputRef}
            hidden
            type="file"
            accept=".doc,.docx,.md,.txt,text/plain"
            onChange={handleFileChange}
          />

          {attachedFile ? (
            <div className="composer-file-card">
              <div className="chat-file-card is-pending">
                <div className="chat-file-icon">{buildAttachment(attachedFile)?.extension ?? "FILE"}</div>
                <div className="chat-file-meta">
                  <strong>{attachedFile.name}</strong>
                  <span>{formatFileSize(attachedFile.size)}</span>
                </div>
              </div>
              <button type="button" className="attachment-clear" onClick={clearAttachedFile}>
                移除
              </button>
            </div>
          ) : null}

          <form className="chat-only-form" onSubmit={handleSubmit}>
            <div className="chat-input-shell">
              <button
                type="button"
                className="chat-plus-btn"
                onClick={triggerAttach}
                aria-label="上传文件"
                disabled={isPreviewMode}
              >
                +
              </button>

              <textarea
                value={composer}
                onChange={(event) => setComposer(event.target.value)}
                placeholder="输入消息，或粘贴纯文字素材"
                disabled={isPreviewMode}
              />

              <button
                type="submit"
                className="chat-send-btn"
                disabled={!canSend || isSending || isExporting || isPreviewMode}
              >
                {isSending ? "处理中" : isExporting ? "导出中" : "发送"}
              </button>
            </div>
          </form>

          <p className="status">{status}</p>
        </footer>
      </section>
    </main>
  );
}
