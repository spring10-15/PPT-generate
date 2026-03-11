"use client";

import { useMemo, useState } from "react";
import type { ChangeEvent, FormEvent } from "react";
import type { OutlineDocument, OutlineSlide } from "@/lib/types";
import { normalizeOutlineDocument } from "@/lib/outline-format";
import { composeSlideSummary } from "@/lib/slide-content";

type FormState = {
  userName: string;
  totalPages: string;
};

type ChatRole = "assistant" | "user";
type FlowPhase = "collect" | "directory" | "summary" | "ready";
type CollectStep = "userName" | "totalPages" | "file";

type ChatMessage = {
  id: string;
  role: ChatRole;
  content: string;
};

const initialForm: FormState = {
  userName: "",
  totalPages: "8"
};

const initialMessages: ChatMessage[] = [
  {
    id: "msg-1",
    role: "assistant",
    content: "欢迎使用 PPT 生成助手。先告诉我汇报人名称。"
  }
];

function makeMessage(role: ChatRole, content: string): ChatMessage {
  return {
    id: `msg-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    role,
    content
  };
}

function patchSlide(
  outline: OutlineDocument,
  slideId: string,
  patch: Partial<OutlineSlide>
): OutlineDocument {
  return normalizeOutlineDocument({
    ...outline,
    slides: outline.slides.map((slide) => (slide.id === slideId ? { ...slide, ...patch } : slide))
  });
}

function patchSlideContent(
  outline: OutlineDocument,
  slideId: string,
  patch: Partial<OutlineSlide["content"]>
): OutlineDocument {
  return normalizeOutlineDocument({
    ...outline,
    slides: outline.slides.map((slide) => {
      if (slide.id !== slideId) {
        return slide;
      }

      const content = { ...slide.content, ...patch };
      return {
        ...slide,
        content,
        summary: composeSlideSummary(content)
      };
    })
  });
}

function patchDirectoryTitle(
  outline: OutlineDocument,
  directoryId: string,
  patch: Partial<OutlineDocument["directory"][number]>
): OutlineDocument {
  return normalizeOutlineDocument({
    ...outline,
    directory: outline.directory.map((item) => (item.id === directoryId ? { ...item, ...patch } : item)),
    slides: outline.slides.map((slide) =>
      slide.directoryId === directoryId && patch.title
        ? { ...slide, directoryTitle: patch.title }
        : slide
    )
  });
}

function removeDirectory(outline: OutlineDocument, directoryId: string): OutlineDocument {
  return normalizeOutlineDocument({
    ...outline,
    directory: outline.directory.filter((item) => item.id !== directoryId),
    slides: outline.slides.filter((slide) => slide.directoryId !== directoryId)
  });
}

function removeSlide(outline: OutlineDocument, slideId: string): OutlineDocument {
  return normalizeOutlineDocument({
    ...outline,
    slides: outline.slides.filter((slide) => slide.id !== slideId)
  });
}

function extractErrorMessage(error: unknown) {
  if (error instanceof Error) {
    return error.message;
  }

  return "处理失败，请稍后重试。";
}

export default function HomePage() {
  const [formState, setFormState] = useState<FormState>(initialForm);
  const [file, setFile] = useState<File | null>(null);
  const [chatInput, setChatInput] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>(initialMessages);
  const [outline, setOutline] = useState<OutlineDocument | null>(null);
  const [phase, setPhase] = useState<FlowPhase>("collect");
  const [collectStep, setCollectStep] = useState<CollectStep>("userName");
  const [activeTab, setActiveTab] = useState<"directory" | "summary">("directory");
  const [isGenerating, setIsGenerating] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [status, setStatus] = useState("按照对话依次收集汇报人、页数和素材文件。");

  const parsedTotalPages = Number(formState.totalPages || 0);
  const canGenerateOutline =
    Boolean(file) &&
    Boolean(formState.userName) &&
    Number.isFinite(parsedTotalPages) &&
    parsedTotalPages >= 3 &&
    !isGenerating;
  const canExport = phase === "ready" && Boolean(outline) && !isExporting;

  const summaryText = useMemo(() => {
    if (!outline) {
      return "当前还没有生成目录与详情摘要。";
    }

    return `当前方案共 ${outline.pageSummary.totalPages} 页，目录 ${outline.directory.length} 项，详情页 ${outline.pageSummary.detailPages} 页。`;
  }, [outline]);

  const stageTitle = useMemo(() => {
    switch (phase) {
      case "directory":
        return "目录架构确认";
      case "summary":
        return "详情摘要确认";
      case "ready":
        return "准备生成终稿";
      default:
        return "对话输入";
    }
  }, [phase]);

  const stageHint = useMemo(() => {
    switch (phase) {
      case "directory":
        return "节点一完成后，先确认目录标题和简要说明。";
      case "summary":
        return "目录确认后，查看详情摘要。摘要区固定高度，支持在框内滚动。";
      case "ready":
        return "目录和详情摘要都已确认，可以生成 PPT 终稿。";
      default:
        return "系统会以聊天形式引导收集信息，并按 SOP 分阶段推进。";
    }
  }, [phase]);

  const pushMessages = (...nextMessages: ChatMessage[]) => {
    setMessages((current) => [...current, ...nextMessages]);
  };

  const resetSession = () => {
    setFormState(initialForm);
    setFile(null);
    setChatInput("");
    setMessages(initialMessages);
    setOutline(null);
    setPhase("collect");
    setCollectStep("userName");
    setActiveTab("directory");
    setStatus("按照对话依次收集汇报人、页数和素材文件。");
  };

  const handleChatSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const value = chatInput.trim();
    if (!value) {
      return;
    }

    if (collectStep === "userName") {
      setFormState((current) => ({ ...current, userName: value }));
      setChatInput("");
      pushMessages(
        makeMessage("user", `汇报人名称：${value}`),
        makeMessage("assistant", "已记录汇报人。接下来告诉我需要生成多少页 PPT。")
      );
      setCollectStep("totalPages");
      setStatus("已收集汇报人名称，等待页数。");
      return;
    }

    const parsed = Number(value.replace(/[^\d]/g, ""));
    if (!Number.isFinite(parsed) || parsed < 3) {
      pushMessages(makeMessage("assistant", "页数请输入不小于 3 的数字。"));
      setChatInput("");
      return;
    }

    setFormState((current) => ({ ...current, totalPages: String(parsed) }));
    setChatInput("");
    pushMessages(
      makeMessage("user", `生成页数：${parsed} 页`),
      makeMessage("assistant", "已记录页数。现在请上传素材文件，我会先生成目录架构。")
    );
    setCollectStep("file");
    setStatus("已收集页数，等待上传素材文件。");
  };

  const handleGenerateOutline = async () => {
    if (!file) {
      setStatus("请先上传 doc 或 docx 源文件。");
      return;
    }

    setIsGenerating(true);
    setStatus("正在执行节点一：文档解析与目录生成。");
    pushMessages(makeMessage("assistant", "开始执行节点一：文档解析与大纲生成。"));

    try {
      const formData = new FormData();
      formData.set("file", file);
      formData.set("userName", formState.userName);
      formData.set("totalPages", formState.totalPages);

      const response = await fetch("/api/generate-outline", {
        method: "POST",
        body: formData
      });
      const payload = (await response.json()) as OutlineDocument | { error: string };

      if (!response.ok || "error" in payload) {
        throw new Error("error" in payload ? payload.error : "生成失败");
      }

      setOutline(normalizeOutlineDocument(payload));
      setPhase("directory");
      setActiveTab("directory");
      setStatus("目录架构已生成，请先确认目录。");
      pushMessages(
        makeMessage("assistant", "目录架构已生成。请先在右侧确认目录标题和简要说明，确认后我再展示详情摘要。")
      );
    } catch (error) {
      const message = extractErrorMessage(error);
      setStatus(message);
      pushMessages(makeMessage("assistant", message));
    } finally {
      setIsGenerating(false);
    }
  };

  const handleConfirmDirectory = () => {
    if (!outline) {
      return;
    }

    setPhase("summary");
    setActiveTab("summary");
    setStatus("目录已确认，请确认详情摘要。");
    pushMessages(
      makeMessage("user", "目录架构已确认。"),
      makeMessage("assistant", "已进入详情摘要确认阶段。请在右侧固定框内滚动查看并修改详情摘要。")
    );
  };

  const handleConfirmSummary = () => {
    if (!outline) {
      return;
    }

    setPhase("ready");
    setStatus("目录和详情摘要都已确认，可以生成 PPT 终稿。");
    pushMessages(
      makeMessage("user", "详情摘要已确认。"),
      makeMessage("assistant", "已完成目录和摘要确认。现在可以生成 PPT 终稿。")
    );
  };

  const handleExport = async () => {
    if (!outline) {
      return;
    }

    setIsExporting(true);
    setStatus("正在执行节点四：数据灌入与文件生成。");
    pushMessages(makeMessage("user", "生成 PPT 终稿。"));

    try {
      const response = await fetch("/api/export-ppt", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(outline)
      });

      if (!response.ok) {
        const payload = (await response.json()) as { error?: string };
        throw new Error(payload.error ?? "导出失败");
      }

      const blob = await response.blob();
      const fileName =
        response.headers
          .get("Content-Disposition")
          ?.match(/filename\*=UTF-8''(.+)$/)?.[1] ?? "固定格式汇报.pptx";
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = decodeURIComponent(fileName);
      anchor.click();
      URL.revokeObjectURL(url);
      setStatus("PPT 终稿已生成并开始下载。");
      pushMessages(makeMessage("assistant", "PPT 终稿已生成，下载已开始。"));
    } catch (error) {
      const message = extractErrorMessage(error);
      setStatus(message);
      pushMessages(makeMessage("assistant", message));
    } finally {
      setIsExporting(false);
    }
  };

  const onFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const nextFile = event.target.files?.[0] ?? null;
    setFile(nextFile);
    if (!nextFile) {
      return;
    }

    pushMessages(
      makeMessage("user", `已上传素材：${nextFile.name}`),
      makeMessage("assistant", "素材已收到。点击下方按钮开始生成目录架构。")
    );
    setStatus(`已上传素材文件：${nextFile.name}`);
  };

  return (
    <main className="app-shell">
      <section className="hero">
        <div className="panel hero-main">
          <div className="eyebrow">固定模板 Demo 2.0</div>
          <h1>PPT 生成</h1>
          <p>根据固定模板生成 PPT</p>
        </div>

        <aside className="panel hero-side hero-side-compact">
          <div>
            <h2 className="section-title">SOP 进度</h2>
            <p className="hint">{summaryText}</p>
          </div>
          <div className="summary-stack">
            <div className={`summary-chip ${phase !== "collect" ? "is-done" : ""}`}>
              1. 文档解析与目录生成
            </div>
            <div className={`summary-chip ${phase === "summary" || phase === "ready" ? "is-done" : ""}`}>
              2. 目录架构确认
            </div>
            <div className={`summary-chip ${phase === "ready" ? "is-done" : ""}`}>
              3. 详情摘要确认
            </div>
            <div className={`summary-chip ${canExport ? "is-ready" : ""}`}>4. 生成 PPT 终稿</div>
          </div>
        </aside>
      </section>

      <section className="workspace">
        <section className="panel chat-panel">
          <div className="chat-header">
            <div>
              <h2 className="section-title">对话输入</h2>
              <p className="hint">
                通过聊天依次收集汇报人名称、页数和素材文件，再分阶段确认目录与摘要。
              </p>
            </div>
            <button type="button" className="btn btn-secondary btn-small" onClick={resetSession}>
              重新开始
            </button>
          </div>

          <div className="chat-thread">
            {messages.map((message) => (
              <div
                key={message.id}
                className={`chat-row ${message.role === "assistant" ? "is-assistant" : "is-user"}`}
              >
                <div className={`chat-bubble ${message.role === "assistant" ? "assistant" : "user"}`}>
                  {message.content}
                </div>
              </div>
            ))}
          </div>

          <div className="chat-compose">
            {phase === "collect" && collectStep !== "file" ? (
              <form className="chat-form" onSubmit={handleChatSubmit}>
                <input
                  value={chatInput}
                  onChange={(event) => setChatInput(event.target.value)}
                  placeholder={
                    collectStep === "userName" ? "输入汇报人名称" : "输入需要生成的 PPT 页数"
                  }
                />
                <button type="submit" className="btn btn-primary">
                  发送
                </button>
              </form>
            ) : phase === "collect" ? (
              <div className="upload-box chat-upload-box">
                <div className="field">
                  <label htmlFor="fileUpload">上传素材文件</label>
                  <input id="fileUpload" type="file" accept=".doc,.docx" onChange={onFileChange} />
                  <p className="hint">
                    优先保证 <code>.docx</code> 解析质量。<code>.doc</code> 会尝试转换，
                    当前环境缺少 LibreOffice 时会退化为文本解析。
                  </p>
                </div>
                <div className="chat-collect-meta">
                  <span>汇报人：{formState.userName || "未填写"}</span>
                  <span>页数：{formState.totalPages || "未填写"}</span>
                  <span>素材：{file?.name || "未上传"}</span>
                </div>
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={!canGenerateOutline}
                  onClick={handleGenerateOutline}
                >
                  {isGenerating ? "生成中..." : "生成目录架构"}
                </button>
              </div>
            ) : (
              <div className="chat-next-step">
                <p className="status">{status}</p>
                {phase === "directory" ? (
                  <button type="button" className="btn btn-primary" onClick={handleConfirmDirectory}>
                    确认目录并查看详情摘要
                  </button>
                ) : phase === "summary" ? (
                  <button type="button" className="btn btn-primary" onClick={handleConfirmSummary}>
                    确认详情摘要
                  </button>
                ) : (
                  <button
                    type="button"
                    className="btn btn-primary"
                    disabled={!canExport}
                    onClick={handleExport}
                  >
                    {isExporting ? "导出中..." : "生成 PPT 终稿"}
                  </button>
                )}
              </div>
            )}
            {phase === "collect" ? <p className="status">{status}</p> : null}
          </div>
        </section>

        <section className="panel editor-panel">
          <div className="editor-header">
            <div>
              <h2 className="section-title">{stageTitle}</h2>
              <p className="hint">{stageHint}</p>
            </div>
          </div>

          {!outline ? (
            <div className="empty-state">
              先在左侧完成对话输入，系统会先生成目录架构，再进入详情摘要确认。
            </div>
          ) : (
            <>
              <div className="tab-row" role="tablist" aria-label="确认内容标签">
                <button
                  type="button"
                  className={`tab-btn ${activeTab === "directory" ? "is-active" : ""}`}
                  onClick={() => setActiveTab("directory")}
                >
                  目录架构
                </button>
                <button
                  type="button"
                  className={`tab-btn ${activeTab === "summary" ? "is-active" : ""}`}
                  onClick={() => setActiveTab("summary")}
                  disabled={phase === "directory"}
                >
                  详情摘要
                </button>
              </div>

              <div className="editor-scroll">
                <div className="editor-grid">
                  {activeTab === "directory" ? (
                    <div className="slide-card">
                      <div className="slide-top">
                        <div>
                          <div className="slide-index">目录</div>
                        </div>
                      </div>
                      <div className="slide-body">
                        {outline.directory.map((item) => (
                          <div className="inline-grid" key={item.id}>
                            <div className="field">
                              <label>目录标题</label>
                              <input
                                value={item.title}
                                onChange={(event) =>
                                  setOutline((current) =>
                                    current
                                      ? patchDirectoryTitle(current, item.id, { title: event.target.value })
                                      : current
                                  )
                                }
                              />
                            </div>
                            <div className="field">
                              <label>简要说明</label>
                              <input
                                value={item.description}
                                onChange={(event) =>
                                  setOutline((current) =>
                                    current
                                      ? patchDirectoryTitle(current, item.id, {
                                          description: event.target.value
                                        })
                                      : current
                                  )
                                }
                              />
                            </div>
                            <div className="card-action">
                              <button
                                type="button"
                                className="btn btn-danger"
                                onClick={() =>
                                  setOutline((current) =>
                                    current ? removeDirectory(current, item.id) : current
                                  )
                                }
                              >
                                删除目录
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : (
                    outline.slides.map((slide, index) => (
                      <div className="slide-card" key={slide.id}>
                        <div className="slide-top">
                          <div>
                            <div className="slide-index">页码 {index + 1}</div>
                          </div>
                        </div>
                        <div className="slide-body">
                          <p className="small-meta">所属目录：{slide.directoryTitle}</p>
                          <div className="field">
                            <label>页标题</label>
                            <input
                              value={slide.title}
                              onChange={(event) =>
                                setOutline((current) =>
                                  current ? patchSlide(current, slide.id, { title: event.target.value }) : current
                                )
                              }
                            />
                          </div>
                          <div className="field">
                            <label>内容简介</label>
                            <textarea
                              value={slide.content.intro}
                              onChange={(event) =>
                                setOutline((current) =>
                                  current
                                    ? patchSlideContent(current, slide.id, { intro: event.target.value })
                                    : current
                                )
                              }
                            />
                          </div>
                          <div className="field">
                            <label>常规标题</label>
                            <input
                              value={slide.content.regularTitle}
                              onChange={(event) =>
                                setOutline((current) =>
                                  current
                                    ? patchSlideContent(current, slide.id, {
                                        regularTitle: event.target.value
                                      })
                                    : current
                                )
                              }
                            />
                          </div>
                          <div className="field">
                            <label>说明性文字</label>
                            <textarea
                              value={slide.content.description}
                              onChange={(event) =>
                                setOutline((current) =>
                                  current
                                    ? patchSlideContent(current, slide.id, {
                                        description: event.target.value
                                      })
                                    : current
                                )
                              }
                            />
                          </div>
                          <div className="card-action">
                            <button
                              type="button"
                              className="btn btn-danger"
                              onClick={() =>
                                setOutline((current) => (current ? removeSlide(current, slide.id) : current))
                              }
                            >
                              删除详情页
                            </button>
                          </div>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </>
          )}
        </section>
      </section>
    </main>
  );
}
