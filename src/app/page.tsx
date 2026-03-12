"use client";

import Image from "next/image";
import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import type { ChangeEvent, FormEvent } from "react";
import type { ChatRouteResponse, ChatSessionState } from "@/lib/types";

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
      "欢迎使用 PPT 生成助手。\n\n我们会通过对话完成信息收集、目录确认、摘要确认和终稿导出。\n\n先告诉我汇报人名称。"
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

export default function HomePage() {
  const [messages, setMessages] = useState<ChatMessage[]>(() => createInitialMessages());
  const [chatState, setChatState] = useState<ChatSessionState>(initialState);
  const [composer, setComposer] = useState("");
  const [attachedFile, setAttachedFile] = useState<File | null>(null);
  const [status, setStatus] = useState("可直接输入文字，或点击左下角 + 上传素材。");
  const [isSending, setIsSending] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const threadRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const node = threadRef.current;
    if (!node) {
      return;
    }

    node.scrollTo({
      top: node.scrollHeight,
      behavior: "smooth"
    });
  }, [messages]);

  const canSend = useMemo(
    () => Boolean(composer.trim()) || Boolean(attachedFile),
    [attachedFile, composer]
  );

  const stageLabel = useMemo(() => {
    switch (chatState.stage) {
      case "directory":
        return "目录确认";
      case "summary":
        return "摘要确认";
      case "ready":
        return "终稿导出";
      default:
        return "信息收集";
    }
  }, [chatState.stage]);

  const handleReset = () => {
    setMessages(createInitialMessages());
    setChatState(initialState);
    setComposer("");
    setAttachedFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setStatus("可直接输入文字，或点击左下角 + 上传素材。");
  };

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
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
      setMessages((current) => [...current, makeMessage("assistant", "PPT 终稿已生成，下载已经开始。")]);
    } catch (error) {
      const message = extractErrorMessage(error);
      setStatus(message);
      setMessages((current) => [...current, makeMessage("assistant", message)]);
    } finally {
      setIsExporting(false);
    }
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!canSend || isSending || isExporting) {
      return;
    }

    const outgoingFile = attachedFile;
    const outgoingMessage = composer.trim();
    const attachment = buildAttachment(outgoingFile);

    setMessages((current) => [...current, makeMessage("user", outgoingMessage || "", attachment)]);
    setComposer("");
    setAttachedFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setIsSending(true);
    setStatus("正在处理中。");

    try {
      const formData = new FormData();
      formData.set("message", outgoingMessage);
      formData.set("state", JSON.stringify(chatState));

      if (outgoingFile) {
        formData.set("file", outgoingFile);
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
      setStatus(
        payload.action === "export"
          ? "准备导出 PPT 终稿。"
          : payload.state.stage === "directory"
            ? "请在对话里确认或修改目录。"
            : payload.state.stage === "summary"
              ? "请在对话里确认或修改摘要。"
              : payload.state.stage === "ready"
                ? "可以直接回复“生成PPT”。"
                : "继续告诉我还缺少的信息。"
      );

      if (payload.state.pipeline) {
        setAttachedFile(null);
      }

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
            <button type="button" className="btn btn-secondary btn-small" onClick={handleReset}>
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

                  {message.content ? (
                    <div className="chat-bubble-text">{message.content}</div>
                  ) : null}
                </div>
              </div>
            </Fragment>
          ))}
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
              <button type="button" className="chat-plus-btn" onClick={triggerAttach} aria-label="上传文件">
                +
              </button>

              <textarea
                value={composer}
                onChange={(event) => setComposer(event.target.value)}
                placeholder="输入消息，或粘贴纯文字素材"
              />

              <button
                type="submit"
                className="chat-send-btn"
                disabled={!canSend || isSending || isExporting}
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
