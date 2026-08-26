import { useEffect, useRef, useState, useMemo } from "react";
import { isSignedIn, signIn } from "../auth";
import { clearTokens } from "../tokenStorage";
// const logoUrl = "https://noah-office.enpointe.io/assets/icon-80.png";
import {
  AddonAgent,
  AddonAnswer,
  Bootstrap,
  Citation,
  ContextPolicy,
  ContextRequiredError,
  DocumentContext,
  XlsxTable,
  askStart,
  cancelJob,
  bootstrap as fetchBootstrap,
  listAgents,
  putIndex,
  putSnapshot,
  renderChartPng,
} from "../addonClient";
import { createConversation, listMessages, type ChatMessage } from "../chatClient";
import { runJob } from "../session/jobs";
import { describeAnchor } from "../document/anchor";
import { capabilities, currentHost, hostKey } from "../document/capabilities";
import {
  buildIndex,
  currentRevision,
  isStale,
  markUploaded,
  watchForChanges,
} from "../document/documentIndex";
import { Selection } from "../document/selection";
import * as snapshotCache from "../document/snapshotCache";
import { OfficeHost } from "./host/OfficeHost";
import { GoogleHost } from "./host/GoogleHost";
import { renderTableImageBase64 } from "./renderTableImage";
import "../styles/tailwind.css";
import { NoahHeader, NoahShell, NoahToolbar } from "./components/NoahShell";
import { SkillSuggestions, WelcomeScreen } from "./components/Welcome";
import { ConversationView } from "./components/Conversation";
import { Composer } from "./components/Composer";
import { SettingsPage } from "./components/Settings";
import { HistoryPage } from "./components/History";
import { MoreMenu } from "./components/MoreMenu";
import { LoginApp } from "../login/LoginApp";

/** The `AddonSurface` the answer is shaped for. Distinct from the host: Word
 * and Google Docs are both `docs`, and shaping only ever keys off this. */
// function surfaceFor(host: string): string {
//   if (host === "Excel") return "sheets";
//   if (host === "PowerPoint") return "slides";
//   return "docs";
// }

/** A stable id for the open file. Office has no document id, and
 * `window.Office?.context.document.url` is the only thing that survives a reopen — so
 * hash it rather than sending a full local path to the server. */
function documentId(): string {
  const url = (window.Office?.context?.document as { url?: string })?.url ?? "";
  let hash = 0;
  for (let i = 0; i < url.length; i++) hash = ((hash << 5) - hash + url.charCodeAt(i)) | 0;
  return `office-${(hash >>> 0).toString(16)}`;
}

function documentTitle(): string {
  const url = (window.Office?.context?.document as { url?: string })?.url ?? "";
  const parts = url.split(/[\\/]/);
  return parts[parts.length - 1] || "the open document";
}

async function blobToBase64(blob: Blob): Promise<string> {
  const buf = await blob.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

export function App() {
  const host = currentHost();
  const documentHost = useMemo(() => {
    // @ts-ignore
    if (typeof google !== "undefined" && google.script) {
      return new GoogleHost(host);
    }
    return new OfficeHost(host);
  }, [host]);
  // const surface = surfaceFor(host);
  const [signedIn, setSignedIn] = useState<boolean | null>(null);
  const [me, setMe] = useState<Bootstrap | null>(null);
  const [policy, setPolicy] = useState<ContextPolicy | null>(null);
  const [orgId, setOrgId] = useState<string>("");
  const [agents, setAgents] = useState<AddonAgent[]>([]);
  const [agentId, setAgentId] = useState<string>("");
  const [conversationId, setConversationId] = useState<string>("");
  const [question, setQuestion] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [latestAnswer, setLatestAnswer] = useState<AddonAnswer | null>(null);
  // Whether to send what the user has selected. On by default — it is the
  // whole point of the feature — but a visible switch, because "what is Noah
  // reading?" should never be a guess.
  const [useSelection, setUseSelection] = useState(true);
  const [selectionHint, setSelectionHint] = useState<string>("");
  const [activity, setActivity] = useState<string>("");
  const cancelRef = useRef<{ cancelled: boolean } | null>(null);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const userScrolledUpRef = useRef<boolean>(false);
  const [showScrollButton, setShowScrollButton] = useState(false);
  const [hasUnread, setHasUnread] = useState(false);

  const scrollToBottom = (smooth = true) => {
    if (messagesEndRef.current) {
      messagesEndRef.current.scrollIntoView({ behavior: smooth ? "smooth" : "auto" });
    } else if (scrollContainerRef.current) {
      scrollContainerRef.current.scrollTo({
        top: scrollContainerRef.current.scrollHeight,
        behavior: smooth ? "smooth" : "auto",
      });
    }
  };

  const handleScroll = () => {
    const container = scrollContainerRef.current;
    if (!container) return;
    const distanceToBottom = container.scrollHeight - container.scrollTop - container.clientHeight;
    const isUp = distanceToBottom > 60;
    userScrolledUpRef.current = isUp;
    setShowScrollButton(isUp);
    if (!isUp) {
      setHasUnread(false);
    }
  };

  useEffect(() => {
    if (!userScrolledUpRef.current) {
      scrollToBottom(true);
    } else if (busy || latestAnswer) {
      setHasUnread(true);
    }
  }, [messages, latestAnswer, busy, activity]);

  const [page, setPage] = useState<"chat" | "settings" | "history">("chat");
  const [menuOpen, setMenuOpen] = useState(false);
  const [textSize, setTextSize] = useState<number>(() => {
    const stored = Number(localStorage.getItem("noah_text_size"));
    return stored >= 12 && stored <= 18 ? stored : 14;
  });
  const [instructions, setInstructions] = useState<string>(
    () => localStorage.getItem("noah_instructions") || ""
  );
  // const [history, setHistory] = useState<HistoryEntry[]>([]);

  useEffect(() => {
    document.documentElement.style.setProperty("--noah-text-base", `${textSize}px`);
    localStorage.setItem("noah_text_size", String(textSize));
  }, [textSize]);

  useEffect(() => {
    localStorage.setItem("noah_instructions", instructions);
  }, [instructions]);

  useEffect(() => {
    void isSignedIn().then(setSignedIn);
  }, []);

  useEffect(() => {
    if (signedIn) void refreshMe();
  }, [signedIn]);

  // listMessages is now called explicitly when selecting a conversation from history

  async function refreshMe() {
    try {
      const boot = await fetchBootstrap();
      setMe(boot);
      setPolicy(boot.context_policy);
      const storedOrg = localStorage.getItem("noah_selected_org");
      const first = boot.organizations.find(o => o.id === storedOrg) || boot.organizations[0];
      if (first) setOrgId(first.id);
    } catch {
      setSignedIn(false);
    }
  }

  useEffect(() => {
    if (orgId) {
      void listAgents(orgId).then((a) => {
        setAgents(a);
        setAgentId(a[0]?.id ?? "");
      });
    }
  }, [orgId]);



  // Ask the host to tell us when the document changes, so the map is rebuilt
  // only when it's actually stale. Best-effort: a host that can't raise the
  // event still works, it just rebuilds less often.
  useEffect(() => {
    void watchForChanges(host);
  }, [host]);

  // Keep a live description of what would be sent. We listen to selection
  // changes continuously. If the user selects something new, we automatically
  // enable the selection context.
  useEffect(() => {
    let cancelled = false;
    let lastHint = "";

    let debounceTimeout: ReturnType<typeof setTimeout>;

    const refresh = () => {
      clearTimeout(debounceTimeout);
      debounceTimeout = setTimeout(() => {
        void documentHost.captureSelection()
          .then((sel: Selection | null) => {
            if (cancelled) return;
            const hint = sel ? describeAnchor(sel.anchor) : "";
            setSelectionHint(hint);

            if (hint && hint !== lastHint) {
              setUseSelection(true);
            }
            lastHint = hint;
          })
          .catch(() => setSelectionHint(""));
      }, 300);
    };

    refresh();

    let interval: ReturnType<typeof setInterval> | undefined;

    if (host.startsWith("Google")) {
      interval = setInterval(refresh, 2000);
    } else {
      try {
        window.Office?.context?.document?.addHandlerAsync(
          window.Office?.EventType.DocumentSelectionChanged,
          refresh
        );
      } catch {
        // Older hosts don't raise it; the hint just won't live-update.
      }
    }

    return () => {
      cancelled = true;
      if (debounceTimeout) clearTimeout(debounceTimeout);
      if (interval) clearInterval(interval);
      if (!host.startsWith("Google")) {
        try {
          window.Office?.context?.document?.removeHandlerAsync(
            window.Office?.EventType.DocumentSelectionChanged,
            refresh
          );
        } catch { }
      }
    };
  }, [host]);

  async function handleSignIn() {
    setError(null);
    const result = await signIn();
    if (result.status === "signed-in") {
      setSignedIn(true);
    } else if (result.status === "error") {
      setError(result.message);
      throw new Error(result.message);
    } else if (result.status === "cancelled") {
      throw new Error("Sign-in was cancelled.");
    }
  }

  async function handleSignOut() {
    await clearTokens();
    snapshotCache.clear();
    setSignedIn(false);
    setMe(null);
  }

  /** Build the context envelope for this ask, uploading the selection only if
   * we don't already hold an id for it. */
  async function buildContext(): Promise<{
    context: DocumentContext | null;
    selection: Selection | null;
  }> {
    const doc = {
      host: hostKey(host),
      document_id: documentId(),
      title: documentTitle(),
    };
    const base: DocumentContext = {
      document: doc,
      scope: "selection",
      capabilities: capabilities(host),
      client_version: "c2",
    };

    // The map, refreshed only when the host told us something changed. This is
    // what lets the agent answer "which sheet has the revenue?" with nothing
    // selected — and it is structure, never cell values.
    if (policy?.allow_index && isStale()) {
      const index = await buildIndex(doc, host);
      if (index) {
        try {
          const stored = await putIndex(orgId, index);
          markUploaded(stored.revision);
        } catch (err) {
          // A missing map is a smaller answer, not a failed question.
          console.warn("Noah: could not upload the document map", err);
        }
      }
    }
    base.index_revision = currentRevision();

    if (!useSelection || !policy?.enabled) {
      // Still send the document so the conversation is bound to this file —
      // that binding is what a per-document chat list keys off later.
      return { context: { ...base, scope: "none" }, selection: null };
    }

    const selection = await documentHost.captureSelection();
    if (!selection) return { context: { ...base, scope: "none" }, selection: null };

    const known = snapshotCache.lookup(selection);
    if (known) return { context: { ...base, snapshot_id: known }, selection };

    const stored = await putSnapshot(orgId, doc, "selection", selection);
    snapshotCache.remember(selection, stored.context_id);
    return { context: { ...base, snapshot_id: stored.context_id }, selection };
  }

  async function handleAsk() {
    if (!question.trim() || !orgId) return;
    const asked = question;
    setQuestion("");
    setBusy(true);
    setError(null);
    setActivity("");
    setLatestAnswer(null);
    userScrolledUpRef.current = false;
    setShowScrollButton(false);
    setHasUnread(false);
    const signal = { cancelled: false };
    cancelRef.current = signal;

    const userMsgId = Date.now().toString();
    setMessages(prev => [
      ...prev,
      { id: userMsgId, role: "user", content: [{ type: "text", text: asked }] }
    ]);
    setTimeout(() => scrollToBottom(true), 50);

    let convIdToUse = conversationId;
    if (!convIdToUse) {
      try {
        const conv = await createConversation(orgId, agentId || "");
        setConversationId(conv.id);
        convIdToUse = conv.id;
      } catch (err) {
        setError("Failed to create conversation: " + String(err));
        setBusy(false);
        return;
      }
    }

    let jobId = "";

    try {
      const { context, selection } = await buildContext();

      const body = {
        text: asked,
        surface: host === "Excel" ? "sheets" : (host === "PowerPoint" ? "slides" : "docs"),
        conversation_id: convIdToUse,
        agent_id: agentId || null,
        context,
      };

      let started;
      try {
        started = await askStart(orgId, body);
      } catch (err: any) {
        if (!(err instanceof ContextRequiredError) || !context || !selection) throw err;
        snapshotCache.forget(err.snapshotId);
        const stored = await putSnapshot(orgId, context.document, "selection", selection);
        snapshotCache.remember(selection, stored.context_id);
        started = await askStart(orgId, {
          ...body,
          context: { ...context, snapshot_id: stored.context_id, selection: null },
        });
      }

      jobId = started.job_id;
      const job = await runJob(jobId, {
        host,
        onActivity: setActivity,
        onAnswerUpdate: setLatestAnswer,
        signal,
      });

      if (job.status === "error" || !job.answer) {
        throw new Error(job.error || "That request failed.");
      }

      const answer = job.answer;
      setLatestAnswer(answer);

      const assistantMsgId = (Date.now() + 1).toString();
      const baseText = answer.markdown || answer.prose || "";
      setMessages(prev => [
        ...prev,
        { id: assistantMsgId, role: "assistant", content: [{ type: "text", text: baseText }] }
      ]);

      let insertedCount = 0;
      let insertedDetails: string[] = [];
      if (answer.tables && answer.tables.length > 0) {
        for (const table of answer.tables) {
          try {
            if (table.chart_type) {
              await handleInsertChart(table);
              insertedCount++;
            } else {
              const res = await handleInsertTable(table, answer.intent?.is_update ?? false);
              insertedCount++;
              if (res && res.address) insertedDetails.push(res.address);
            }
          } catch (e) {
            console.error("Failed to insert table:", e);
          }
        }
      }

      if (insertedCount > 0) {
        setMessages(prev => prev.map(msg => {
          if (msg.id === assistantMsgId) {
            const currentText = (msg.content[0] as any).text;
            let suffix = `*Inserted ${insertedCount} item(s) directly into your document.*`;
            if (insertedDetails.length > 0) {
              suffix = `*Inserted ${insertedCount} item(s) directly into your document at ${insertedDetails.join(", ")}.*`;
            }
            const newText = currentText
              ? `${currentText}\n\n${suffix}`
              : suffix;
            return { ...msg, content: [{ type: "text", text: newText }] };
          }
          return msg;
        }));
      }
    } catch (e) {
      if (!signal.cancelled) {
        setError(e instanceof Error ? e.message : "Something went wrong.");
      }
    } finally {
      if (jobId && signal.cancelled) void cancelJob(jobId).catch(() => undefined);
      cancelRef.current = null;
      setActivity("");
      setLatestAnswer(null);
      setBusy(false);
    }
  }

  function handleStop() {
    if (cancelRef.current) cancelRef.current.cancelled = true;
    setBusy(false);
    setActivity("");
  }

  function handleNewChat() {
    setConversationId("");
    setQuestion("");
    setMessages([]);
    setError(null);
    setActivity("");
    setLatestAnswer(null);
    userScrolledUpRef.current = false;
    setShowScrollButton(false);
    setHasUnread(false);
  }

  // function handleSelectHistory(entry: HistoryEntry) {
  //   setQuestion(entry.question);
  //   setPage("chat");
  // }

  /** Select the region a claim came from. The single highest-trust feature
   * here: it turns "the variance is 82,044" into something checkable. */
  async function handleCitation(citation: Citation) {
    try {
      await documentHost.handleCitation(citation);
    } catch {
      setError("Couldn't jump to that — it may have moved.");
    }
  }

  async function handleInsertProse(textToInsert?: string) {
    let content = textToInsert;
    if (!content) {
      if (messages.length === 0) return;
      const lastMsg = messages[messages.length - 1];
      if (lastMsg.role !== "assistant") return;

      const textBlocks = lastMsg.content.filter((b) => b.type === "text" && typeof b.text === "string");
      content = textBlocks.map((b) => (b as any).text).join("\n\n");
    }
    if (!content) return;

    await documentHost.insertProse(content);
  }

  async function handleInsertTable(table: XlsxTable, isUpdate: boolean = false) {
    try {
      if (host === "PowerPoint" || host === "GoogleSlides") {
        await documentHost.insertImageBase64(await renderTableImageBase64(table));
      } else {
        const res = await documentHost.insertTable(table, isUpdate);
        return res;
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not insert the table.");
      throw e;
    }
  }

  async function handleInsertChart(table: XlsxTable) {
    if (!table.chart_type) return;
    try {
      if (host === "Excel") {
        // Excel builds a real native chart directly from the range it just
        // wrote — write the table first, then chart the resulting range.
        const result = await documentHost.insertTable(table);
        if (result && result.address) {
          await documentHost.insertChart(table.chart_type, result.address);
        }
      } else {
        const png = await renderChartPng(orgId, table.chart_type, table);
        const base64 = await blobToBase64(png);
        await documentHost.insertImageBase64(base64);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not render the chart.");
    }
  }

  if (signedIn === null) {
    return (
      <NoahShell>
        <WelcomeScreen subtitle="Loading…" />
      </NoahShell>
    );
  }

  if (!signedIn) {
    const isGoogle = host.startsWith("Google");
    if (typeof window !== "undefined" && (!window.Office || !window.Office.context || !window.Office.context.ui) && !isGoogle) {
      window.location.href = "/login";
      return null;
    }
    return (
      <div className="w-full h-screen px-4 box-border bg-canvas overflow-y-auto flex flex-col justify-center items-center">
        {/* <img src={logoUrl} alt="Noah" className="logo mb-4 h-8" onError={(e) => (e.currentTarget.style.display = 'none')} /> */}
        <LoginApp
          onSuccess={async () => {
            setSignedIn(true);
          }}
          onSsoClick={handleSignIn}
        />
      </div>
    );
  }

  if (page === "settings") {
    return (
      <NoahShell>
        <SettingsPage
          onBack={() => setPage("chat")}
          email={me?.email ?? ""}
          useSelectionDefault={useSelection}
          onToggleSelectionDefault={setUseSelection}
          textSize={textSize}
          onTextSizeChange={setTextSize}
          instructions={instructions}
          onInstructionsChange={setInstructions}
          organizations={me?.organizations || []}
          orgId={orgId}
          onOrgChange={(id) => {
            setOrgId(id);
            localStorage.setItem("noah_selected_org", id);
          }}
          onSignOut={handleSignOut}
        />
      </NoahShell>
    );
  }

  if (page === "history") {
    return (
      <NoahShell>
        <HistoryPage
          orgId={orgId}
          currentId={conversationId}
          onBack={() => setPage("chat")}
          onNewChat={() => {
            handleNewChat();
            setPage("chat");
          }}
          onSelect={(conv) => {
            setConversationId(conv.id);
            setMessages([]);
            setQuestion("");
            setActivity("");
            setError(null);
            setPage("chat");
            userScrolledUpRef.current = false;
            setShowScrollButton(false);
            setHasUnread(false);
            listMessages(orgId, conv.id)
              .then((msgs) => {
                setMessages(msgs);
                setTimeout(() => scrollToBottom(false), 50);
              })
              .catch(console.error);
          }}
        />
      </NoahShell>
    );
  }

  const hasExchange = busy || messages.length > 0 || !!error;

  return (
    <NoahShell>
      <NoahHeader
        title="Noah"
        actions={
          <>
            <button
              className="w-[30px] h-[30px] inline-flex items-center justify-center rounded-lg border-none bg-transparent text-ink-secondary cursor-pointer hover:bg-surface-hover hover:text-ink transition-colors duration-150"
              onClick={handleNewChat}
              aria-label="New chat"
              title="New chat"
            >
              <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
                <path d="M8 3.5v9M3.5 8h9" />
              </svg>
            </button>
            <button
              className={`w-[30px] h-[30px] inline-flex items-center justify-center rounded-lg border-none bg-transparent text-ink-secondary cursor-pointer hover:bg-surface-hover hover:text-ink transition-colors duration-150 ${menuOpen ? 'bg-surface-hover text-ink' : ''}`}
              onClick={() => setMenuOpen((prev) => !prev)}
              aria-label="More options"
              title="More options"
            >
              <svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor">
                <circle cx="3" cy="8" r="1.25" />
                <circle cx="8" cy="8" r="1.25" />
                <circle cx="13" cy="8" r="1.25" />
              </svg>
            </button>
          </>
        }
      />

      {agents.length > 1 ? (
        <NoahToolbar>
          <select className="h-[26px] border border-border rounded-full bg-canvas text-ink-secondary text-[12px] px-2 max-w-full" value={agentId} onChange={(e) => setAgentId(e.target.value)}>
            {agents.map((a) => (
              <option key={a.id} value={a.id}>
                {a.name}
              </option>
            ))}
          </select>
        </NoahToolbar>
      ) : null}

      <div className="relative flex-1 min-h-0 flex flex-col overflow-hidden">
        <div
          ref={scrollContainerRef}
          onScroll={handleScroll}
          className="flex-1 overflow-y-auto min-h-0"
        >
          {hasExchange ? (
            <ConversationView
              messages={messages}
              latestAnswer={latestAnswer}
              error={error}
              busy={busy}
              activity={activity}
              host={host}
              onInsertProse={handleInsertProse}
              onInsertTable={handleInsertTable}
              onCitation={handleCitation}
              bottomRef={messagesEndRef}
            />
          ) : (
            <WelcomeScreen subtitle={`Signed in as ${me?.email ?? ""} · ${host}`}>
              <SkillSuggestions onPick={setQuestion} />
            </WelcomeScreen>
          )}
        </div>

        {/* ChatGPT-style floating scroll-to-bottom indicator */}
        {hasExchange && showScrollButton && (
          <button
            onClick={() => {
              userScrolledUpRef.current = false;
              setShowScrollButton(false);
              setHasUnread(false);
              scrollToBottom(true);
            }}
            className="absolute bottom-3 left-1/2 -translate-x-1/2 z-20 w-8 h-8 rounded-full bg-white/95 backdrop-blur border border-zinc-200 shadow-md text-zinc-600 hover:text-emerald-700 hover:bg-zinc-50 flex items-center justify-center cursor-pointer transition-all duration-200 animate-fadeIn hover:scale-105 active:scale-95"
            aria-label="Scroll to bottom"
            title="Scroll to bottom"
          >
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12 5v14M19 12l-7 7-7-7" />
            </svg>
            {hasUnread && (
              <span className="absolute -top-0.5 -right-0.5 flex h-2.5 w-2.5">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500 border border-white"></span>
              </span>
            )}
          </button>
        )}
      </div>

      <Composer
        value={question}
        onChange={setQuestion}
        onSend={handleAsk}
        onStop={handleStop}
        busy={busy}
        showSelectionToggle={true}
        useSelection={useSelection}
        onToggleSelection={setUseSelection}
        selectionHint={selectionHint}
      />

      {menuOpen && (
        <MoreMenu
          onClose={() => setMenuOpen(false)}
          onSettings={() => setPage("settings")}
          onHistory={() => setPage("history")}
          onSignOut={handleSignOut}
        />
      )}
    </NoahShell>
  );
}
