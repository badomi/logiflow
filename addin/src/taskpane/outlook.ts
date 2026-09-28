/* global document, Office */

import { finalizeDraft, findPendingDraft } from "../shared/finalize";
import { savePendingDraft } from "../shared/handoff";
import {
  ApiError,
  apiUrl,
  createCase,
  getCase,
  listCases,
  matchReply,
  mergeReply,
  quoteDraft,
  recordMailEvent,
  supplementDraft,
} from "./api";
import type { Draft, MailInput, ReplyCandidate } from "./api";
import { currentUserEmail, readSelectedMail, readSelectedMailRaw } from "./mail";
import type { MailSnapshot } from "./mail";

/** 본문 미리보기로 보여줄 글자 수 */
const PREVIEW_LENGTH = 300;

const REASON_LABELS: Record<string, string> = {
  SUBJECT_CASE_ID: "제목의 케이스 ID",
  HEADER: "회신 헤더",
  CONVERSATION: "같은 대화",
};

/** 지금 패널에 떠 있는 메일의 요청 데이터 (버튼을 누를 때 다시 읽지 않도록 보관) */
let current: MailInput | null = null;
/** 지금 메일이 속한 케이스 ID (등록된 경우에만) */
let currentCaseId: string | null = null;
/** 지금 메일이 담당자가 보낸 케이스 메일일 때 그 케이스 ID */
let sentCaseId: string | null = null;

const CASE_ID_PATTERN = /LQ-\d{4}-\d{4}-\d{3}/;
/** 메일을 빠르게 바꿀 때, 늦게 끝난 이전 메일 결과가 화면을 덮어쓰지 않게 하는 번호 */
let loadToken = 0;

// Office.js 준비가 끝나면 한 번 실행된다. Outlook 안에서 열렸을 때만 패널을 보여준다.
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "block";

  // 작성 창에서 연 패널이면 초안 마무리만 한다
  if (isComposeMode()) {
    document.getElementById("read-body")!.hidden = true;
    document.getElementById("compose-section")!.hidden = false;
    document.getElementById("compose-retry")!.onclick = runComposeFinalize;
    runComposeFinalize();
    return;
  }

  document.getElementById("create-case")!.onclick = onCreateCase;
  document.getElementById("merge-selected")!.onclick = () => {
    const select = document.getElementById("case-select") as HTMLSelectElement;
    if (select.value) onMerge(select.value);
  };
  document.getElementById("record-sent")!.onclick = onRecordSent;
  document.getElementById("supplement-draft")!.onclick = onSupplementDraft;
  document.getElementById("quote-draft")!.onclick = onQuoteDraft;

  // 패널을 고정(pin)해 두면, 담당자가 다른 메일을 선택할 때마다 새로 읽는다 (사용자 선택 이벤트 — 폴링 아님)
  if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => loadCurrentItem());
  }
  loadCurrentItem();
});

/** 작성 중인 메일이면 제목을 쓸 수 있는(setAsync) 객체가 있다 */
function isComposeMode(): boolean {
  const item = Office.context.mailbox.item as Office.MessageCompose | undefined;
  return typeof item?.subject?.setAsync === "function";
}

/** [작성 창] 패널이 준비한 초안이면 제목 케이스 ID·첨부를 채우고 초안함에 저장한다 (FR-304·505). */
async function runComposeFinalize() {
  const item = Office.context.mailbox.item as Office.MessageCompose;
  const retry = document.getElementById("compose-retry")!;
  retry.hidden = true;
  setStatus("", "info");
  try {
    const pending = await findPendingDraft(item);
    if (!pending) {
      setText("compose-state", "LEONA 패널에서 연 초안이 아닙니다. 받은 메일의 LEONA 패널에서 [초안 열기]를 먼저 눌러 주세요.");
      return;
    }
    await finalizeDraft(item, pending, (step) => setText("compose-state", step));
    setText("compose-state", `${pending.caseId} 초안을 초안함에 저장했습니다.`);
    setStatus("내용을 확인한 뒤 직접 [보내기]를 눌러 주세요. 자동으로 보내지 않습니다.", "success");
  } catch (error) {
    setText("compose-state", "초안을 마무리하지 못했습니다.");
    setStatus(messageOf(error), "error");
    retry.hidden = false;
  }
}

async function loadCurrentItem() {
  const token = ++loadToken;
  current = null;
  resetCaseArea();

  if (!Office.context.mailbox.item) {
    setStatus("메일을 선택하면 정보가 표시됩니다.", "info");
    return;
  }

  setStatus("메일 정보를 읽는 중…", "info");
  try {
    const snapshot = await readSelectedMail();
    if (token !== loadToken) return;
    renderMail(snapshot);

    const emlBase64 = await readSelectedMailRaw().catch(() => null);
    if (token !== loadToken) return;
    current = { snapshot, emlBase64, actor: currentUserEmail() };

    setStatus("케이스 등록 여부를 확인하는 중…", "info");
    const match = await matchReply(current);
    if (token !== loadToken) return;
    renderCaseState(match.alreadyRegisteredCaseId, match.candidates);
    renderSentState(snapshot);
    setStatus("", "info");
  } catch (error) {
    if (token !== loadToken) return;
    setStatus(messageOf(error), "error");
    // 메일은 읽었지만 백엔드가 꺼져 있는 경우에도 케이스 생성은 다시 시도할 수 있게 둔다
    if (current) showCreateButton(true);
  }
}

// ---------------------------------------------------------------- 버튼 동작

async function onCreateCase() {
  if (!current) return;
  setBusy(true);
  setStatus("케이스를 만드는 중…", "info");
  try {
    const result = await createCase(current);
    const caseId = result.case.caseId;
    renderCaseState(caseId, []);
    setStatus(
      result.duplicate
        ? `이미 등록된 메일입니다. 기존 케이스 ${caseId}를 표시합니다.`
        : `케이스 ${caseId}를 생성했습니다.`,
      "success"
    );
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
}

async function onMerge(caseId: string) {
  if (!current) return;
  setBusy(true);
  setStatus(`${caseId} 케이스에 회신을 병합하는 중…`, "info");
  try {
    await mergeReply(caseId, current);
    renderCaseState(caseId, []);
    setStatus(`${caseId} 케이스에 회신을 병합했습니다. 추출·검증을 다시 실행했습니다.`, "success");
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
}

/**
 * 담당자 본인이 보낸 [케이스ID] 메일이면 [발송 기록] 버튼을 보여준다 (FR-505).
 * 보내기 이벤트(OnMessageSend)가 동작하지 않는 환경을 위한 보조 경로다. 기록은 메일의 실제 발송 시각으로 남는다.
 */
function renderSentState(mail: MailSnapshot) {
  const caseId = mail.subject.match(CASE_ID_PATTERN)?.[0];
  const mine = mail.from.email.toLowerCase() === currentUserEmail().toLowerCase();
  const section = document.getElementById("sent-section")!;
  section.hidden = !(caseId && mine);
  if (!section.hidden) {
    setText("sent-state", `담당자가 보낸 ${caseId} 케이스 메일입니다. 발송 이력에 남기려면 누르세요.`);
    sentCaseId = caseId!;
    // 보낸 메일은 새 케이스 대상이 아니다. 이미 케이스에 등록된 메일이면 그 안내는 그대로 둔다
    showCreateButton(false);
    if (!currentCaseId) setText("case-state", "");
  }
}

async function onRecordSent() {
  if (!current || !sentCaseId) return;
  const mail = current.snapshot;
  // 보낸 메일의 '생성 시각'은 초안을 처음 만든 시각이다. 실제 발송 시각은 원문 Date 헤더(백엔드가 읽음),
  // 원문이 없으면 Outlook의 마지막 수정 시각(발송 시점)을 쓴다.
  setBusy(true);
  try {
    const item = Office.context.mailbox.item as Office.MessageRead;
    const fallback = toIsoOrNow(item.dateTimeModified ?? item.dateTimeCreated);
    const result = await recordMailEvent(sentCaseId, "MAIL_SENT", mail.subject, currentUserEmail(), undefined, {
      internetMessageId: mail.internetMessageId,
      occurredAt: fallback,
      emlBase64: current.emlBase64,
    });
    const sent = [...result.events].reverse().find((e) => e.eventType === "MAIL_SENT");
    const sentAt = sent?.detail?.sentAt ?? fallback;
    setStatus(`${sentCaseId} 발송 이력을 기록했습니다 (발송 시각 ${new Date(sentAt).toLocaleString("ko-KR")}).`, "success");
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
}

/** Outlook 날짜 값(Date 또는 문자열)을 ISO 문자열로. 읽을 수 없으면 지금 시각. */
function toIsoOrNow(value: unknown): string {
  const date = value instanceof Date ? value : new Date(String(value));
  return Number.isNaN(date.getTime()) ? new Date().toISOString() : date.toISOString();
}

async function onSupplementDraft() {
  if (!currentCaseId) return;
  const caseId = currentCaseId;
  const questions = (document.getElementById("questions") as HTMLTextAreaElement).value.split("\n");
  await openDraft(
    () => supplementDraft(caseId, questions, currentUserEmail()),
    // 보완 요청은 원래 메일에 대한 '회신'으로 연다 → 스레드가 이어져 화주 회신을 헤더로도 찾을 수 있다 (FR-104)
    (draft) => {
      savePendingDraft({ caseId, kind: "SUPPLEMENT", subject: `RE: ${draft.subject}`, attachments: [] });
      Office.context.mailbox.item!.displayReplyForm({ htmlBody: draft.htmlBody });
    }
  );
}

async function onQuoteDraft() {
  if (!currentCaseId) return;
  const caseId = currentCaseId;
  await openDraft(
    () => quoteDraft(caseId, currentUserEmail()),
    (draft) => {
      // 첨부는 작성 창이 열린 뒤 이벤트 처리기가 붙인다 (Outlook 웹은 localhost 파일 주소를 직접 못 가져온다)
      savePendingDraft({ caseId, kind: "QUOTE", subject: draft.subject, attachments: draft.attachments });
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: draft.to.map((a) => a.email),
        subject: draft.subject,
        htmlBody: draft.htmlBody,
      });
    }
  );
}

/**
 * 백엔드가 만든 초안 내용으로 Outlook 작성 창을 띄운다.
 * 작성 창이 열리면 이벤트 처리기(commands)가 제목·첨부를 마무리하고 초안함에 저장한다.
 * 보내기 버튼은 담당자가 직접 누른다 (자동 발송 없음).
 */
async function openDraft(load: () => Promise<Draft>, open: (draft: Draft) => void) {
  setBusy(true);
  setStatus("초안을 준비하는 중…", "info");
  try {
    open(await load());
    setStatus("작성 창을 열었습니다. 초안함에 저장되며, 내용을 확인한 뒤 직접 보내기를 눌러 주세요.", "success");
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
}

// ---------------------------------------------------------------- 화면 그리기

function renderMail(mail: MailSnapshot) {
  setText("subject", mail.subject || "(제목 없음)");
  setText("from", mail.from.name ? `${mail.from.name} <${mail.from.email}>` : mail.from.email);
  setText("received-at", mail.receivedAt ? new Date(mail.receivedAt).toLocaleString("ko-KR") : "-");
  setText("body-preview", toPreview(mail.bodyText));
  setText("conversation-id", mail.conversationId || "-");
  setText("internet-message-id", mail.internetMessageId || "-");
  setText("snapshot-json", JSON.stringify(mail, null, 2));
}

function resetCaseArea() {
  currentCaseId = null;
  sentCaseId = null;
  document.getElementById("sent-section")!.hidden = true;
  document.getElementById("files-section")!.hidden = true;
  setText("case-state", "");
  showCreateButton(false);
  document.getElementById("draft-section")!.hidden = true;
  document.getElementById("reply-section")!.hidden = true;
  document.getElementById("reply-candidates")!.replaceChildren();
}

/** 등록 여부·회신 후보에 따라 [케이스 생성] / [병합] 버튼을 보여준다. */
function renderCaseState(registeredCaseId: string | null, candidates: ReplyCandidate[]) {
  resetCaseArea();

  if (registeredCaseId) {
    currentCaseId = registeredCaseId;
    setText("case-state", `이 메일은 케이스 ${registeredCaseId}에 등록되어 있습니다.`);
    document.getElementById("draft-section")!.hidden = false;
    renderCaseFiles(registeredCaseId);
    return;
  }

  showCreateButton(true);
  setText(
    "case-state",
    candidates.length
      ? "기존 케이스의 회신으로 보입니다. 맞는 케이스에 병합하거나, 새 요청이면 케이스를 생성하세요."
      : "아직 케이스로 등록되지 않은 메일입니다."
  );
  renderReplySection(candidates);
}

/** 케이스에 저장된 메일 원문·첨부를 다시 열 수 있게 링크로 보여준다 (FR-103). */
async function renderCaseFiles(caseId: string) {
  const token = loadToken;
  const section = document.getElementById("files-section")!;
  const list = document.getElementById("case-files")!;
  try {
    const detail = await getCase(caseId);
    if (token !== loadToken || currentCaseId !== caseId) return;
    list.replaceChildren();
    for (const mail of detail.mails) {
      const item = document.createElement("li");
      const meta = document.createElement("div");
      meta.className = "file-meta";
      const when = mail.snapshot.receivedAt ? new Date(mail.snapshot.receivedAt).toLocaleString("ko-KR") : "";
      meta.textContent = `${mail.index}. ${mail.role === "ORIGINAL" ? "요청 메일" : "회신"} · ${when}`;
      item.appendChild(meta);
      if (mail.rawUrl) item.appendChild(fileLink("원문(.eml)", mail.rawUrl));
      for (const file of mail.attachmentFiles) item.appendChild(fileLink(file.filename, file.url));
      list.appendChild(item);
    }
    section.hidden = detail.mails.length === 0;
  } catch {
    section.hidden = true; // 자료 목록을 못 불러와도 다른 기능은 그대로 쓴다
  }
}

function fileLink(label: string, url: string): HTMLAnchorElement {
  const link = document.createElement("a");
  link.textContent = label;
  link.href = apiUrl(url);
  link.target = "_blank";
  link.rel = "noopener";
  return link;
}

function renderReplySection(candidates: ReplyCandidate[]) {
  const section = document.getElementById("reply-section")!;
  const list = document.getElementById("reply-candidates")!;
  section.hidden = false;
  // 담당자 자신이 보낸 메일(참조 사본 등)을 화주 회신으로 잘못 병합하지 않도록 경고한다
  const fromMe = current?.snapshot.from.email.toLowerCase() === currentUserEmail().toLowerCase();
  document.getElementById("reply-warning")!.hidden = !(fromMe && candidates.length > 0);

  for (const candidate of candidates) {
    const item = document.createElement("li");

    const title = document.createElement("div");
    title.className = "candidate__title";
    title.textContent = `${candidate.caseId} · ${candidate.subject}`;

    const reason = document.createElement("div");
    reason.className = "candidate__reason";
    reason.textContent = `근거: ${candidate.reasons.map((r) => REASON_LABELS[r] ?? r).join(", ")}`;

    const button = document.createElement("button");
    button.className = "button";
    button.textContent = `${candidate.caseId}에 병합`;
    button.onclick = () => onMerge(candidate.caseId);

    item.append(title, reason, button);
    list.appendChild(item);
  }
  fillCaseSelect();
}

/** 자동 식별이 실패했을 때 담당자가 직접 고를 수 있도록 최근 케이스 목록을 채운다 (FR-104). */
async function fillCaseSelect() {
  const select = document.getElementById("case-select") as HTMLSelectElement;
  select.replaceChildren(new Option("케이스를 선택하세요", ""));
  try {
    for (const c of await listCases()) {
      select.appendChild(new Option(`${c.caseId} · ${c.subject}`, c.caseId));
    }
  } catch {
    // 목록을 못 불러와도 자동 후보·케이스 생성은 쓸 수 있으므로 조용히 넘어간다
  }
}

/** 줄바꿈·연속 공백을 한 칸으로 줄이고 앞부분만 잘라낸다. */
function toPreview(text: string): string {
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length > PREVIEW_LENGTH ? `${flat.slice(0, PREVIEW_LENGTH)}…` : flat || "(본문 없음)";
}

// ---------------------------------------------------------------- 공통

function showCreateButton(visible: boolean) {
  document.getElementById("create-case")!.hidden = !visible;
}

function setBusy(busy: boolean) {
  document.querySelectorAll<HTMLButtonElement>("button").forEach((b) => (b.disabled = busy));
}

/** 메일 내용은 외부에서 온 값이므로 HTML로 해석하지 않고 textContent로만 넣는다(스크립트 삽입 방지). */
function setText(id: string, value: string) {
  document.getElementById(id)!.textContent = value;
}

function setStatus(message: string, kind: "info" | "success" | "error") {
  const status = document.getElementById("status")!;
  status.textContent = message;
  status.className = `status status--${kind}`;
  status.hidden = message === "";
}

function messageOf(error: unknown): string {
  if (error instanceof ApiError) return error.message;
  return `메일 정보를 읽지 못했습니다: ${error instanceof Error ? error.message : String(error)}`;
}
