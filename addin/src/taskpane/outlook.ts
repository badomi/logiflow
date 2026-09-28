/* global document, Office */

import { ApiError, createCase, listCases, matchReply, mergeReply } from "./api";
import type { MailInput, ReplyCandidate } from "./api";
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
/** 메일을 빠르게 바꿀 때, 늦게 끝난 이전 메일 결과가 화면을 덮어쓰지 않게 하는 번호 */
let loadToken = 0;

// Office.js 준비가 끝나면 한 번 실행된다. Outlook 안에서 열렸을 때만 패널을 보여준다.
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "block";
  document.getElementById("create-case")!.onclick = onCreateCase;
  document.getElementById("merge-selected")!.onclick = () => {
    const select = document.getElementById("case-select") as HTMLSelectElement;
    if (select.value) onMerge(select.value);
  };

  // 패널을 고정(pin)해 두면, 담당자가 다른 메일을 선택할 때마다 새로 읽는다 (사용자 선택 이벤트 — 폴링 아님)
  if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => loadCurrentItem());
  }
  loadCurrentItem();
});

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
  setText("case-state", "");
  showCreateButton(false);
  document.getElementById("reply-section")!.hidden = true;
  document.getElementById("reply-candidates")!.replaceChildren();
}

/** 등록 여부·회신 후보에 따라 [케이스 생성] / [병합] 버튼을 보여준다. */
function renderCaseState(registeredCaseId: string | null, candidates: ReplyCandidate[]) {
  resetCaseArea();

  if (registeredCaseId) {
    setText("case-state", `이 메일은 케이스 ${registeredCaseId}에 등록되어 있습니다.`);
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

function renderReplySection(candidates: ReplyCandidate[]) {
  const section = document.getElementById("reply-section")!;
  const list = document.getElementById("reply-candidates")!;
  section.hidden = false;

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
