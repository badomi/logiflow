/* global document, Office */

import { call, finalizeDraft, findPendingDraft } from "../shared/finalize";
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
  rerunExtraction,
  fetchFileBase64,
} from "./api";
import type { Draft, ExtractionState, MailInput, ReplyCandidate } from "./api";
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
  document.getElementById("record-sent")!.onclick = () => onRecordSent(false);
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

/**
 * [작성 창] 담당자가 Outlook [회신]으로 직접 연 작성 창.
 * 제목의 케이스 ID·같은 대화로 케이스 메일인지 확인하고, 맞으면 읽기 화면과 같은 '메일 초안' 카드를 보여준다.
 * 같은 버튼이 여기서는 새 창을 여는 대신 지금 쓰는 메일을 채운다. 담당자가 눌러야만 바꾸고, 케이스와 관계없는 메일은 건드리지 않는다.
 */
async function offerDirectCompose(item: Office.MessageCompose) {
  const card = document.getElementById("draft-section")!;
  card.hidden = true;
  const subject = await call<string>((done) => item.subject.getAsync(done));
  const match = await matchReply({
    snapshot: {
      subject,
      from: { name: "", email: "" },
      to: [],
      cc: [],
      receivedAt: "",
      bodyText: "",
      conversationId: item.conversationId ?? "",
      internetMessageId: "",
      inReplyTo: "",
      references: [],
      attachments: [],
    },
    emlBase64: null,
    actor: currentUserEmail(),
  });
  const caseId = match.candidates[0]?.caseId;
  if (!caseId) {
    setText("compose-state", "케이스와 연결된 메일이 아니라서 바꾸지 않았습니다. 그대로 쓰고 보내면 됩니다.");
    return;
  }
  // 읽기 화면과 같은 카드(케이스 상태·케이스 자료·메일 초안)를 숨겨진 읽기 영역 밖으로 옮겨 그대로 보여준다
  const compose = document.getElementById("compose-section")!;
  compose.hidden = true;
  const caseCard = document.getElementById("case-section")!;
  const files = document.getElementById("files-section")!;
  compose.after(caseCard, files, card);
  currentCaseId = caseId;
  setText("case-state", `작성 중인 이 메일은 케이스 ${caseId}에 대한 회신입니다. 아래 버튼을 누르기 전에는 메일을 바꾸지 않습니다.`);
  renderCaseFiles(caseId);
  card.hidden = false;
  document.getElementById("supplement-draft")!.onclick = () => runDirectCompose(item, caseId, "SUPPLEMENT");
  document.getElementById("quote-draft")!.onclick = () => runDirectCompose(item, caseId, "QUOTE");
}

/** 직접 연 작성 창을 보완 요청·견적서 송부 초안으로 만든다: 제목·본문·첨부를 채우고 초안함에 저장 (FR-304·505). 보내지 않는다. */
async function runDirectCompose(item: Office.MessageCompose, caseId: string, kind: "SUPPLEMENT" | "QUOTE") {
  setBusy(true);
  setStatus("초안을 준비하는 중…", "info");
  try {
    const actor = currentUserEmail();
    const questions = (document.getElementById("questions") as HTMLTextAreaElement).value.split("\n");
    const draft = kind === "SUPPLEMENT" ? await supplementDraft(caseId, questions, actor) : await quoteDraft(caseId, actor);
    const current = await call<string>((done) => item.subject.getAsync(done));
    const isReply = /^(re|회신)\s*:/i.test(current);
    // 회신 창이면 받는 사람은 Outlook이 이미 채웠다. 비어 있을 때만 케이스의 화주 주소를 넣는다
    const to = await call<Office.EmailAddressDetails[]>((done) => item.to.getAsync(done));
    if (to.length === 0) await call<void>((done) => item.to.setAsync(draft.to.map((a) => a.email), done));
    // 담당자가 이미 쓴 내용·인용문은 그대로 두고 맨 위에 넣는다
    await call<void>((done) => item.body.prependAsync(draft.htmlBody, { coercionType: Office.CoercionType.Html }, done));
    await finalizeDraft(
      item,
      { caseId, kind, subject: isReply ? `RE: ${draft.subject}` : draft.subject, attachments: draft.attachments, createdAt: Date.now() },
      (step) => setStatus(step, "info")
    );
    document.getElementById("draft-section")!.hidden = true;
    setText("case-state", `${caseId} ${kind === "SUPPLEMENT" ? "보완 요청" : "견적서 송부"} 초안을 초안함에 저장했습니다.`);
    setStatus("내용을 확인한 뒤 직접 [보내기]를 눌러 주세요. 자동으로 보내지 않습니다.", "success");
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
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
      await offerDirectCompose(item);
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
    const autoRecord = renderSentState(snapshot);
    setStatus("", "info");
    if (autoRecord) await onRecordSent(true);
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
    setStatus(`${caseId} 케이스에 회신을 병합했습니다. 추출·검증을 다시 실행합니다.`, "success");
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
}

/**
 * 담당자 본인이 보낸 [케이스ID] 메일이면 발송 이력을 남긴다 (FR-505).
 * 보내기 이벤트(OnMessageSend)가 동작하지 않는 환경을 위한 경로다. 기록은 메일의 실제 발송 시각으로 남는다.
 * 돌려주는 값: true면 메일을 열자마자 자동 기록한다 (같은 메일은 서버가 한 번만 기록).
 * 받는 사람에 본인이 있으면(계정 1개로 화주·담당자를 함께 테스트) 화주 회신과 구분할 수 없어 [발송 기록] 버튼으로 남긴다.
 */
function renderSentState(mail: MailSnapshot): boolean {
  const me = currentUserEmail().toLowerCase();
  const caseId = mail.subject.match(CASE_ID_PATTERN)?.[0];
  const mine = mail.from.email.toLowerCase() === me;
  const section = document.getElementById("sent-section")!;
  section.hidden = !(caseId && mine);
  if (section.hidden) return false;

  const toMe = [...mail.to, ...mail.cc].some((a) => a.email.toLowerCase() === me);
  sentCaseId = caseId!;
  // 보낸 메일은 새 케이스 대상이 아니다. 이미 케이스에 등록된 메일이면 그 안내는 그대로 둔다
  showCreateButton(false);
  // 케이스 등록 안내가 없으면 빈 카드가 남으므로 카드째 숨긴다
  if (!currentCaseId) {
    setText("case-state", "");
    document.getElementById("case-section")!.hidden = true;
  }
  document.getElementById("record-sent")!.hidden = !toMe;
  setText(
    "sent-state",
    toMe
      ? `담당자가 보낸 ${caseId} 케이스 메일입니다. 받는 사람에 본인이 있어 화주 회신과 구분할 수 없으므로, 보낸 메일이 맞으면 누르세요.`
      : `담당자가 보낸 ${caseId} 케이스 메일입니다. 발송 이력을 기록하는 중…`
  );
  return !toMe;
}

async function onRecordSent(auto = false) {
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
    const sentEvents = [...result.events].reverse().filter((e) => e.eventType === "MAIL_SENT");
    const sent = sentEvents.find((e) => e.detail?.messageId === mail.internetMessageId) ?? sentEvents[0];
    const sentAt = sent?.detail?.sentAt ?? fallback;
    const who = sent?.actor ?? currentUserEmail();
    // 보낸 메일의 Outlook 생성 시각은 초안을 만든 시각이라, 실제 발송 시각(원문 Date 헤더)으로 바꿔 보여준다
    setText("received-at-label", "발송 시각");
    setText("received-at", new Date(sentAt).toLocaleString("ko-KR"));
    setText("sent-state", `${sentCaseId} 발송 이력: ${new Date(sentAt).toLocaleString("ko-KR")} · ${who}`);
    setStatus(`${sentCaseId} 발송 이력을 ${auto ? "자동으로 " : ""}기록했습니다 (발송 시각 ${new Date(sentAt).toLocaleString("ko-KR")}).`, "success");
    // LEONA가 보낸 메일 자체는 회신이 아니므로 병합 안내를 숨긴다 (서버도 병합을 거부한다)
    if (auto) document.getElementById("reply-section")!.hidden = true;
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
  // Mailbox 1.15 이상(현재 Outlook 웹): 고객 메일에 대한 '회신'으로 열면서 견적서를 바로 첨부한다 → 버튼 한 번으로 끝
  if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
    await openQuoteReplyWithAttachments(caseId);
    return;
  }
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
 * 견적서 송부 (FR-505): 지금 열린 고객 메일에 대한 회신 창을 견적서 파일을 붙인 채로 연다.
 * 새 메일 창은 공개 URL로만 첨부할 수 있지만(localhost 불가), 회신 창은 파일 내용(base64)을 직접 넘길 수 있다.
 * 보내기는 담당자가 직접 누른다 (자동 발송 없음).
 */
async function openQuoteReplyWithAttachments(caseId: string) {
  setBusy(true);
  setStatus("견적서를 불러오는 중…", "info");
  try {
    const draft = await quoteDraft(caseId, currentUserEmail());
    const attachments: Office.ReplyFormAttachment[] = [];
    for (const file of draft.attachments) {
      attachments.push({
        type: Office.MailboxEnums.AttachmentType.Base64,
        name: file.filename,
        base64file: await fetchFileBase64(file.url),
      });
    }
    // 작성 창에서 LEONA를 열면 제목에 케이스 번호를 넣고 초안함에 저장한다 (이미 붙은 견적서는 다시 붙이지 않음)
    savePendingDraft({ caseId, kind: "QUOTE", subject: `RE: ${draft.subject}`, attachments: draft.attachments });
    const item = Office.context.mailbox.item as Office.MessageRead;
    await call<void>((done) => item.displayReplyFormAsync({ htmlBody: draft.htmlBody, attachments }, done));
    setStatus(
      `견적서 ${attachments.length}개를 첨부한 회신 창을 열었습니다. 첨부·내용을 확인하고 직접 보내기를 눌러 주세요. ` +
        "(제목에 케이스 번호를 넣고 초안함에 저장하려면 그 창에서 [앱] → LEONA)",
      "success"
    );
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
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
    const draft = await load();
    open(draft);
    // 작성 창만 열어서는 첨부·초안함 저장이 되지 않는다 (Outlook 웹이 localhost 파일을 못 가져옴) → 한 단계 더 안내
    const what = draft.attachments.length ? "견적서가 첨부되고 초안함에 저장됩니다" : "초안함에 저장됩니다";
    setStatus(`새로 열린 메일 창에서 [앱] → LEONA를 열어야 ${what}. 그다음 내용을 확인하고 직접 보내기를 눌러 주세요.`, "info");
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
  setText("received-at-label", "수신 시각");
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
  document.getElementById("extraction")!.hidden = true;
  setText("case-state", "");
  document.getElementById("case-section")!.hidden = false;
  showCreateButton(false);
  document.getElementById("draft-section")!.hidden = true;
  document.getElementById("reply-section")!.hidden = true;
  document.getElementById("reply-candidates")!.replaceChildren();
  const questions = document.getElementById("questions") as HTMLTextAreaElement;
  questions.value = ""; // 다른 메일로 넘어가면 이전 케이스 문항·'직접 고침' 표시를 지운다
  delete questions.dataset.edited;
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
  // 회신 후보가 있으면 [병합]이 주 버튼이다. [케이스 생성]은 '예전 대화로 온 새 요청'일 때만 쓰므로 보조 버튼으로 내린다
  const hasCandidate = candidates.length > 0;
  const create = document.getElementById("create-case")!;
  create.classList.toggle("button--primary", !hasCandidate);
  create.textContent = hasCandidate ? "새 요청이면: 케이스 생성" : "케이스 생성";
  setText(
    "case-state",
    hasCandidate
      ? `기존 케이스 ${candidates[0].caseId}의 회신으로 보입니다. 아래 [병합]을 누르세요. 예전 대화로 온 새 요청일 때만 케이스를 새로 만드세요.`
      : "아직 케이스로 등록되지 않은 메일입니다."
  );
  renderReplySection(candidates);
}

/** 추출 상태를 다시 확인하는 간격과 최대 대기 시간 — 담당자가 버튼을 누른 뒤에만, 우리 백엔드에만 묻는다 (사서함 폴링 아님) */
const EXTRACTION_CHECK_MS = 2000;
const EXTRACTION_WAIT_LIMIT_MS = 90_000;

/**
 * 케이스 정보(추출 상태·자료)를 불러와 그린다. 추출이 도는 중이면 끝날 때까지 몇 초마다 다시 확인한다.
 */
async function renderCaseFiles(caseId: string, waitedMs = 0) {
  const token = loadToken;
  const section = document.getElementById("files-section")!;
  const list = document.getElementById("case-files")!;
  try {
    const detail = await getCase(caseId);
    if (token !== loadToken || currentCaseId !== caseId) return;
    renderExtraction(detail.extraction, waitedMs);
    if (detail.extraction.state === "running" && waitedMs < EXTRACTION_WAIT_LIMIT_MS) {
      setTimeout(() => renderCaseFiles(caseId, waitedMs + EXTRACTION_CHECK_MS), EXTRACTION_CHECK_MS);
    }
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

/** 항목 추출 상태를 보여준다 (FR-101: 추출 수행, 60초 이내). */
function renderExtraction(extraction: ExtractionState, waitedMs: number) {
  const box = document.getElementById("extraction")!;
  const result = document.getElementById("extraction-result")!;
  result.replaceChildren();
  box.hidden = extraction.state === "none";

  const seconds = extraction.elapsedMs != null ? (extraction.elapsedMs / 1000).toFixed(1) : null;
  const limit = extraction.withinLimit === false ? " · ⚠ 60초 기준 초과" : "";
  const messages: Record<ExtractionState["state"], string> = {
    none: "",
    running:
      waitedMs >= EXTRACTION_WAIT_LIMIT_MS
        ? "항목 추출이 오래 걸리고 있습니다. 잠시 뒤 패널을 다시 열어 확인해 주세요."
        : `항목 추출 중… (${Math.round(waitedMs / 1000)}초)`,
    done: `항목 추출 완료 (${seconds}초${limit})`,
    "not-connected": `항목 추출이 꺼져 있어 실행만 기록했습니다 (${seconds}초) — server\\.env의 EXTRACTION_MODE를 확인하세요`,
    failed: `항목 추출 실패: ${extraction.error ?? "알 수 없는 오류"} — 케이스와 원문은 저장되어 있습니다`,
  };
  setText("extraction-state", messages[extraction.state]);
  const rerun = document.getElementById("rerun-extraction") as HTMLButtonElement;
  rerun.hidden = extraction.state === "none" || extraction.state === "running";
  rerun.onclick = () => onRerunExtraction();

  const fields = extractedFields(extraction.result);
  for (const [field, info] of Object.entries(fields)) {
    if (info.value == null || info.status === "missing") continue;
    addRow(result, field, info.status === "review" ? `${String(info.value)} (확인 필요)` : String(info.value), info.evidence);
  }

  const validation = (extraction.validation ?? {}) as ValidationView;
  const hold = validation.hold ?? [];
  if (hold.length) addRow(result, "보류 사유", hold.map((h) => h.message).join(" / "));
  const questions = validation.questions ?? [];
  if (questions.length) addRow(result, "보완 필요", `${questions.length}개 항목 — 아래 '보완 요청 문항'에 채워 두었습니다`);
  if (validation.quote?.totals) {
    const totals = Object.entries(validation.quote.totals).map(([ccy, v]) => `${ccy} ${v.toLocaleString()}`).join(" + ");
    addRow(result, "견적서", `v${validation.quote.version} 생성 (${totals}) — [견적서 송부 초안]으로 보낼 수 있습니다`);
  }
  const llm = (extraction.result as { llm?: { used?: boolean; error?: string } } | null)?.llm;
  if (llm?.error) addRow(result, "LLM", "사용 못 함 — 키워드 규칙 결과만 반영 (" + llm.error.slice(0, 60) + ")");

  fillQuestions(questions);
}

interface FieldView {
  value: unknown;
  status?: "filled" | "review" | "missing";
  evidence?: string | null;
}

interface ValidationView {
  questions?: string[];
  hold?: { code: string; message: string }[];
  quote?: { version: number; totals: Record<string, number> } | null;
}

/** B트랙 결과 { fields: { 필드: {value, status, evidence} } } — 예전 형식 { 필드: 값 }도 받는다 */
function extractedFields(result: Record<string, unknown> | null): Record<string, FieldView> {
  if (!result) return {};
  const source = (result.fields && typeof result.fields === "object" ? result.fields : result) as Record<string, unknown>;
  const out: Record<string, FieldView> = {};
  for (const [field, raw] of Object.entries(source)) {
    out[field] = raw && typeof raw === "object" && "value" in raw ? (raw as FieldView) : { value: raw, status: "filled" };
  }
  return out;
}

function addRow(list: HTMLElement, label: string, text: string, title?: string | null) {
  const dt = document.createElement("dt");
  dt.textContent = label;
  const dd = document.createElement("dd");
  dd.textContent = text;
  if (title) dd.title = `근거: ${title}`; // 마우스를 올리면 추출 근거 원문 (FR-202)
  list.append(dt, dd);
}

/** [다시 추출] 설정을 바꾼 뒤나 LLM 시간 초과로 규칙 결과만 나왔을 때 추출→검증을 다시 돌린다 (NFR-03) */
async function onRerunExtraction() {
  if (!currentCaseId) return;
  const caseId = currentCaseId;
  setBusy(true);
  try {
    await rerunExtraction(caseId, currentUserEmail());
    const questions = document.getElementById("questions") as HTMLTextAreaElement;
    delete questions.dataset.edited; // 새 검증 결과 문항으로 다시 채운다
    setStatus("다시 추출을 시작했습니다. 끝나면 결과가 바뀝니다.", "info");
    await renderCaseFiles(caseId);
  } catch (error) {
    setStatus(messageOf(error), "error");
  } finally {
    setBusy(false);
  }
}

/** 검증 결과 문항을 입력 칸에 채운다. 담당자가 직접 고친 뒤에는 덮어쓰지 않는다. */
function fillQuestions(questions: string[]) {
  const box = document.getElementById("questions") as HTMLTextAreaElement | null;
  if (!box || box.dataset.edited === "1") return;
  box.value = questions.join("\n");
  if (!box.dataset.watch) {
    box.dataset.watch = "1";
    box.addEventListener("input", () => (box.dataset.edited = "1"));
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

  candidates.forEach((candidate, index) => {
    const item = document.createElement("li");

    const title = document.createElement("div");
    title.className = "candidate__title";
    title.textContent = `${candidate.caseId} · ${candidate.subject}`;

    const reason = document.createElement("div");
    reason.className = "candidate__reason";
    reason.textContent = `근거: ${candidate.reasons.map((r) => REASON_LABELS[r] ?? r).join(", ")}`;

    const button = document.createElement("button");
    // 가장 근거가 강한 첫 후보를 주 버튼으로. 담당자 본인이 보낸 메일이면 주 버튼으로 권하지 않는다
    button.className = index === 0 && !fromMe ? "button button--primary" : "button";
    button.textContent = `${candidate.caseId}에 병합`;
    button.onclick = () => onMerge(candidate.caseId);

    item.append(title, reason, button);
    list.appendChild(item);
  });
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
