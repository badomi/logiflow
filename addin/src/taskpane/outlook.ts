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
  searchCases,
  getCaseFields,
} from "./api";
import type { CaseDetail, CaseField, CaseSearchFilters, CaseSearchResult, Draft, ExtractionState, MailInput, ReplyCandidate } from "./api";
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
  initializeCaseConsole();

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
  compose.after(caseCard, card, files);
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
  document.getElementById("mail-overview")!.hidden = true;
  document.getElementById("mail-source-section")!.hidden = true;

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
  document.getElementById("mail-overview")!.hidden = false;
  document.getElementById("mail-source-section")!.hidden = false;
}

function resetCaseArea() {
  currentCaseId = null;
  sentCaseId = null;
  document.getElementById("sent-section")!.hidden = true;
  document.getElementById("files-section")!.hidden = true;
  document.getElementById("extraction")!.hidden = true;
  document.getElementById("mail-validation-issues")!.replaceChildren();
  document.getElementById("mail-validation-issues")!.hidden = true;
  document.getElementById("mail-case-status")!.hidden = true;
  (document.getElementById("mail-extracted-details") as HTMLDetailsElement).open = false;
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
    const badge = document.getElementById("mail-case-status")!;
    badge.textContent = detail.status;
    badge.className = consoleBadge(detail.status).className;
    badge.hidden = false;
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
  const summary = document.getElementById("mail-validation-summary")!;
  result.replaceChildren();
  summary.replaceChildren();
  box.hidden = extraction.state === "none";

  const seconds = extraction.elapsedMs != null ? (extraction.elapsedMs / 1000).toFixed(1) : null;
  const limit = extraction.withinLimit === false ? " · ⚠ 60초 기준 초과" : "";
  const messages: Record<ExtractionState["state"], string> = {
    none: "",
    running:
      waitedMs >= EXTRACTION_WAIT_LIMIT_MS
        ? "항목 추출이 오래 걸리고 있습니다. 잠시 뒤 패널을 다시 열어 확인해 주세요."
        : `항목 추출 중… (${Math.round(waitedMs / 1000)}초)`,
    done: `항목 추출 완료${seconds !== null ? ` (${seconds}초${limit})` : limit}`,
    "not-connected": `항목 추출이 꺼져 있어 실행만 기록했습니다 (${seconds}초) — server\\.env의 EXTRACTION_MODE를 확인하세요`,
    failed: `항목 추출 실패: ${extraction.error ?? "알 수 없는 오류"} — 케이스와 원문은 저장되어 있습니다`,
  };
  setText("extraction-state", messages[extraction.state]);
  box.dataset.state = extraction.state;
  const rerun = document.getElementById("rerun-extraction") as HTMLButtonElement;
  rerun.hidden = extraction.state === "none" || extraction.state === "running";
  rerun.onclick = () => onRerunExtraction();

  const fields = extractedFields(extraction.result);
  for (const [field, info] of Object.entries(fields)) {
    if (info.value == null || info.status === "missing") continue;
    addRow(result, consoleFieldLabels[field] ?? field, info.status === "review" ? `${String(info.value)} (확인 필요)` : String(info.value), info.evidence);
  }
  const count = result.querySelectorAll("dt").length;
  setText("mail-field-count", `${count}개`);
  document.getElementById("mail-extracted-details")!.hidden = count === 0;

  const validation = (extraction.validation ?? {}) as ValidationView;
  const hold = validation.hold ?? [];
  const issues = document.getElementById("mail-validation-issues")!;
  const issuePanel = validationIssuesPanel(extraction.validation);
  issues.replaceChildren();
  if (issuePanel) issues.append(issuePanel);
  issues.hidden = !issuePanel;
  const questions = validation.questions ?? [];
  if (questions.length) addRow(summary, "보완 필요", `${questions.length}개 항목 — 아래 보완 요청 문항을 확인하세요.`);
  if (validation.quote?.totals) {
    const totals = Object.entries(validation.quote.totals).map(([ccy, v]) => `${ccy} ${v.toLocaleString()}`).join(" + ");
    addRow(summary, "견적서", `v${validation.quote.version} 생성 · ${totals}`);
  }
  const llm = (extraction.result as { llm?: { used?: boolean; error?: string } } | null)?.llm;
  if (llm?.error) addRow(summary, "LLM", "사용 못 함 — 키워드 규칙 결과만 반영 (" + llm.error.slice(0, 60) + ")");
  summary.hidden = summary.children.length === 0;
  summary.classList.toggle("mail-validation-attention", hold.length > 0 || questions.length > 0 || !!llm?.error);

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

/** 원문은 바꾸지 않고, 미리보기에서 줄바꿈을 유지한 채 앞부분만 보여준다. */
function toPreview(text: string): string {
  const flat = text.replace(/\r\n?/g, "\n").replace(/[^\S\n]+/g, " ").trim();
  return flat.length > PREVIEW_LENGTH ? `${flat.slice(0, PREVIEW_LENGTH)}…` : flat || "(본문 없음)";
}

// ---------------------------------------------------------------- 공통

function showCreateButton(visible: boolean) {
  document.getElementById("create-case")!.hidden = !visible;
}

function setBusy(busy: boolean) {
  document.querySelectorAll<HTMLButtonElement>("#mail-workspace button").forEach((b) => (b.disabled = busy));
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

// ---------------------------------------------------------------- D트랙: 케이스 조회 (FR-601·107·103·602)
// 조회 대상은 currentCaseId와 독립적이다. 목록에서 선택해도 A의 메일/초안 처리 대상은 바꾸지 않는다.
const consoleFilters: CaseSearchFilters = {
  customer: "", status: "", pol: "", pod: "", startDate: "", endDate: "", sort: "createdAt", order: "desc",
};
let consoleResult: CaseSearchResult | null = null;
let consoleCaseId: string | null = null;
let consoleListToken = 0;
let consoleDetailToken = 0;
let consolePage = 1;
let consoleTab = "fields";
let consoleListBusy = false;

const consoleFieldLabels: Record<string, string> = {
  customerName: "고객 회사", contactName: "고객 담당자", contactEmail: "회신 주소",
  commodity: "품목", qty: "수량", qtyUnit: "수량 단위", packing: "포장",
  boxL: "길이 (mm)", boxW: "너비 (mm)", boxH: "높이 (mm)", totalCbm: "전체 부피 (CBM)",
  grossWeightKg: "총중량 (kg)", pol: "출발항 (POL)", pod: "도착항 (POD)",
  cargoReadyDate: "화물 준비일", incoterms: "거래조건", containerType: "컨테이너",
  invoiceValue: "인보이스 금액", ccy: "인보이스 통화", paymentTerm: "결제조건",
  containerCount: "컨테이너 대수", carrier: "선사",
  volume: "부피·박스 규격",
};

/** 담당자용 검증 문구에서 요구사항 번호만 숨긴다. 서버 응답·원본 판단 로그는 변경하지 않는다. */
function validationDisplayMessage(message: string): string {
  return message
    .replace(/\(\s*N?FR-\d{2,3}(?:\s*[·,/]\s*(?:N?FR-)?\d{2,3})*\s*\)/gi, "")
    .replace(/\bN?FR-\d{2,3}\b(?:\s*[·,/]\s*(?:N?FR-)?\d{2,3}\b)*/gi, "")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

/** FR-301: C의 판정을 표시한다. 요구사항 번호만 숨기고 심각도나 보완 문항은 다시 판정하지 않는다. */
function validationIssuesPanel(value: unknown): HTMLElement | null {
  const validation = consoleRecord(value);
  const groups: [string, Record<string, unknown>[]][] = [
    ["누락·확인 항목", validation.issues], ["보류 사유", validation.hold],
  ].map(([label, items]) => [String(label), Array.isArray(items)
    ? items.map(consoleRecord).filter((item) => Object.keys(item).length > 0) : []]);
  const entries = groups.flatMap(([, items]) => items);
  if (!entries.length) return null;

  const levels = ["Block", "Warn", "Info"];
  const labels: Record<string, string> = { Block: "차단", Warn: "경고", Info: "참고", "미기록": "심각도 미기록" };
  const levelOf = (item: Record<string, unknown>) => levels.includes(String(item.severity)) ? String(item.severity) : "미기록";
  const badge = (level: string, count?: number) => consoleNode("span", `validation-level validation-level-${level === "미기록" ? "unknown" : level.toLowerCase()}`,
    count == null ? (level === "미기록" ? labels[level] : `${level} · ${labels[level]}`) : `${level} ${count}`);
  const section = consoleNode("section", "validation-results");
  section.setAttribute("aria-label", "누락 항목 및 검증 심각도");
  section.append(consoleNode("h3", "validation-title", "누락 항목·검증 결과"));
  const counts = consoleNode("div", "validation-counts");
  counts.setAttribute("aria-label", "심각도별 건수");
  for (const level of [...levels, "미기록"]) {
    const count = entries.filter((item) => levelOf(item) === level).length;
    if (count) counts.append(badge(level, count));
  }
  section.append(counts);
  for (const [title, items] of groups) {
    if (!items.length) continue;
    section.append(consoleNode("h4", "validation-group-title", title));
    const list = consoleNode("ul", "validation-list");
    for (const item of items) {
      const row = consoleNode("li", "validation-item");
      const heading = consoleNode("div", "validation-item-heading");
      const fields = typeof item.field === "string" ? [item.field]
        : Array.isArray(item.fields) ? item.fields.filter((field): field is string => typeof field === "string") : [];
      const label = fields.map((field) => consoleFieldLabels[field] || field).join(" · ");
      if (label) heading.append(consoleNode("strong", "validation-field", label));
      heading.append(badge(levelOf(item)));
      row.append(heading, consoleNode("p", "validation-message", typeof item.message === "string" ? validationDisplayMessage(item.message) : "사유 미기록"));
      list.append(row);
    }
    section.append(list);
  }
  return section;
}

function initializeCaseConsole() {
  document.getElementById("console-nav-mail")!.onclick = () => showConsoleView("mail");
  document.getElementById("console-nav-cases")!.onclick = () => loadConsoleList(consolePage);
  document.getElementById("console-return-mail")!.onclick = () => showConsoleView("mail");
  document.getElementById("console-back")!.onclick = () => {
    ++consoleDetailToken;
    showConsoleView("list");
    const selected = [...document.querySelectorAll<HTMLButtonElement>(".console-case-row")]
      .find((button) => button.dataset.caseId === consoleCaseId);
    selected?.focus();
  };
  document.getElementById("console-refresh")!.onclick = () => loadConsoleList(consolePage);
  document.getElementById("console-detail-refresh")!.onclick = () => {
    if (consoleCaseId) loadConsoleDetail(consoleCaseId, false);
  };
  document.getElementById("console-prev")!.onclick = () => {
    if (!consoleListBusy && consolePage > 1) loadConsoleList(consolePage - 1);
  };
  document.getElementById("console-next")!.onclick = () => {
    if (!consoleListBusy && consoleResult && consolePage * consoleResult.pageSize < consoleResult.total) {
      loadConsoleList(consolePage + 1);
    }
  };
  const form = document.getElementById("console-search") as HTMLFormElement;
  form.onsubmit = (event) => {
    event.preventDefault();
    const input = new FormData(form);
    const values = Object.fromEntries(Object.keys(consoleFilters).map((key) => [key, String(input.get(key) ?? "").trim()])) as unknown as CaseSearchFilters;
    if (values.startDate && values.endDate && values.startDate > values.endDate) {
      consoleMessage("console-list-message", "검색 시작일은 종료일보다 늦을 수 없습니다.", true);
      return;
    }
    Object.assign(consoleFilters, values);
    loadConsoleList(1);
  };
  form.onreset = () => {
    Object.assign(consoleFilters, { customer: "", status: "", pol: "", pod: "", startDate: "", endDate: "", sort: "createdAt", order: "desc" });
    loadConsoleList(1);
  };
  const tabs = [...document.querySelectorAll<HTMLButtonElement>("[data-console-tab]")];
  tabs.forEach((button, index) => {
    button.onclick = () => selectConsoleTab(button.dataset.consoleTab!);
    button.onkeydown = (event) => {
      let next = index;
      if (event.key === "ArrowRight") next = (index + 1) % tabs.length;
      else if (event.key === "ArrowLeft") next = (index + tabs.length - 1) % tabs.length;
      else if (event.key === "Home") next = 0;
      else if (event.key === "End") next = tabs.length - 1;
      else return;
      event.preventDefault();
      selectConsoleTab(tabs[next].dataset.consoleTab!);
      tabs[next].focus();
    };
  });
}

function showConsoleView(view: "mail" | "list" | "detail") {
  document.getElementById("mail-workspace")!.hidden = view !== "mail";
  document.getElementById("console-workspace")!.hidden = view === "mail";
  document.getElementById("console-list-view")!.hidden = view !== "list";
  document.getElementById("console-detail-view")!.hidden = view !== "detail";
  document.getElementById("console-nav-mail")!.setAttribute("aria-pressed", String(view === "mail"));
  document.getElementById("console-nav-cases")!.setAttribute("aria-pressed", String(view !== "mail"));
  if (view === "mail") ++consoleDetailToken;
}

async function loadConsoleList(page: number) {
  const token = ++consoleListToken;
  ++consoleDetailToken;
  consoleListBusy = true;
  showConsoleView("list");
  document.getElementById("console-case-list")!.replaceChildren();
  consoleMessage("console-list-message", "케이스를 불러오는 중…");
  updateConsolePagination();
  try {
    const result = await searchCases({ ...consoleFilters }, page);
    if (token !== consoleListToken) return;
    consoleResult = result;
    consolePage = result.page;
    const states = document.getElementById("console-status-filter") as HTMLSelectElement;
    const chosen = states.value;
    states.replaceChildren(new Option("전체 상태", ""), ...result.statuses.map((s) => new Option(s, s)));
    states.value = result.statuses.includes(chosen) ? chosen : "";
    const list = document.getElementById("console-case-list")!;
    for (const item of result.items) {
      const row = consoleNode("button", "console-case-row");
      row.type = "button";
      row.dataset.caseId = item.caseId;
      row.append(consoleNode("span", "console-case-id", item.caseId), consoleBadge(item.status),
        consoleNode("span", "console-case-customer", item.customerName || "고객 미확인"),
        consoleNode("span", "console-case-route", `${item.pol || "미확보"} → ${item.pod || "미확보"}`),
        consoleNode("span", "console-case-subject", item.subject || "(제목 없음)"),
        consoleNode("span", "console-case-date", `등록 ${consoleDate(item.createdAt)}`));
      row.onclick = () => loadConsoleDetail(item.caseId);
      const li = document.createElement("li");
      li.append(row);
      list.append(li);
    }
    consoleMessage("console-list-message", result.total
      ? `총 ${result.total.toLocaleString()}건 · ${(result.page - 1) * result.pageSize + 1}–${Math.min(result.page * result.pageSize, result.total)}번째`
      : "검색 조건에 맞는 케이스가 없습니다. 조건을 변경하거나 초기화해 주세요.");
  } catch (error) {
    if (token !== consoleListToken) return;
    consoleResult = null;
    consoleMessage("console-list-message", consoleError(error), true);
  } finally {
    if (token === consoleListToken) {
      consoleListBusy = false;
      updateConsolePagination();
    }
  }
}

function updateConsolePagination() {
  const pages = consoleResult ? Math.max(1, Math.ceil(consoleResult.total / consoleResult.pageSize)) : 1;
  setText("console-page-label", consoleListBusy ? "조회 중…" : `${consoleResult ? consolePage : 1} / ${pages} 페이지`);
  (document.getElementById("console-prev") as HTMLButtonElement).disabled = consoleListBusy || !consoleResult || consolePage <= 1;
  (document.getElementById("console-next") as HTMLButtonElement).disabled = consoleListBusy || !consoleResult || consolePage >= pages;
}

async function loadConsoleDetail(caseId: string, resetTab = true) {
  const token = ++consoleDetailToken;
  consoleCaseId = caseId;
  if (resetTab) consoleTab = "fields";
  showConsoleView("detail");
  document.getElementById("console-detail-content")!.hidden = true;
  consoleMessage("console-detail-message", "케이스 상세를 불러오는 중…");
  const [detailResult, fieldsResult] = await Promise.allSettled([getCase(caseId), getCaseFields(caseId)]);
  if (token !== consoleDetailToken) return;
  if (detailResult.status === "rejected") {
    consoleMessage("console-detail-message", consoleError(detailResult.reason), true);
    return;
  }
  const detail = detailResult.value;
  const fields = fieldsResult.status === "fulfilled" ? fieldsResult.value : [];
  const customer = fields.find((f) => f.field === "customerName")?.value;
  const heading = consoleNode("div", "console-detail-summary");
  const meta = consoleNode("div", "console-detail-id");
  meta.append(consoleNode("span", "console-case-id", detail.caseId), consoleBadge(detail.status));
  heading.append(meta, consoleNode("h1", "", customer ? String(customer) : "고객 미확인"),
    consoleNode("p", "console-detail-subject", detail.subject || "(제목 없음)"),
    consoleNode("div", "console-detail-meta", `등록 ${consoleDate(detail.createdAt)} · ${detail.fromEmail || "발신 주소 없음"}`));
  document.getElementById("console-detail-heading")!.replaceChildren(heading);
  renderConsoleFields(detail, fields, fieldsResult.status === "rejected");
  renderConsoleQuote(detail);
  renderConsoleFiles(detail);
  renderConsoleHistory(detail);
  setText("console-mail-hint", "메일 초안·회신 병합은 ‘선택한 메일’에서 처리합니다. 다른 케이스를 조회해도 Outlook에서 열린 메일은 바뀌지 않습니다.");
  selectConsoleTab(consoleTab);
  document.getElementById("console-detail-content")!.hidden = false;
  consoleMessage("console-detail-message", fieldsResult.status === "rejected"
    ? "필드 값을 불러오지 못했습니다. 다른 자료는 조회할 수 있으며, 새로고침으로 다시 시도해 주세요." : "", fieldsResult.status === "rejected");
}

function selectConsoleTab(tab: string) {
  consoleTab = tab;
  for (const button of document.querySelectorAll<HTMLButtonElement>("[data-console-tab]")) {
    const active = button.dataset.consoleTab === tab;
    button.setAttribute("aria-selected", String(active));
    button.tabIndex = active ? 0 : -1;
    document.getElementById(`console-panel-${button.dataset.consoleTab}`)!.hidden = !active;
  }
}

function renderConsoleFields(detail: CaseDetail, fields: CaseField[], failed: boolean) {
  const panel = document.getElementById("console-panel-fields")!;
  // 같은 케이스를 새로고침할 때만 펼친 위치를 유지한다. 저장 데이터에는 영향을 주지 않는다.
  const expanded = new Set(panel.dataset.caseId === detail.caseId
    ? Array.from(panel.querySelectorAll<HTMLDetailsElement>("details[data-console-disclosure][open]"))
      .map((node) => node.dataset.consoleDisclosure) : []);
  panel.replaceChildren();
  panel.dataset.caseId = detail.caseId;
  const extraction = detail.extraction;
  const messages: Record<string, string> = {
    none: "아직 추출 결과가 없습니다.", running: "추출·검증 중입니다. 잠시 후 새로고침해 주세요.",
    failed: `추출 실패: ${extraction.error || "오류 상세 없음"}`, "not-connected": "추출기가 연결되지 않았습니다.",
    done: `추출·검증 결과${extraction.elapsedMs == null ? "" : ` · ${(extraction.elapsedMs / 1000).toFixed(1)}초`}`,
  };
  panel.append(consoleNode("p", "console-note", messages[extraction.state] || extraction.state));
  const validation = consoleRecord(extraction.validation);
  const issuePanel = validationIssuesPanel(validation);
  if (issuePanel) panel.append(issuePanel);
  const questions = Array.isArray(validation.questions) ? validation.questions : [];
  if (questions.length) {
    const note = consoleNode("div", "console-note console-note-warning");
    note.append(consoleNode("strong", "", `보완 요청 문항 · ${questions.length}개`));
    const ol = document.createElement("ol");
    for (const question of questions) ol.append(consoleNode("li", "", String(question)));
    note.append(ol);
    panel.append(note);
  }
  if (failed) return;
  const byName = new Map<string, Partial<CaseField>>(fields.map((field) => [field.field, field]));
  const extra = consoleRecord(consoleRecord(extraction.result).fields);
  const extras = ["carrier", "containerCount"].filter((key) => byName.get(key)?.value != null || consoleRecord(extra[key]).value != null);
  for (const key of extras) {
    if (byName.has(key)) continue;
    const field = consoleRecord(extra[key]);
    byName.set(key, {
      value: field.value, candidate: field.candidate,
      status: typeof field.status === "string" ? field.status : undefined,
      score: typeof field.score === "number" ? field.score : undefined,
      evidence: typeof field.evidence === "string" ? field.evidence : null,
      origin: typeof field.origin === "string" ? field.origin : null,
      method: typeof field.method === "string" ? field.method : null,
    });
  }
  const brief = (name: string, unit = "") => {
    const field = byName.get(name);
    const label = consoleFieldLabels[name] || name;
    if (field?.status === "review") return `${label} 확인 필요`;
    if (field?.value == null || field.value === "") return `${label} 미확보`;
    return `${consoleValue(field.value)}${unit ? ` ${unit}` : ""}`;
  };
  const groups: [string, string, string[], string[]][] = [
    ["customer", "고객 정보", ["customerName", "contactName", "contactEmail"], [brief("customerName"), brief("contactName")]],
    ["transport", "운송 정보", ["pol", "pod", "cargoReadyDate", "incoterms", "containerType", ...extras],
      [`${brief("pol")} → ${brief("pod")}`, brief("incoterms"), brief("containerType"),
        ...(extras.includes("containerCount") ? [brief("containerCount", "대")] : [])]],
    ["cargo", "화물 정보", ["commodity", "qty", "qtyUnit", "packing", "boxL", "boxW", "boxH", "totalCbm", "grossWeightKg"],
      [brief("commodity"), `${brief("qty")} ${brief("qtyUnit")}`, brief("totalCbm", "CBM"), brief("grossWeightKg", "kg")]],
    ["payment", "금액·결제조건", ["invoiceValue", "ccy", "paymentTerm"],
      [`${brief("ccy")} ${brief("invoiceValue")}`, brief("paymentTerm")]],
  ];
  const stack = consoleNode("div", "console-field-groups");
  for (const [key, title, names, preview] of groups) {
    const section = consoleNode("details", "console-field-group");
    section.dataset.consoleDisclosure = key;
    section.open = expanded.has(key);
    const summary = consoleNode("summary", "console-field-summary");
    const toggle = consoleNode("span", "console-field-toggle");
    toggle.append(consoleNode("span", "console-field-expand", "상세 보기 ＋"), consoleNode("span", "console-field-collapse", "접기 −"));
    summary.append(consoleNode("strong", "console-field-title", title), toggle,
      consoleNode("span", "console-field-preview", preview.join(" · ")));
    const review = names.filter((name) => byName.get(name)?.status === "review");
    if (review.length) summary.append(consoleNode("span", "console-field-review", `확인 필요: ${review.map((name) => consoleFieldLabels[name] || name).join(", ")}`));
    const dl = consoleNode("dl", "console-data");
    const evidenceRows = consoleNode("div", "console-field-evidence-rows");
    for (const name of names) {
      const field = byName.get(name);
      const dd = consoleNode("dd", field?.value == null ? "console-muted" : "", consoleValue(field?.value));
      if (field?.status === "review") {
        dd.append(consoleNode("div", "console-evidence", `확인 필요${field.candidate == null ? "" : ` · 후보: ${consoleValue(field.candidate)}`}`));
      }
      if (field?.evidence || field?.method) {
        const evidence = consoleNode("section", "console-field-evidence-row");
        evidence.append(consoleNode("h3", "console-field-evidence-title", consoleFieldLabels[name] || name));
        if (field.method === "manual") evidence.append(consoleNode("p", "console-field-evidence-meta", "담당자 보정 기록"));
        evidence.append(consoleNode("p", "console-field-evidence-text", field.evidence || "근거 원문 없음"),
          consoleNode("p", "console-field-evidence-meta", `점수 ${field.score ?? "—"}${field.origin ? ` · ${field.origin}` : ""}`));
        evidenceRows.append(evidence);
      }
      dl.append(consoleNode("dt", "", consoleFieldLabels[name] || name), dd);
    }
    section.append(summary, dl);
    if (evidenceRows.childElementCount) {
      const evidence = consoleNode("details", "console-field-evidence");
      evidence.dataset.consoleDisclosure = `${key}-evidence`;
      evidence.open = expanded.has(`${key}-evidence`);
      evidence.append(consoleNode("summary", "", "추출 근거·보정 기록 보기"), evidenceRows);
      section.append(evidence);
    }
    stack.append(section);
  }
  panel.append(stack);
}

function renderConsoleQuote(detail: CaseDetail) {
  const panel = document.getElementById("console-panel-quote")!;
  panel.replaceChildren();
  const quote = consoleRecord(consoleRecord(detail.extraction.validation).quote);
  if (!Object.keys(quote).length) {
    panel.append(consoleNode("p", "console-note", "현재 검증 결과에 생성된 견적서가 없습니다. 항목·검증 탭에서 보완 항목이나 보류 사유를 확인해 주세요."));
    return;
  }
  const section = consoleSection(`견적서 · Rev. ${consoleValue(quote.version)}`);
  const totals = Object.entries(consoleRecord(quote.totals)).map(([ccy, value]) => `${ccy} ${typeof value === "number" ? value.toLocaleString("en-US", { maximumFractionDigits: 2 }) : consoleValue(value)}`).join(" + ");
  section.append(consoleNode("p", "console-note console-note-success", totals || "합계 정보 없음"));
  const links = consoleNode("div", "console-links");
  for (const [key, label] of [["pdf", "견적서 PDF 열기"], ["file", "견적서 Excel 열기"]]) {
    const link = consoleQuoteLink(detail.caseId, quote[key], label);
    if (link) links.append(link);
  }
  section.append(links, consoleData([
    ["ETD / ETA", `${consoleValue(quote.etd)} / ${consoleValue(quote.eta)}`],
    ["선사", consoleValue(quote.carrier)], ["견적 유효일", consoleValue(quote.validUntil)],
    ["생성 소요 시간", typeof quote.elapsedMs === "number" ? `${(quote.elapsedMs / 1000).toFixed(2)}초${quote.withinLimit === false ? " · 10초 기준 초과" : ""}` : "미기록"],
  ]));
  panel.append(section, consoleNode("p", "console-muted", "보완 요청·견적서 송부 초안은 해당 고객 메일을 선택한 뒤 ‘선택한 메일’에서 만듭니다."));
}

const consoleEventLabels: Record<string, string> = {
  CASE_CREATED: "케이스 생성", PIPELINE_STARTED: "추출·검증 시작", PIPELINE_RUN: "추출·검증 결과",
  PIPELINE_FAILED: "처리 실패", DRAFT_PREPARED: "메일 초안 준비", DRAFT_SAVED: "초안함 저장",
  MAIL_SENT: "발송 기록", REPLY_MERGED: "회신 병합", FIELD_CORRECTED: "필드 보정",
};

/** 같은 시각·종류의 기록도 구분하도록 케이스와 원본 배열 순번으로 연결한다. */
function consoleEventRecords(detail: CaseDetail) {
  const eventTime = (event: CaseDetail["events"][number]) => {
    const time = Date.parse(event.createdAt);
    return Number.isFinite(time) ? time : Infinity;
  };
  return detail.events.map((event, index) => ({ event, index, logId: `console-audit-${encodeURIComponent(detail.caseId)}-${index}` }))
    .sort((a, b) => eventTime(a.event) - eventTime(b.event) || a.index - b.index);
}

function renderConsoleFiles(detail: CaseDetail) {
  const panel = document.getElementById("console-panel-files")!;
  panel.replaceChildren();
  panel.append(consoleNode("h2", "", "메일 자료"));
  if (!detail.mails.length) panel.append(consoleNode("p", "console-note", "저장된 메일 자료가 없습니다."));
  for (const mail of detail.mails) {
    const section = consoleNode("section", "console-file");
    section.append(consoleNode("h3", "", `${mail.index}. ${mail.role === "ORIGINAL" ? "요청 메일" : "회신 메일"}`),
      consoleNode("p", "", mail.snapshot.subject || "(제목 없음)"),
      consoleNode("p", "console-muted", `${mail.snapshot.from?.email || "발신 주소 없음"} · ${consoleDate(mail.snapshot.receivedAt)}`));
    const links = consoleNode("div", "console-links");
    if (mail.rawUrl) links.append(fileLink("원문 (.eml)", mail.rawUrl));
    for (const file of mail.attachmentFiles) links.append(fileLink(file.filename, file.url));
    if (!mail.rawUrl && !mail.attachmentFiles.length) links.append(consoleNode("span", "console-muted", "원문·첨부 파일 없음"));
    section.append(links);
    const body = (mail.snapshot as MailSnapshot).bodyText;
    if (body) {
      const disclosure = document.createElement("details");
      disclosure.append(consoleNode("summary", "", "저장된 메일 본문"), consoleNode("div", "console-mail-body", body));
      section.append(disclosure);
    }
    panel.append(section);
  }
  const logs = consoleSection("판단 로그");
  logs.classList.add("console-audit-logs");
  logs.append(consoleNode("p", "console-history-help", "처리 단계별 원본 기록입니다. 이력의 ‘판단 로그 보기’에서 해당 기록으로 이동할 수 있습니다."));
  if (!detail.events.length) logs.append(consoleNode("p", "console-note", "저장된 판단 로그가 없습니다."));
  for (const { event, logId } of consoleEventRecords(detail)) {
    const disclosure = consoleNode("details", "console-audit-log");
    disclosure.id = logId;
    const summary = consoleNode("summary", "", consoleEventLabels[event.eventType] || event.eventType);
    summary.append(consoleNode("span", "console-audit-time", `기록 ${consoleDate(event.createdAt, true)}`));
    disclosure.append(summary, consoleData([
      ["수행 주체", event.actor || "미기록"], ["기록 종류", event.eventType],
    ]));
    if (event.detail != null) {
      disclosure.append(consoleNode("pre", "", JSON.stringify(event.detail, null, 2)));
    } else {
      disclosure.append(consoleNode("p", "console-history-help", "상세 내용이 기록되지 않았습니다."));
    }
    logs.append(disclosure);
  }
  panel.append(logs);
}

function renderConsoleHistory(detail: CaseDetail) {
  const panel = document.getElementById("console-panel-history")!;
  panel.replaceChildren();
  if (!detail.events.length) {
    panel.append(consoleNode("p", "console-note", "저장된 이력이 없습니다."));
    return;
  }
  const list = consoleNode("ol", "console-history");
  const records = consoleEventRecords(detail);
  const events = records.map(({ event }) => event);
  panel.append(consoleNode("p", "console-history-help", "기록 시각순입니다. 발송 이력에는 실제 발송 시각을 함께 표시합니다."));
  for (const { event, logId } of records) {
    const entry = consoleNode("li", "");
    const data = consoleRecord(event.detail);
    entry.append(consoleNode("time", "", `기록 ${consoleDate(event.createdAt, true)}`), consoleNode("h3", "", consoleEventLabels[event.eventType] || event.eventType),
      consoleNode("p", "console-muted", `수행 주체: ${event.actor || "미기록"}`));
    if (event.eventType === "MAIL_SENT") entry.append(consoleNode("p", "", `실제 발송: ${consoleDate(typeof data.sentAt === "string" ? data.sentAt : null, true)}`));
    // 기록에 없는 상태는 현재 상태나 주변 이력으로 채워 넣지 않는다.
    const from = typeof data.statusFrom === "string" && data.statusFrom ? data.statusFrom : null;
    const after = data.statusTo ?? consoleRecord(data.validation).status ?? data.status;
    const to = typeof after === "string" && after ? after : null;
    if (from || to) {
      const transition = from && to ? (from === to ? `상태 유지: ${to}` : `상태 변경: ${from} → ${to}`)
        : from ? `변경 전: ${from} · 변경 후: 미기록` : `변경 전: 미기록 · 처리 후: ${to}`;
      entry.append(consoleNode("p", "console-history-state", transition));
    }
    const links = consoleNode("div", "console-links");
    const mail = historyMail(detail, events, event);
    if (mail?.rawUrl) links.append(fileLink("메일 원문", mail.rawUrl));
    for (const file of mail?.attachmentFiles || []) links.append(fileLink(file.filename, file.url));
    if ((event.eventType === "CASE_CREATED" || event.eventType === "REPLY_MERGED") && !mail) {
      entry.append(consoleNode("p", "console-history-help", "해당 이력의 메일을 특정할 수 없습니다."));
      const files = consoleNode("button", "console-history-link", "케이스 자료에서 확인");
      files.type = "button";
      files.onclick = () => {
        selectConsoleTab("files");
        document.getElementById("console-tab-files")!.focus();
      };
      links.append(files);
    }
    const recordedQuote = consoleRecord(consoleRecord(data.validation).quote);
    // 각 이력에 저장된 당시 파일만 연결한다. 경로가 없으면 최신 견적으로 대신하지 않는다.
    if (recordedQuote.version != null) {
      let linkedFiles = 0;
      for (const [key, label] of [["pdf", "견적서 PDF"], ["file", "견적서 Excel"]]) {
        const link = consoleQuoteLink(detail.caseId, recordedQuote[key], `${label} · Rev. ${consoleValue(recordedQuote.version)}`);
        if (link) {
          links.append(link);
          linkedFiles++;
        }
      }
      if (!linkedFiles) {
        entry.append(consoleNode("p", "console-history-help", `Rev. ${consoleValue(recordedQuote.version)} 생성 기록입니다. 당시 견적 파일 경로가 기록되지 않았습니다.`));
      }
    }
    const output = historyOutput(event, mail);
    if (output) {
      output.id = `console-history-output-${list.children.length}`;
      const link = consoleNode("a", "console-history-link", output.querySelector("summary")!.textContent || "산출물 기록 보기");
      link.href = `#${output.id}`;
      link.onclick = (click) => {
        click.preventDefault();
        output.open = true;
        output.querySelector("summary")!.focus();
        output.scrollIntoView({ block: "nearest" });
      };
      links.append(link);
    }
    const logLink = consoleNode("a", "console-history-link", "판단 로그 보기");
    logLink.href = `#${logId}`;
    logLink.onclick = (click) => {
      click.preventDefault();
      const log = document.getElementById(logId);
      if (!(log instanceof HTMLDetailsElement)) return;
      selectConsoleTab("files");
      log.open = true;
      log.querySelector("summary")!.focus();
      log.scrollIntoView({ block: "nearest" });
    };
    links.append(logLink);
    entry.append(links);
    if (output) entry.append(output);
    list.append(entry);
  }
  panel.append(list);
}

/** 회신 기록이 일부 빠져도 다른 메일을 같은 순번이라는 이유로 연결하지 않는다. */
function historyMail(detail: CaseDetail, events: CaseDetail["events"], event: CaseDetail["events"][number]): CaseDetail["mails"][number] | undefined {
  if (event.eventType === "CASE_CREATED") {
    const originals = detail.mails.filter((mail) => mail.role === "ORIGINAL");
    return originals.length === 1 ? originals[0] : undefined;
  }
  if (event.eventType !== "REPLY_MERGED") return undefined;
  const replies = detail.mails.filter((mail) => mail.role === "REPLY").sort((a, b) => a.index - b.index);
  const merges = events.filter((item) => item.eventType === "REPLY_MERGED");
  const subject = consoleRecord(event.detail).subject;
  if (typeof subject !== "string" || !subject) return undefined;
  // 현재 저장 규격은 병합마다 메일 1건과 제목 이력 1건을 같은 순서로 남긴다.
  if (merges.length === replies.length && merges.every((item, i) => consoleRecord(item.detail).subject === replies[i].snapshot.subject)) {
    return replies[merges.indexOf(event)];
  }
  const matching = replies.filter((mail) => mail.snapshot.subject === subject);
  return matching.length === 1 && merges.filter((item) => consoleRecord(item.detail).subject === subject).length === 1 ? matching[0] : undefined;
}

/** 당시 이력에 저장된 산출물만 읽는다. 최신 추출 결과나 새 초안을 가져오지 않는다. */
function historyOutput(event: CaseDetail["events"][number], mail?: CaseDetail["mails"][number]): HTMLDetailsElement | null {
  const data = consoleRecord(event.detail);
  const names: Record<string, string> = {
    CASE_CREATED: "요청 메일 내용 보기", REPLY_MERGED: "병합한 회신 내용 보기",
    PIPELINE_RUN: "당시 추출·검증 결과 보기", PIPELINE_FAILED: "실패 사유 보기",
    DRAFT_PREPARED: "초안 준비 내역 보기", DRAFT_SAVED: "초안 저장 내역 보기",
    MAIL_SENT: "발송 내역 보기", FIELD_CORRECTED: "필드 보정 내역 보기",
  };
  if (!names[event.eventType] || (["CASE_CREATED", "REPLY_MERGED"].includes(event.eventType) && !mail)) return null;
  const output = consoleNode("details", "console-history-output");
  output.append(consoleNode("summary", "", names[event.eventType]));
  if (mail) {
    output.append(consoleData([["제목", consoleValue(mail.snapshot.subject)], ["발신자", consoleValue(mail.snapshot.from?.email)]]));
    const body = (mail.snapshot as MailSnapshot).bodyText;
    output.append(consoleNode("div", "console-mail-body", body || "저장된 본문이 없습니다."));
  } else if (event.eventType === "PIPELINE_RUN") {
    const fields = Object.entries(extractedFields(consoleRecord(data.extraction)))
      .filter(([field, info]) => field in consoleFieldLabels && info.value != null);
    if (fields.length) output.append(consoleData(fields.map(([field, info]) => [consoleFieldLabels[field], `${consoleValue(info.value)}${info.status === "review" ? " (확인 필요)" : ""}`])));
    const validation = consoleRecord(data.validation);
    const issues = validationIssuesPanel(validation);
    if (issues) output.append(issues);
    const questions = Array.isArray(validation.questions) ? validation.questions : [];
    if (questions.length) {
      output.append(consoleNode("h4", "", "당시 보완 문항"));
      const list = consoleNode("ol", "console-history-questions");
      for (const question of questions) list.append(consoleNode("li", "", String(question)));
      output.append(list);
    }
    const quote = consoleRecord(validation.quote);
    if (quote.version != null) output.append(consoleData([
      ["견적 버전", `Rev. ${consoleValue(quote.version)}`],
      ["당시 견적 합계", Object.entries(consoleRecord(quote.totals)).map(([ccy, amount]) => `${ccy} ${consoleValue(amount)}`).join(" + ") || "미기록"],
    ]));
    if (typeof data.elapsedMs === "number") output.append(consoleNode("p", "console-history-help", `추출·검증 처리 시간: ${(data.elapsedMs / 1000).toFixed(2)}초`));
    if (data.extraction == null && data.validation == null) output.append(consoleNode("p", "console-history-help", "추출·검증 상세 결과가 기록되지 않았습니다."));
  } else if (event.eventType === "PIPELINE_FAILED") {
    output.append(consoleNode("p", "console-history-error", typeof data.error === "string" ? data.error : "실패 사유 미기록"));
  } else if (event.eventType === "FIELD_CORRECTED") {
    const field = String(data.field || "");
    output.append(consoleData([["항목", consoleFieldLabels[field] || field || "미기록"], ["수정 전", consoleValue(data.before)], ["수정 후", consoleValue(data.after)]]));
  } else {
    const kind = data.kind === "SUPPLEMENT" ? "보완 요청" : data.kind === "QUOTE" ? "견적서 송부" : "미기록";
    output.append(consoleData([["메일 구분", kind], ["제목", typeof data.subject === "string" ? data.subject : "미기록"]]));
    if (Array.isArray(data.questions)) {
      output.append(consoleNode("h4", "", "초안에 사용한 보완 문항"));
      const list = consoleNode("ol", "console-history-questions");
      for (const question of data.questions) list.append(consoleNode("li", "", String(question)));
      output.append(list);
    }
    if (Array.isArray(data.files)) {
      output.append(consoleNode("h4", "", "기록된 첨부파일명"));
      const list = consoleNode("ul", "console-history-attachments");
      for (const file of data.files) list.append(consoleNode("li", "", String(file)));
      output.append(list);
    }
    output.append(consoleNode("p", "console-history-help", "저장된 처리 내역입니다. 실제 메일 내용과 첨부파일은 Outlook에서 확인하세요."));
  }
  return output;
}

function consoleQuoteLink(caseId: string, path: unknown, label: string): HTMLAnchorElement | null {
  if (typeof path !== "string" || !path) return null;
  const filename = path.replace(/\\/g, "/").split("/").pop();
  if (!filename || !/\.(pdf|xlsx)$/i.test(filename)) return null;
  return fileLink(label, `/api/cases/${encodeURIComponent(caseId)}/quotes/${encodeURIComponent(filename)}`);
}

function consoleSection(title: string): HTMLElement {
  const section = consoleNode("section", "console-section");
  section.append(consoleNode("h2", "", title));
  return section;
}

function consoleData(rows: string[][]): HTMLDListElement {
  const dl = consoleNode("dl", "console-data");
  for (const [label, value] of rows) dl.append(consoleNode("dt", "", label), consoleNode("dd", "", value));
  return dl;
}

function consoleBadge(status: string): HTMLSpanElement {
  const kind = ["계산완료", "발송완료"].includes(status) ? "ready"
    : ["정보부족", "보완대기", "승인대기"].includes(status) ? "attention"
      : ["보류", "실패"].includes(status) ? "error" : "working";
  return consoleNode("span", `console-badge console-badge-${kind}`, status);
}

function consoleNode<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text?: string): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function consoleRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

function consoleValue(value: unknown): string {
  if (value == null || value === "") return "미확보";
  return typeof value === "object" ? JSON.stringify(value) : String(value);
}

function consoleDate(value: string | null, seconds = false): string {
  if (!value) return "시각 미기록";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "시각 미기록";
  return new Intl.DateTimeFormat("ko-KR", { timeZone: "Asia/Seoul", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: seconds ? "2-digit" : undefined, hour12: false }).format(date);
}

function consoleMessage(id: string, message: string, error = false) {
  const node = document.getElementById(id)!;
  node.textContent = message;
  node.className = `console-message${error ? " console-message-error" : ""}`;
  node.hidden = !message;
}

function consoleError(error: unknown): string {
  return error instanceof ApiError ? error.message : "케이스 조회 중 오류가 발생했습니다. 다시 시도해 주세요.";
}
