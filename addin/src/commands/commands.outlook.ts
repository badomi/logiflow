/* global Office */

/*
 * 작성 창에서 실행되는 동작. 모두 담당자의 동작(버튼·보내기)에 반응할 때만 실행된다 — 폴링 아님.
 *
 * 1) [LEONA 초안 마무리] 버튼(finalizeDraft): 패널이 남긴 쪽지대로
 *    제목(케이스 ID)·견적서 첨부를 채우고 초안함에 저장한다 (FR-304·505).
 * 2) 이벤트 기반 실행(사내 관리자 배포용):
 *    - 작성 창이 열릴 때(OnNewMessageCompose): 1)과 같은 일을 자동으로 한다.
 *    - 보내기를 누를 때(OnMessageSend): 발송 시각·수행자를 기록한다 (FR-505). 보내기를 막거나 대신 보내지 않는다.
 *    개인이 사이드로드한 Outlook 웹에서는 이 이벤트가 전달되지 않는 것을 확인했다 (2026-09-28 테스트).
 */

import { takePendingDraft } from "../shared/handoff";
import type { PendingDraft } from "../shared/handoff";
import { fetchFileBase64, recordMailEvent } from "../taskpane/api";

const CASE_ID_PATTERN = /LQ-\d{4}-\d{4}-\d{3}/;

/** [LEONA 초안 마무리] 버튼 */
async function finalizeDraft(event: Office.AddinCommands.Event) {
  const item = Office.context.mailbox.item as Office.MessageCompose;
  const pending = takePendingDraft();
  try {
    if (!pending) {
      notify(item, "LEONA: 패널에서 연 초안이 아닙니다. 패널의 [초안 열기] 버튼으로 다시 열어 주세요.", true);
      return;
    }
    await finalize(item, pending);
  } finally {
    event.completed();
  }
}

/** 작성 창이 열릴 때 자동 실행 (이벤트 기반 실행을 지원하는 배포 환경에서만) */
async function onNewMessageComposeHandler(event: Office.AddinCommands.Event) {
  const pending = takePendingDraft();
  try {
    if (pending) await finalize(Office.context.mailbox.item as Office.MessageCompose, pending);
  } finally {
    event.completed(); // 쪽지가 없으면 담당자가 직접 연 일반 메일 — 아무것도 하지 않는다
  }
}

/** 보내기를 누를 때 자동 실행 (이벤트 기반 실행을 지원하는 배포 환경에서만) */
async function onMessageSendHandler(event: Office.AddinCommands.Event) {
  try {
    const item = Office.context.mailbox.item as Office.MessageCompose;
    const subject = await call<string>((done) => item.subject.getAsync(done));
    const caseId = subject.match(CASE_ID_PATTERN)?.[0];
    if (caseId) {
      await recordMailEvent(caseId, "MAIL_SENT", subject, userEmail());
    }
  } catch {
    // 백엔드가 꺼져 있어도 담당자의 발송을 막지 않는다
  } finally {
    event.completed({ allowEvent: true });
  }
}

/** 제목·첨부를 채우고 초안함에 저장한 뒤 이력에 남긴다. */
async function finalize(item: Office.MessageCompose, pending: PendingDraft) {
  try {
    await call<void>((done) => item.subject.setAsync(pending.subject, done));
    for (const file of pending.attachments) {
      const base64 = await fetchFileBase64(file.url);
      await call<string>((done) => item.addFileAttachmentFromBase64Async(base64, file.filename, done));
    }
    await call<string>((done) => item.saveAsync(done)); // 초안함에 저장
    await recordMailEvent(pending.caseId, "DRAFT_SAVED", pending.subject, userEmail(), pending.kind);
    notify(item, `LEONA: ${pending.caseId} 초안을 초안함에 저장했습니다. 확인 후 직접 보내기를 눌러 주세요.`, false);
  } catch (error) {
    notify(item, `LEONA: 초안을 마무리하지 못했습니다 — ${error instanceof Error ? error.message : error}`, true);
  }
}

/** Office.js 콜백 API를 await로 쓸 수 있게 감싼다. */
function call<T>(start: (done: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    start((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function userEmail(): string {
  return Office.context.mailbox.userProfile?.emailAddress ?? "";
}

function notify(item: Office.MessageCompose, message: string, isError: boolean) {
  item.notificationMessages.replaceAsync("leona", {
    type: isError
      ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message.slice(0, 150),
    ...(isError ? {} : { icon: "Icon.16x16", persistent: false }),
  });
}

// 매니페스트의 FunctionName과 함수를 연결한다
Office.actions.associate("finalizeDraft", finalizeDraft);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
