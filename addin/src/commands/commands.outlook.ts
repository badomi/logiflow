/* global Office */

/*
 * Outlook 이벤트 처리기 (이벤트 기반 실행). 사용자의 동작에 반응할 때만 실행된다 — 폴링 아님.
 *
 * 1) 작성 창이 열릴 때(OnNewMessageCompose): 패널이 남긴 쪽지가 있으면
 *    제목(케이스 ID)·견적서 첨부를 채우고 초안함에 저장한다 (FR-304·505).
 * 2) 담당자가 보내기를 누를 때(OnMessageSend): 제목에 케이스 ID가 있으면 발송 시각·수행자를 기록한다 (FR-505).
 *    보내기를 막거나 대신 보내지 않는다. 기록에 실패해도 메일은 그대로 나간다.
 */

import { takePendingDraft } from "../shared/handoff";
import { fetchFileBase64, recordMailEvent } from "../taskpane/api";

const CASE_ID_PATTERN = /LQ-\d{4}-\d{4}-\d{3}/;

async function onNewMessageComposeHandler(event: Office.AddinCommands.Event) {
  const pending = takePendingDraft();
  if (!pending) {
    event.completed(); // 담당자가 직접 연 일반 메일 작성 창 — 아무것도 하지 않는다
    return;
  }

  const item = Office.context.mailbox.item as Office.MessageCompose;
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
  } finally {
    event.completed();
  }
}

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

// 매니페스트의 LaunchEvent FunctionName과 함수를 연결한다
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
