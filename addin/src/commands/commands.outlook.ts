/* global Office */

/*
 * 이벤트 기반 실행 처리기 — 사내 관리자 배포 환경용.
 * 사용자의 동작(작성 창 열기·보내기)에 반응할 때만 실행된다 — 폴링 아님.
 *
 * - 작성 창이 열릴 때(OnNewMessageCompose): 패널이 준비한 초안이면 제목·첨부를 채우고 초안함에 저장 (FR-304·505)
 * - 보내기를 누를 때(OnMessageSend): 발송 시각·수행자 기록 (FR-505). 보내기를 막거나 대신 보내지 않는다.
 *
 * 2026-09-28 테스트: 개인이 사이드로드한 Outlook 웹(학교 계정)에서는 이 백그라운드 실행 환경이 동작하지 않았다.
 * 그 환경에서는 작성 창의 [LEONA 초안 마무리] 버튼이 여는 패널(taskpane)이 같은 일을 한다.
 */

import { call, finalizeDraft, userEmail } from "../shared/finalize";
import { takePendingDraft } from "../shared/handoff";
import { recordMailEvent } from "../taskpane/api";

const CASE_ID_PATTERN = /LQ-\d{4}-\d{4}-\d{3}/;

async function onNewMessageComposeHandler(event: Office.AddinCommands.Event) {
  const item = Office.context.mailbox.item as Office.MessageCompose;
  try {
    // 자동 실행에서는 쪽지만 본다 (담당자가 직접 연 일반 메일 작성 창에는 아무것도 하지 않는다)
    const note = takePendingDraft();
    if (note) {
      await finalizeDraft(item, note);
      notify(item, `LEONA: ${note.caseId} 초안을 초안함에 저장했습니다. 확인 후 직접 보내기를 눌러 주세요.`);
    }
  } catch {
    // 실패하면 작성 창의 [LEONA 초안 마무리] 버튼으로 다시 할 수 있다
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

function notify(item: Office.MessageCompose, message: string) {
  item.notificationMessages.replaceAsync("leona", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message.slice(0, 150),
    icon: "Icon.16x16",
    persistent: false,
  });
}

// 매니페스트의 LaunchEvent FunctionName과 함수를 연결한다
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
