/* global Office */

/**
 * 선택한 메일 1통의 정보 — A트랙 "메일 표준 JSON" 초안(v0).
 *
 * 패널 표시, 이후 "케이스 생성" 전달, .eml 리더 출력이 모두 이 모양을 쓰도록 한 곳에 정의한다.
 * 필드 구성은 B트랙과 합의 후 확정한다.
 */
export interface MailSnapshot {
  subject: string;
  from: { name: string; email: string };
  /** 수신 시각, ISO 8601(UTC) 문자열 */
  receivedAt: string;
  /** 본문 전체(서식 없는 텍스트) */
  bodyText: string;
  /** 같은 대화(스레드)에 속한 메일끼리 공유하는 ID — 회신 식별(FR-104)에 사용 */
  conversationId: string;
  /** 메일 1통마다 고유한 인터넷 표준 ID — 중복 등록 방지(FR-105)에 사용 */
  internetMessageId: string;
}

/** 지금 Outlook에서 선택(열람) 중인 메일을 읽어 MailSnapshot으로 돌려준다. */
export async function readSelectedMail(): Promise<MailSnapshot> {
  const item = Office.context.mailbox.item as Office.MessageRead | undefined;
  if (!item) {
    throw new Error("선택된 메일이 없습니다.");
  }

  return {
    subject: item.subject ?? "",
    from: {
      name: item.from?.displayName ?? "",
      email: item.from?.emailAddress ?? "",
    },
    // Office.js에는 '수신 시각' 속성이 따로 없다. 받은 메일은 사서함에 생성된 시각이 곧 수신 시각이다.
    receivedAt: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : "",
    bodyText: await getBodyText(item),
    conversationId: item.conversationId ?? "",
    internetMessageId: item.internetMessageId ?? "",
  };
}

/** 본문은 비동기 API(콜백 방식)라서 Promise로 감싸 await로 쓸 수 있게 한다. */
function getBodyText(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}
