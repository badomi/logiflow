/* global Office */

/**
 * 선택한 메일 1통의 정보 — A트랙 "메일 표준 JSON"(v1).
 *
 * 백엔드(server/app/schemas.py의 MailSnapshot)와 같은 모양이다. 패널, .eml 리더, 백엔드가 모두 이 형식을 쓴다.
 * 헤더(inReplyTo·references)와 첨부 목록은 원문(.eml)을 받은 백엔드가 채운다.
 */
export interface Address {
  name: string;
  email: string;
}

export interface MailSnapshot {
  subject: string;
  from: Address;
  to: Address[];
  cc: Address[];
  /** 수신 시각, ISO 8601(UTC) 문자열 */
  receivedAt: string;
  /** 본문 전체(서식 없는 텍스트) */
  bodyText: string;
  /** 같은 대화(스레드)에 속한 메일끼리 공유하는 ID — 회신 식별(FR-104)에 사용 */
  conversationId: string;
  /** 메일 1통마다 고유한 인터넷 표준 ID — 중복 등록 방지(FR-105)에 사용 */
  internetMessageId: string;
  inReplyTo: string | null;
  references: string[];
  attachments: { filename: string; contentType: string | null; size: number; sha256: string }[];
}

/** 지금 Outlook에서 선택(열람) 중인 메일을 읽어 MailSnapshot으로 돌려준다. */
export async function readSelectedMail(): Promise<MailSnapshot> {
  const item = currentItem();
  const toAddress = (a: Office.EmailAddressDetails): Address => ({ name: a.displayName ?? "", email: a.emailAddress ?? "" });

  return {
    subject: item.subject ?? "",
    from: {
      name: item.from?.displayName ?? "",
      email: item.from?.emailAddress ?? "",
    },
    to: (item.to ?? []).map(toAddress),
    cc: (item.cc ?? []).map(toAddress),
    // Office.js에는 '수신 시각' 속성이 따로 없다. 받은 메일은 사서함에 생성된 시각이 곧 수신 시각이다.
    receivedAt: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : "",
    bodyText: await getBodyText(item),
    conversationId: item.conversationId ?? "",
    internetMessageId: item.internetMessageId ?? "",
    inReplyTo: null,
    references: [],
    attachments: [],
  };
}

/**
 * 선택한 메일의 원문(.eml)을 Base64로 받는다. 첨부파일·헤더가 모두 들어 있어 원문 보관(FR-103)에 쓴다.
 * 요구사항 세트 1.14 이상에서만 되므로, 지원하지 않는 Outlook에서는 null을 돌려준다.
 */
export function readSelectedMailRaw(): Promise<string | null> {
  const item = currentItem();
  if (!Office.context.requirements.isSetSupported("Mailbox", "1.14") || !item.getAsFileAsync) {
    return Promise.resolve(null);
  }
  return new Promise((resolve, reject) => {
    item.getAsFileAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/** 지금 Outlook을 쓰는 담당자 메일 주소 — 이력에 "누가 실행했는지" 남긴다. */
export function currentUserEmail(): string {
  return Office.context.mailbox.userProfile?.emailAddress ?? "";
}

function currentItem(): Office.MessageRead {
  const item = Office.context.mailbox.item as Office.MessageRead | undefined;
  if (!item) {
    throw new Error("선택된 메일이 없습니다.");
  }
  return item;
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
