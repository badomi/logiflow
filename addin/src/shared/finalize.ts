/* global Office */

/*
 * 작성 중인 초안 마무리 (FR-304·505): 제목에 케이스 ID, 견적서 첨부, 초안함 저장, 이력 기록.
 * 작성 창에서 연 LEONA 패널과 이벤트 처리기(사내 배포용)가 함께 쓴다.
 * 메일을 보내는 코드는 없다. 보내기는 담당자가 직접 누른다.
 */

import { takePendingDraft } from "./handoff";
import type { PendingDraft } from "./handoff";
import { fetchFileBase64, pendingDraft, recordMailEvent } from "../taskpane/api";

export function userEmail(): string {
  return Office.context.mailbox.userProfile?.emailAddress ?? "";
}

/** Office.js 콜백 API를 await로 쓸 수 있게 감싼다. */
export function call<T>(start: (done: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
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

/** 패널이 남긴 쪽지 → 없으면 백엔드에 "이 담당자가 방금 준비한 초안"을 묻는다. */
export async function findPendingDraft(item: Office.MessageCompose): Promise<PendingDraft | null> {
  const note = takePendingDraft();
  if (note) return note;

  const draft = await pendingDraft(userEmail());
  if (!draft) return null;
  const current = await call<string>((done) => item.subject.getAsync(done));
  // 다른 케이스 ID가 제목에 있는 작성 창(담당자가 직접 연 회신 등)에는 붙이지 않는다
  const subjectCaseId = current.match(/LQ-\d{4}-\d{4}-\d{3}/)?.[0];
  if (subjectCaseId && subjectCaseId !== draft.caseId) return null;
  const isReply = /^(re|회신)\s*:/i.test(current);
  return {
    caseId: draft.caseId,
    kind: draft.kind,
    subject: draft.kind === "SUPPLEMENT" && isReply ? `RE: ${draft.subject}` : draft.subject,
    attachments: draft.attachments,
    createdAt: Date.now(),
  };
}

/** 제목·첨부를 채우고 초안함에 저장한 뒤 이력에 남긴다. 실패하면 예외를 던진다. */
export async function finalizeDraft(
  item: Office.MessageCompose,
  pending: PendingDraft,
  onStep: (message: string) => void = () => undefined
): Promise<void> {
  onStep("제목에 케이스 ID를 넣는 중…");
  await call<void>((done) => item.subject.setAsync(pending.subject, done));

  for (const file of pending.attachments) {
    onStep(`첨부 중: ${file.filename}`);
    const base64 = await fetchFileBase64(file.url);
    await call<string>((done) => item.addFileAttachmentFromBase64Async(base64, file.filename, done));
  }

  onStep("초안함에 저장하는 중…");
  await call<string>((done) => item.saveAsync(done));
  await recordMailEvent(pending.caseId, "DRAFT_SAVED", pending.subject, userEmail(), pending.kind);
}
