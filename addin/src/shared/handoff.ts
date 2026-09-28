/* global localStorage */

/*
 * 패널 → 작성 창 전달 쪽지.
 * 패널이 초안 작성 창을 띄우기 직전에 "어느 케이스의 어떤 초안인지"를 남기고,
 * 작성 창이 열릴 때 실행되는 이벤트 처리기(commands)가 이를 꺼내 제목·첨부를 채우고 초안함에 저장한다.
 * 둘 다 https://localhost:3000 에서 실행되므로 같은 localStorage를 본다.
 */

export interface PendingDraft {
  caseId: string;
  kind: "SUPPLEMENT" | "QUOTE";
  subject: string;
  attachments: { filename: string; url: string }[];
  createdAt: number;
}

const KEY = "leona.pendingDraft";
/** 이보다 오래된 쪽지는 무시한다 (담당자가 직접 연 다른 메일 작성 창에 잘못 붙지 않게) */
const MAX_AGE_MS = 2 * 60 * 1000;

export function savePendingDraft(draft: Omit<PendingDraft, "createdAt">): void {
  try {
    localStorage.setItem(KEY, JSON.stringify({ ...draft, createdAt: Date.now() }));
  } catch {
    // 저장소를 못 쓰는 환경이면 작성 창에서 제목의 케이스 ID로 대신 판별한다
  }
}

/** 쪽지를 꺼내고 지운다. 한 번만 쓰인다. */
export function takePendingDraft(): PendingDraft | null {
  try {
    const raw = localStorage.getItem(KEY);
    localStorage.removeItem(KEY);
    if (!raw) return null;
    const draft = JSON.parse(raw) as PendingDraft;
    return Date.now() - draft.createdAt <= MAX_AGE_MS ? draft : null;
  } catch {
    return null;
  }
}
