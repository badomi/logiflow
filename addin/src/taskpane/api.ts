/*
 * 백엔드(A트랙 API) 호출 모음.
 * 개발 중에는 애드인 개발 서버(https://localhost:3000)가 /api 요청을 Python 서버(localhost:8000)로 넘겨준다.
 * 메일 본문은 이 로컬 서버로만 가며, 외부 서비스로는 보내지 않는다 (NFR-04).
 */

import type { MailSnapshot } from "./mail";

export interface MailInput {
  emlBase64: string | null;
  snapshot: MailSnapshot;
  actor: string;
}

export interface CaseSummary {
  caseId: string;
  status: string;
  subject: string;
  fromEmail: string;
  createdAt: string;
}

export interface CaseDetail extends CaseSummary {
  mails: { role: string; source: string; hasRaw: boolean }[];
}

export interface ReplyCandidate {
  caseId: string;
  subject: string;
  score: number;
  reasons: string[];
}

export interface ReplyMatchResult {
  alreadyRegisteredCaseId: string | null;
  candidates: ReplyCandidate[];
}

export class ApiError extends Error {}

async function request<T>(method: "GET" | "POST", path: string, body?: unknown): Promise<T> {
  let response: Response;
  try {
    response = await fetch(`/api${path}`, {
      method,
      headers: body ? { "Content-Type": "application/json" } : undefined,
      body: body ? JSON.stringify(body) : undefined,
    });
  } catch {
    throw new ApiError("백엔드 서버에 연결할 수 없습니다. server 폴더에서 API 서버가 실행 중인지 확인하세요.");
  }
  if (!response.ok) {
    const detail = await response
      .json()
      .then((data) => (typeof data.detail === "string" ? data.detail : JSON.stringify(data.detail)))
      .catch(() => response.statusText);
    throw new ApiError(`요청 실패 (${response.status}): ${detail}`);
  }
  return (await response.json()) as T;
}

/** [케이스 생성] FR-101·102·103·105 */
export function createCase(input: MailInput) {
  return request<{ duplicate: boolean; case: CaseDetail }>("POST", "/cases", input);
}

/** [회신 식별] FR-104 — 저장하지 않고 후보만 받는다 */
export function matchReply(input: MailInput) {
  return request<ReplyMatchResult>("POST", "/replies/match", input);
}

/** [회신 병합] FR-104·207 — 담당자가 고른 케이스에 합친다 */
export function mergeReply(caseId: string, input: MailInput) {
  return request<CaseDetail>("POST", `/cases/${encodeURIComponent(caseId)}/replies`, input);
}

/** 최근 케이스 목록 — 자동 식별 실패 시 직접 선택용 */
export function listCases() {
  return request<CaseSummary[]>("GET", "/cases?limit=30");
}
