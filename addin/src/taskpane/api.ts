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

export interface FileLink {
  filename: string;
  size: number;
  url: string;
}

export interface CaseDetail extends CaseSummary {
  validityDays: number | null;
  mails: {
    index: number;
    role: string;
    source: string;
    hasRaw: boolean;
    rawUrl: string | null;
    attachmentFiles: FileLink[];
    snapshot: { subject: string; receivedAt: string | null; from: { name: string; email: string } };
  }[];
  events: { eventType: string; actor: string | null; detail: Record<string, string> | null; createdAt: string }[];
  extraction: ExtractionState;
}

/** 가장 최근 항목 추출 실행 상태 (FR-101: 추출 수행, 60초 이내) */
export interface ExtractionState {
  state: "none" | "running" | "done" | "not-connected" | "failed";
  trigger: string | null;
  startedAt: string | null;
  finishedAt: string | null;
  elapsedMs: number | null;
  withinLimit: boolean | null;
  result: Record<string, unknown> | null;
  validation: Record<string, unknown> | null;
  error: string | null;
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

/**
 * API 주소. 패널·작성 창(HTML)에서는 자기 주소(https://localhost:3000)를 쓰고,
 * 주소가 없는 실행 환경(데스크톱 Outlook의 JavaScript 전용 런타임)에서는 개발 서버 주소를 쓴다.
 */
const API_ORIGIN =
  typeof location !== "undefined" && location.origin.startsWith("http") ? location.origin : "https://localhost:3000";

export function apiUrl(path: string): string {
  return path.startsWith("/api") ? `${API_ORIGIN}${path}` : `${API_ORIGIN}/api${path}`;
}

async function request<T>(method: "GET" | "POST", path: string, body?: unknown): Promise<T> {
  let response: Response;
  try {
    response = await fetch(apiUrl(path), {
      method,
      headers: body ? { "Content-Type": "application/json" } : undefined,
      body: body ? JSON.stringify(body) : undefined,
    });
  } catch {
    throw new ApiError("백엔드 서버에 연결할 수 없습니다. server 폴더에서 API 서버가 실행 중인지 확인하세요.");
  }
  if (!response.ok) {
    const data = await response.json().catch(() => null);
    // 400·404·409 + 문장 안내는 백엔드가 담당자에게 보여주려고 쓴 문구다 → 그대로 보여준다 (NFR-08)
    if (response.status < 500 && typeof data?.detail === "string") throw new ApiError(data.detail);
    const detail = data?.detail !== undefined ? JSON.stringify(data.detail) : response.statusText;
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

export interface Draft {
  kind: "SUPPLEMENT" | "QUOTE";
  caseId: string;
  to: { name: string; email: string }[];
  subject: string;
  htmlBody: string;
  attachments: { filename: string; url: string; size: number }[];
}

/** [다시 추출] NFR-03 — 추출→검증을 다시 실행한다. 담당자가 고친 필드는 유지된다 */
export function rerunExtraction(caseId: string, actor: string) {
  return request<CaseDetail>("POST", `/cases/${encodeURIComponent(caseId)}/extract`, { actor });
}

export function setQuoteValidity(caseId: string, validityDays: number, actor: string) {
  return request<CaseDetail>("POST", `/cases/${encodeURIComponent(caseId)}/quote-validity`, { validityDays, actor });
}

/** [보완 요청 초안] FR-304·305 — 초안 내용만 받는다. 발송 기능은 없다 */
export function supplementDraft(caseId: string, questions: string[], actor: string) {
  return request<Draft>("POST", `/cases/${encodeURIComponent(caseId)}/drafts/supplement`, { questions, actor });
}

/** [견적서 송부 초안] FR-505 */
export function quoteDraft(caseId: string, actor: string) {
  return request<Draft>("POST", `/cases/${encodeURIComponent(caseId)}/drafts/quote`, { actor });
}

/** 이 담당자가 방금 패널에서 준비한 초안 (작성 창 버튼의 보조 경로). 없으면 null */
export function pendingDraft(actor: string) {
  return request<Draft | null>("GET", `/drafts/pending?actor=${encodeURIComponent(actor)}`);
}

/** [초안 저장·발송 기록] FR-505 — 기록만 한다. 발송은 담당자가 Outlook에서 직접 누른다 */
export function recordMailEvent(
  caseId: string,
  eventType: "DRAFT_SAVED" | "MAIL_SENT",
  subject: string,
  actor: string,
  kind?: string,
  sent?: { internetMessageId: string; occurredAt: string; emlBase64: string | null }
) {
  return request<CaseDetail>("POST", `/cases/${encodeURIComponent(caseId)}/mail-events`, {
    eventType,
    kind,
    subject,
    actor,
    ...sent,
  });
}

/** 견적서 파일을 Base64로 받아 온다 (작성 창에 첨부하기 위해) */
export async function fetchFileBase64(url: string): Promise<string> {
  const response = await fetch(apiUrl(url));
  if (!response.ok) throw new ApiError(`첨부파일을 받지 못했습니다 (${response.status})`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

/** 케이스 상세 (메일·첨부 다시 열기, 추출 상태) */
export function getCase(caseId: string) {
  return request<CaseDetail>("GET", `/cases/${encodeURIComponent(caseId)}`);
}

/** 최근 케이스 목록 — 자동 식별 실패 시 직접 선택용 */
export function listCases() {
  return request<CaseSummary[]>("GET", "/cases?limit=30");
}
