/* global document, Office */

import { readSelectedMail } from "./mail";
import type { MailSnapshot } from "./mail";

/** 본문 미리보기로 보여줄 글자 수 */
const PREVIEW_LENGTH = 300;

// Office.js 준비가 끝나면 한 번 실행된다. Outlook 안에서 열렸을 때만 패널을 보여준다.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "block";
    showSelectedMail();
  }
});

async function showSelectedMail() {
  setStatus("메일 정보를 읽는 중…", false);
  try {
    const mail = await readSelectedMail();
    render(mail);
    setStatus("", false);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`메일 정보를 읽지 못했습니다: ${message}`, true);
  }
}

function render(mail: MailSnapshot) {
  setText("subject", mail.subject || "(제목 없음)");
  setText("from", mail.from.name ? `${mail.from.name} <${mail.from.email}>` : mail.from.email);
  setText("received-at", mail.receivedAt ? new Date(mail.receivedAt).toLocaleString("ko-KR") : "-");
  setText("body-preview", toPreview(mail.bodyText));
  setText("conversation-id", mail.conversationId || "-");
  setText("internet-message-id", mail.internetMessageId || "-");
  setText("snapshot-json", JSON.stringify(mail, null, 2));
}

/** 줄바꿈·연속 공백을 한 칸으로 줄이고 앞부분만 잘라낸다. */
function toPreview(text: string): string {
  const flat = text.replace(/\s+/g, " ").trim();
  return flat.length > PREVIEW_LENGTH ? `${flat.slice(0, PREVIEW_LENGTH)}…` : flat || "(본문 없음)";
}

/** 메일 내용은 외부에서 온 값이므로 HTML로 해석하지 않고 textContent로만 넣는다(스크립트 삽입 방지). */
function setText(id: string, value: string) {
  document.getElementById(id)!.textContent = value;
}

function setStatus(message: string, isError: boolean) {
  const status = document.getElementById("status")!;
  status.textContent = message;
  status.className = isError ? "status status--error" : "status";
  status.hidden = message === "";
}
