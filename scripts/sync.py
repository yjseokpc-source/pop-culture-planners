#!/usr/bin/env python3
"""대중문화예술기획업 인허가 데이터 동기화 + 변동 알림 메일.

1. 공공데이터포털 API에서 전체 목록을 내려받는다.
2. 직전 스냅샷(data/planners.json)과 비교해 변동(상태 변경·신규·삭제)을 찾는다.
3. 변동이 있으면 메일을 보내고, 스냅샷과 변동 이력(data/changes.json)을 갱신한다.

표준 라이브러리만 사용한다 (Python 3.9+).

환경변수
  DATA_GO_KR_KEY  공공데이터포털 인증키 (Encoding/Decoding 모두 가능)   [필수]
  SMTP_USER       SMTP 로그인 계정 (예: Gmail 주소)                      [메일 발송 시 필수]
  SMTP_PASS       SMTP 비밀번호 (Gmail은 '앱 비밀번호')                  [메일 발송 시 필수]
  MAIL_TO         받는 사람. 여러 명이면 쉼표로 구분                     [메일 발송 시 필수]
  SMTP_HOST       기본 smtp.gmail.com
  SMTP_PORT       기본 465 (SSL). 587이면 STARTTLS
  MAIL_FROM       기본 SMTP_USER
  SITE_URL        메일 본문에 넣을 조회 페이지 주소 (선택)

옵션
  --no-mail   메일을 보내지 않고 데이터만 갱신
  --dry-run   아무것도 저장/발송하지 않고 결과와 메일 본문만 출력
"""
import argparse
import html
import json
import math
import os
import smtplib
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path

API_URL = os.environ.get("API_URL") or "https://apis.data.go.kr/1741000/pop_culture_art_planners/info"
PAGE_SIZE = 100  # API 허용 최대값
DATA_DIR = Path(os.environ.get("DATA_DIR") or Path(__file__).resolve().parent.parent / "data")
SNAPSHOT = DATA_DIR / "planners.json"
CHANGELOG = DATA_DIR / "changes.json"
CHANGELOG_LIMIT = 2000   # 변동 이력 보관 건수
MAIL_ROW_LIMIT = 300     # 메일 한 구역에 싣는 최대 건수
KST = timezone(timedelta(hours=9))

# 스냅샷에 보관할 항목 (API 응답 항목 중 조회 페이지에 쓰는 것만)
KEEP_FIELDS = [
    "OPN_ATMY_GRP_CD", "MNG_NO", "BPLC_NM",
    "SALS_STTS_CD", "SALS_STTS_NM", "DTL_SALS_STTS_CD", "DTL_SALS_STTS_NM",
    "LCPMT_YMD", "LCPMT_RTRCN_YMD", "CLSBIZ_YMD", "TCBIZ_BGNG_YMD", "TCBIZ_END_YMD",
    "ROAD_NM_ADDR", "LOTNO_ADDR", "TELNO", "CULTR_BZMN_SE_NM",
    "DAT_UPDT_PNT", "LAST_MDFCN_PNT",
]

# 변동을 감시할 항목 → 표시 이름. 여기에 항목을 추가하면 그 항목이 바뀔 때도 알림이 간다.
WATCH_FIELDS = {
    "SALS_STTS_NM": "영업상태",
    "DTL_SALS_STTS_NM": "상세영업상태",
    "BPLC_NM": "사업장명",
    "ROAD_NM_ADDR": "도로명주소",
    "LOTNO_ADDR": "지번주소",
    "TELNO": "전화번호",
}
STATUS_FIELDS = {"SALS_STTS_NM", "DTL_SALS_STTS_NM"}


class SyncError(Exception):
    pass


# ---------------------------------------------------------------- API

def normalize_key(key):
    """Encoding 키(%2B, %3D 포함)가 들어와도 이중 인코딩되지 않게 원래 값으로 되돌린다."""
    key = key.strip()
    return urllib.parse.unquote(key) if "%" in key else key


def gateway_error(text):
    """게이트웨이 오류 응답(JSON 또는 XML)에서 메시지를 뽑는다. 오류가 아니면 None."""
    try:
        head = json.loads(text)["OpenAPI_ServiceResponse"]["cmmMsgHeader"]
        return f"{head.get('returnAuthMsg', '')} ({head.get('errMsg', '')})"
    except (ValueError, KeyError, TypeError):
        pass
    if "<errMsg>" in text:
        code = text.split("<errMsg>")[1].split("<")[0]
        msg = text.split("<returnAuthMsg>")[1].split("<")[0] if "<returnAuthMsg>" in text else ""
        return f"{msg} ({code})"
    return None


def fetch_page(key, page):
    query = urllib.parse.urlencode({
        "serviceKey": key, "pageNo": page, "numOfRows": PAGE_SIZE, "returnType": "json",
    })
    request = urllib.request.Request(f"{API_URL}?{query}", headers={"User-Agent": "planner-sync/1.0"})

    last_error = None
    for attempt in range(4):
        if attempt:
            time.sleep(3 * attempt)
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                text = response.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            text = e.read().decode("utf-8", "replace")
            message = gateway_error(text) or f"HTTP {e.code}"
            if e.code in (400, 401, 403):  # 키/권한 문제는 재시도해도 소용없다
                raise SyncError(f"API 호출 거부: {message}")
            last_error = message
            continue
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            last_error = str(e)
            continue

        message = gateway_error(text)
        if message:
            raise SyncError(f"API 오류: {message}")
        try:
            body = json.loads(text)["response"]["body"]
            items = body.get("items") or {}
            rows = items.get("item", []) if isinstance(items, dict) else items
            if isinstance(rows, dict):
                rows = [rows]
            return int(body.get("totalCount") or 0), rows or []
        except (ValueError, KeyError, TypeError, AttributeError):
            last_error = "응답 형식이 예상과 다릅니다: " + text[:200]

    raise SyncError(f"{page}페이지 조회 실패: {last_error}")


def row_id(row):
    return f"{row.get('OPN_ATMY_GRP_CD', '')}:{row.get('MNG_NO', '')}"


def fetch_all(key):
    total, rows = fetch_page(key, 1)
    if not total or not rows:
        raise SyncError("전체 목록이 0건으로 조회되었습니다. (API 일시 장애 가능)")
    pages = math.ceil(total / PAGE_SIZE)
    print(f"전체 {total:,}건 / {pages}페이지")
    for page in range(2, pages + 1):
        time.sleep(0.2)
        _, more = fetch_page(key, page)
        if not more:
            break
        rows.extend(more)

    if len(rows) < total * 0.98:
        raise SyncError(f"일부만 수집되었습니다 ({len(rows):,}/{total:,}건). 이번 실행은 반영하지 않습니다.")

    result = {}
    for raw in rows:
        row = {k: str(raw[k]).strip() for k in KEEP_FIELDS if raw.get(k) not in (None, "")}
        result[row_id(row)] = row
    if len(result) < len(rows):
        print(f"::warning::관리번호가 겹치는 행 {len(rows) - len(result)}건은 하나로 합쳐졌습니다.")
    return result


# ---------------------------------------------------------------- 비교

def address(row):
    return row.get("ROAD_NM_ADDR") or row.get("LOTNO_ADDR") or ""


def diff(old, new, date):
    events = []
    for key, row in new.items():
        before = old.get(key)
        base = {"date": date, "id": key, "name": row.get("BPLC_NM", ""), "addr": address(row)}
        if before is None:
            events.append({**base, "type": "new", "status": row.get("SALS_STTS_NM", ""),
                           "licensed": row.get("LCPMT_YMD", "")})
            continue
        changes = [
            {"field": f, "label": label, "before": before.get(f, ""), "after": row.get(f, "")}
            for f, label in WATCH_FIELDS.items() if before.get(f, "") != row.get(f, "")
        ]
        if changes:
            events.append({**base, "type": "changed", "changes": changes})
    for key, row in old.items():
        if key not in new:
            events.append({"date": date, "id": key, "name": row.get("BPLC_NM", ""), "addr": address(row),
                           "type": "removed", "status": row.get("SALS_STTS_NM", "")})
    return events


def split_events(events):
    status = [e for e in events if e["type"] == "changed" and any(c["field"] in STATUS_FIELDS for c in e["changes"])]
    other = [e for e in events if e["type"] == "changed" and e not in status]
    new = [e for e in events if e["type"] == "new"]
    removed = [e for e in events if e["type"] == "removed"]
    return status, other, new, removed


# ---------------------------------------------------------------- 저장

def load_snapshot():
    if not SNAPSHOT.exists():
        return None
    data = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    return {row_id(row): row for row in data["items"]}


def dump_lines(rows):
    return ",\n".join(json.dumps(r, ensure_ascii=False, separators=(",", ":")) for r in rows)


def save(rows, events, now):
    """한 줄에 한 건씩, 관리번호 순으로 저장해 git 변경분이 실제 변동만큼만 생기게 한다."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ordered = [rows[k] for k in sorted(rows)]
    SNAPSHOT.write_text(
        '{"checkedAt":%s,"total":%d,"items":[\n%s\n]}\n' % (json.dumps(now), len(ordered), dump_lines(ordered)),
        encoding="utf-8",
    )
    history = json.loads(CHANGELOG.read_text(encoding="utf-8")) if CHANGELOG.exists() else []
    history = (events + history)[:CHANGELOG_LIMIT]
    CHANGELOG.write_text("[\n%s\n]\n" % dump_lines(history) if history else "[]\n", encoding="utf-8")


# ---------------------------------------------------------------- 메일

def fmt_date(value):
    return f"{value[:4]}-{value[4:6]}-{value[6:8]}" if len(value) == 8 and value.isdigit() else value


def describe(event):
    if event["type"] == "changed":
        return " / ".join(f"{c['label']}: {c['before'] or '(없음)'} → {c['after'] or '(없음)'}" for c in event["changes"])
    if event["type"] == "new":
        licensed = fmt_date(event.get("licensed", ""))
        return f"{event.get('status') or '-'}" + (f" (인허가일 {licensed})" if licensed else "")
    return f"직전 상태: {event.get('status') or '-'}"


def build_change_mail(events, total, date):
    status, other, new, removed = split_events(events)
    sections = [
        ("영업상태 변경", status), ("기타 정보 변경", other),
        ("신규 등록", new), ("목록에서 삭제", removed),
    ]
    counts = ", ".join(f"{title} {len(rows)}건" for title, rows in sections if rows)
    subject = f"[기획업 모니터링] {date} 변동 {len(events)}건 ({counts})"

    text = [f"대중문화예술기획업 등록 정보에 변동이 있습니다. (전체 {total:,}건)", ""]
    parts = [f"<p>대중문화예술기획업 등록 정보에 변동이 있습니다. (전체 {total:,}건)</p>"]
    cell = 'style="padding:6px 10px;border:1px solid #d5d9e0;text-align:left;vertical-align:top"'
    for title, rows in sections:
        if not rows:
            continue
        shown = rows[:MAIL_ROW_LIMIT]
        text.append(f"■ {title} ({len(rows)}건)")
        parts.append(f'<h3 style="margin:20px 0 8px">{title} ({len(rows)}건)</h3>'
                     '<table style="border-collapse:collapse;font-size:14px">'
                     f"<tr><th {cell}>사업장명</th><th {cell}>내용</th><th {cell}>소재지</th></tr>")
        for e in shown:
            text.append(f"- {e['name']} | {describe(e)} | {e['addr']}")
            parts.append(f"<tr><td {cell}><b>{html.escape(e['name'])}</b></td>"
                         f"<td {cell}>{html.escape(describe(e))}</td><td {cell}>{html.escape(e['addr'])}</td></tr>")
        parts.append("</table>")
        if len(rows) > len(shown):
            more = f"외 {len(rows) - len(shown)}건은 조회 페이지에서 확인하세요."
            text.append(more)
            parts.append(f"<p>{more}</p>")
        text.append("")
    return subject, text, parts


def build_start_mail(rows, date):
    counts = {}
    for row in rows.values():
        name = row.get("SALS_STTS_NM") or "(상태 없음)"
        counts[name] = counts.get(name, 0) + 1
    summary = [f"{name} {count:,}건" for name, count in sorted(counts.items(), key=lambda x: -x[1])]
    subject = f"[기획업 모니터링] 모니터링을 시작했습니다 (전체 {len(rows):,}건)"
    lead = f"{date} 기준 전체 {len(rows):,}건을 기준 목록으로 저장했습니다. 이후 변동이 생기면 이 주소로 알림 메일이 발송됩니다."
    text = [lead, ""] + [f"- {s}" for s in summary]
    parts = [f"<p>{lead}</p><ul>" + "".join(f"<li>{html.escape(s)}</li>" for s in summary) + "</ul>"]
    return subject, text, parts


def make_message(subject, text, parts):
    site = os.environ.get("SITE_URL", "").strip()
    if site:
        text.append(f"조회 페이지: {site}")
        parts.append(f'<p><a href="{html.escape(site)}">조회 페이지 열기</a></p>')
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = os.environ.get("MAIL_FROM") or os.environ.get("SMTP_USER", "")
    message["To"] = os.environ.get("MAIL_TO", "")
    message.set_content("\n".join(text))
    message.add_alternative(
        '<div style="font-family:\'Malgun Gothic\',sans-serif;font-size:14px;color:#16181d">%s</div>' % "".join(parts),
        subtype="html",
    )
    return message


def send(message):
    host = os.environ.get("SMTP_HOST") or "smtp.gmail.com"
    port = int(os.environ.get("SMTP_PORT") or 465)
    context = ssl.create_default_context()
    if port == 465:
        server = smtplib.SMTP_SSL(host, port, timeout=60, context=context)
    else:
        server = smtplib.SMTP(host, port, timeout=60)
        server.starttls(context=context)
    with server:
        server.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
        server.send_message(message)


# ---------------------------------------------------------------- 실행

def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--no-mail", action="store_true", help="메일을 보내지 않고 데이터만 갱신")
    parser.add_argument("--dry-run", action="store_true", help="저장/발송 없이 결과만 출력")
    args = parser.parse_args()
    mail_enabled = not (args.no_mail or args.dry_run)

    key = normalize_key(os.environ.get("DATA_GO_KR_KEY", ""))
    if not key:
        raise SyncError("환경변수 DATA_GO_KR_KEY(공공데이터포털 인증키)가 없습니다.")
    if mail_enabled:
        missing = [name for name in ("SMTP_USER", "SMTP_PASS", "MAIL_TO") if not os.environ.get(name, "").strip()]
        if missing:
            raise SyncError(f"메일 설정이 없습니다: {', '.join(missing)}. (데이터만 갱신하려면 --no-mail)")

    now = datetime.now(KST)
    date, stamp = now.strftime("%Y-%m-%d"), now.strftime("%Y-%m-%d %H:%M")

    old = load_snapshot()
    new = fetch_all(key)
    if old and len(new) < len(old) * 0.8:
        raise SyncError(f"목록이 비정상적으로 줄었습니다 ({len(old):,} → {len(new):,}건). 이번 실행은 반영하지 않습니다.")

    if old is None:
        events = []
        print(f"첫 실행: {len(new):,}건을 기준 목록으로 저장합니다.")
        mail = build_start_mail(new, date)
    else:
        events = diff(old, new, date)
        status, other, added, removed = split_events(events)
        print(f"변동 {len(events)}건 (상태 변경 {len(status)}, 기타 변경 {len(other)}, 신규 {len(added)}, 삭제 {len(removed)})")
        for e in events[:50]:
            print(f"  [{e['type']}] {e['name']} | {describe(e)}")
        mail = build_change_mail(events, len(new), date) if events else None

    message = make_message(*mail) if mail else None
    if args.dry_run:
        if message:
            print("\n----- 메일 미리보기 -----")
            print("제목:", message["Subject"])
            print(message.get_body(preferencelist=("plain",)).get_content())
        return

    # 메일이 실패하면 스냅샷을 갱신하지 않는다 → 다음 실행에서 같은 변동을 다시 알린다.
    if message and mail_enabled:
        send(message)
        print(f"메일 발송 완료 → {message['To']}")
    save(new, events, stamp)
    print(f"저장 완료: {SNAPSHOT}")


if __name__ == "__main__":
    try:
        main()
    except SyncError as e:
        print(f"::error::{e}")
        sys.exit(1)
    except smtplib.SMTPException as e:
        print(f"::error::메일 발송 실패: {e}")
        sys.exit(1)
