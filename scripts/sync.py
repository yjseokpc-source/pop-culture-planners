#!/usr/bin/env python3
"""대중문화예술기획업 인허가 데이터 동기화 + 변동 알림 메일.

1. 공공데이터포털 API에서 인허가 전체 목록을 내려받는다.
2. 직전 스냅샷(data/planners.json)과 비교해 변동(상태 변경·신규·삭제)을 찾는다.
3. private/biznos.enc(암호화된 사업자번호 목록)가 있으면 국세청 사업자등록 상태도 조회해
   직전 결과(data/nts.json)와 비교한다. 공개 파일과 실행 로그에는 사업자번호를 남기지 않는다.
4. 변동이 있으면 메일을 보내고, 스냅샷과 변동 이력(data/changes.json)을 갱신한다.

표준 라이브러리만 사용한다 (Python 3.9+).

환경변수
  DATA_GO_KR_KEY  공공데이터포털 인증키 (Encoding/Decoding 모두 가능)   [필수]
  BIZNO_PASSPHRASE  사업자번호 목록 암호                                 [국세청 조회 시 필수]
  NTS_API_KEY     국세청 API 인증키. 없으면 DATA_GO_KR_KEY 를 쓴다
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

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bizlist  # noqa: E402

API_URL = os.environ.get("API_URL") or "https://apis.data.go.kr/1741000/pop_culture_art_planners/info"
PAGE_SIZE = 100  # API 허용 최대값
DATA_DIR = Path(os.environ.get("DATA_DIR") or Path(__file__).resolve().parent.parent / "data")
SNAPSHOT = DATA_DIR / "planners.json"
CHANGELOG = DATA_DIR / "changes.json"
NTS_URL = os.environ.get("NTS_API_URL") or "https://api.odcloud.kr/api/nts-businessman/v1/status"
NTS_BATCH = 100          # 국세청 API 1회 최대 조회 건수
NTS_SNAPSHOT = DATA_DIR / "nts.json"
NTS_UNREGISTERED = "미등록(번호 확인 필요)"
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

# 국세청 조회 결과 중 변동을 감시할 항목
NTS_WATCH_FIELDS = {"status": "사업자 상태", "tax": "과세유형"}


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
    """인허가 변동을 종류별로 나눈다. 국세청 변동(source == 'nts')은 여기서 다루지 않는다."""
    events = [e for e in events if e.get("source") != "nts"]
    status =[e for e in events if e["type"] == "changed" and any(c["field"] in STATUS_FIELDS for c in e["changes"])]
    other = [e for e in events if e["type"] == "changed" and e not in status]
    new = [e for e in events if e["type"] == "new"]
    removed = [e for e in events if e["type"] == "removed"]
    return status, other, new, removed


# ---------------------------------------------------------------- 국세청

def fetch_nts_batch(url, numbers, label):
    """사업자번호 묶음의 상태를 조회한다. 오류 메시지에 사업자번호가 섞이지 않게 한다 (실행 로그가 공개됨)."""
    request = urllib.request.Request(
        url, data=json.dumps({"b_no": numbers}).encode("utf-8"),
        headers={"Content-Type": "application/json", "Accept": "application/json", "User-Agent": "planner-sync/1.0"},
    )
    last_error = None
    for attempt in range(4):
        if attempt:
            time.sleep(3 * attempt)
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = json.loads(response.read().decode("utf-8", "replace"))
        except urllib.error.HTTPError as e:
            try:
                body = json.loads(e.read().decode("utf-8", "replace"))
                message = str(body.get("msg") or body.get("status_code") or f"HTTP {e.code}")
            except ValueError:
                message = f"HTTP {e.code}"
            if e.code in (400, 401, 403):
                raise SyncError(f"국세청 API 호출 거부: {message}")
            last_error = message
            continue
        except (urllib.error.URLError, TimeoutError, OSError, ValueError) as e:
            last_error = type(e).__name__
            continue
        rows = payload.get("data")
        if payload.get("status_code") == "OK" and isinstance(rows, list) and len(rows) == len(numbers):
            return rows
        last_error = f"응답 상태 {payload.get('status_code')}"
    raise SyncError(f"국세청 조회 실패 ({label}): {last_error}")


def sync_nts(date):
    """국세청 상태를 조회해 직전 결과와 비교한다. 사업자번호 목록 파일이 없으면 None."""
    if not bizlist.ENC_FILE.exists():
        return None
    secret = bizlist.passphrase()
    if not secret:
        raise SyncError("사업자번호 목록은 있는데 BIZNO_PASSPHRASE 가 없습니다.")
    try:
        entries = bizlist.decrypt(bizlist.ENC_FILE.read_bytes(), secret)
    except ValueError as e:
        raise SyncError(str(e))

    key = normalize_key(os.environ.get("NTS_API_KEY") or os.environ.get("DATA_GO_KR_KEY", ""))
    url = f"{NTS_URL}?{urllib.parse.urlencode({'serviceKey': key})}"
    batches = math.ceil(len(entries) / NTS_BATCH)
    print(f"국세청 조회 {len(entries):,}건 / {batches}회")

    rows, numbers = {}, {}
    for index in range(batches):
        chunk = entries[index * NTS_BATCH:(index + 1) * NTS_BATCH]
        answers = {a.get("b_no", ""): a
                   for a in fetch_nts_batch(url, [e["b_no"] for e in chunk], f"{index + 1}/{batches}")}
        for entry in chunk:
            answer = answers.get(entry["b_no"]) or {}
            registered = bool(answer.get("b_stt"))
            record = {
                "id": bizlist.public_id(entry["b_no"], secret),
                "name": entry["name"],
                "status": answer.get("b_stt") if registered else NTS_UNREGISTERED,
                "tax": answer.get("tax_type", "") if registered else "",
                "end": answer.get("end_dt", ""),
            }
            rows[record["id"]] = {k: v for k, v in record.items() if v}
            numbers[record["id"]] = entry["b_no"]
        if (index + 1) % 10 == 0:
            print(f"  국세청 조회 진행 {index + 1}/{batches}")
        time.sleep(0.1)

    old = None
    if NTS_SNAPSHOT.exists():
        old = {r["id"]: r for r in json.loads(NTS_SNAPSHOT.read_text(encoding="utf-8"))["items"]}
    # 직전 결과가 없거나(첫 실행) 식별자가 하나도 겹치지 않으면(암호 교체) 비교 없이 기준만 다시 잡는다.
    baseline = not old or not (old.keys() & rows.keys())
    events = []
    if not baseline:
        for key_id, row in rows.items():
            before = old.get(key_id)
            if before is None:  # 목록에 새로 추가한 번호는 알림 대상이 아니다
                continue
            changes = [
                {"field": f, "label": label, "before": before.get(f, ""), "after": row.get(f, "")}
                for f, label in NTS_WATCH_FIELDS.items() if before.get(f, "") != row.get(f, "")
            ]
            if changes:
                events.append({"date": date, "id": key_id, "name": row.get("name", ""), "source": "nts",
                               "type": "changed", "changes": changes})
    return {"rows": rows, "events": events, "baseline": baseline, "numbers": numbers}


def count_by(rows, field):
    counts = {}
    for row in rows:
        name = row.get(field) or "(상태 없음)"
        counts[name] = counts.get(name, 0) + 1
    return [f"{name} {count:,}건" for name, count in sorted(counts.items(), key=lambda x: -x[1])]


# ---------------------------------------------------------------- 저장

def load_snapshot():
    if not SNAPSHOT.exists():
        return None
    data = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    return {row_id(row): row for row in data["items"]}


def dump_lines(rows):
    return ",\n".join(json.dumps(r, ensure_ascii=False, separators=(",", ":")) for r in rows)


def write_snapshot(path, rows, now):
    """한 줄에 한 건씩, 식별자 순으로 저장해 git 변경분이 실제 변동만큼만 생기게 한다."""
    ordered = [rows[k] for k in sorted(rows)]
    path.write_text(
        '{"checkedAt":%s,"total":%d,"items":[\n%s\n]}\n' % (json.dumps(now), len(ordered), dump_lines(ordered)),
        encoding="utf-8",
    )


def save(rows, nts, events, now):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    write_snapshot(SNAPSHOT, rows, now)
    if nts:
        write_snapshot(NTS_SNAPSHOT, nts["rows"], now)
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


def fmt_bizno(value):
    return f"{value[:3]}-{value[3:5]}-{value[5:]}" if len(value) == 10 else value


class Mail:
    """본문을 일반 텍스트와 HTML 두 가지로 함께 쌓는다."""
    CELL = 'style="padding:6px 10px;border:1px solid #d5d9e0;text-align:left;vertical-align:top"'

    def __init__(self):
        self.text, self.parts = [], []

    def paragraph(self, line):
        self.text += [line, ""]
        self.parts.append(f"<p>{html.escape(line)}</p>")

    def bullets(self, items):
        self.text += [f"- {item}" for item in items] + [""]
        self.parts.append("<ul>" + "".join(f"<li>{html.escape(item)}</li>" for item in items) + "</ul>")

    def table(self, title, events, last_header, last_value):
        cell, shown = self.CELL, events[:MAIL_ROW_LIMIT]
        self.text.append(f"■ {title} ({len(events)}건)")
        self.parts.append(f'<h3 style="margin:20px 0 8px">{title} ({len(events)}건)</h3>'
                          '<table style="border-collapse:collapse;font-size:14px">'
                          f"<tr><th {cell}>사업장명</th><th {cell}>내용</th><th {cell}>{last_header}</th></tr>")
        for e in shown:
            last = last_value(e)
            self.text.append(f"- {e['name']} | {describe(e)} | {last}")
            self.parts.append(f"<tr><td {cell}><b>{html.escape(e['name'])}</b></td>"
                              f"<td {cell}>{html.escape(describe(e))}</td><td {cell}>{html.escape(last)}</td></tr>")
        self.parts.append("</table>")
        if len(events) > len(shown):
            more = f"외 {len(events) - len(shown)}건은 조회 페이지에서 확인하세요."
            self.text.append(more)
            self.parts.append(f"<p>{more}</p>")
        self.text.append("")


def build_mail(date, planners, planner_events, first_run, nts):
    """보낼 내용이 없으면 None, 있으면 (제목, Mail)."""
    status, other, added, removed = split_events(planner_events)
    nts_events = nts["events"] if nts else []
    nts_status = [e for e in nts_events if any(c["field"] == "status" for c in e["changes"])]
    nts_tax = [e for e in nts_events if e not in nts_status]

    def addr(e):
        return e.get("addr", "")

    def number(e):  # 사업자번호는 메일에만 싣는다 (공개 파일에는 없음)
        return fmt_bizno(nts["numbers"].get(e["id"], ""))

    sections = [s for s in [
        ("인허가 영업상태 변경", status, "소재지", addr),
        ("인허가 기타 정보 변경", other, "소재지", addr),
        ("인허가 신규 등록", added, "소재지", addr),
        ("인허가 목록에서 삭제", removed, "소재지", addr),
        ("국세청 사업자 상태 변경", nts_status, "사업자번호", number),
        ("국세청 과세유형 변경", nts_tax, "사업자번호", number),
    ] if s[1]]
    changed = sum(len(s[1]) for s in sections)
    nts_started = bool(nts and nts["baseline"])

    mail = Mail()
    if first_run:
        subject = f"[기획업 모니터링] 모니터링을 시작했습니다 (전체 {len(planners):,}건)"
        mail.paragraph(f"{date} 기준 인허가 전체 {len(planners):,}건을 기준 목록으로 저장했습니다. "
                       "이후 변동이 생기면 이 주소로 알림 메일이 발송됩니다.")
        mail.bullets(count_by(planners.values(), "SALS_STTS_NM"))
    elif changed:
        counts = ", ".join(f"{title} {len(rows)}건" for title, rows, _, _ in sections)
        subject = f"[기획업 모니터링] {date} 변동 {changed}건 ({counts})"
        mail.paragraph(f"대중문화예술기획업 등록 정보에 변동이 있습니다. (인허가 전체 {len(planners):,}건)")
    elif nts_started:
        subject = f"[기획업 모니터링] 국세청 사업자 상태 모니터링을 시작했습니다 ({len(nts['rows']):,}건)"
    else:
        return None

    if nts_started:
        mail.paragraph(f"{date} 기준 국세청 사업자등록 상태 {len(nts['rows']):,}건을 기준 목록으로 저장했습니다. "
                       "이후 사업자 상태나 과세유형이 바뀌면 알림 메일이 발송됩니다.")
        mail.bullets(count_by(nts["rows"].values(), "status"))
    for section in sections:
        mail.table(*section)
    return subject, mail


def make_message(subject, mail):
    site = os.environ.get("SITE_URL", "").strip()
    if site:
        mail.text.append(f"조회 페이지: {site}")
        mail.parts.append(f'<p><a href="{html.escape(site)}">조회 페이지 열기</a></p>')
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = os.environ.get("MAIL_FROM") or os.environ.get("SMTP_USER", "")
    message["To"] = os.environ.get("MAIL_TO", "")
    message.set_content("\n".join(mail.text))
    message.add_alternative(
        '<div style="font-family:\'Malgun Gothic\',sans-serif;font-size:14px;color:#16181d">%s</div>'
        % "".join(mail.parts),
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

    # 1) 인허가 목록
    old = load_snapshot()
    new = fetch_all(key)
    if old and len(new) < len(old) * 0.8:
        raise SyncError(f"목록이 비정상적으로 줄었습니다 ({len(old):,} → {len(new):,}건). 이번 실행은 반영하지 않습니다.")
    first_run = old is None
    events = [] if first_run else diff(old, new, date)
    if first_run:
        print(f"첫 실행: {len(new):,}건을 기준 목록으로 저장합니다.")
    else:
        status, other, added, removed = split_events(events)
        print(f"인허가 변동 {len(events)}건 (상태 변경 {len(status)}, 기타 변경 {len(other)}, "
              f"신규 {len(added)}, 삭제 {len(removed)})")

    # 2) 국세청 사업자 상태. 여기서 실패해도 인허가 쪽 알림은 그대로 보낸다.
    nts, nts_error = None, None
    try:
        nts = sync_nts(date)
    except SyncError as e:
        nts_error = str(e)
        print(f"::error::{e}")
    if nts:
        events += nts["events"]
        print("국세청: 기준 목록 저장" if nts["baseline"] else f"국세청 변동 {len(nts['events'])}건")

    for e in events[:50]:
        print(f"  [{e.get('source', 'permit')}:{e['type']}] {e['name']} | {describe(e)}")

    built = build_mail(date, new, events, first_run, nts)
    message = make_message(*built) if built else None
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
    save(new, nts, events, stamp)
    print(f"저장 완료: {DATA_DIR}")
    if nts_error:
        raise SyncError(f"국세청 조회에 실패했습니다 (인허가 데이터는 정상 반영됨): {nts_error}")


if __name__ == "__main__":
    try:
        main()
    except SyncError as e:
        print(f"::error::{e}")
        sys.exit(1)
    except smtplib.SMTPException as e:
        print(f"::error::메일 발송 실패: {e}")
        sys.exit(1)
