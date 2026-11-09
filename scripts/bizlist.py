#!/usr/bin/env python3
"""사업자번호 목록(비공개) 다루기.

이 저장소는 공개이므로 사업자번호는 평문으로 올리지 않는다.
엑셀 목록을 암호화한 private/biznos.enc 만 저장소에 두고, 암호(BIZNO_PASSPHRASE)는
GitHub Secret 과 로컬 키 파일에만 둔다.

목록을 새로 만들거나 갱신할 때 (로컬에서 실행, openssl 필요):
    BIZNO_PASSPHRASE=... python scripts/bizlist.py 사업자번호.xlsx

엑셀/CSV 형식: 한 행에 기업명과 사업자번호(하이픈 유무 무관)가 있으면 된다.
시트가 여러 개면 모두 읽고, 머리글 행은 자동으로 건너뛴다.
"""
import csv
import gzip
import hashlib
import hmac
import json
import os
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENC_FILE = Path(os.environ.get("BIZNO_FILE") or ROOT / "private" / "biznos.enc")
CIPHER = ["openssl", "enc", "-aes-256-cbc", "-pbkdf2", "-iter", "600000", "-md", "sha256"]
SHEET_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def passphrase():
    return os.environ.get("BIZNO_PASSPHRASE", "").strip()


# ---------------------------------------------------------------- 암호화

def _openssl(args, data, secret):
    result = subprocess.run(
        CIPHER + args + ["-pass", "env:BIZNO_PASSPHRASE"],
        input=data, capture_output=True, env={**os.environ, "BIZNO_PASSPHRASE": secret},
    )
    if result.returncode:
        raise ValueError("openssl 실행 실패")
    return result.stdout


def encrypt(entries, secret):
    raw = gzip.compress(json.dumps(entries, ensure_ascii=False).encode("utf-8"), mtime=0)
    return _openssl(["-salt"], raw, secret)


def decrypt(blob, secret):
    try:
        return json.loads(gzip.decompress(_openssl(["-d"], blob, secret)).decode("utf-8"))
    except Exception:
        raise ValueError("사업자번호 목록을 복호화하지 못했습니다. BIZNO_PASSPHRASE 값을 확인하세요.") from None


def public_id(b_no, secret):
    """공개 데이터에 쓰는 식별자. 암호 없이는 사업자번호로 되돌리거나 대조할 수 없다."""
    return hmac.new(secret.encode("utf-8"), b_no.encode("ascii"), hashlib.sha256).hexdigest()[:16]


# ---------------------------------------------------------------- 목록 읽기

def read_xlsx(path):
    rows = []
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for item in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(f"{SHEET_NS}si"):
                shared.append("".join(t.text or "" for t in item.iter(f"{SHEET_NS}t")))
        for name in sorted(n for n in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)):
            for row in ET.fromstring(z.read(name)).iter(f"{SHEET_NS}row"):
                cells = []
                for cell in row.findall(f"{SHEET_NS}c"):
                    value = cell.find(f"{SHEET_NS}v")
                    if cell.get("t") == "s" and value is not None:
                        cells.append(shared[int(value.text)])
                    elif cell.get("t") == "inlineStr":
                        cells.append("".join(t.text or "" for t in cell.iter(f"{SHEET_NS}t")))
                    else:
                        cells.append(value.text if value is not None and value.text else "")
                rows.append(cells)
    return rows


def read_csv(path):
    raw = Path(path).read_bytes()
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return list(csv.reader(raw.decode(encoding).splitlines()))
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 인코딩을 알 수 없습니다.")


def as_business_number(cell):
    text = str(cell).strip()
    if re.fullmatch(r"\d+(\.\d+)?[eE]\+?\d+", text):  # 숫자 셀이 지수 표기로 저장된 경우
        text = str(int(float(text)))
    digits = re.sub(r"[\s-]", "", text)
    return digits if re.fullmatch(r"\d{10}", digits) else None


def checksum_ok(b_no):
    weights = [1, 3, 7, 1, 3, 7, 1, 3, 5]
    total = sum(int(b_no[i]) * weights[i] for i in range(9)) + int(b_no[8]) * 5 // 10
    return (10 - total % 10) % 10 == int(b_no[9])


def parse_entries(rows):
    """행마다 사업자번호(10자리) 칸과 이름 칸을 찾는다. 번호가 없는 행(머리글 등)은 건너뛴다."""
    entries, seen, skipped = [], set(), 0
    for row in rows:
        number = next((n for n in map(as_business_number, row) if n), None)
        name = next((str(c).strip() for c in row if str(c).strip() and not as_business_number(c)), "")
        if not number:
            skipped += bool(name)
            continue
        if number in seen:
            continue
        seen.add(number)
        entries.append({"name": name, "b_no": number})
    return entries, skipped


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    secret = passphrase()
    if not secret:
        sys.exit("환경변수 BIZNO_PASSPHRASE 가 필요합니다.")
    source = Path(sys.argv[1])
    rows = read_csv(source) if source.suffix.lower() == ".csv" else read_xlsx(source)
    entries, skipped = parse_entries(rows)
    if not entries:
        sys.exit("사업자번호(10자리)가 있는 행을 찾지 못했습니다.")

    ENC_FILE.parent.mkdir(parents=True, exist_ok=True)
    ENC_FILE.write_bytes(encrypt(entries, secret))
    assert decrypt(ENC_FILE.read_bytes(), secret) == entries

    bad = [e["name"] for e in entries if not checksum_ok(e["b_no"])]
    print(f"{len(entries):,}건을 암호화해 {ENC_FILE} 에 저장했습니다. (번호 없는 행 {skipped}건 제외)")
    if bad:
        print(f"번호 검증(체크섬)에 맞지 않는 {len(bad)}건 - 오타 여부 확인 필요:")
        for name in bad:
            print("  -", name)


if __name__ == "__main__":
    main()
