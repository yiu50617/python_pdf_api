from __future__ import annotations

import io
import json
import mimetypes
import os
import re
import sys
import uuid
from base64 import b64decode, b64encode
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[2]
APP_ROOT = Path(__file__).resolve().parent
STATIC_ROOT = APP_ROOT / "static"
OUTPUT_ROOT = APP_ROOT / "outputs"
CLAIM_TEMPLATE = ROOT / "【新】共済金請求書.pdf"
CERT_TEMPLATE = ROOT / "【新】申告書兼団体証明書.pdf"
PDF_WIDTH = 595.276
PDF_HEIGHT = 841.89
PREVIEW_WIDTH = 1240
PREVIEW_HEIGHT = 1754

FONT_NAME = "PrototypeJP"
FONT_BOLD_NAME = "PrototypeJP-Bold"
FONT_PATHS = [
    Path(r"C:\Windows\Fonts\meiryo.ttc"),
    Path(r"C:\Windows\Fonts\NotoSansJP-VF.ttf"),
    Path(r"C:\Windows\Fonts\YuGothM.ttc"),
    Path(r"C:\Windows\Fonts\msgothic.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"),
]
FONT_BOLD_PATHS = [
    Path(r"C:\Windows\Fonts\meiryob.ttc"),
    Path(r"C:\Windows\Fonts\YuGothB.ttc"),
    Path(r"C:\Windows\Fonts\NotoSansJP-VF.ttf"),
    Path(r"C:\Windows\Fonts\msgothic.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc"),
    Path("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"),
]

CLAIM_TYPES = {
    "parent_death": {"label": "親死亡給付", "code": "12", "amount": 20000},
    "sickness": {"label": "傷病見舞金", "code": "84", "amount": 20000},
    "marriage": {"label": "結婚祝金", "code": "31", "amount": 50000},
    "birth": {"label": "出生祝金", "code": "41", "amount": 30000},
    "long_service": {"label": "勤続祝金", "code": "73", "amount": 12000},
    "retirement": {"label": "退職餞別金", "code": "61", "amount": 10000},
}

BANK_KINDS = {
    "bank": "銀行",
    "shinkin": "金庫",
    "kumiai": "組合",
}

BRANCH_KINDS = {
    "branch": "店",
    "subbranch": "出張所",
}


def register_font() -> None:
    registered_regular = False
    registered_bold = False
    for path in FONT_PATHS:
        if path.exists():
            pdfmetrics.registerFont(TTFont(FONT_NAME, str(path)))
            registered_regular = True
            break
    for path in FONT_BOLD_PATHS:
        if path.exists():
            pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, str(path)))
            registered_bold = True
            break
    if not registered_regular or not registered_bold:
        raise RuntimeError("日本語フォントが見つかりません。C:\\Windows\\Fonts を確認してください。")


def normalize_date(value: str) -> tuple[str, str, str]:
    if not value:
        return "", "", ""
    digits = re.findall(r"\d+", value)
    if len(digits) >= 3:
        return digits[0], str(int(digits[1])), str(int(digits[2]))
    return value, "", ""


def normalize_date_range(value: str) -> tuple[tuple[str, str, str], tuple[str, str, str]]:
    digits = re.findall(r"\d+", value or "")
    if len(digits) >= 6:
        start = (digits[0], str(int(digits[1])), str(int(digits[2])))
        end = (digits[3], str(int(digits[4])), str(int(digits[5])))
        return start, end
    return normalize_date(value), ("", "", "")


def yen(value: int | str | None) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return str(value)


def full_name(data: dict, prefix: str) -> str:
    family = data.get(f"{prefix}FamilyName", "").strip()
    given = data.get(f"{prefix}GivenName", "").strip()
    return f"{family} {given}".strip()


def full_kana(data: dict, prefix: str) -> str:
    family = data.get(f"{prefix}FamilyKana", "").strip()
    given = data.get(f"{prefix}GivenKana", "").strip()
    return f"{family} {given}".strip()


def bank_name(data: dict) -> str:
    base = data.get("bankNameBase", "").strip()
    return f"{base}{BANK_KINDS.get(data.get('bankKind'), '')}" if base else ""


def branch_name(data: dict) -> str:
    base = data.get("branchNameBase", "").strip()
    return f"{base}{BRANCH_KINDS.get(data.get('branchKind'), '')}" if base else ""


def normalized_account_number(data: dict) -> str:
    account = re.sub(r"\D", "", data.get("accountNumber", ""))
    return account[-7:].zfill(7) if account else ""


def safe_filename_part(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', "", value or "")
    cleaned = re.sub(r"\s+", "", cleaned)
    return cleaned or "未入力"


def validate_input(data: dict) -> None:
    member_number = data.get("memberNumber", "")
    phone = data.get("phone", "")
    account = data.get("accountNumber", "")
    if not re.fullmatch(r"\d{7}", member_number):
        raise ValueError("組合員番号は半角数字7桁で入力してください。")
    if phone and not re.fullmatch(r"\d{11}", phone):
        raise ValueError("日中連絡先はハイフン不要の半角数字11桁で入力してください。")
    if account and not re.fullmatch(r"\d{1,7}", account):
        raise ValueError("口座番号は半角数字最大7桁で入力してください。")
    if data.get("eventDate") and data["eventDate"] <= "2025-03-31":
        raise ValueError("事由発生日が2025年3月31日以前のため、旧制度での申請手続きとなります。分会役員を通じて中央本部または地方本部へ連絡してください。")


def draw_text(c: canvas.Canvas, x: float, y: float, text: str, size: int = 10, color: str = "#111827", bold: bool = True) -> None:
    if not text:
        return
    c.setFillColor(HexColor(color))
    c.setFont(FONT_BOLD_NAME if bold else FONT_NAME, size)
    c.drawString(x, y, str(text))


def draw_center(c: canvas.Canvas, x: float, y: float, text: str, size: int = 10, bold: bool = True) -> None:
    if not text:
        return
    c.setFillColor(HexColor("#111827"))
    c.setFont(FONT_BOLD_NAME if bold else FONT_NAME, size)
    c.drawCentredString(x, y, str(text))


def draw_check(c: canvas.Canvas, x: float, y: float) -> None:
    c.setFillColor(HexColor("#111827"))
    c.setFont(FONT_BOLD_NAME, 13)
    c.drawString(x, y, "✓")


def draw_boxed_digits(c: canvas.Canvas, centers: list[float], y: float, text: str, size: int = 9) -> None:
    digits = re.sub(r"\D", "", text)
    c.setFillColor(HexColor("#111827"))
    c.setFont(FONT_BOLD_NAME, size)
    for x, digit in zip(centers[-len(digits) :], digits):
        c.drawCentredString(x, y, digit)


def preview_to_pdf(x: float, y: float) -> tuple[float, float]:
    """Convert 150dpi preview coordinates with top-left origin to PDF points."""
    return x * PDF_WIDTH / PREVIEW_WIDTH, PDF_HEIGHT - (y * PDF_HEIGHT / PREVIEW_HEIGHT)


def draw_text_preview(c: canvas.Canvas, x: float, y: float, text: str, size: int = 8, bold: bool = True) -> None:
    px, py = preview_to_pdf(x, y)
    draw_text(c, px, py, text, size=size, bold=bold)


def draw_center_preview(c: canvas.Canvas, x: float, y: float, text: str, size: int = 8, bold: bool = True) -> None:
    px, py = preview_to_pdf(x, y)
    draw_center(c, px, py, text, size=size, bold=bold)


def draw_check_preview(c: canvas.Canvas, x: float, y: float) -> None:
    px, py = preview_to_pdf(x, y)
    draw_check(c, px, py)


def draw_digits_preview(c: canvas.Canvas, positions: list[tuple[float, float]], text: str, size: int = 8) -> None:
    digits = re.sub(r"\D", "", text)
    for (x, y), digit in zip(positions[-len(digits) :], digits):
        draw_center_preview(c, x, y, digit, size=size)


def draw_digits_at(c: canvas.Canvas, positions: list[tuple[float, float]], text: str, size: int = 8) -> None:
    digits = re.sub(r"\D", "", text)
    for (x, y), digit in zip(positions[-len(digits) :], digits):
        draw_center(c, x, y, digit, size=size)


def overlay_pdf(width: float, height: float, draw_fn) -> PdfReader:
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    draw_fn(c)
    c.save()
    packet.seek(0)
    return PdfReader(packet)


def merge_template(template: Path, overlays: list[PdfReader | None], output: Path) -> None:
    reader = PdfReader(str(template))
    writer = PdfWriter()
    for index, page in enumerate(reader.pages):
        if index < len(overlays) and overlays[index] is not None:
            page.merge_page(overlays[index].pages[0])
        writer.add_page(page)
    with output.open("wb") as f:
        writer.write(f)


def draw_claim(data: dict):
    claim_type = CLAIM_TYPES.get(data.get("claimType"), CLAIM_TYPES["marriage"])
    claim_year, claim_month, claim_day = normalize_date(data.get("claimDate", ""))
    event_year, event_month, event_day = normalize_date(data.get("eventDate", ""))
    amount = str(claim_type["amount"]).zfill(5)
    code = claim_type["code"].zfill(2)

    def _draw(c: canvas.Canvas) -> None:
        # Coordinates below follow the user's latest convention:
        # actual PDF x = provided y, actual PDF y = provided x.
        draw_text(c, 410, 720, data.get("representativeName", ""), 9)

        draw_center(c, 95, 687, claim_year, 9)
        draw_center(c, 135, 687, claim_month, 9)
        draw_center(c, 165, 687, claim_day, 9)

        draw_text(c, 245, 520, data.get("memberFamilyName", ""), 8)
        draw_text(c, 325, 520, data.get("memberGivenName", ""), 8)
        draw_text(c, 235, 542, data.get("memberFamilyKana", ""), 6.5)
        draw_text(c, 325, 542, data.get("memberGivenKana", ""), 6.5)

        draw_center(c, 385, 510, event_year, 9)
        draw_center(c, 417, 510, event_month, 9)
        draw_center(c, 442, 510, event_day, 9)

        draw_digits_at(c, [(465, 515), (475, 515)], code, size=8)
        draw_digits_at(c, [(501, 515), (512, 515), (523, 515), (535, 515), (546, 515)], amount, size=8)
        draw_digits_at(c, [(477, 227), (495, 227), (512, 227), (527, 227), (543, 227)], amount, size=8)

        draw_text(c, 210, 185, data.get("bankNameBase", ""), 7.5)
        if data.get("bankKind") == "bank":
            draw_check(c, 248, 191)
        elif data.get("bankKind") == "shinkin":
            draw_check(c, 248, 181)
        elif data.get("bankKind") == "kumiai":
            draw_check(c, 248, 171)

        draw_text(c, 320, 185, data.get("branchNameBase", ""), 7.5)
        if data.get("branchKind") == "branch":
            draw_check(c, 352, 188)
        elif data.get("branchKind") == "subbranch":
            draw_check(c, 352, 174)

        if data.get("accountType") == "ordinary":
            draw_check(c, 387, 188)
        elif data.get("accountType") == "current":
            draw_check(c, 387, 174)
        draw_digits_at(
            c,
            [(443, 185), (462, 185), (479, 185), (496, 185), (512, 185), (528, 185), (544, 185)],
            normalized_account_number(data),
            size=8,
        )
        draw_text(c, 350, 161, data.get("accountHolderKana", ""), 7.5)
        draw_text(c, 350, 137, data.get("accountHolderName", ""), 8)

    return _draw


def draw_certificate(data: dict):
    member_birth = normalize_date(data.get("memberBirthDate", ""))

    def _draw(c: canvas.Canvas) -> None:
        draw_text(c, 210, 687, data.get("memberFamilyName", ""), 8)
        draw_text(c, 325, 687, data.get("memberGivenName", ""), 8)
        draw_text(c, 220, 710, data.get("memberFamilyKana", ""), 6.5)
        draw_text(c, 325, 710, data.get("memberGivenKana", ""), 6.5)
        draw_center(c, 415, 687, member_birth[0], 8)
        draw_center(c, 470, 687, member_birth[1], 8)
        draw_center(c, 520, 687, member_birth[2], 8)

        claim = data.get("claimType")
        if claim == "parent_death":
            draw_text(c, 200, 610, data.get("targetFamilyName", ""), 8)
            draw_text(c, 295, 610, data.get("targetGivenName", ""), 8)
            draw_text(c, 200, 627, data.get("targetFamilyKana", ""), 6.5)
            draw_text(c, 295, 627, data.get("targetGivenKana", ""), 6.5)
            draw_check(c, 160, 562)
            birth = normalize_date(data.get("targetBirthDate", ""))
            death = normalize_date(data.get("deathDate") or data.get("eventDate", ""))
            draw_center(c, 375, 610, birth[0], 8)
            draw_center(c, 415, 610, birth[1], 8)
            draw_center(c, 445, 610, birth[2], 8)
            if data.get("targetGender") == "male":
                draw_check(c, 481, 612)
            elif data.get("targetGender") == "female":
                draw_check(c, 517, 612)
            draw_center(c, 467, 557, death[0], 8)
            draw_center(c, 497, 557, death[1], 8)
            draw_center(c, 523, 557, death[2], 8)
        elif claim == "sickness":
            draw_text(c, 360, 537, data.get("sicknessName", ""), 8)
            for key, y in [("leavePeriod1", 513), ("leavePeriod2", 488), ("leavePeriod3", 466)]:
                start, end = normalize_date_range(data.get(key, ""))
                draw_center(c, 252, y, start[0], 8)
                draw_center(c, 295, y, start[1], 8)
                draw_center(c, 335, y, start[2], 8)
                draw_center(c, 435, y, end[0], 8)
                draw_center(c, 475, y, end[1], 8)
                draw_center(c, 515, y, end[2], 8)
            draw_text(c, 290, 442, data.get("hospitalName", ""), 8)
            draw_text(c, 495, 442, data.get("hospitalPhone", ""), 8)
        elif claim == "marriage":
            spouse_birth = normalize_date(data.get("spouseBirthDate", ""))
            marriage_date = normalize_date(data.get("marriageDate") or data.get("eventDate", ""))
            draw_text(c, 200, 391, data.get("spouseFamilyName", ""), 8)
            draw_text(c, 295, 391, data.get("spouseGivenName", ""), 8)
            draw_text(c, 200, 412, data.get("spouseFamilyKana", ""), 6.5)
            draw_text(c, 295, 412, data.get("spouseGivenKana", ""), 6.5)
            draw_check(c, 187, 352)
            draw_center(c, 360, 391, spouse_birth[0], 8)
            draw_center(c, 395, 391, spouse_birth[1], 8)
            draw_center(c, 425, 391, spouse_birth[2], 8)
            draw_center(c, 465, 391, marriage_date[0], 8)
            draw_center(c, 500, 391, marriage_date[1], 8)
            draw_center(c, 527, 391, marriage_date[2], 8)
            draw_text(c, 450, 352, data.get("marriageOffice", ""), 8)
        elif claim == "birth":
            child_birth = normalize_date(data.get("childBirthDate") or data.get("eventDate", ""))
            draw_text(c, 200, 302, data.get("childFamilyName", ""), 8)
            draw_text(c, 295, 302, data.get("childGivenName", ""), 8)
            draw_text(c, 200, 322, data.get("childFamilyKana", ""), 6.5)
            draw_text(c, 295, 322, data.get("childGivenKana", ""), 6.5)
            draw_center(c, 370, 302, child_birth[0], 8)
            draw_center(c, 407, 302, child_birth[1], 8)
            draw_center(c, 430, 302, child_birth[2], 8)
        elif claim == "long_service":
            start = normalize_date(data.get("membershipStartDate", ""))
            draw_center(c, 235, 185, start[0], 8)
            draw_center(c, 277, 185, start[1], 8)
            draw_center(c, 314, 185, start[2], 8)
            draw_check(c, 500, 190)
        elif claim == "retirement":
            retire = normalize_date(data.get("retirementDate") or data.get("eventDate", ""))
            if data.get("overThreeYears", "yes") == "yes":
                draw_check(c, 257, 160)
            else:
                draw_check(c, 302, 160)
            draw_center(c, 450, 157, retire[0], 8)
            draw_center(c, 497, 157, retire[1], 8)
            draw_center(c, 527, 157, retire[2], 8)
            if data.get("retirementReason") == "regular":
                draw_check(c, 222, 137)
            else:
                draw_check(c, 286, 137)

        cert_date = normalize_date(data.get("certificationDate") or data.get("claimDate", ""))
        draw_center(c, 100, 75, cert_date[0], 8)
        draw_center(c, 140, 75, cert_date[1], 8)
        draw_center(c, 175, 75, cert_date[2], 8)
        draw_text(c, 425, 50, data.get("representativeName", ""), 8)

    return _draw


def generate_pdfs(data: dict) -> dict:
    register_font()
    validate_input(data)
    claim_type = CLAIM_TYPES.get(data.get("claimType"), CLAIM_TYPES["marriage"])
    data["amount"] = claim_type["amount"]
    data["paymentMethod"] = "custom"
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    created = datetime.now()
    created_at = created.strftime("%Y%m%d_%H%M%S")
    unique = uuid.uuid4().hex[:8]
    applicant = safe_filename_part(f"{data.get('memberFamilyName', '')}{data.get('memberGivenName', '')}")
    benefit = safe_filename_part(claim_type["label"])
    claim_name = f"{created_at}_共済金申請書_{applicant}_{benefit}_{unique}.pdf"
    cert_name = f"{created_at}_申告書兼団体証明書_{applicant}_{benefit}_{unique}.pdf"
    claim_output = OUTPUT_ROOT / claim_name
    cert_output = OUTPUT_ROOT / cert_name

    claim_page = PdfReader(str(CLAIM_TEMPLATE)).pages[0]
    claim_overlay = overlay_pdf(float(claim_page.mediabox.width), float(claim_page.mediabox.height), draw_claim(data))
    merge_template(CLAIM_TEMPLATE, [claim_overlay], claim_output)

    cert_reader = PdfReader(str(CERT_TEMPLATE))
    cert_page = cert_reader.pages[0]
    cert_overlay = overlay_pdf(float(cert_page.mediabox.width), float(cert_page.mediabox.height), draw_certificate(data))
    merge_template(CERT_TEMPLATE, [cert_overlay, None], cert_output)

    return {
        "claimPdf": f"/outputs/{claim_output.name}",
        "certificatePdf": f"/outputs/{cert_output.name}",
        "createdAt": created.isoformat(timespec="seconds"),
    }


class Handler(BaseHTTPRequestHandler):
    def authenticate(self) -> bool:
        api_token = os.getenv("PDF_API_TOKEN")
        if api_token:
            bearer = self.headers.get("Authorization", "")
            if bearer == f"Bearer {api_token}":
                return True
        user = os.getenv("APP_BASIC_USER")
        password = os.getenv("APP_BASIC_PASSWORD")
        if not user or not password:
            return True
        header = self.headers.get("Authorization", "")
        if not header.startswith("Basic "):
            return False
        try:
            supplied = b64decode(header.removeprefix("Basic "), validate=True).decode("utf-8")
        except Exception:
            return False
        return supplied == f"{user}:{password}"

    def require_auth(self) -> bool:
        if self.authenticate():
            return True
        self.send_response(401)
        self.send_header("WWW-Authenticate", 'Basic realm="kyosai-prototype"')
        self.send_header("Content-Length", "0")
        self.end_headers()
        return False

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            self.send_json({"ok": True})
            return
        if not self.require_auth():
            return
        if parsed.path == "/":
            self.serve_file(STATIC_ROOT / "index.html")
            return
        if parsed.path.startswith("/static/"):
            self.serve_file(STATIC_ROOT / unquote(parsed.path.removeprefix("/static/")))
            return
        if parsed.path.startswith("/outputs/"):
            self.serve_file(OUTPUT_ROOT / unquote(parsed.path.removeprefix("/outputs/")))
            return
        self.send_error(404)

    def do_POST(self) -> None:
        if not self.require_auth():
            return
        path = urlparse(self.path).path
        if path not in ["/api/generate", "/api/generate-base64"]:
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            data = json.loads(self.rfile.read(length).decode("utf-8"))
            result = generate_pdfs(data)
            if path == "/api/generate-base64":
                claim_path = OUTPUT_ROOT / result["claimPdf"].removeprefix("/outputs/")
                cert_path = OUTPUT_ROOT / result["certificatePdf"].removeprefix("/outputs/")
                result.update(
                    {
                        "claimPdfBase64": b64encode(claim_path.read_bytes()).decode("ascii"),
                        "certificatePdfBase64": b64encode(cert_path.read_bytes()).decode("ascii"),
                        "claimPdfName": claim_path.name,
                        "certificatePdfName": cert_path.name,
                    }
                )
            self.send_json({"ok": True, **result})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, status=500)

    def serve_file(self, path: Path) -> None:
        try:
            path = path.resolve()
            allowed = [STATIC_ROOT.resolve(), OUTPUT_ROOT.resolve()]
            if not any(os.path.commonpath([str(path), str(root)]) == str(root) for root in allowed):
                self.send_error(403)
                return
            if not path.exists() or not path.is_file():
                self.send_error(404)
                return
            content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(path.stat().st_size))
            self.end_headers()
            with path.open("rb") as f:
                self.wfile.write(f.read())
        except Exception:
            self.send_error(500)

    def send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    port = int(os.getenv("PORT") or (sys.argv[1] if len(sys.argv) > 1 else 8787))
    host = os.getenv("HOST", "127.0.0.1")
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Prototype server: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
