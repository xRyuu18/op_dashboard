"""
Client tipis ke OpenProject API v3.

Ini setara "SendOpenProjectRequest" + fungsi-fungsi ekstraksi ID di macro
VBA. Bedanya, di sini kita pakai `response.json()` beneran (bukan regex
nebak-nebak pola teks), jadi semua bug ambiguitas parsing yang kejadian
di VBA (ID sub-project ketuker ID project parent, dll) otomatis nggak
ada lagi -- kita ambil field langsung dari path yang benar.
"""
from __future__ import annotations

import re
from typing import Any, Optional

import requests


class OpenProjectError(Exception):
    """Dilempar kalau request ke OpenProject gagal di level fatal (network dll)."""


class OpenProjectClient:
    def __init__(self, base_url: str, api_key: str, timeout: float = 30.0):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.auth = ("apikey", api_key)
        self.session.headers.update({"Content-Type": "application/json"})

    def request(self, method: str, path: str, json_body: Optional[dict] = None,
                params: Optional[dict] = None) -> requests.Response:
        url = path if path.startswith("http") else f"{self.base_url}{path}"
        try:
            return self.session.request(
                method, url, json=json_body, params=params, timeout=self.timeout
            )
        except requests.RequestException as e:
            raise OpenProjectError(f"HTTP error: {e}") from e

    def get(self, path: str, params: Optional[dict] = None) -> requests.Response:
        return self.request("GET", path, params=params)

    def post(self, path: str, json_body: dict) -> requests.Response:
        return self.request("POST", path, json_body=json_body)

    def patch(self, path: str, json_body: dict) -> requests.Response:
        return self.request("PATCH", path, json_body=json_body)


def extract_error_detail(resp: requests.Response) -> str:
    """
    Setara ExtractErrorDetail di VBA: OpenProject error 422 "MultipleErrors"
    nyimpen alasan sebenarnya di dalam _embedded.errors[].message, bukan
    cuma di message utama. Kumpulin SEMUA pesan "message" yang ketemu
    (level manapun) biar kelihatan field mana yang bermasalah.
    """
    try:
        data = resp.json()
    except ValueError:
        return resp.text[:300]

    messages: list[str] = []

    def walk(obj: Any):
        if isinstance(obj, dict):
            if "message" in obj and isinstance(obj["message"], str):
                messages.append(obj["message"])
            for v in obj.values():
                walk(v)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(data)
    if not messages:
        return resp.text[:300]
    # unik-in tapi jaga urutan
    seen = set()
    uniq = []
    for m in messages:
        if m not in seen:
            seen.add(m)
            uniq.append(m)
    return " | ".join(uniq)


def extract_id_from_href(href: Optional[str]) -> int:
    """Ambil angka ID di paling belakang URL, misal '/api/v3/projects/53' -> 53."""
    if not href:
        return 0
    m = re.search(r"/(\d+)/?$", href)
    return int(m.group(1)) if m else 0


def build_unique_identifier(raw_name: str, session_tag: str, counter: int) -> str:
    """
    Slugify nama project + tempel tag sesi + nomor urut, setara
    BuildUniqueIdentifier di VBA. Identifier OpenProject harus huruf
    kecil/angka/strip, mulai dengan huruf.
    """
    s = re.sub(r"[^a-z0-9]", "-", raw_name.strip().lower())
    s = re.sub(r"-+", "-", s).strip("-")
    if not s or not s[0].isalpha():
        s = f"sp-{s}"
    s = s[:40].rstrip("-")
    return f"{s}-{session_tag}-{counter}"
