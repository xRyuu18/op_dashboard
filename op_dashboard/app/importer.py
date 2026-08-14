"""
Port dari macro VBA "CreateProjectAndImportSubitemsSimple".

Alurnya PERSIS sama:
1. Bikin project utama.
2. Baca file task (xlsx/xls/csv), deteksi kolom: Subject, Type, Status,
   Assignee, Priority, Project (opsional), Start Date, Due Date, dan
   ID+Parent (opsional).
3. Kalau ada kolom Project -> bikin sub-project asli per nilai unik,
   nested di bawah project utama, lalu enable semua Type di situ.
4. Kalau ada kolom ID+Parent -> hierarki dari situ (2-pass: create dulu
   semua sesuai urutan file, baru pasang relasi parent belakangan).
   Kalau tidak ada -> hierarki dari indentasi 4-spasi di depan Subject.
5. Task-task masuk ke project utama atau sub-project sesuai kolom Project.

Bedanya dari VBA: semua ekstraksi ID pakai `response.json()` asli (Python
punya JSON parser bawaan), jadi nggak perlu trik "generate identifier
sendiri" atau "baca header Location" segala macam kayak di VBA -- di sini
`resp.json()["id"]` sudah pasti benar dan aman.
"""
from __future__ import annotations

import csv
import datetime as dt
import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import openpyxl

from .openproject_client import (
    OpenProjectClient,
    build_unique_identifier,
    extract_error_detail,
)

HEADER_ALIASES = {
    "id": "id",
    "parent": "parent",
    "subject": "subject",
    "type": "type",
    "status": "status",
    "assignee": "assignee",
    "priority": "priority",
    "start date": "start_date",
    "startdate": "start_date",
    "start": "start_date",
    "due date": "due_date",
    "duedate": "due_date",
    "due": "due_date",
    "project": "project",
    "sub project": "project",
    "subproject": "project",
    "sub-project": "project",
}


@dataclass
class ImportResult:
    project_name: str = ""
    project_id: int = 0
    task_success: int = 0
    task_fail: int = 0
    fail_log: list[str] = field(default_factory=list)
    subproject_log: list[str] = field(default_factory=list)
    subproject_map: dict[str, str] = field(default_factory=dict)  # nama -> id/identifier
    per_project_count: dict[str, int] = field(default_factory=dict)


def _read_rows(file_path: str) -> tuple[list[str], list[list[Any]]]:
    """Baca file xlsx/xls/csv, kembalikan (header, rows) apa adanya."""
    ext = Path(file_path).suffix.lower()
    if ext == ".csv":
        with open(file_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]

    if not rows:
        return [], []
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    data_rows = rows[1:]
    return header, data_rows


def _map_columns(header: list[str]) -> dict[str, int]:
    """Cocokkan header ke nama field kanonik (case-insensitive), setara
    Select Case di VBA."""
    col: dict[str, int] = {}
    for idx, h in enumerate(header):
        key = HEADER_ALIASES.get(h.strip().lower())
        if key:
            col[key] = idx
    return col


def _cell(row: list[Any], col: dict[str, int], key: str) -> str:
    idx = col.get(key)
    if idx is None or idx >= len(row):
        return ""
    v = row[idx]
    if v is None:
        return ""
    return str(v).strip()


def _format_date(row: list[Any], col: dict[str, int], key: str) -> str:
    idx = col.get(key)
    if idx is None or idx >= len(row):
        return ""
    v = row[idx]
    if v is None or v == "":
        return ""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    # coba parse teks tanggal umum
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def _extract_id_from_parent_text(parent_text: str) -> str:
    """Setara ExtractIDFromParentLocal: 'Task #99: Nama' -> '99'."""
    if "#" not in parent_text:
        return ""
    after_hash = parent_text.split("#", 1)[1]
    num = ""
    for ch in after_hash:
        if ch.isdigit():
            num += ch
        else:
            break
    return num


def create_main_project(client: OpenProjectClient, name: str) -> int:
    resp = client.post("/api/v3/projects", {"name": name})
    if resp.status_code != 201:
        raise RuntimeError(f"Gagal membuat project. Status {resp.status_code}: {extract_error_detail(resp)}")
    return resp.json()["id"]


def create_subproject(client: OpenProjectClient, parent_id: int, name: str,
                       forced_identifier: Optional[str] = None) -> tuple[Optional[str], str]:
    """Return (identifier_atau_id_string, error_message)."""
    body: dict[str, Any] = {
        "name": name,
        "_links": {"parent": {"href": f"/api/v3/projects/{parent_id}"}},
    }
    if forced_identifier:
        body["identifier"] = forced_identifier

    resp = client.post("/api/v3/projects", body)
    if resp.status_code != 201:
        return None, f"Status {resp.status_code}: {extract_error_detail(resp)}"
    return str(resp.json()["id"]), ""


def enable_project_types(client: OpenProjectClient, project_ref: str, type_map: dict[str, int]) -> str:
    if not type_map:
        return ""
    links = [{"href": f"/api/v3/types/{tid}"} for tid in type_map.values()]
    resp = client.patch(f"/api/v3/projects/{project_ref}", {"_links": {"types": links}})
    if resp.status_code != 200:
        return f"Status {resp.status_code}: {extract_error_detail(resp)}"
    return ""


def fetch_name_id_map(client: OpenProjectClient, endpoint: str) -> dict[str, int]:
    resp = client.get(endpoint, params={"pageSize": 200})
    result: dict[str, int] = {}
    if resp.status_code != 200:
        return result
    for el in resp.json().get("_embedded", {}).get("elements", []):
        name = el.get("name")
        if name:
            result[name.strip().lower()] = el["id"]
    return result


def set_work_package_parent(client: OpenProjectClient, wp_id: int, parent_id: int) -> str:
    get_resp = client.get(f"/api/v3/work_packages/{wp_id}")
    if get_resp.status_code != 200:
        return f"GET gagal (status {get_resp.status_code}): {get_resp.text[:150]}"
    lock_version = get_resp.json().get("lockVersion", 0)

    patch_resp = client.patch(
        f"/api/v3/work_packages/{wp_id}",
        {"lockVersion": lock_version, "_links": {"parent": {"href": f"/api/v3/work_packages/{parent_id}"}}},
    )
    if patch_resp.status_code != 200:
        return f"PATCH gagal (status {patch_resp.status_code}): {extract_error_detail(patch_resp)}"
    return ""


def create_task(client: OpenProjectClient, project_ref: str, subject: str,
                 row: list[Any], col: dict[str, int],
                 type_map: dict[str, int], status_map: dict[str, int],
                 priority_map: dict[str, int], assignee_map: dict[str, int]) -> tuple[int, str]:
    """Return (new_wp_id_atau_0, error_message)."""
    type_name = _cell(row, col, "type")
    type_id = type_map.get(type_name.lower())
    if not type_id:
        # FIX yang sama kayak VBA: jangan biarin type kosong (bisa kena
        # default type project yang salah, misal Milestone).
        type_id = type_map.get("task")

    status_id = status_map.get(_cell(row, col, "status").lower())
    priority_id = priority_map.get(_cell(row, col, "priority").lower())
    assignee_name = _cell(row, col, "assignee")
    assignee_id = None
    if assignee_name and assignee_name != "-":
        assignee_id = assignee_map.get(assignee_name.lower())

    start_date = _format_date(row, col, "start_date")
    due_date = _format_date(row, col, "due_date")

    body: dict[str, Any] = {"subject": subject}

    is_milestone = type_name.strip().lower() == "milestone"
    if is_milestone:
        milestone_date = due_date or start_date
        if milestone_date:
            body["date"] = milestone_date
    else:
        if start_date:
            body["startDate"] = start_date
        if due_date:
            body["dueDate"] = due_date

    links: dict[str, Any] = {}
    if type_id:
        links["type"] = {"href": f"/api/v3/types/{type_id}"}
    if status_id:
        links["status"] = {"href": f"/api/v3/statuses/{status_id}"}
    if priority_id:
        links["priority"] = {"href": f"/api/v3/priorities/{priority_id}"}
    if assignee_id:
        links["assignee"] = {"href": f"/api/v3/users/{assignee_id}"}
    if links:
        body["_links"] = links

    resp = client.post(f"/api/v3/projects/{project_ref}/work_packages", body)
    if resp.status_code != 201:
        return 0, extract_error_detail(resp)
    return resp.json()["id"], ""


def get_target_project_ref(row: list[Any], col: dict[str, int], main_project_id: int,
                            subproject_map: dict[str, str]) -> str:
    if "project" not in col:
        return str(main_project_id)
    p_name = _cell(row, col, "project")
    if not p_name:
        return str(main_project_id)
    ref = subproject_map.get(p_name.lower())
    return ref if ref else str(main_project_id)


def run_import(base_url: str, api_key: str, project_name: str, file_path: str) -> ImportResult:
    client = OpenProjectClient(base_url, api_key)
    result = ImportResult(project_name=project_name)

    result.project_id = create_main_project(client, project_name)

    header, rows = _read_rows(file_path)
    col = _map_columns(header)

    if "subject" not in col:
        raise RuntimeError("Kolom 'Subject' tidak ditemukan di file task. Project sudah terlanjur dibuat "
                            f"(ID {result.project_id}).")

    type_map = fetch_name_id_map(client, "/api/v3/types")
    status_map = fetch_name_id_map(client, "/api/v3/statuses")
    priority_map = fetch_name_id_map(client, "/api/v3/priorities")
    assignee_map = fetch_name_id_map(client, "/api/v3/users")

    # ==== Bikin sub-project dari kolom Project (kalau ada) ====
    subproject_map: dict[str, str] = {}
    if "project" in col:
        session_tag = dt.datetime.now().strftime("%y%m%d%H%M%S")
        counter = 0
        seen_names: dict[str, str] = {}
        for row in rows:
            p_name = _cell(row, col, "project")
            if p_name and p_name.lower() not in seen_names:
                counter += 1
                identifier = build_unique_identifier(p_name, session_tag, counter)
                ref, err = create_subproject(client, result.project_id, p_name, identifier)
                if ref:
                    seen_names[p_name.lower()] = ref
                    enable_err = enable_project_types(client, ref, type_map)
                    if enable_err:
                        result.subproject_log.append(
                            f"Sub-project '{p_name}' dibuat tapi GAGAL enable types: {enable_err}"
                        )
                else:
                    seen_names[p_name.lower()] = ""
                    result.subproject_log.append(f"Sub-project '{p_name}': {err}")
        subproject_map = seen_names
        result.subproject_map = subproject_map

    use_id_parent_mode = "id" in col and "parent" in col

    if use_id_parent_mode:
        id_clean = [_cell(row, col, "id") for row in rows]
        parent_id_clean = [_extract_id_from_parent_text(_cell(row, col, "parent")) for row in rows]
        parent_row_idx = [None] * len(rows)
        for i, pid in enumerate(parent_id_clean):
            if pid:
                for j, own_id in enumerate(id_clean):
                    if own_id == pid:
                        parent_row_idx[i] = j
                        break

        new_wp_id: list[Optional[int]] = [None] * len(rows)

        # PASS 1: create sesuai urutan file, tanpa parent dulu
        for i, row in enumerate(rows):
            subject = _cell(row, col, "subject")
            if not subject:
                continue
            target_ref = get_target_project_ref(row, col, result.project_id, subproject_map)
            new_id, err = create_task(client, target_ref, subject, row, col,
                                       type_map, status_map, priority_map, assignee_map)
            new_wp_id[i] = new_id or None
            if new_id:
                result.task_success += 1
                result.per_project_count[target_ref] = result.per_project_count.get(target_ref, 0) + 1
            else:
                result.task_fail += 1
                result.fail_log.append(f"[{subject}] {err}")

        # PASS 2: pasang parent belakangan
        for i, row in enumerate(rows):
            if new_wp_id[i] and parent_row_idx[i] is not None:
                parent_new_id = new_wp_id[parent_row_idx[i]]
                if parent_new_id:
                    err = set_work_package_parent(client, new_wp_id[i], parent_new_id)
                    if err:
                        subject = _cell(row, col, "subject")
                        result.fail_log.append(f"[{subject}] dibuat OK (ID {new_wp_id[i]}) tapi GAGAL set parent: {err}")

    else:
        # Mode indentasi
        last_id_at_level: dict[int, int] = {}
        for row in rows:
            raw_subject = _cell(row, col, "subject")
            if not raw_subject:
                continue
            leading_spaces = len(raw_subject) - len(raw_subject.lstrip(" "))
            level = min(leading_spaces // 4, 50)
            subject = raw_subject.strip()

            parent_new_id = last_id_at_level.get(level - 1) if level > 0 else None
            target_ref = get_target_project_ref(row, col, result.project_id, subproject_map)

            new_id, err = create_task(client, target_ref, subject, row, col,
                                       type_map, status_map, priority_map, assignee_map)
            if new_id:
                result.task_success += 1
                last_id_at_level[level] = new_id
                result.per_project_count[target_ref] = result.per_project_count.get(target_ref, 0) + 1
                if parent_new_id:
                    perr = set_work_package_parent(client, new_id, parent_new_id)
                    if perr:
                        result.fail_log.append(f"[{subject}] dibuat OK (ID {new_id}) tapi GAGAL set parent: {perr}")
            else:
                result.task_fail += 1
                result.fail_log.append(f"[{subject}] {err}")

    return result
