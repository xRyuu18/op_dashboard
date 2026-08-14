"""
Port dari macro VBA "ExportProjectToExcel".

Cari project berdasarkan NAMA (harus persis sama, biar nggak ketuker
kalau ada nama mirip), cari semua sub-project di bawahnya (rekursif),
tarik semua work package dari project utama + sub-project, susun
hierarki, tulis ke Excel dengan format kolom sama persis kayak yang
dipakai importer:
    Subject | Type | Status | Assignee | Priority | Project | Start Date | Due Date
Kolom Project dikosongin untuk task di project utama, diisi nama
sub-project untuk task yang ada di sub-project.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

import openpyxl

from .openproject_client import OpenProjectClient, extract_error_detail, extract_id_from_href


class AmbiguousProjectName(Exception):
    def __init__(self, matches: list[dict]):
        self.matches = matches
        names = "\n".join(f"- {m['name']} (ID {m['id']})" for m in matches)
        super().__init__(f"Ada {len(matches)} project dengan nama persis sama:\n{names}")


class ProjectNotFound(Exception):
    pass


@dataclass
class ExportResult:
    project_name: str = ""
    project_id: int = 0
    subproject_count: int = 0
    task_count: int = 0
    fetch_errors: list[str] = field(default_factory=list)
    file_path: str = ""


def find_project_by_name(client: OpenProjectClient, search_name: str) -> dict:
    """Setara FindProjectByName di VBA: harus PERSIS sama (case-insensitive)."""
    filters = f'[{{"name":{{"operator":"=","values":["{search_name}"]}}}}]'
    resp = client.get("/api/v3/projects", params={"filters": filters, "pageSize": 200})
    if resp.status_code != 200:
        raise RuntimeError(f"Gagal mencari project '{search_name}'. Status {resp.status_code}: "
                            f"{extract_error_detail(resp)}")

    elements = resp.json().get("_embedded", {}).get("elements", [])
    exact = [e for e in elements if e.get("name", "").strip().lower() == search_name.strip().lower()]

    if not exact:
        raise ProjectNotFound(
            f"Project dengan nama persis '{search_name}' tidak ditemukan. "
            "Cek lagi penulisan namanya (harus sama persis, termasuk spasi)."
        )
    if len(exact) > 1:
        raise AmbiguousProjectName(exact)
    return exact[0]


def collect_descendant_projects(client: OpenProjectClient, parent_id: int,
                                 result: list[dict], safety_counter: list[int]) -> None:
    """Rekursif cari SEMUA sub-project di bawah parent_id, berapa level pun."""
    if safety_counter[0] > 500:
        return
    filters = f'[{{"parent":{{"operator":"=","values":["{parent_id}"]}}}}]'
    resp = client.get("/api/v3/projects", params={"filters": filters, "pageSize": 200})
    if resp.status_code != 200:
        return
    for el in resp.json().get("_embedded", {}).get("elements", []):
        safety_counter[0] += 1
        child_id = el.get("id")
        child_name = el.get("name", "")
        if child_id:
            result.append({"id": child_id, "name": child_name})
            collect_descendant_projects(client, child_id, result, safety_counter)


def fetch_work_packages_for_project(client: OpenProjectClient, project_id: int) -> tuple[list[dict], str]:
    """
    Ambil SEMUA work package sebuah project (pagination), lewat endpoint
    global /api/v3/work_packages + filter project=ID (bukan endpoint
    /projects/{id}/work_packages yang berpotensi ikut narik sub-project
    dan bikin dobel-hitung, karena kita sudah loop tiap project sendiri).
    Filter custom juga otomatis narik SEMUA status, bukan cuma yang open.
    """
    result: list[dict] = []
    page = 1
    page_size = 100
    total = None
    filters = f'[{{"project":{{"operator":"=","values":["{project_id}"]}}}}]'

    while True:
        resp = client.get("/api/v3/work_packages",
                           params={"filters": filters, "pageSize": page_size, "offset": page})
        if resp.status_code != 200:
            return result, f"Status {resp.status_code}: {extract_error_detail(resp)}"

        data = resp.json()
        if total is None:
            total = data.get("total", 0)

        elements = data.get("_embedded", {}).get("elements", [])
        if not elements:
            break
        result.extend(elements)

        page += 1
        if len(result) >= total or page > 1000:
            break

    return result, ""


def _link_title(wp: dict, key: str) -> str:
    link = (wp.get("_links") or {}).get(key)
    if not isinstance(link, dict):
        return ""
    return link.get("title") or ""


def build_workbook(all_wps: list[dict], main_project_name: str) -> openpyxl.Workbook:
    id_to_index = {wp["id"]: i for i, wp in enumerate(all_wps)}
    children_of: dict[int, list[int]] = {}
    root_indices: list[int] = []

    for i, wp in enumerate(all_wps):
        parent_href = ((wp.get("_links") or {}).get("parent") or {}).get("href")
        parent_id = extract_id_from_href(parent_href)
        if parent_id and parent_id in id_to_index:
            children_of.setdefault(parent_id, []).append(i)
        else:
            root_indices.append(i)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"
    ws.append(["Subject", "Type", "Status", "Assignee", "Priority", "Project", "Start Date", "Due Date"])

    def write_row(idx: int, depth: int):
        wp = all_wps[idx]
        subject = wp.get("subject") or ""
        type_title = _link_title(wp, "type")
        status_title = _link_title(wp, "status")
        assignee_title = _link_title(wp, "assignee")
        priority_title = _link_title(wp, "priority")

        start_date = wp.get("startDate") or ""
        due_date = wp.get("dueDate") or ""
        date_field = wp.get("date") or ""
        if type_title.strip().lower() == "milestone" and date_field:
            due_date = date_field

        task_project_name = _link_title(wp, "project")
        project_col = "" if (not task_project_name or
                              task_project_name.strip().lower() == main_project_name.strip().lower()) \
            else task_project_name

        ws.append([
            ("    " * depth) + subject,
            type_title, status_title, assignee_title, priority_title,
            project_col, start_date, due_date,
        ])

        for child_idx in children_of.get(wp["id"], []):
            write_row(child_idx, depth + 1)

    for idx in root_indices:
        write_row(idx, 0)

    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

    return wb


def run_export(base_url: str, api_key: str, project_name: str, output_path: str) -> ExportResult:
    client = OpenProjectClient(base_url, api_key)
    result = ExportResult(project_name=project_name)

    main = find_project_by_name(client, project_name)
    result.project_id = main["id"]
    result.project_name = main["name"]

    descendants: list[dict] = []
    collect_descendant_projects(client, result.project_id, descendants, [0])
    result.subproject_count = len(descendants)

    all_projects = [{"id": result.project_id, "name": result.project_name}] + descendants

    all_wps: list[dict] = []
    seen_ids: set[int] = set()
    for p in all_projects:
        wps, err = fetch_work_packages_for_project(client, p["id"])
        if err:
            result.fetch_errors.append(f"Project '{p['name']}': {err}")
            continue
        for wp in wps:
            if wp["id"] not in seen_ids:
                seen_ids.add(wp["id"])
                all_wps.append(wp)

    result.task_count = len(all_wps)

    if all_wps:
        wb = build_workbook(all_wps, result.project_name)
        wb.save(output_path)
        result.file_path = output_path

    return result
