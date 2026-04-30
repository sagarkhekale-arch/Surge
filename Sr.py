import json
import re
import time
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Surge Runner", page_icon="⚡", layout="wide")

ENDPOINT_URL = "https://hobbit-app.shadowfax.in/central_allocation/upload-cluster-surge-mode-details/"
IST = timezone(timedelta(hours=5, minutes=30))
ALL_CLUSTER_DETAILS = [{"cluster_name": "Select All Clusters", "id": -1}]
_DEFAULT_MASTER_XLSX = str(Path(__file__).resolve().parent / "surge_master.xlsx")

CITY_ALIAS = {
    "delhi": "del",
    "new delhi": "del",
    "gurgaon": "ggn",
    "gurugram": "ggn",
    "bangalore": "blr",
    "bengaluru": "blr",
    "mumbai": "bom",
    "chennai": "maa",
    "hyderabad": "hyd",
    "kolkata": "ccu",
    "pune": "pnq",
    "ahmedabad": "amd",
}


def read_csv_upload(file_obj: Any) -> Tuple[pd.DataFrame, str]:
    try:
        file_obj.seek(0)
        df = pd.read_csv(file_obj)
    except Exception as exc:
        return pd.DataFrame(), f"CSV read error: {exc}"
    if df.empty and len(df.columns) == 0:
        return pd.DataFrame(), "CSV has no columns."
    return df, ""


def parse_time(value: Any) -> str:
    raw = str(value).strip()
    fmts = (
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
    )
    dt = None
    for fmt in fmts:
        try:
            dt = datetime.strptime(raw, fmt)
            break
        except ValueError:
            pass
    if dt is None:
        parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)
        if pd.isna(parsed):
            raise ValueError(f"Invalid start_time: {raw}")
        dt = parsed.to_pydatetime()
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    return dt.isoformat(timespec="seconds")


LOG_SHEET_NAME = "Log"


def run_timestamp_ist() -> str:
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")


def sheet_date_tab_from_row(csv_row: Dict[str, Any]) -> str:
    """Excel tab name YYYY-MM-DD from CSV start_time, else today (IST)."""
    raw = csv_row.get("start_time")
    if raw is None or str(raw).strip() == "":
        return datetime.now(IST).strftime("%Y-%m-%d")
    try:
        return parse_time(raw)[:10]
    except Exception:
        return datetime.now(IST).strftime("%Y-%m-%d")


def sanitize_excel_sheet_name(name: str) -> str:
    invalid = "[]:*?/\\"
    s = "".join(c if c not in invalid else "-" for c in str(name).strip())[:31]
    return s or "Data"


def load_workbook_all_sheets(path: Path) -> Dict[str, pd.DataFrame]:
    if not path.exists():
        return {}
    xl = pd.ExcelFile(path)
    return {sn: pd.read_excel(path, sheet_name=sn) for sn in xl.sheet_names}


def merge_master_run_to_excel(
    path: Path,
    log_df: pd.DataFrame,
    inputs_by_date: Dict[str, pd.DataFrame],
) -> Tuple[bool, str]:
    """Append logs to sheet Log. Append inputs to one sheet per date; skip duplicate _dedupe_key."""
    try:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        sheets = load_workbook_all_sheets(path) if path.exists() else {}

        old_log = sheets.get(LOG_SHEET_NAME, pd.DataFrame())
        if log_df is not None and not log_df.empty:
            sheets[LOG_SHEET_NAME] = pd.concat([old_log, log_df], ignore_index=True)
        elif LOG_SHEET_NAME not in sheets:
            sheets[LOG_SHEET_NAME] = pd.DataFrame()

        for date_key, new_part in (inputs_by_date or {}).items():
            if new_part is None or new_part.empty:
                continue
            sh = sanitize_excel_sheet_name(date_key)
            old = sheets.get(sh, pd.DataFrame())
            if "_dedupe_key" not in old.columns and not old.empty:
                old = old.copy()
                old["_dedupe_key"] = ""
            if "_dedupe_key" not in new_part.columns:
                new_part = new_part.copy()
                new_part["_dedupe_key"] = ""
            comb = pd.concat([old, new_part], ignore_index=True)
            comb = comb.drop_duplicates(subset=["_dedupe_key"], keep="first")
            if "_dedupe_key" in comb.columns:
                cols = [c for c in comb.columns if c != "_dedupe_key"] + ["_dedupe_key"]
                comb = comb[cols]
            sheets[sh] = comb

        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
            for sn, df in sheets.items():
                sheet_name = sanitize_excel_sheet_name(sn)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        return True, ""
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"


def _flatten_csv_row_for_excel(csv_row: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for k, v in csv_row.items():
        key = str(k).strip() or "column"
        if v is None or (isinstance(v, float) and pd.isna(v)):
            out[key] = ""
        else:
            out[key] = v
    return out


def build_input_dedupe_key(
    flat: Dict[str, Any],
    city_id_val: str,
    clusters_sent: str,
) -> str:
    parts = [
        str(city_id_val).strip(),
        str(flat.get("start_time", "")).strip(),
        str(flat.get("duration", "")).strip(),
        str(flat.get("amount", "")).strip(),
        clusters_sent.strip(),
        str(flat.get("chain_ids", "")).strip(),
        str(flat.get("cluster_ids", "")).strip(),
        str(flat.get("cluster_names", "")).strip(),
        str(flat.get("surge_type", "")).strip(),
        str(flat.get("vehicle_type", "")).strip(),
    ]
    return "|".join(parts)


def build_input_row_for_master(
    csv_row: Dict[str, Any],
    city_key: str,
    city_id_raw: Any,
    cluster_ids: List[int] | None,
    payload: Dict[str, Any] | None,
    success: bool,
    status: int,
    msg: str,
    cluster_note: str,
) -> Dict[str, Any]:
    flat = _flatten_csv_row_for_excel(csv_row)
    clusters_sent = ",".join(str(x) for x in cluster_ids) if cluster_ids else ""
    pst = (payload or {}).get("start_time", "")
    rec: Dict[str, Any] = {**flat}
    rec["resolved_city_key"] = city_key
    cid_str = ""
    if city_id_raw is not None and str(city_id_raw).strip() != "":
        cid_str = str(city_id_raw).strip()
    rec["city_id_sent"] = cid_str or str(flat.get("city", "")).strip()
    rec["clusters_sent"] = clusters_sent
    rec["api_start_time_iso"] = pst
    rec["success"] = success
    rec["http_status"] = status
    rec["response_snippet"] = (msg or "")[:2000]
    rec["cluster_note"] = (cluster_note or "")[:500]
    rec["logged_at_ist"] = run_timestamp_ist()
    rec["_dedupe_key"] = build_input_dedupe_key(flat, rec["city_id_sent"], clusters_sent)
    return rec


def parse_int_list(value: Any) -> List[int]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    if text.startswith("["):
        try:
            return [int(x) for x in json.loads(text)]
        except Exception:
            return []
    out = []
    for p in text.replace(",", "|").split("|"):
        p = p.strip()
        if not p:
            continue
        try:
            out.append(int(p))
        except ValueError:
            pass
    return out


def parse_chains(row: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw_json = str(row.get("selected_chains_json", "")).strip()
    if raw_json:
        try:
            parsed = json.loads(raw_json)
            if isinstance(parsed, list):
                return parsed
        except json.JSONDecodeError:
            pass
    ids = parse_int_list(row.get("chain_ids"))
    names = [x.strip() for x in str(row.get("chain_names", "")).replace(",", "|").split("|") if x.strip()]
    return [{"id": cid, "chain_name": names[i] if i < len(names) else f"Chain {cid}"} for i, cid in enumerate(ids)]


def collect_cluster_ids(obj: Any, out: Set[int]) -> None:
    def to_int(value: Any) -> int | None:
        try:
            if isinstance(value, str):
                value = value.strip()
                if not value:
                    return None
                if "." in value:
                    return int(float(value))
            return int(value)
        except Exception:
            return None

    if isinstance(obj, list):
        for item in obj:
            collect_cluster_ids(item, out)
        return
    if isinstance(obj, dict):
        cluster_hints = (
            "cluster",
            "cluster_name",
            "cluster_id",
            "clusterId",
            "cluster_identifier",
            "latitude",
            "longitude",
        )
        candidate_keys = ("id", "cluster_id", "clusterId", "clusterID")

        if any(k in obj for k in cluster_hints):
            for ck in candidate_keys:
                if ck in obj:
                    cid = to_int(obj.get(ck))
                    if cid is not None and cid > 0:
                        out.add(cid)

        # common response shape: {"data": [{"id": ...}, ...]}
        if "data" in obj and isinstance(obj["data"], list):
            for item in obj["data"]:
                if isinstance(item, dict):
                    for ck in candidate_keys:
                        if ck in item:
                            cid = to_int(item.get(ck))
                            if cid is not None and cid > 0:
                                out.add(cid)

        for v in obj.values():
            collect_cluster_ids(v, out)


def collect_ids_from_data_section(obj: Any, out: Set[int]) -> None:
    """Loose extractor specifically for top-level 'data' payloads."""
    def to_int(value: Any) -> int | None:
        try:
            if isinstance(value, str):
                value = value.strip()
                if not value:
                    return None
                if "." in value:
                    return int(float(value))
            return int(value)
        except Exception:
            return None

    if isinstance(obj, list):
        for item in obj:
            collect_ids_from_data_section(item, out)
        return

    if isinstance(obj, dict):
        for key in ("id", "cluster_id", "clusterId", "clusterID"):
            if key in obj:
                cid = to_int(obj.get(key))
                if cid is not None and cid > 0:
                    out.add(cid)
        for v in obj.values():
            collect_ids_from_data_section(v, out)


def collect_cluster_name_map(obj: Any, out: Dict[str, int]) -> None:
    def normalize_name(name: str) -> str:
        s = name.strip().lower()
        s = s.replace("&", " and ")
        s = re.sub(r"[^a-z0-9]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    if isinstance(obj, list):
        for item in obj:
            collect_cluster_name_map(item, out)
        return
    if isinstance(obj, dict):
        raw_name = obj.get("cluster_name") or obj.get("name")
        raw_id = obj.get("id") if "id" in obj else obj.get("cluster_id")
        if raw_name is not None and raw_id is not None:
            try:
                cid = int(float(str(raw_id).strip()))
                name_key = normalize_name(str(raw_name))
                if name_key and cid > 0:
                    out[name_key] = cid
            except Exception:
                pass
        for v in obj.values():
            collect_cluster_name_map(v, out)


def fetch_city_clusters(
    headers: Dict[str, str],
    city_id: int | None,
    city_name_api: str,
    timeout: int,
) -> Tuple[List[int], Dict[str, int], str]:
    param_candidates: List[Dict[str, Any]] = []
    if city_id is not None and city_id > 0:
        param_candidates.append({"city_id": city_id})
    if city_name_api:
        param_candidates.append({"city_name": city_name_api})

    if not param_candidates:
        return [], {}, "Both city_id and city_name_api are missing."

    last_error = ""
    for params in param_candidates:
        resp = None
        for attempt in range(1, 4):
            try:
                resp = requests.get(ENDPOINT_URL, headers=headers, params=params, timeout=timeout)
            except requests.RequestException as exc:
                last_error = str(exc)
                if attempt < 3:
                    time.sleep(0.8 * attempt)
                continue

            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                wait_seconds = 1.0
                if retry_after:
                    try:
                        wait_seconds = max(1.0, float(retry_after))
                    except ValueError:
                        wait_seconds = 1.5
                if attempt < 3:
                    time.sleep(wait_seconds)
                    continue
                last_error = f"HTTP 429 for params={params}"
            break

        if resp is None:
            continue
        if not (200 <= resp.status_code < 300):
            last_error = f"HTTP {resp.status_code} for params={params}"
            continue
        try:
            data = resp.json()
        except ValueError:
            last_error = f"Non-JSON response for params={params}"
            continue

        ids: Set[int] = set()
        collect_cluster_ids(data, ids)
        cluster_name_map: Dict[str, int] = {}
        collect_cluster_name_map(data, cluster_name_map)

        if isinstance(data, dict) and "data" in data:
            collect_ids_from_data_section(data.get("data"), ids)
            collect_cluster_name_map(data.get("data"), cluster_name_map)

        vals = sorted([x for x in ids if x > 0])
        if vals:
            return vals, cluster_name_map, ""

        if isinstance(data, dict):
            preview = f"top-level keys: {list(data.keys())[:8]}"
        elif isinstance(data, list):
            preview = f"top-level list length: {len(data)}"
        else:
            preview = f"top-level type: {type(data).__name__}"
        last_error = f"No cluster IDs found ({preview}) for params={params}"

    return [], {}, last_error or "No cluster IDs found for provided city inputs."


def choose_cluster_ids(row: Dict[str, Any], all_city_cluster_ids: List[int], cluster_name_map: Dict[str, int]) -> Tuple[List[int], str]:
    # Highest priority: explicit IDs from CSV.
    direct_ids = parse_int_list(row.get("cluster_ids"))
    if direct_ids:
        unique_direct_ids = list(dict.fromkeys([x for x in direct_ids if x > 0]))
        if unique_direct_ids:
            return unique_direct_ids, ""

    def normalize_name(name: str) -> str:
        s = name.strip().lower()
        s = s.replace("&", " and ")
        s = re.sub(r"[^a-z0-9]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    raw = str(row.get("cluster_names", "")).strip()
    if not raw:
        return all_city_cluster_ids, ""

    wanted = [normalize_name(x) for x in raw.replace("|", ",").split(",") if x.strip()]
    selected: List[int] = []
    missing: List[str] = []
    for name in wanted:
        cid = cluster_name_map.get(name)
        if cid is None:
            # fuzzy fallback for minor naming differences
            for api_name, api_id in cluster_name_map.items():
                if name in api_name or api_name in name:
                    cid = api_id
                    break
        if cid is None:
            missing.append(name)
        else:
            selected.append(cid)

    # de-duplicate while preserving order
    unique_selected = list(dict.fromkeys(selected))
    if unique_selected:
        note = ""
        if missing:
            note = f"Some cluster names not found: {', '.join(missing[:5])}"
        return unique_selected, note
    available = sorted(cluster_name_map.keys())
    sample_available = ", ".join(available[:8]) if available else "none"
    return [], (
        f"No provided cluster_names matched for city. Missing: {', '.join(missing[:5])}. "
        f"Available examples: {sample_available}"
    )


def resolve_city_key(row: Dict[str, Any]) -> str:
    direct = str(row.get("city_name_api", "")).strip().lower()
    if direct:
        return direct

    name = str(row.get("city_name", "")).strip().lower()
    if name:
        if name in CITY_ALIAS:
            return CITY_ALIAS[name]
        if len(name) <= 4:
            return name
        return name[:3]
    return ""


def build_payload(row: Dict[str, Any], cluster_ids: List[int]) -> Dict[str, Any]:
    if not cluster_ids:
        raise ValueError("No cluster IDs for city.")
    chains = parse_chains(row)
    chain_ids = [int(c["id"]) for c in chains if "id" in c]
    return {
        "activate_for_marketplace_chains": int(row.get("activate_for_marketplace_chains", 0)),
        "is_active": str(row.get("is_active", "true")).lower() != "false",
        "duration": int(float(row["duration"])),
        "amount": float(row["amount"]),
        "start_time": parse_time(row["start_time"]),
        "city": int(float(row["city"])),
        "cluster_details": ALL_CLUSTER_DETAILS,
        "selected_chains": chains,
        "surge_type": int(float(row.get("surge_type", 2))),
        "vehicle_type": int(float(row.get("vehicle_type", 100))),
        "chains": json.dumps(chain_ids),
        "cluster": cluster_ids,
    }


def post_payload(headers: Dict[str, str], payload: Dict[str, Any], timeout: int) -> Tuple[bool, int, str]:
    try:
        r = requests.post(ENDPOINT_URL, headers=headers, json=payload, timeout=timeout)
        return 200 <= r.status_code < 300, r.status_code, r.text[:800]
    except requests.RequestException as exc:
        return False, -1, str(exc)


def request_json_method(
    url: str,
    method: str,
    headers: Dict[str, str],
    payload: Dict[str, Any],
    timeout: int,
) -> Tuple[bool, int, str]:
    m = method.strip().upper()
    try:
        if m == "POST":
            r = requests.post(url, headers=headers, json=payload, timeout=timeout)
        elif m == "PUT":
            r = requests.put(url, headers=headers, json=payload, timeout=timeout)
        elif m == "PATCH":
            r = requests.patch(url, headers=headers, json=payload, timeout=timeout)
        else:
            return False, -1, f"Unsupported method: {method}"
        return 200 <= r.status_code < 300, r.status_code, r.text[:800]
    except requests.RequestException as exc:
        return False, -1, str(exc)


def build_deactivate_payload(template_str: str, surge_id: int) -> Dict[str, Any]:
    s = template_str.strip().replace("{{surge_id}}", str(int(surge_id)))
    parsed = json.loads(s)
    if not isinstance(parsed, dict):
        raise ValueError("Deactivate body template must be a JSON object.")
    return parsed


def expand_deactivate_url(url_template: str, surge_id: int) -> str:
    return url_template.strip().replace("{{surge_id}}", str(int(surge_id)))


def extract_csrf_from_html(html: str) -> str | None:
    patterns = [
        r'name=["\']csrfmiddlewaretoken["\']\s+value=["\']([^"\']+)["\']',
        r'value=["\']([^"\']+)["\']\s+name=["\']csrfmiddlewaretoken["\']',
        r'name=csrfmiddlewaretoken\s+value=["\']([^"\']+)["\']',
    ]
    for pat in patterns:
        m = re.search(pat, html, re.IGNORECASE)
        if m:
            return m.group(1)
    return None


def apply_deactivate_template(template_str: str, surge_id: int, csrf_token: str | None = None) -> str:
    s = template_str.replace("{{surge_id}}", str(int(surge_id)))
    if csrf_token is not None:
        s = s.replace("{{csrf_token}}", csrf_token)
    return s


def _html_attr(attrs: str, key: str) -> str | None:
    m = re.search(rf"{re.escape(key)}\s*=\s*[\"']([^\"']*)[\"']", attrs, re.I)
    if m:
        return m.group(1)
    m = re.search(rf"{re.escape(key)}\s*=\s*([^\s>]+)", attrs, re.I)
    return m.group(1) if m else None


def _find_main_change_form(html: str) -> str | None:
    """Largest POST form that contains Django admin CSRF (the model change form)."""
    lower = html.lower()
    idx = 0
    best: str | None = None
    best_len = 0
    while True:
        start = lower.find("<form", idx)
        if start == -1:
            break
        end = lower.find("</form>", start)
        if end == -1:
            break
        chunk = html[start : end + 7]
        lchunk = chunk.lower()
        if "csrfmiddlewaretoken" not in lchunk:
            idx = start + 5
            continue
        if re.search(r"method\s*=\s*[\"']get[\"']", lchunk, re.I):
            idx = start + 5
            continue
        if not re.search(r"method\s*=\s*[\"']post[\"']", lchunk, re.I):
            idx = start + 5
            continue
        if len(chunk) > best_len:
            best_len = len(chunk)
            best = chunk
        idx = start + 5
    return best


def _inactive_value_to_set(form_chunk: str, field_name: str) -> str | None:
    """Pick a value that means inactive/off (Django admin Yes/No or true/false select)."""
    esc = re.escape(field_name)
    m = re.search(
        rf"<select[^>]*name=[\"']{esc}[\"'][^>]*>(.*?)</select>",
        form_chunk,
        re.I | re.DOTALL,
    )
    if m:
        inner = m.group(1)
        candidates: List[Tuple[str, str]] = []
        for om in re.finditer(r"<option\s([^>]+)>([^<]*)", inner, re.I):
            val = _html_attr(om.group(1), "value")
            if val is None:
                val = ""
            label = (om.group(2) or "").strip().lower()
            vl = val.lower()
            candidates.append((val, label))
        for val, label in candidates:
            lab = label.lower()
            vl = val.lower()
            if lab in ("no", "false", "inactive", "off") or vl in ("false", "0", "no", "f", "off"):
                return val
        if len(candidates) >= 2:
            return candidates[-1][0]
        return candidates[0][0] if candidates else None

    for rm in re.finditer(r"<input\s([^>]+)>", form_chunk, re.I):
        at = rm.group(1)
        if (_html_attr(at, "type") or "").lower() != "radio":
            continue
        if _html_attr(at, "name") != field_name:
            continue
        v = _html_attr(at, "value")
        if v is not None and v.lower() in ("0", "false", "no", "f", "off"):
            return v
    vals: List[str] = []
    for rm in re.finditer(r"<input\s([^>]+)>", form_chunk, re.I):
        at = rm.group(1)
        if (_html_attr(at, "type") or "").lower() != "radio":
            continue
        if _html_attr(at, "name") != field_name:
            continue
        v = _html_attr(at, "value")
        if v is not None:
            vals.append(v)
    if len(vals) >= 2:
        return vals[-1]
    return None


def _parse_change_form_pairs_ordered(form_html: str) -> List[Tuple[str, str]]:
    """Walk tags in document order so hidden+checkbox Django patterns keep order."""
    pairs: List[Tuple[str, str]] = []
    pos = 0
    while pos < len(form_html):
        lt = form_html.find("<", pos)
        if lt == -1:
            break
        gt = form_html.find(">", lt)
        if gt == -1:
            break
        raw_tag = form_html[lt : gt + 1]
        pos = gt + 1
        tag_l = raw_tag.lower()
        if tag_l.startswith("<input"):
            m_in = re.match(r"<input\s([^>]+)/?\s*>", raw_tag, re.I)
            if not m_in:
                continue
            attrs = m_in.group(1)
            name = _html_attr(attrs, "name")
            if not name:
                continue
            typ = (_html_attr(attrs, "type") or "text").lower()
            val = _html_attr(attrs, "value") or ""
            if typ == "hidden":
                pairs.append((name, val))
            elif typ == "checkbox":
                if re.search(r"\bchecked\b", attrs, re.I):
                    pairs.append((name, val if val else "on"))
            elif typ == "radio":
                if re.search(r"\bchecked\b", attrs, re.I):
                    pairs.append((name, val))
            elif typ in ("submit", "button", "image", "file"):
                continue
            else:
                pairs.append((name, val))
        elif tag_l.startswith("<textarea"):
            m_open = re.match(r"<textarea\s([^>]+)>", raw_tag, re.I)
            if not m_open:
                continue
            tname = _html_attr(m_open.group(1), "name")
            close = form_html.lower().find("</textarea>", pos)
            if close == -1:
                break
            inner = form_html[pos:close]
            if tname:
                pairs.append((tname, inner))
            pos = close + len("</textarea>")
        elif tag_l.startswith("<select"):
            m_open = re.match(r"<select\s([^>]+)>", raw_tag, re.I)
            if not m_open:
                continue
            sname = _html_attr(m_open.group(1), "name")
            close = form_html.lower().find("</select>", pos)
            if close == -1:
                break
            inner = form_html[pos:close]
            pos = close + len("</select>")
            if not sname:
                continue
            selected_val: str | None = None
            for om in re.finditer(r"<option\s([^>]+)>", inner, re.I):
                oat = om.group(1)
                if re.search(r"\bselected\b", oat, re.I):
                    v = _html_attr(oat, "value")
                    selected_val = "" if v is None else v
                    break
            if selected_val is None:
                first = re.search(r"<option\s+([^>]+)>", inner, re.I)
                if first:
                    v = _html_attr(first.group(1), "value")
                    selected_val = "" if v is None else v
                else:
                    selected_val = ""
            pairs.append((sname, selected_val))
    save_name: str | None = None
    save_val: str | None = None
    for m in re.finditer(r"<input\s([^>]+)/?\s*>", form_html, re.I):
        attrs = m.group(1)
        if (_html_attr(attrs, "type") or "").lower() != "submit":
            continue
        save_name = _html_attr(attrs, "name")
        save_val = _html_attr(attrs, "value") or "Save"
        if save_name:
            break
    if save_name:
        pairs.append((save_name, save_val))
    else:
        pairs.append(("_save", "Save"))
    return pairs


def build_django_admin_inactive_body_from_html(html: str, csrf_token: str) -> Tuple[str | None, str]:
    form_chunk = _find_main_change_form(html)
    if not form_chunk:
        return None, "Could not find the change form (is the Cookie valid and surge_id correct?)."
    names = list(
        dict.fromkeys(
            re.findall(r"name=[\"']([^\"']*is_active[^\"']*)[\"']", form_chunk, re.I),
        )
    )
    if not names:
        return None, "This form has no is_active field name; use Manual mode (paste Payload from Network)."
    pairs = _parse_change_form_pairs_ordered(form_chunk)
    pairs = [(k, csrf_token if k == "csrfmiddlewaretoken" else v) for k, v in pairs]
    for fname in names:
        new_val = _inactive_value_to_set(form_chunk, fname)
        if new_val is None:
            new_val = "false"
        pairs = [p for p in pairs if p[0] != fname]
        pairs.append((fname, new_val))
    body = urllib.parse.urlencode(pairs, doseq=True)
    return body, ""


def post_form_urlencoded(
    url: str,
    headers: Dict[str, str],
    body_str: str,
    timeout: int,
) -> Tuple[bool, int, str]:
    h = {**headers, "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    try:
        r = requests.post(url, headers=h, data=body_str.encode("utf-8"), timeout=timeout)
        # Django admin "Save" usually responds with 302/303 to the changelist or same change page.
        ok = r.status_code in (200, 201, 204, 302, 303) or (200 <= r.status_code < 300)
        return ok, r.status_code, r.text[:800]
    except requests.RequestException as exc:
        return False, -1, str(exc)


def django_admin_deactivate_one(
    url_template: str,
    surge_id: int,
    cookie_header: str,
    form_body_template: str,
    timeout: int,
) -> Tuple[bool, int, str]:
    """GET change page → CSRF → POST same URL as form (Django admin)."""
    url = expand_deactivate_url(url_template, surge_id)
    cookie_val = cookie_header.strip()
    if cookie_val.lower().startswith("cookie:"):
        cookie_val = cookie_val.split(":", 1)[1].strip()

    get_headers = {
        "Cookie": cookie_val,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    try:
        g = requests.get(url, headers=get_headers, timeout=timeout)
    except requests.RequestException as exc:
        return False, -1, f"GET failed: {exc}"
    if not (200 <= g.status_code < 300):
        return False, g.status_code, f"GET change page failed: {g.text[:400]}"

    csrf = extract_csrf_from_html(g.text)
    if not csrf:
        return False, -1, "Could not find csrfmiddlewaretoken in GET HTML (login cookie valid?)"

    body = apply_deactivate_template(form_body_template, surge_id, csrf)
    post_url = g.url
    post_headers = {
        "Cookie": cookie_val,
        "Referer": post_url,
        "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
        "Origin": "https://api.shadowfax.in",
    }
    return post_form_urlencoded(post_url, post_headers, body, timeout)


def django_admin_deactivate_auto_one(
    url_template: str,
    surge_id: int,
    cookie_header: str,
    timeout: int,
) -> Tuple[bool, int, str]:
    """GET change page → parse form → set is_active inactive → POST (no Payload paste)."""
    url = expand_deactivate_url(url_template, surge_id)
    cookie_val = cookie_header.strip()
    if cookie_val.lower().startswith("cookie:"):
        cookie_val = cookie_val.split(":", 1)[1].strip()

    get_headers = {
        "Cookie": cookie_val,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    try:
        g = requests.get(url, headers=get_headers, timeout=timeout)
    except requests.RequestException as exc:
        return False, -1, f"GET failed: {exc}"
    if not (200 <= g.status_code < 300):
        return False, g.status_code, f"GET change page failed: {g.text[:400]}"

    csrf = extract_csrf_from_html(g.text)
    if not csrf:
        return False, -1, "Could not find csrfmiddlewaretoken in GET HTML (login cookie valid?)"

    body, berr = build_django_admin_inactive_body_from_html(g.text, csrf)
    if not body:
        return False, -1, berr

    post_url = g.url
    post_headers = {
        "Cookie": cookie_val,
        "Referer": post_url,
        "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
        "Origin": "https://api.shadowfax.in",
    }
    return post_form_urlencoded(post_url, post_headers, body, timeout)


st.title("⚡ Surge Runner")
st.caption("Create surges or deactivate existing ones (by surge_id), then re-create with new amount.")

jwt = st.text_input("JWT token", type="password", key="jwt")
timeout = st.number_input("Timeout sec", min_value=5, max_value=240, value=30, key="timeout")

tab_create, tab_deactivate = st.tabs(["Create / apply surges", "Deactivate surges (by surge_id)"])

with tab_create:
    st.markdown("**Create:** Required CSV columns: `city`, `start_time`, `duration`, `amount`")
    st.markdown(
        "City lookup: `city` is enough. Optional: `cluster_ids`, `chain_ids`. "
        "If duplicate error shows `surge_id:…`, use **Deactivate** tab first, then create again with new amount."
    )

    with st.expander("Sample CSV format (create)", expanded=True):
        sample_df = pd.DataFrame(
            [
                {
                    "city": 8,
                    "start_time": "2026-04-09 07:00",
                    "duration": 1,
                    "amount": 1,
                    "surge_type": 2,
                    "vehicle_type": 100,
                    "is_active": True,
                    "activate_for_marketplace_chains": 0,
                    "chain_ids": "6557",
                    "cluster_ids": "152,683",
                }
            ]
        )
        st.dataframe(sample_df, use_container_width=True)
        st.download_button(
            "Download sample CSV",
            sample_df.to_csv(index=False).encode("utf-8"),
            file_name="surge_runner_sample.csv",
            mime="text/csv",
            key="download_sample_csv",
        )

    delay_between_surges_sec = st.number_input(
        "Wait between surges (sec, after each successful apply)",
        min_value=0,
        max_value=120,
        value=15,
        help="Extra pause after a successful surge before the next row. Use 10–20 to reduce 429 / server load. Set 0 to disable.",
        key="delay_create",
    )
    upload = st.file_uploader("Upload CSV (create)", type=["csv"], key="upload_create")

    st.markdown(
        "**Master Excel (optional):** one `.xlsx` file that grows every run. "
        f"Sheet **{LOG_SHEET_NAME}** = full API log. **One sheet per date** (from each row’s `start_time`) = your inputs + outcome; **duplicate surge lines are skipped** (same city/time/amount/clusters/chains)."
    )
    st.checkbox("Append this run to the master workbook", value=True, key="master_excel_enable")
    st.text_input("Master workbook path", value=_DEFAULT_MASTER_XLSX, key="master_excel_path")
    st.caption(
        "No need to create the file yourself: the first run creates it; every later run updates that same path."
    )

    run_create = st.button("Apply Surges", type="primary", key="run_create")

    if upload is not None:
        df_preview, err_preview = read_csv_upload(upload)
        if err_preview:
            st.error(err_preview)
        else:
            st.caption(f"CSV preview: **{len(df_preview)}** rows — scroll the table to see all rows.")
            st.dataframe(df_preview, use_container_width=True, height=600)

    if run_create:
        if not jwt.strip():
            st.error("JWT token required")
            st.stop()
        if upload is None:
            st.error("Upload CSV first")
            st.stop()

        df, err = read_csv_upload(upload)
        if err:
            st.error(err)
            st.stop()

        req = {"city", "start_time", "duration", "amount"}
        miss = [c for c in req if c not in df.columns]
        if miss:
            st.error(f"Missing columns: {', '.join(miss)}")
            st.stop()

        headers = {
            "Authorization": f"JWT {jwt.strip()}",
            "Content-Type": "application/json",
            "Accept": "application/json, text/plain, */*",
        }

        city_cache: Dict[str, List[int]] = {}
        city_cluster_name_cache: Dict[str, Dict[str, int]] = {}
        rows: List[Dict[str, Any]] = []
        batch_id = datetime.now(IST).strftime("%Y%m%d_%H%M%S")
        master_log_rows: List[Dict[str, Any]] = []
        master_inputs_by_date: Dict[str, List[Dict[str, Any]]] = {}

        def record_master_excel_row(
            d: Dict[str, Any],
            city_key: str,
            city_id_r: Any,
            cluster_ids: List[int] | None,
            payload: Dict[str, Any] | None,
            success: bool,
            status: int,
            msg: str,
            cluster_note: str,
            row_idx: int,
        ) -> None:
            master_log_rows.append(
                {
                    "run_batch_id": batch_id,
                    "csv_row_index": row_idx,
                    "resolved_city_key": city_key,
                    "start_time_csv": str(d.get("start_time", "")),
                    "cluster_count": len(payload["cluster"]) if payload and "cluster" in payload else "",
                    "success": success,
                    "http_status": status,
                    "msg": (msg or "")[:4000],
                    "logged_at_ist": run_timestamp_ist(),
                }
            )
            date_tab = sheet_date_tab_from_row(d)
            inp = build_input_row_for_master(
                d, city_key, city_id_r, cluster_ids, payload, success, status, msg, cluster_note
            )
            master_inputs_by_date.setdefault(date_tab, []).append(inp)

        prog = st.progress(0)
        total = len(df)
        wait_status = st.empty()

        for pos, (i, row) in enumerate(df.iterrows()):
            d = row.to_dict()
            city_key = resolve_city_key(d)
            city_id_raw = d.get("city")
            city_id: int | None = None
            try:
                if city_id_raw is not None and str(city_id_raw).strip() != "":
                    city_id = int(float(city_id_raw))
            except Exception:
                city_id = None

            cache_key = f"id:{city_id}" if city_id is not None else f"name:{city_key}"
            if city_id is None and not city_key:
                msg_e = "city is missing or invalid"
                rows.append({"row": i + 1, "success": False, "status": -1, "msg": msg_e})
                record_master_excel_row(d, city_key, city_id_raw, None, None, False, -1, msg_e, "", i + 1)
                prog.progress((i + 1) / total if total else 1)
                continue

            if cache_key not in city_cache:
                ids, name_map, ferr = fetch_city_clusters(headers, city_id, city_key, int(timeout))
                if ferr:
                    msg_e = f"cluster fetch: {ferr}"
                    rows.append({"row": i + 1, "city": city_key, "success": False, "status": -1, "msg": msg_e})
                    record_master_excel_row(d, city_key, city_id_raw, None, None, False, -1, msg_e, "", i + 1)
                    prog.progress((i + 1) / total if total else 1)
                    continue
                city_cache[cache_key] = ids
                city_cluster_name_cache[cache_key] = name_map

            cluster_ids: List[int] | None = None
            cluster_note = ""
            payload: Dict[str, Any] | None = None
            try:
                cluster_ids, cluster_note = choose_cluster_ids(
                    d,
                    city_cache[cache_key],
                    city_cluster_name_cache.get(cache_key, {}),
                )
                if not cluster_ids:
                    msg_e = f"cluster select: {cluster_note}"
                    rows.append(
                        {
                            "row": i + 1,
                            "city": city_key,
                            "success": False,
                            "status": -1,
                            "msg": msg_e,
                        }
                    )
                    record_master_excel_row(
                        d, city_key, city_id_raw, None, None, False, -1, msg_e, cluster_note, i + 1
                    )
                    prog.progress((i + 1) / total if total else 1)
                    continue

                payload = build_payload(d, cluster_ids)
                ok, code, msg = post_payload(headers, payload, int(timeout))
                msg_out = msg if not cluster_note else f"{cluster_note} | {msg}"
                rows.append(
                    {
                        "row": i + 1,
                        "city": city_key,
                        "cluster_count": len(payload["cluster"]),
                        "success": ok,
                        "status": code,
                        "msg": msg_out,
                    }
                )
                record_master_excel_row(
                    d, city_key, city_id_raw, cluster_ids, payload, ok, code, msg_out, cluster_note, i + 1
                )
                if ok and float(delay_between_surges_sec) > 0 and pos < total - 1:
                    wait_status.info(f"Waiting {float(delay_between_surges_sec):.0f}s before next surge…")
                    time.sleep(float(delay_between_surges_sec))
                    wait_status.empty()
            except Exception as exc:
                msg_e = f"payload: {exc}"
                rows.append({"row": i + 1, "city": city_key, "success": False, "status": -1, "msg": msg_e})
                record_master_excel_row(
                    d, city_key, city_id_raw, cluster_ids, payload, False, -1, msg_e, cluster_note, i + 1
                )
            prog.progress((i + 1) / total if total else 1)

        out = pd.DataFrame(rows)
        st.dataframe(out, use_container_width=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "Download result CSV",
            out.to_csv(index=False).encode("utf-8"),
            f"surge_result_{ts}.csv",
            "text/csv",
            key="download_create_result",
        )

        if st.session_state.get("master_excel_enable", True):
            raw_path = st.session_state.get("master_excel_path", _DEFAULT_MASTER_XLSX)
            mpath = Path(str(raw_path).strip() or _DEFAULT_MASTER_XLSX)
            if not str(mpath).lower().endswith(".xlsx"):
                st.warning("Master path should end with .xlsx — skipped.")
            else:
                log_df = pd.DataFrame(master_log_rows)
                inputs_by_date_df = {k: pd.DataFrame(v) for k, v in master_inputs_by_date.items()}
                ok_m, err_m = merge_master_run_to_excel(mpath, log_df, inputs_by_date_df)
                if ok_m:
                    st.success(f"Master workbook updated: `{mpath}`")
                else:
                    st.warning(f"Could not update master workbook: {err_m}")


DJANGO_SIMPLE_URL = (
    "https://api.shadowfax.in/admin/configuration/multipleclustersurgemode/{{surge_id}}/change/"
)


with tab_deactivate:
    st.subheader("Turn surges inactive (simple)")
    st.markdown(
        """
**Easy path (default): two things only**

1. **Cookie** — While logged into **admin** on `api.shadowfax.in`: **F12** → **Network** → click any request → **Request headers** → copy the full **Cookie** value (or at least `sessionid=…` and `csrftoken=…`).
2. **CSV** — One column **`surge_id`**: the number in the admin URL (`…/multipleclustersurgemode/**77217**/change/` → `77217`).

The app opens each surge’s change page, fills **Active = No**, and clicks **Save** for you. No DevTools Payload paste.

**If that ever fails** (unusual admin layout), switch to **Manual** below and paste form data once from Network like before.
        """
    )

    admin_cookie = st.text_area(
        "1) Cookie (from admin browser)",
        value="",
        height=72,
        placeholder="sessionid=...; csrftoken=...",
        key="admin_cookie_deactivate",
    )

    deact_mode = st.radio(
        "2) How to fill the Save form",
        (
            "Easy — Cookie + CSV only (no Payload paste)",
            "Manual — paste form data from Network (Payload)",
        ),
        index=0,
        key="deact_mode",
    )
    use_easy_deactivate = deact_mode.startswith("Easy")

    if use_easy_deactivate:
        st.caption("Easy mode reads each surge’s admin page and sets inactive automatically.")
    else:
        st.text_area(
            "2b) Form data — edit one surge → Active = No → Save → Network → that POST → Payload → paste here. "
            "Replace the csrf value with exactly: {{csrf_token}}",
            value="",
            height=160,
            key="django_form_template",
        )

    st.number_input(
        "Wait between each surge (seconds, after success)",
        min_value=0,
        max_value=120,
        value=10,
        key="delay_deactivate",
    )

    upload_deact = st.file_uploader("CSV — column name must be: surge_id", type=["csv"], key="upload_deactivate")
    run_deact = st.button("Turn OFF surges (inactive)", type="primary", key="run_deactivate")

    sample_deact = pd.DataFrame([{"surge_id": 77217}])
    st.download_button(
        "Download example CSV",
        sample_deact.to_csv(index=False).encode("utf-8"),
        file_name="surge_deactivate_sample.csv",
        mime="text/csv",
        key="download_deactivate_sample",
    )

    with st.expander("Advanced only: deactivate with JWT + JSON (not admin)", expanded=False):
        st.checkbox(
            "Use this instead of Cookie + form above",
            value=False,
            key="use_adv_json_deact",
        )
        st.text_input(
            "API URL (use {{surge_id}} in path if needed)",
            key="deactivate_url",
            placeholder="https://...",
        )
        st.selectbox("HTTP method", ["POST", "PUT", "PATCH"], index=0, key="deactivate_method")
        st.text_area(
            "JSON body",
            value='{"surge_id": {{surge_id}}, "is_active": false}',
            height=80,
            key="deactivate_body_template",
        )

    if upload_deact is not None:
        dprev, derr = read_csv_upload(upload_deact)
        if not derr:
            st.dataframe(dprev, use_container_width=True, height=400)

    if run_deact:
        if upload_deact is None:
            st.error("Upload a CSV with column surge_id.")
            st.stop()

        df_d, err_d = read_csv_upload(upload_deact)
        if err_d:
            st.error(err_d)
            st.stop()
        if "surge_id" not in df_d.columns:
            st.error("CSV must have a column named exactly: surge_id")
            st.stop()

        use_simple = not st.session_state.get("use_adv_json_deact", False)

        if use_simple:
            if not admin_cookie.strip():
                st.error("Paste the admin Cookie (step 1).")
                st.stop()
            easy = st.session_state.get("deact_mode", "Easy").startswith("Easy")
            if not easy and not st.session_state.get("django_form_template", "").strip():
                st.error("Manual mode: paste the form data from Save (Payload), with {{csrf_token}} for the csrf value.")
                st.stop()
            url_tpl = DJANGO_SIMPLE_URL
        else:
            if not jwt.strip():
                st.error("Advanced mode needs JWT at top of page.")
                st.stop()
            if not st.session_state.get("deactivate_url", "").strip():
                st.error("Advanced mode: set URL in the Advanced section.")
                st.stop()
            url_tpl = st.session_state["deactivate_url"].strip()
            use_django = False

        delay_sec = float(st.session_state.get("delay_deactivate", 10))

        rows_d: List[Dict[str, Any]] = []
        prog_d = st.progress(0)
        total_d = len(df_d)
        wait_d = st.empty()

        for pos, (idx, row) in enumerate(df_d.iterrows()):
            sid_raw = row.get("surge_id")
            try:
                sid = int(float(sid_raw))
            except Exception:
                rows_d.append({"row": pos + 1, "surge_id": sid_raw, "success": False, "status": -1, "msg": "invalid surge_id"})
                prog_d.progress((pos + 1) / total_d if total_d else 1)
                continue

            if use_simple:
                if st.session_state.get("deact_mode", "Easy").startswith("Easy"):
                    ok, code, msg = django_admin_deactivate_auto_one(
                        url_tpl,
                        sid,
                        admin_cookie.strip(),
                        int(timeout),
                    )
                else:
                    ok, code, msg = django_admin_deactivate_one(
                        url_tpl,
                        sid,
                        admin_cookie.strip(),
                        st.session_state.get("django_form_template", ""),
                        int(timeout),
                    )
            else:
                try:
                    body = build_deactivate_payload(st.session_state.get("deactivate_body_template", "{}"), sid)
                except Exception as exc:
                    rows_d.append({"row": pos + 1, "surge_id": sid, "success": False, "status": -1, "msg": f"json: {exc}"})
                    prog_d.progress((pos + 1) / total_d if total_d else 1)
                    continue
                final_url = expand_deactivate_url(st.session_state.get("deactivate_url", "").strip(), sid)
                headers = {
                    "Authorization": f"JWT {jwt.strip()}",
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/plain, */*",
                }
                ok, code, msg = request_json_method(
                    final_url,
                    st.session_state.get("deactivate_method", "POST"),
                    headers,
                    body,
                    int(timeout),
                )

            rows_d.append({"row": pos + 1, "surge_id": sid, "success": ok, "status": code, "msg": msg})
            if ok and delay_sec > 0 and pos < total_d - 1:
                wait_d.info(f"Waiting {delay_sec:.0f}s…")
                time.sleep(delay_sec)
                wait_d.empty()
            prog_d.progress((pos + 1) / total_d if total_d else 1)

        out_d = pd.DataFrame(rows_d)
        st.dataframe(out_d, use_container_width=True)
        ts_d = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "Download result CSV",
            out_d.to_csv(index=False).encode("utf-8"),
            f"surge_deactivate_result_{ts_d}.csv",
            "text/csv",
            key="download_deactivate_result",
        )
