"""
조달청 특정품목조달내역 API 호출 모듈
g2b_api.py - 속도 개선 버전

변경점 (이전 버전 대비):
- 한 번에 가져오는 건수: 100 → 999 (약 10배 빠름)
- 호출 간 대기 시간: 0.1초 → 0.05초 (서버 부담 없이 약간 절약)

API 키 읽는 우선순위:
1. Streamlit Cloud의 Secrets (배포 환경)
2. 로컬 .env 파일 (개발 환경)
"""

import os
import time
import requests
from dotenv import load_dotenv

# 1) 로컬 .env가 있으면 먼저 읽기
load_dotenv()

# 2) Streamlit Cloud Secrets에서 읽기 시도 (배포 환경)
API_KEY = os.getenv("G2B_API_KEY")
if not API_KEY:
    try:
        import streamlit as st
        API_KEY = st.secrets.get("G2B_API_KEY")
    except Exception:
        API_KEY = None

ENDPOINT = "https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService/getSpcifyPrdlstPrcureInfoList"

# ⚡ 한 번에 가져올 건수 (최대 999까지 가능)
PAGE_SIZE = 999


def fetch_page(params, page_no, num_rows=PAGE_SIZE, timeout=30):
    """단일 페이지를 호출해서 (items, total_count)를 반환"""
    q = dict(params)
    q.update({
        "serviceKey": API_KEY,
        "pageNo": str(page_no),
        "numOfRows": str(num_rows),
        "type": "json",
    })
    r = requests.get(ENDPOINT, params=q, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    resp = data.get("response", {})
    header = resp.get("header", {})
    if header.get("resultCode") not in ("00", "0", None):
        if header.get("resultCode") == "03":
            return [], 0
        raise RuntimeError(f"API 오류: {header.get('resultMsg')} (코드 {header.get('resultCode')})")

    body = resp.get("body", {})
    total = int(body.get("totalCount", 0) or 0)
    items = body.get("items", [])
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    return items or [], total


def fetch_all(inqry_div, bgn_date, end_date, prdct_div,
              prdct_clsfc_no=None, prdct_clsfc_no_nm=None,
              dtil_prdct_clsfc_no=None, dtil_prdct_clsfc_no_nm=None,
              prdct_idnt_no=None, prdct_idnt_no_nm=None,
              extra=None, num_rows=PAGE_SIZE, max_pages=200, progress=None):
    """조건에 맞는 모든 데이터를 페이지를 넘기며 수집"""
    base = {
        "inqryDiv": inqry_div,
        "inqryBgnDate": bgn_date,
        "inqryEndDate": end_date,
        "inqryPrdctDiv": prdct_div,
    }
    if prdct_clsfc_no:
        base["prdctClsfcNo"] = prdct_clsfc_no
    if prdct_clsfc_no_nm:
        base["prdctClsfcNoNm"] = prdct_clsfc_no_nm
    if dtil_prdct_clsfc_no:
        base["dtilPrdctClsfcNo"] = dtil_prdct_clsfc_no
    if dtil_prdct_clsfc_no_nm:
        base["dtilPrdctClsfcNoNm"] = dtil_prdct_clsfc_no_nm
    if prdct_idnt_no:
        base["prdctIdntNo"] = prdct_idnt_no
    if prdct_idnt_no_nm:
        base["prdctIdntNoNm"] = prdct_idnt_no_nm

    if extra:
        base.update({k: v for k, v in extra.items() if v})

    all_items = []
    items, total = fetch_page(base, 1, num_rows)
    all_items.extend(items)

    if total == 0:
        return [], 0

    total_pages = (total + num_rows - 1) // num_rows
    total_pages = min(total_pages, max_pages)

    if progress:
        progress(1, total_pages, len(all_items), total)

    for p in range(2, total_pages + 1):
        time.sleep(0.05)  # ⚡ 0.1초 → 0.05초로 단축
        items, _ = fetch_page(base, p, num_rows)
        all_items.extend(items)
        if progress:
            progress(p, total_pages, len(all_items), total)

    return all_items, total