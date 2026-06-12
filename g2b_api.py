"""
조달청 특정품목조달내역 API 호출 모듈
g2b_api.py
"""

import os
import time
import requests
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("G2B_API_KEY")
ENDPOINT = "https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService/getSpcifyPrdlstPrcureInfoList"


def fetch_page(params, page_no, num_rows=100, timeout=30):
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
        # No Data(03)는 정상적으로 빈 결과 처리
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
              extra=None, num_rows=100, max_pages=200, progress=None):
    """
    조건에 맞는 모든 데이터를 페이지를 넘기며 수집.

    inqry_div: '1'(계약납품요구일자) 또는 '2'(최초계약납품요구일자)
    prdct_div: '1'(품명) '2'(세부품명) '3'(물품규격명)
    progress: 진행 상황을 알려주는 콜백 함수(선택)
    """
    base = {
        "inqryDiv": inqry_div,
        "inqryBgnDate": bgn_date,
        "inqryEndDate": end_date,
        "inqryPrdctDiv": prdct_div,
    }
    # 품목 검색 조건 (구분에 맞는 것만 채움)
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

    # 추가 필터 (우수제품여부, 수요기관, 업체, 지역, 계약방법 등)
    if extra:
        base.update({k: v for k, v in extra.items() if v})

    all_items = []
    # 첫 페이지로 총 건수 파악
    items, total = fetch_page(base, 1, num_rows)
    all_items.extend(items)

    if total == 0:
        return [], 0

    total_pages = (total + num_rows - 1) // num_rows
    total_pages = min(total_pages, max_pages)

    if progress:
        progress(1, total_pages, len(all_items), total)

    for p in range(2, total_pages + 1):
        time.sleep(0.1)  # 서버 배려
        items, _ = fetch_page(base, p, num_rows)
        all_items.extend(items)
        if progress:
            progress(p, total_pages, len(all_items), total)

    return all_items, total