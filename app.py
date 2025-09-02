# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re, io, calendar
from datetime import datetime, date

st.set_page_config(page_title="제네시스 정산 생성기", layout="wide")

# =========================
# Helpers
# =========================
def pick(cols, *keys):
    for k in keys:
        hit = [c for c in cols if k in str(c)]
        if hit: return hit[0]
    return None

def normalize_invoice(x):
    if pd.isna(x): return None
    s = re.sub(r"[^0-9\-]", "", str(x).strip())
    d = re.sub(r"\D", "", s)
    return f"{d[0:4]}-{d[4:8]}-{d[8:12]}" if len(d)==12 else (s or None)

def contains(x, pat): return bool(re.search(pat, str(x)))

def is_yeonmakgi(x):   return contains(x, r"연막기")
def is_beoksdelta(x):  return contains(x, r"벅스델타")
def is_hwaksan(x):     return contains(x, r"확산제")

def to_date(x):
    try: return pd.to_datetime(x).date()
    except: return pd.NaT

def gt0(x):
    try: return float(x) > 0
    except: return False

def rocket_from_spec(txt):
    # 롯데 '특기사항'에서 뱀이싹 라벨 파싱 (띄어쓰기/변형 허용)
    s = str(txt)
    has20  = bool(re.search(r"뱀이싹\s*20\s*개(입)?", s))
    has2ea = bool(re.search(r"뱀이싹\s*2\s*개입", s))
    if has20 and has2ea: return "혼합"
    if has2ea: return "작업"   # 로켓그로스 작업
    if has20:  return "단순"   # 로켓그로스 미작업
    return ""

def classify_type(row):
    # 최신 규칙:
    # - 연막기 ≥2 → B
    # - 연막기 ==1 & (부자재 0, 1, (1+0), (0+1), (1+1)) → B
    # - 연막기 ==1 & (부자재 합계 ≥2) 단, 1+1은 B이므로 제외 → C
    # - 연막기 ==0 → A
    y, b, h = row["연막기개수"], row["벅스델타개수"], row["확산제개수"]
    if y >= 2:
        return "B"
    if y == 1:
        total = b + h
        if (total == 0) or (b==1 and h==0) or (b==0 and h==1) or (b==1 and h==1):
            return "B"
        if total >= 2:
            return "C"
        return "A"
    return "A"

weekday_map = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}

# =========================
# UI
# =========================
st.title("📦 제네시스 정산 생성기")
st.caption("롯데/사방넷 엑셀을 드래그&드롭 → 버튼 클릭 → 최종 엑셀 다운로드")

col1, col2 = st.columns(2)
with col1:
    lotte_file = st.file_uploader("롯데.xlsx (시트0 사용)", type=["xlsx"])
with col2:
    sabang_file = st.file_uploader("사방넷.xlsx (시트0 사용)", type=["xlsx"])

with st.expander("⚙️ 옵션(선택)"):
    forced_year = st.number_input("연도 강제 설정 (미설정=자동)", min_value=2000, max_value=2100, value=0, step=1)
    forced_month = st.number_input("월 강제 설정 (미설정=자동)", min_value=0, max_value=12, value=0, step=1)
    out_name = st.text_input("파일명 접두어 (미입력시 자동)", value="제네시스_정산")

go = st.button("정산 파일 생성")

# =========================
# Processing
# =========================
if go:
    if not lotte_file or not sabang_file:
        st.error("롯데/사방넷 파일을 모두 업로드하세요.")
        st.stop()

    try:
        df_l = pd.ExcelFile(lotte_file).parse(0).copy()
        df_s = pd.ExcelFile(sabang_file).parse(0).copy()
    except Exception as e:
        st.error(f"엑셀 로딩 오류: {e}")
        st.stop()

    # --- Column map
    col_l_inv   = pick(df_l.columns, "운송장번호","송장")
    col_l_date  = pick(df_l.columns, "집하일자")
    col_l_spec  = pick(df_l.columns, "특기사항")
    col_l_jeju  = pick(df_l.columns, "제주연계")
    col_l_dose1 = pick(df_l.columns, "집하도선료")
    col_l_dose2 = pick(df_l.columns, "배달도선료")

    col_s_inv   = pick(df_s.columns, "송장번호","운송장")
    col_s_name  = pick(df_s.columns, "출고상품명","상품명")
    col_s_qty   = pick(df_s.columns, "수량")

    for need, nm in [("롯데 운송장번호", col_l_inv), ("사방넷 송장번호", col_s_inv),
                     ("사방넷 상품명", col_s_name), ("사방넷 수량", col_s_qty)]:
        if not nm:
            st.error(f"필수 컬럼 탐지 실패: {need}")
            st.stop()

    # --- Normalize keys
    df_l["invoice"] = df_l[col_l_inv].map(normalize_invoice)
    df_s["invoice"] = df_s[col_s_inv].map(normalize_invoice)
    df_s[col_s_qty] = pd.to_numeric(df_s[col_s_qty], errors="coerce").fillna(0).astype(int)

    if not df_s["invoice"].notna().any():
        st.error("사방넷 invoice 전부 결측입니다. 원본을 확인하세요.")
        st.stop()

    # --- Aggregate per invoice
    agg = (
        df_s.groupby("invoice", dropna=True)
            .apply(lambda g: pd.Series({
                "연막기개수":   int((g[col_s_name].apply(is_yeonmakgi)   * g[col_s_qty]).sum()),
                "벅스델타개수": int((g[col_s_name].apply(is_beoksdelta) * g[col_s_qty]).sum()),
                "확산제개수":   int((g[col_s_name].apply(is_hwaksan)    * g[col_s_qty]).sum()),
            }))
    ).reset_index()

    agg["type"] = agg.apply(classify_type, axis=1)

    # --- Merge Lotte & Rocket
    df_l_use = df_l[["invoice", col_l_date, col_l_spec, col_l_jeju, col_l_dose1, col_l_dose2]].copy()
    df = agg.merge(df_l_use, on="invoice", how="left")
    df["rocket_label"] = df[col_l_spec].apply(rocket_from_spec)
    df["연막기_그로스"] = np.where((df["연막기개수"]>0) & (df["rocket_label"]!=""), 1, 0)

    # --- Flags & Dates
    df["집하일자"] = df[col_l_date].apply(to_date) if col_l_date else pd.NaT
    df["요일"] = df["집하일자"].apply(lambda d: weekday_map[d.weekday()] if pd.notna(d) else "")
    df["일"] = df["집하일자"].apply(lambda d: d.day if pd.notna(d) else np.nan)
    df["is_jeju"]   = df[col_l_jeju].apply(gt0)  if col_l_jeju  else False
    dose1 = df[col_l_dose1].apply(gt0) if col_l_dose1 else False
    dose2 = df[col_l_dose2].apply(gt0) if col_l_dose2 else False
    df["is_doseon"] = dose1 | dose2

    # --- Detail
    detail_cols = [
        "invoice","집하일자","요일","일",
        "연막기개수","벅스델타개수","확산제개수",
        "type","rocket_label","연막기_그로스","is_jeju","is_doseon"
    ]
    detail_cols = [c for c in detail_cols if c in df.columns]
    주문상세 = (df[detail_cols]
                .drop_duplicates(subset=["invoice"])
                .sort_values(["집하일자","invoice"], na_position="last")
                .reset_index(drop=True))

    # --- Year/Month
    if forced_year and forced_month:
        year, month = int(forced_year), int(forced_month)
    else:
        if 주문상세["집하일자"].notna().any():
            base_date = pd.to_datetime(주문상세["집하일자"].dropna().iloc[0])
            year, month = base_date.year, base_date.month
        else:
            year, month = 2025, 8

    # --- Daily summary (nunique invoices)
    last_day = calendar.monthrange(year, month)[1]
    days = pd.date_range(f"{year}-{month:02d}-01", f"{year}-{month:02d}-{last_day}", freq="D")
    cal_df = pd.DataFrame({"집하일자": days.date,
                           "일": [d.day for d in days],
                           "요일": [weekday_map[d.weekday()] for d in days]})

    by_day_type = (주문상세.groupby(["집하일자","type"])["invoice"]
                   .nunique().unstack(fill_value=0))
    for t in ["A","B","C"]:
        if t not in by_day_type.columns: by_day_type[t] = 0
    by_day_type = by_day_type.reset_index()

    그작 = (주문상세.assign(_=(주문상세["rocket_label"]=="작업").astype(int))
            .groupby("집하일자")["_"].sum().reset_index().rename(columns={"_":"그로스(작업)"}))
    그단 = (주문상세.assign(_=(주문상세["rocket_label"]=="단순").astype(int))
            .groupby("집하일자")["_"].sum().reset_index().rename(columns={"_":"그로스(단순)"}))
    연그 = 주문상세.groupby("집하일자")["연막기_그로스"].sum().reset_index()

    현황 = (cal_df.merge(by_day_type, on="집하일자", how="left")
                   .merge(그작, on="집하일자", how="left")
                   .merge(그단, on="집하일자", how="left")
                   .merge(연그, on="집하일자", how="left")).fillna(0)
    현황["D타입(일반)"] = 0
    현황["E타입(일반)"] = 0

    타입별_배송현황 = 현황[[
        "일","요일","A","B","C","D타입(일반)","E타입(일반)","그로스(작업)","그로스(단순)","연막기_그로스"
    ]].rename(columns={
        "A":"A타입(일반)", "B":"B타입(일반)", "C":"C타입(일반)",
        "연막기_그로스":"연막기 그로스"
    })

    # --- Monthly summaries
    sum_type = (주문상세.groupby("type")["invoice"]
                .nunique().reindex(["A","B","C"]).fillna(0).astype(int)
                .rename("건수").reset_index())

    sum_rocket = (주문상세["rocket_label"]
                  .replace({"":"(없음)"})
                  .value_counts().rename_axis("rocket_label").reset_index(name="건수"))

    # --- Validation
    total_sabang = df_s["invoice"].dropna().nunique()
    total_detail = 주문상세["invoice"].nunique()
    sum_types_total = int(sum_type["건수"].sum())

    검증리포트 = pd.DataFrame([
        ("사방넷 원천 송장수", total_sabang),
        ("주문상세 송장수", total_detail),
        ("타입별 합계 송장수(월간)", sum_types_total),
        ("일치 여부(원천=상세)", total_sabang == total_detail),
        ("일치 여부(상세=타입합계)", total_detail == sum_types_total),
    ], columns=["항목","값"])

    miss_set = set(df_s["invoice"].dropna()) - set(df_l["invoice"].dropna())
    미매칭 = pd.DataFrame({"invoice(사방넷 기준, 롯데 미존재)": sorted(list(miss_set))})

    # --- Preview
    st.success(f"월 자동판정 → {year}년 {month}월")
    st.dataframe(주문상세.head(20))
    st.dataframe(타입별_배송현황.head(20))
    st.dataframe(sum_type)
    st.dataframe(sum_rocket)
    st.dataframe(검증리포트)
    if len(미매칭):
        st.warning(f"미매칭 {len(미매칭)}건 존재")
        st.dataframe(미매칭.head(30))

    # --- Build Excel in-memory
    bio = io.BytesIO()
    out_filename = f"{out_name or '제네시스_정산'}_{year}년{month}월_산출.xlsx"
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        주문상세.to_excel(w, index=False, sheet_name="주문상세")
        타입별_배송현황.to_excel(w, index=False, sheet_name="타입별 배송현황")
        sum_type.to_excel(w, index=False, sheet_name="월간 요약(타입)")
        sum_rocket.to_excel(w, index=False, sheet_name="월간 요약(로켓)")
        검증리포트.to_excel(w, index=False, sheet_name="검증리포트")
        미매칭.to_excel(w, index=False, sheet_name="미매칭")
    bio.seek(0)

    # --- Download
    st.download_button(
        label="⬇️ 정산 엑셀 다운로드",
        data=bio,
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )