"""EP 데이터 변환기 — 사내 원본 파일을 대시보드용 CSV로 변환합니다.

지원 파일:
1. EP채널 데이터 (Data.xlsx / Data.csv) → ep_data_long.csv
2. EP실적 데이터 (1_EP실적.csv) → ep_traffic.csv
"""
import datetime
import io
import streamlit as st
import pandas as pd

st.set_page_config(page_title="EP 데이터 변환기", page_icon="🔄", layout="centered")

# ─── 공통 상수 ───
METRIC_ORDER = [
    "평균 EP 전시 상품수", "평균 원부매칭 상품수", "원부매칭율(%)",
    "평균 최저가 상품수", "최저가율(%)", "평균 EP 거래액(순결제)",
    "평균 EP 거래액(총결제)", "평균 EP 고객수(총결제)",
    "평균 EP 첫구매 거래액(총결제)", "평균 EP 첫구매 고객수(총결제)",
    "첫구매거래액(%)", "평균 EP UV", "평균 EP 비회원UV",
    "EP 전시 상품당 유입수", "평균 EP 신규가입수", "신규가입율",
    "구매전환율(%)", "첫구매 전환율(%)",
]
PERCENT_COLS = {"원부매칭율(%)", "최저가율(%)", "첫구매거래액(%)",
                "신규가입율", "구매전환율(%)", "첫구매 전환율(%)"}
KEEP_BPU = {"Total", "e-영업1", "e-영업2", "e-영업3", "e-영업4"}
KEEP_MATCH = {"Total", "매칭"}
KEEP_LOWEST = {"Total", "최저가"}


def _parse_date(year_val, md_val):
    try:
        y = int(str(year_val).strip())
        m, d = str(md_val).strip().split("/")
        return datetime.date(y, int(m), int(d))
    except (ValueError, AttributeError):
        return None


# ─── 파일 타입 자동 판별 ───
def detect_file_type(uploaded_file):
    """파일 내용을 보고 EP채널 / EP실적 자동 판별."""
    name = uploaded_file.name.lower()

    # 엑셀이면 무조건 EP채널
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return "ep_channel"

    # CSV: 내용으로 판별
    raw = uploaded_file.read(10000)
    uploaded_file.seek(0)

    if raw[:2] == b"\xff\xfe":
        text = raw.decode("utf-16-le", errors="ignore")
    else:
        text = raw.decode("utf-8", errors="ignore")

    # EP실적 파일 특징: 헤더에 "회원구분"이 있음 (EP채널에는 없는 컬럼)
    if "회원구분" in text:
        return "ep_traffic"
    # EP채널 파일 특징: "평균 EP" 또는 "원부매칭" 이 들어있음
    if "평균 EP" in text or "원부매칭" in text:
        return "ep_channel"
    # BPU가 첫 컬럼이면 EP실적 (상세)
    if text.strip().startswith("BPU") or "e-영업" in text[:200]:
        return "ep_traffic"

    return "unknown"


# ─── EP채널 변환 ───
def convert_ep_channel(uploaded_file, file_name):
    ext = file_name.lower().split(".")[-1]
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        pct_is_fraction = True
    else:
        raw = uploaded_file.read()
        uploaded_file.seek(0)
        if raw[:2] == b"\xff\xfe":
            text = raw.decode("utf-16-le")
            df = pd.read_csv(io.StringIO(text), sep="\t", header=None, low_memory=False)
        else:
            df = pd.read_csv(uploaded_file, encoding="utf-8-sig", sep=None,
                             engine="python", header=None, low_memory=False)
        pct_is_fraction = False

    metric_row = df.iloc[0].ffill()
    year_row = df.iloc[1].ffill()
    monthday_row = df.iloc[2]
    label_cols = df.iloc[4:, 0:3].ffill()

    metric_col_dates = {}
    for metric in METRIC_ORDER:
        cols = [c for c in df.columns if c >= 3 and metric_row[c] == metric]
        cd = []
        for c in cols:
            dt = _parse_date(year_row[c], monthday_row[c])
            if dt is not None:
                cd.append((c, dt))
        metric_col_dates[metric] = cd

    rows = []
    for r in range(4, len(df)):
        bpu = label_cols.loc[r, 0]
        match = label_cols.loc[r, 1]
        lowest = label_cols.loc[r, 2]
        if pd.isna(bpu) or pd.isna(match) or pd.isna(lowest):
            continue
        if bpu not in KEEP_BPU or match not in KEEP_MATCH or lowest not in KEEP_LOWEST:
            continue
        by_date = {}
        for metric in METRIC_ORDER:
            for c, dt in metric_col_dates[metric]:
                raw_val = df.iloc[r, c]
                if isinstance(raw_val, str):
                    raw_val = raw_val.replace(",", "").replace("%", "")
                val = float(raw_val) if pd.notna(raw_val) and raw_val != "" else None
                if val is not None and metric in PERCENT_COLS and pct_is_fraction:
                    val = val * 100
                by_date.setdefault(dt, {})[metric] = val
        for dt, metrics in by_date.items():
            row = {"날짜": dt, "BPU": bpu, "원부매칭여부": match, "최저가여부": lowest}
            for metric in METRIC_ORDER:
                row[metric] = metrics.get(metric)
            rows.append(row)

    out = pd.DataFrame(rows).sort_values(["BPU", "원부매칭여부", "최저가여부", "날짜"])
    out["날짜"] = pd.to_datetime(out["날짜"]).dt.strftime("%Y-%m-%d")
    return out.reset_index(drop=True)


# ─── EP실적 변환 ───
def convert_ep_traffic(uploaded_file):
    raw = uploaded_file.read()
    uploaded_file.seek(0)
    if raw[:2] == b"\xff\xfe":
        text = raw.decode("utf-16-le")
        df = pd.read_csv(io.StringIO(text), sep="\t", header=None, low_memory=False)
    else:
        df = pd.read_csv(uploaded_file, encoding="utf-8-sig", sep=None,
                         engine="python", header=None, low_memory=False)

    # --- 구조 자동 판별 ---
    header_row0 = str(df.iloc[0, 0]).strip()
    if header_row0 == "BPU" or header_row0 in ("e-영업1", "e-영업2", "e-영업3", "e-영업4"):
        # 새 구조: 열0=BPU, 열1=지표, 열2=회원구분, 열3=신규구분1, 열4=신규구분2, 열5=카테고리, 열6=브랜드, 열7~=날짜
        return _convert_traffic_new(df)  # (ep_traffic_df, ep_category_df_or_None)
    else:
        # 기존 구조: 열0=지표, 열1=회원구분, 열2=신규구분1, 열3=신규구분2, 열4=구분, 열5=BPU, 열6~=날짜
        return _convert_traffic_old(df), None


def _convert_traffic_old(df):
    """기존 EP실적 구조 (열0=지표, 열5=BPU, 열4=구분)."""
    col0 = df.iloc[:, 0].ffill()
    date_cols = {}
    for c in range(6, df.shape[1]):
        v = str(df.iloc[0, c])
        if v.startswith("20"):
            date_cols[v] = c

    targets = [
        ("Total", "전체", "전체"),
        ("e-영업1", "기본", "e-영업1"),
        ("e-영업2", "기본", "e-영업2"),
        ("e-영업3", "기본", "e-영업3"),
        ("e-영업4", "기본", "e-영업4"),
    ]
    SEGMENTS = [
        ("전체", "전체", "전체"),
        ("회원", "회원", "전체"),
        ("비회원", "비회원", "전체"),
        ("신규", "전체", "신규"),
        ("기존", "전체", "기존"),
    ]

    rows = []
    for bpu_label, gubun, bpu_val in targets:
        for metric in ["트래픽", "거래액", "구매객수", "CR", "객단가"]:
            for seg_label, member_filter, sinew_filter in SEGMENTS:
                mask = ((col0 == metric) & (df.iloc[:, 1] == member_filter) &
                        (df.iloc[:, 2] == sinew_filter) & (df.iloc[:, 3] == "전체") &
                        (df.iloc[:, 4] == gubun) & (df.iloc[:, 5] == bpu_val))
                matched = df[mask]
                if matched.empty:
                    continue
                row_data = matched.iloc[0]
                for date_str, col_idx in date_cols.items():
                    val = str(row_data.iloc[col_idx]).replace(",", "").replace("%", "")
                    try:
                        val = float(val)
                    except:
                        val = None
                    rows.append({"날짜": date_str, "BPU": bpu_label,
                                 "회원구분": seg_label, "지표": metric, "값": val})
    return _pivot_traffic(rows)


def _convert_traffic_new(df):
    """새 EP실적 구조 (열0=BPU, 열1=지표, 열5=카테고리, 열6=브랜드).
    Total 행이 없으므로 BPU 합산으로 만든다.
    반환: (ep_traffic 형태 DataFrame, ep_category 형태 DataFrame or None)
    """
    bpu_col = df.iloc[:, 0].ffill()

    date_cols = {}
    for c in range(7, df.shape[1]):
        v = str(df.iloc[0, c])
        if v.startswith("20"):
            date_cols[v] = c
    date_list = list(date_cols.items())

    SEGMENTS = [
        ("전체", "전체", "전체"),
        ("회원", "회원", "전체"),
        ("비회원", "비회원", "전체"),
        ("신규", "전체", "신규"),
        ("기존", "전체", "기존"),
    ]

    # ── 1. ep_traffic (카테고리=전체,브랜드=전체, 5개 세그먼트) ──
    rows = []
    for bpu in ["e-영업1", "e-영업2", "e-영업3", "e-영업4"]:
        for metric in ["트래픽", "거래액", "구매객수", "CR", "객단가"]:
            for seg_label, member_filter, sinew_filter in SEGMENTS:
                mask = ((bpu_col == bpu) & (df.iloc[:, 1] == metric) &
                        (df.iloc[:, 2] == member_filter) & (df.iloc[:, 3] == sinew_filter) &
                        (df.iloc[:, 4] == "전체") & (df.iloc[:, 5] == "전체") & (df.iloc[:, 6] == "전체"))
                matched = df[mask]
                if matched.empty:
                    continue
                row_data = matched.iloc[0]
                for date_str, col_idx in date_list:
                    val = str(row_data.iloc[col_idx]).replace(",", "").replace("%", "")
                    try:
                        val = float(val)
                    except:
                        val = None
                    rows.append({"날짜": date_str, "BPU": bpu,
                                 "회원구분": seg_label, "지표": metric, "값": val})

    pivot = _pivot_traffic(rows)

    # Total 행 생성 (BPU 합산, CR/객단가는 재계산)
    total_rows = []
    for date_str in sorted(date_cols.keys()):
        for seg_label, _, _ in SEGMENTS:
            sub = pivot[(pivot["날짜"] == date_str) & (pivot["회원구분"] == seg_label)]
            if sub.empty:
                continue
            row = {"날짜": date_str, "BPU": "Total", "회원구분": seg_label}
            for col in ["트래픽", "거래액", "구매객수"]:
                if col in sub.columns:
                    row[col] = sub[col].sum()
            if row.get("트래픽", 0) > 0:
                row["CR"] = row.get("구매객수", 0) / row["트래픽"] * 100
            else:
                row["CR"] = 0
            if row.get("구매객수", 0) > 0:
                row["객단가"] = row.get("거래액", 0) / row["구매객수"]
            else:
                row["객단가"] = 0
            total_rows.append(row)

    total_df = pd.DataFrame(total_rows)
    ep_traffic_result = pd.concat([pivot, total_df], ignore_index=True)
    ep_traffic_result = ep_traffic_result.sort_values(["BPU", "회원구분", "날짜"]).reset_index(drop=True)

    # ── 2. ep_category (카테고리/브랜드 전체 조합, 세그먼트=전체만, groupby로 빠르게) ──
    combos = df.iloc[1:, [5, 6]].drop_duplicates().values.tolist()
    ep_category_result = None
    if len(combos) > 1:  # 카테고리 breakdown이 실제로 있는 파일일 때만
        cat_mask = (df.iloc[:, 2] == "전체") & (df.iloc[:, 3] == "전체") & (df.iloc[:, 4] == "전체") \
                   & df.iloc[:, 0].notna()
        sub_df = df[cat_mask].copy()
        sub_df["_bpu"] = bpu_col[cat_mask]

        date_col_indices = [idx for _, idx in date_list]
        date_strs = [d for d, _ in date_list]

        cat_rows = []
        for metric in ["트래픽", "거래액", "구매객수", "CR", "객단가"]:
            metric_rows = sub_df[sub_df.iloc[:, 1] == metric]
            for _, r in metric_rows.iterrows():
                bpu_val = r["_bpu"]
                cat_val = r.iloc[5]
                brand_val = r.iloc[6]
                for date_str, col_idx in date_list:
                    v = str(r.iloc[col_idx]).replace(",", "").replace("%", "")
                    try:
                        v = float(v)
                    except:
                        v = None
                    cat_rows.append({
                        "날짜": date_str, "BPU": bpu_val, "카테고리": cat_val,
                        "브랜드": brand_val, "지표": metric, "값": v,
                    })
        cat_long = pd.DataFrame(cat_rows)
        cat_pivot = cat_long.pivot_table(
            index=["날짜", "BPU", "카테고리", "브랜드"], columns="지표", values="값", aggfunc="first"
        ).reset_index()
        cat_pivot.columns.name = None
        ep_category_result = cat_pivot.sort_values(["BPU", "카테고리", "브랜드", "날짜"]).reset_index(drop=True)

    return ep_traffic_result, ep_category_result


def _pivot_traffic(rows):
    """rows 리스트를 피벗해서 정리된 DataFrame 반환."""
    long = pd.DataFrame(rows)
    if long.empty:
        return long
    pivot = long.pivot_table(index=["날짜", "BPU", "회원구분"],
                             columns="지표", values="값", aggfunc="first").reset_index()
    pivot.columns.name = None
    return pivot.sort_values(["BPU", "회원구분", "날짜"]).reset_index(drop=True)


# ─── UI ───
st.markdown("## 🔄 EP 데이터 변환기")
st.markdown("사내에서 받은 원본 파일을 대시보드용 CSV로 변환합니다.")

st.markdown(
    "<div style='background:#f0f4ff;border-radius:8px;padding:12px 16px;margin:8px 0 16px;font-size:0.88rem;'>"
    "📁 <b>EP채널 데이터</b> (Data.xlsx / Data.csv) → <code>ep_data_long.csv</code><br/>"
    "📁 <b>EP실적 데이터</b> (1_EP실적.csv) → <code>ep_traffic.csv</code><br/>"
    "📁 <b>EP실적(카테고리 포함)</b> 파일이면 → <code>ep_category.csv</code>도 함께 생성<br/>"
    "파일을 올리면 자동으로 종류를 판별합니다."
    "</div>",
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("사내 원본 파일을 올려주세요", type=["csv", "xlsx", "xls"])

if uploaded is not None:
    file_type = detect_file_type(uploaded)

    if file_type == "unknown":
        st.error("파일 종류를 판별할 수 없습니다. EP채널 또는 EP실적 파일인지 확인해주세요.")
        st.stop()

    type_label = {"ep_channel": "EP채널 데이터", "ep_traffic": "EP실적 데이터"}[file_type]
    out_name = {"ep_channel": "ep_data_long.csv", "ep_traffic": "ep_traffic.csv"}[file_type]

    st.info(f"🔍 **{type_label}**로 판별됨 → `{out_name}` 생성")

    with st.spinner("변환 중... (파일 크기에 따라 최대 1~2분 소요될 수 있어요)"):
        try:
            if file_type == "ep_channel":
                result = convert_ep_channel(uploaded, uploaded.name)
                category_result = None
            else:
                result, category_result = convert_ep_traffic(uploaded)
        except Exception as e:
            st.error(f"변환 실패: {e}")
            st.stop()

    date_min = result["날짜"].min()
    date_max = result["날짜"].max()
    n_days = result["날짜"].nunique()

    st.success("변환 완료!")
    c1, c2, c3 = st.columns(3)
    c1.metric("기간", f"{date_min} ~ {date_max}")
    c2.metric("일수", f"{n_days}일")
    c3.metric("행수", f"{len(result):,}")

    # 검증
    if file_type == "ep_channel":
        total = result[(result["BPU"] == "Total") & (result["원부매칭여부"] == "Total") & (result["최저가여부"] == "Total")]
        if not total.empty:
            pct_val = total["원부매칭율(%)"].iloc[-1]
            if pct_val > 200:
                st.warning(f"⚠️ 원부매칭율이 {pct_val:.0f}%로 비정상적입니다.")
            else:
                st.caption(f"✅ 원부매칭율 {pct_val:.1f}% — 정상")
    else:
        total = result[(result["BPU"] == "Total") & (result["회원구분"] == "전체")]
        if not total.empty:
            last_uv = total["트래픽"].iloc[-1]
            last_gmv = total["거래액"].iloc[-1]
            st.caption(f"✅ 최신 Total — UV: {last_uv:,.0f} / 거래액: {last_gmv:,.0f}")

    csv_data = result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        f"⬇️ {out_name} 다운로드",
        csv_data, file_name=out_name, mime="text/csv",
        use_container_width=True, type="primary",
    )

    # 카테고리/브랜드 데이터가 함께 추출됐으면 별도 다운로드 제공
    if category_result is not None and not category_result.empty:
        st.divider()
        n_cats = category_result["카테고리"].nunique()
        n_brands = category_result["브랜드"].nunique()
        st.info(f"🗂️ 카테고리/브랜드 상세 데이터도 함께 발견됐어요 (카테고리 {n_cats}개, 브랜드 {n_brands}개) → `ep_category.csv` 생성")
        cat_csv = category_result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "⬇️ ep_category.csv 다운로드",
            cat_csv, file_name="ep_category.csv", mime="text/csv",
            use_container_width=True,
        )

    st.divider()
    st.markdown(f"**다운로드한 파일을 GitHub에 덮어쓰면 대시보드가 갱신됩니다.**")
    with st.expander("미리보기 (처음 10행)"):
        st.dataframe(result.head(10), use_container_width=True, hide_index=True)
