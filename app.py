"""EP 데이터 변환기 — 사내 원본 파일을 대시보드용 CSV로 변환합니다.

지원 파일:
1. EP채널 데이터 (Data.xlsx / Data.csv) → ep_data_long.csv
2. EP실적 데이터 (1_EP실적.csv) → ep_traffic.csv
3. EP실적(카테고리 포함) 데이터 → ep_category.csv
4. 쿠폰 일자별 원본 (플러스/일반) → ep_coupon_daily.csv
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
    """파일 내용을 보고 EP채널 / EP실적 자동 판별.
    예전엔 '엑셀이면 무조건 EP채널'이었는데, EP실적/EP상세실적도 엑셀로 올 수 있어서
    (예: 소수점이 CSV에서 잘리는 문제 때문에 엑셀 원본을 그대로 쓰고 싶은 경우) 잘못
    판정되고 있었음 — 이제 엑셀도 CSV와 똑같이 내용을 읽어서 판별한다."""
    name = uploaded_file.name.lower()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        try:
            preview = pd.read_excel(uploaded_file, sheet_name=0, header=None, nrows=5)
        except Exception:
            return "unknown"
        finally:
            uploaded_file.seek(0)
        text = " ".join(str(v) for v in preview.values.flatten() if pd.notna(v))
    else:
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
    name = getattr(uploaded_file, "name", "").lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        # 엑셀 원본 그대로 사용 — CSV로 미리 변환하면 소수점이 잘리는 문제가 있어서,
        # 엑셀을 직접 읽으면 원본 정밀도가 그대로 보존된다.
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    else:
        raw = uploaded_file.read()
        uploaded_file.seek(0)
        if raw[:2] == b"\xff\xfe":
            text = raw.decode("utf-16-le")
            df = pd.read_csv(io.StringIO(text), sep="\t", header=None, low_memory=False)
        else:
            df = pd.read_csv(uploaded_file, encoding="utf-8-sig", sep=None,
                             engine="python", header=None, low_memory=False)

    # --- 구조 자동 판별 ---
    is_xlsx = name.endswith(".xlsx") or name.endswith(".xls")
    header_row0 = str(df.iloc[0, 0]).strip()
    if header_row0 == "BPU" or header_row0 in ("e-영업1", "e-영업2", "e-영업3", "e-영업4"):
        # 새 구조: 열0=BPU, 열1=지표, 열2=회원구분, 열3=신규구분1, 열4=신규구분2, 열5=카테고리, 열6=브랜드, 열7~=날짜
        return _convert_traffic_new(df, is_xlsx)  # (ep_traffic_df, ep_category_df_or_None)
    else:
        # 기존 구조: 열0=지표, 열1=회원구분, 열2=신규구분1, 열3=신규구분2, 열4=구분, 열5=BPU, 열6~=날짜
        return _convert_traffic_old(df, is_xlsx), None


def _convert_traffic_old(df, is_xlsx=False):
    """기존 EP실적 구조 (열0=지표, 열5=BPU, 열4=구분).
    is_xlsx: 엑셀에서 읽은 경우, CR(%) 셀이 엑셀의 '%서식' 때문에 0.033같은 분수로
    읽힌다(CSV/기존 파이프라인은 3.3처럼 이미 100배 된 값). 그대로 두면 대시보드 전체의
    구매전환율이 100배 작게 나오게 되므로, 엑셀 소스일 때만 CR을 ×100 보정한다."""
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
                    if val is not None and metric == "CR" and is_xlsx:
                        val = val * 100
                    rows.append({"날짜": date_str, "BPU": bpu_label,
                                 "회원구분": seg_label, "지표": metric, "값": val})
    return _pivot_traffic(rows)


def _convert_traffic_new(df, is_xlsx=False):
    """새 EP실적 구조 (열0=BPU, 열1=지표, 열5=카테고리, 열6=브랜드).
    Total 행이 없으므로 BPU 합산으로 만든다.
    is_xlsx: CR(%) 값이 엑셀 %서식 때문에 분수(0.033)로 읽히는 문제 보정용 (아래 참고).
    반환: (ep_traffic 형태 DataFrame, ep_category 형태 DataFrame or None)
    """
    # 라벨 컬럼(BPU/지표/회원구분/신규구분1/신규구분2/카테고리)이 전부 엑셀 원본에서
    # 병합 셀 형태다 — 각 그룹의 첫 행에만 값이 있고 나머지 행은 비어있다(엑셀에서
    # 시각적으로는 "이어지는 것처럼" 보이지만 실제 셀 값은 빈 칸). 전부 ffill로
    # 앞의 값을 채워야 카테고리/브랜드별 세부 행들이 올바르게 매칭된다 — 이게 안 돼서
    # 카테고리별 상세 데이터 대부분이 누락되는 문제가 있었음. 브랜드(열6)만 매 행마다
    # 자기 값이 있어서 ffill이 필요 없다(확인함).
    bpu_col = df.iloc[:, 0].ffill()
    metric_col = df.iloc[:, 1].ffill()
    member_col = df.iloc[:, 2].ffill()
    sinew1_col = df.iloc[:, 3].ffill()
    sinew2_col = df.iloc[:, 4].ffill()
    category_col = df.iloc[:, 5].ffill()

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
                mask = ((bpu_col == bpu) & (metric_col == metric) &
                        (member_col == member_filter) & (sinew1_col == sinew_filter) &
                        (sinew2_col == "전체") & (category_col == "전체") & (df.iloc[:, 6] == "전체"))
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
                    if val is not None and metric == "CR" and is_xlsx:
                        val = val * 100
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

    # ── 2. ep_category (카테고리/브랜드 전체 조합) ──
    # 5개 세그먼트를 다 넣으면 카테고리×브랜드 조합 수 때문에 파일이 너무 커져서(25MB+,
    # 깃허브 웹 업로드 제한 초과) 전체/회원/신규 3개만 담는다. 비회원=전체-회원,
    # 기존=전체-신규 로 나중에 필요하면 역산 가능하고, 필요하면 원본에서 다시 뽑을 수도 있음.
    # (ep_traffic.csv는 카테고리/브랜드로 곱해지지 않아 용량이 작으므로 5개 세그먼트 그대로 유지)
    CATEGORY_SEGMENTS = [s for s in SEGMENTS if s[0] in ("전체", "회원", "신규")]
    combos = pd.DataFrame({"카테고리": category_col.iloc[1:], "브랜드": df.iloc[1:, 6]}).drop_duplicates().values.tolist()
    ep_category_result = None
    if len(combos) > 1:  # 카테고리 breakdown이 실제로 있는 파일일 때만
        date_strs = [d for d, _ in date_list]
        date_col_positions = [idx for _, idx in date_list]

        cat_frames = []
        for seg_label, member_filter, sinew_filter in CATEGORY_SEGMENTS:
            seg_mask = (
                (member_col == member_filter) & (sinew1_col == sinew_filter)
                & (sinew2_col == "전체")
            )
            sub_df = df[seg_mask]
            if sub_df.empty:
                continue
            melt_src = pd.DataFrame({
                "BPU": bpu_col[seg_mask].values,
                "지표": metric_col[seg_mask].values,
                "카테고리": category_col[seg_mask].values,
                "브랜드": sub_df.iloc[:, 6].values,
            })
            for date_str, col_idx in zip(date_strs, date_col_positions):
                melt_src[date_str] = sub_df.iloc[:, col_idx].values
            long_df = melt_src.melt(
                id_vars=["BPU", "지표", "카테고리", "브랜드"], var_name="날짜", value_name="값"
            )
            long_df["값"] = pd.to_numeric(
                long_df["값"].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )
            if is_xlsx:
                _cr_mask = long_df["지표"] == "CR"
                long_df.loc[_cr_mask, "값"] = long_df.loc[_cr_mask, "값"] * 100
            long_df["회원구분"] = seg_label
            cat_frames.append(long_df)

        if cat_frames:
            cat_long = pd.concat(cat_frames, ignore_index=True)
            cat_pivot = cat_long.pivot_table(
                index=["날짜", "BPU", "카테고리", "브랜드", "회원구분"], columns="지표", values="값", aggfunc="first"
            ).reset_index()
            cat_pivot.columns.name = None
            ep_category_result = cat_pivot.sort_values(
                ["BPU", "카테고리", "브랜드", "회원구분", "날짜"]
            ).reset_index(drop=True)

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


def merge_other_dept(traffic_df, other_file):
    """기타부서 원본(xlsx)을 읽어 Total(회원구분=전체) 행의 트래픽/거래액에 더하고
    CR/객단가를 재계산한다. 구매객수는 원본에 없으므로 그대로 둔다."""
    odf = pd.read_excel(other_file, sheet_name=0)
    date_cols = [c for c in odf.columns if str(c).startswith("20")]
    if not date_cols:
        raise ValueError("날짜 컬럼을 찾지 못했습니다.")
    metric_col = odf.columns[0]

    other_daily = {}
    for metric in ["트래픽", "거래액"]:
        match = odf[odf[metric_col] == metric]
        if match.empty:
            continue
        row = match.iloc[0]
        for c in date_cols:
            v = row[c]
            try:
                v = float(v)
            except (TypeError, ValueError):
                v = 0.0
            other_daily.setdefault(pd.Timestamp(c).strftime("%Y-%m-%d"), {})[metric] = v

    merged = traffic_df.copy()
    mask_total = (merged["BPU"] == "Total") & (merged["회원구분"] == "전체")
    n_adjusted = 0
    for idx in merged[mask_total].index:
        date_str = merged.at[idx, "날짜"]
        add = other_daily.get(str(date_str))
        if not add:
            continue
        add_uv = add.get("트래픽", 0.0)
        add_gmv = add.get("거래액", 0.0)
        if add_uv == 0 and add_gmv == 0:
            continue
        merged.at[idx, "트래픽"] = merged.at[idx, "트래픽"] + add_uv
        merged.at[idx, "거래액"] = merged.at[idx, "거래액"] + add_gmv
        # CR/객단가 재계산 (구매객수는 기타부서 데이터가 없어 그대로 유지)
        buyers = merged.at[idx, "구매객수"]
        new_uv = merged.at[idx, "트래픽"]
        new_gmv = merged.at[idx, "거래액"]
        merged.at[idx, "CR"] = (buyers / new_uv * 100) if new_uv > 0 else 0
        merged.at[idx, "객단가"] = (new_gmv / buyers) if buyers > 0 else 0
        n_adjusted += 1

    return merged, n_adjusted


def parse_coupon_daily_wide(uploaded_file, coupon_type, extra_dims):
    """'결제_일자' 일자별 와이드 포맷(거래액/쿠폰할인 두 블록, 월별 컬럼)을
    날짜 단위 long 형태로 변환. extra_dims의 마지막 항목은 반드시 '결제_일자'.
    """
    raw = uploaded_file.read()
    text = raw.decode("utf-16-le")
    df = pd.read_csv(io.StringIO(text), sep="\t", header=None, low_memory=False)

    n_meta = len(extra_dims)
    header_row = df.iloc[2]
    date_cols_idx = [i for i in range(n_meta + 2, len(df.columns)) if str(header_row.iloc[i]).isdigit()]
    n_months = len(date_cols_idx) // 2
    gmv_cols = date_cols_idx[:n_months]
    coupon_cols = date_cols_idx[n_months:]
    month_labels = [str(header_row.iloc[i]) for i in gmv_cols]

    data = df.iloc[3:].reset_index(drop=True)
    date_col_idx = n_meta - 1  # 결제_일자는 extra_dims의 마지막 컬럼

    rows = []
    for _, r in data.iterrows():
        day_str = str(r.iloc[date_col_idx])
        if day_str == "총계" or day_str == "nan":
            continue  # 총계행은 날짜별이 아니라 전체기간 합계라 건너뜀
        day_month = day_str[:6]
        if day_month not in month_labels:
            continue
        mi = month_labels.index(day_month)
        gmv_raw = str(r.iloc[gmv_cols[mi]]).replace(",", "")
        coupon_raw = str(r.iloc[coupon_cols[mi]]).replace(",", "")
        try:
            gmv = float(gmv_raw)
        except (TypeError, ValueError):
            gmv = 0.0
        try:
            coupon = float(coupon_raw)
        except (TypeError, ValueError):
            coupon = 0.0
        if gmv == 0 and coupon == 0:
            continue
        meta = {dim: r.iloc[i] for i, dim in enumerate(extra_dims[:-1])}
        row = dict(meta)
        row["날짜"] = pd.to_datetime(day_str, format="%Y%m%d")
        row["자체거래액"] = gmv
        row["쿠폰할인"] = coupon
        row["쿠폰유형"] = coupon_type
        rows.append(row)
    return pd.DataFrame(rows)


def _plus_bpu_final(row):
    """플러스 쿠폰 파일의 '자사입점구분'+'BPU' 2개 컬럼을 하나의 BPU로 통합.
    (총계,총계)->Total, (자사/입점,총계)->자사/입점 소계, (자사/입점,e-영업N)->e-영업N"""
    if row["자사입점구분"] == "총계" and row["BPU"] == "총계":
        return "Total"
    if row["BPU"] == "총계":
        return row["자사입점구분"]
    return row["BPU"]


def build_coupon_daily(plus_file, general_file):
    """일자별 쿠폰 원본(플러스+일반) → ep_coupon_daily.csv 형태.
    e-영업1~4 세부 BPU 레벨만 남긴다 (Total/자사/입점 소계는 대시보드에서
    BPU_GROUPS로 즉석 합산하므로 여기서는 만들지 않음)."""
    plus_daily = parse_coupon_daily_wide(
        plus_file, "플러스", ["AF대분류명", "자사입점구분", "BPU", "쿠폰ID", "쿠폰명", "결제_일자"]
    )
    general_daily = parse_coupon_daily_wide(
        general_file, "일반", ["AF대분류명", "BPU", "쿠폰ID", "쿠폰명", "결제_일자"]
    )

    plus_ep = plus_daily[plus_daily["AF대분류명"] == "EP"].copy()
    general_ep = general_daily[general_daily["AF대분류명"] == "EP"].copy()

    plus_ep["BPU_final"] = plus_ep.apply(_plus_bpu_final, axis=1)
    general_ep["BPU_final"] = general_ep["BPU"].replace({"총계": "Total"})

    plus_clean = plus_ep[["날짜", "BPU_final", "쿠폰ID", "쿠폰명", "쿠폰유형", "쿠폰할인"]].rename(columns={"BPU_final": "BPU"})
    general_clean = general_ep[["날짜", "BPU_final", "쿠폰ID", "쿠폰명", "쿠폰유형", "쿠폰할인"]].rename(columns={"BPU_final": "BPU"})

    result = pd.concat([plus_clean, general_clean], ignore_index=True)
    result = result[result["쿠폰할인"].notna() & (result["쿠폰할인"] != 0)]
    return result.sort_values(["날짜", "쿠폰유형", "BPU"]).reset_index(drop=True)


# ─── UI ───
st.markdown("## 🔄 EP 데이터 변환기")
st.markdown("사내에서 받은 원본 파일을 대시보드용 CSV로 변환합니다.")

st.markdown(
    "<div style='background:#f0f4ff;border-radius:8px;padding:12px 16px;margin:8px 0 16px;font-size:0.86rem;'>"
    "각 파일을 <b>맞는 슬롯에</b> 올려주세요. 슬롯마다 결과 파일이 정해져 있어서, "
    "카테고리 포함 파일을 잘못된 슬롯에 올려도 섞이지 않아요.<br/><br/>"
    "① <b>EP채널 데이터</b> (Data.xlsx/csv) → <code>ep_data_long.csv</code><br/>"
    "② <b>EP실적 데이터 (카테고리 구분 없는 집계 파일)</b> → <code>ep_traffic.csv</code> "
    "· <span style='color:#dc2626'>이 슬롯엔 반드시 카테고리/브랜드 구분이 없는 순수 집계 파일만 올려주세요</span><br/>"
    "③ <b>EP실적(카테고리·브랜드 포함) 데이터</b> → <code>ep_category.csv</code> "
    "· <span style='color:#dc2626'>이 슬롯에서 나오는 ep_traffic.csv는 사용하지 않습니다 (중복 카운트가 섞여 있어서 부정확함)</span>"
    "</div>",
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4 = st.tabs(["① EP채널", "② EP실적 (집계)", "③ EP실적 (카테고리 포함)", "④ 쿠폰 데이터"])

# ── ① EP채널 ──
with tab1:
    st.caption("Data.xlsx / Data.csv — 원부매칭율/최저가율 등")
    up_channel = st.file_uploader("EP채널 원본 업로드", type=["csv", "xlsx", "xls"], key="up_channel")
    if up_channel is not None:
        detected = detect_file_type(up_channel)
        if detected != "ep_channel":
            st.error("EP채널 데이터로 보이지 않아요. ②나 ③ 탭에 올리려던 파일은 아닌가요?")
        else:
            with st.spinner("변환 중..."):
                try:
                    result = convert_ep_channel(up_channel, up_channel.name)
                except Exception as e:
                    st.error(f"변환 실패: {e}")
                    st.stop()
            date_min, date_max = result["날짜"].min(), result["날짜"].max()
            st.success("변환 완료!")
            c1, c2, c3 = st.columns(3)
            c1.metric("기간", f"{date_min} ~ {date_max}")
            c2.metric("일수", f"{result['날짜'].nunique()}일")
            c3.metric("행수", f"{len(result):,}")

            total = result[(result["BPU"] == "Total") & (result["원부매칭여부"] == "Total") & (result["최저가여부"] == "Total")]
            if not total.empty:
                pct_val = total["원부매칭율(%)"].iloc[-1]
                if pct_val > 200:
                    st.warning(f"⚠️ 원부매칭율이 {pct_val:.0f}%로 비정상적입니다.")
                else:
                    st.caption(f"✅ 원부매칭율 {pct_val:.1f}% — 정상")

            st.download_button(
                "⬇️ ep_data_long.csv 다운로드",
                result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                file_name="ep_data_long.csv", mime="text/csv",
                use_container_width=True, type="primary",
            )
            with st.expander("미리보기 (처음 10행)"):
                st.dataframe(result.head(10), use_container_width=True, hide_index=True)

# ── ② EP실적 (집계, 카테고리 없음) → ep_traffic.csv ──
with tab2:
    st.caption("카테고리/브랜드 구분이 없는, BPU별 순수 집계 파일만 올려주세요.")
    up_traffic = st.file_uploader("EP실적(집계) 원본 업로드", type=["csv", "xlsx", "xls"], key="up_traffic")
    other_dept_file = st.file_uploader(
        "기타부서 파일 (선택 · 있으면 Total에 자동 반영됩니다)",
        type=["xlsx", "xls"], key="other_dept_main",
    )
    if up_traffic is not None:
        detected = detect_file_type(up_traffic)
        if detected != "ep_traffic":
            st.error("EP실적 데이터로 보이지 않아요. ①이나 ③ 탭에 올리려던 파일은 아닌가요?")
        else:
            with st.spinner("변환 중... (파일 크기에 따라 최대 1~2분 소요될 수 있어요)"):
                try:
                    result, category_result = convert_ep_traffic(up_traffic)
                except Exception as e:
                    st.error(f"변환 실패: {e}")
                    st.stop()

            if category_result is not None and not category_result.empty:
                st.warning(
                    "⚠️ 이 파일에 카테고리/브랜드 구분이 있는 것 같아요! "
                    "이 슬롯에서 나온 Total은 카테고리 집계 방식이라 부정확할 수 있어요. "
                    "③ 탭에 이 파일을 올려서 ep_category.csv만 받아 쓰시고, "
                    "이 ②탭에는 순수 집계 파일을 따로 올려주세요."
                )

            other_dept_applied = False
            if other_dept_file is not None:
                try:
                    result, n_adj = merge_other_dept(result, other_dept_file)
                    other_dept_applied = True
                    st.success(f"✅ 기타부서 데이터를 Total에 반영했어요 ({n_adj}일 조정됨)")
                except Exception as e:
                    st.warning(f"기타부서 반영 실패(원본 EP실적 결과는 정상 생성됨): {e}")

            date_min, date_max = result["날짜"].min(), result["날짜"].max()
            st.success("변환 완료!")
            c1, c2, c3 = st.columns(3)
            c1.metric("기간", f"{date_min} ~ {date_max}")
            c2.metric("일수", f"{result['날짜'].nunique()}일")
            c3.metric("행수", f"{len(result):,}")
            if not other_dept_applied:
                st.caption("ℹ️ 기타부서 파일을 위에 같이 올리면, Total(전체)에 자동으로 더해서 반영할 수 있어요 (선택 사항).")

            total = result[(result["BPU"] == "Total") & (result["회원구분"] == "전체")]
            if not total.empty:
                last_uv = total["트래픽"].iloc[-1]
                last_gmv = total["거래액"].iloc[-1]
                st.caption(f"✅ 최신 Total — UV: {last_uv:,.0f} / 거래액: {last_gmv:,.0f}")

            st.download_button(
                "⬇️ ep_traffic.csv 다운로드",
                result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                file_name="ep_traffic.csv", mime="text/csv",
                use_container_width=True, type="primary",
            )
            with st.expander("미리보기 (처음 10행)"):
                st.dataframe(result.head(10), use_container_width=True, hide_index=True)

# ── ③ EP실적 (카테고리·브랜드 포함) → ep_category.csv ──
with tab3:
    st.caption("카테고리/브랜드 구분이 있는 상세 파일을 올려주세요. 여기서는 ep_category.csv만 사용합니다.")
    up_cat = st.file_uploader("EP실적(카테고리 포함) 원본 업로드", type=["csv", "xlsx", "xls"], key="up_cat")
    if up_cat is not None:
        detected = detect_file_type(up_cat)
        if detected != "ep_traffic":
            st.error("EP실적 데이터로 보이지 않아요. ①이나 ② 탭에 올리려던 파일은 아닌가요?")
        else:
            with st.spinner("변환 중... (파일 크기에 따라 최대 1~2분 소요될 수 있어요)"):
                try:
                    _ignored_traffic, category_result = convert_ep_traffic(up_cat)
                except Exception as e:
                    st.error(f"변환 실패: {e}")
                    st.stop()

            if category_result is None or category_result.empty:
                st.error(
                    "이 파일에서 카테고리/브랜드 구분을 찾지 못했어요. "
                    "혹시 ② 탭에 올리려던 순수 집계 파일 아닌가요?"
                )
            else:
                n_cats = category_result["카테고리"].nunique()
                n_brands = category_result["브랜드"].nunique()
                date_min, date_max = category_result["날짜"].min(), category_result["날짜"].max()
                st.success(f"변환 완료! 카테고리 {n_cats}개, 브랜드 {n_brands}개")
                c1, c2, c3 = st.columns(3)
                c1.metric("기간", f"{date_min} ~ {date_max}")
                c2.metric("일수", f"{category_result['날짜'].nunique()}일")
                c3.metric("행수", f"{len(category_result):,}")
                st.caption("ℹ️ 이 슬롯에서 나온 Total(트래픽/거래액)은 카테고리 합산 방식이라, "
                          "①UV처럼 개인이 여러 카테고리를 봤을 때 중복 카운트될 수 있어요. "
                          "EP실적 페이지용 Total은 ② 탭 결과를 사용하세요.")

                st.download_button(
                    "⬇️ ep_category.csv 다운로드",
                    category_result.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                    file_name="ep_category.csv", mime="text/csv",
                    use_container_width=True, type="primary",
                )
                with st.expander("미리보기 (처음 10행)"):
                    st.dataframe(category_result.head(10), use_container_width=True, hide_index=True)

# ── ④ 쿠폰 데이터 → ep_coupon_daily.csv ──
with tab4:
    st.caption(
        "사내에서 받은 쿠폰 일자별 원본 2종(플러스/일반)을 올리면 ep_coupon_daily.csv 하나로 변환돼요. "
        "월별 집계·쿠폰명 상세는 대시보드에서 이 파일 하나로 자동 계산되므로, "
        "따로 ep_coupon.csv나 ep_coupon_detail.csv를 만들 필요는 없어요."
    )

    cc1, cc2 = st.columns(2)
    with cc1:
        coupon_plus_daily_file = st.file_uploader("쿠폰 거래액_플러스 일자별 원본", type=["csv"], key="coupon_plus_daily")
    with cc2:
        coupon_general_daily_file = st.file_uploader("쿠폰 거래액_일반 일자별 원본", type=["csv"], key="coupon_general_daily")

    if coupon_plus_daily_file is not None and coupon_general_daily_file is not None:
        with st.spinner("쿠폰 일자별 데이터 변환 중... (쿠폰명이 많아 시간이 좀 걸릴 수 있어요)"):
            try:
                coupon_daily_result = build_coupon_daily(coupon_plus_daily_file, coupon_general_daily_file)
            except Exception as e:
                st.error(f"쿠폰 일자별 변환 실패: {e}")
                coupon_daily_result = None
        if coupon_daily_result is not None:
            date_min = coupon_daily_result["날짜"].min().strftime("%Y-%m-%d")
            date_max = coupon_daily_result["날짜"].max().strftime("%Y-%m-%d")
            st.success(f"변환 완료! ({date_min} ~ {date_max}, 고유 쿠폰 {coupon_daily_result['쿠폰ID'].nunique()}개)")
            c1, c2, c3 = st.columns(3)
            c1.metric("기간", f"{date_min} ~ {date_max}")
            c2.metric("일수", f"{coupon_daily_result['날짜'].nunique()}일")
            c3.metric("행수", f"{len(coupon_daily_result):,}")
            st.caption(f"BPU: {', '.join(sorted(coupon_daily_result['BPU'].unique()))}")

            _export = coupon_daily_result.copy()
            _export["날짜"] = _export["날짜"].dt.strftime("%Y-%m-%d")
            st.download_button(
                "⬇️ ep_coupon_daily.csv 다운로드",
                _export.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                file_name="ep_coupon_daily.csv", mime="text/csv",
                use_container_width=True, type="primary", key="dl_coupon_daily",
            )
            with st.expander("미리보기 (처음 10행)"):
                st.dataframe(coupon_daily_result.head(10), use_container_width=True, hide_index=True)
    else:
        st.info("위 2개(거래액_플러스 일자별, 거래액_일반 일자별) 파일을 모두 올리면 ep_coupon_daily.csv가 생성돼요.")

st.divider()
st.markdown("**다운로드한 파일을 GitHub에 각각의 파일명으로 덮어쓰면 대시보드가 갱신됩니다.**")




# ============================================================
# 기타부서 보정값 미리보기 (빠르게 월별 영향만 확인하고 싶을 때)
# ============================================================
st.divider()
with st.expander("🔍 기타부서 영향 미리보기만 하기 (선택 사항)"):
    st.markdown(
        "위 메인 업로드에 기타부서 파일을 같이 올리면 **자동으로 Total에 반영**돼요. "
        "이 아래는 EP실적 원본 없이 **기타부서 파일만으로 월별 영향만 빠르게 확인**하고 싶을 때 쓰세요."
    )
    _other_file = st.file_uploader(
        "기타부서 원본 파일 업로드 (xlsx)", type=["xlsx", "xls"], key="other_dept_upload",
    )
    if _other_file is not None:
        try:
            _odf = pd.read_excel(_other_file, sheet_name=0)
            _date_cols = [c for c in _odf.columns if str(c).startswith("20")]
            if not _date_cols:
                st.error("날짜 컬럼(2025-01-01 형태)을 찾지 못했습니다. 파일 구조를 확인해주세요.")
            else:
                _metric_col = _odf.columns[0]
                _available_metrics = _odf[_metric_col].dropna().unique().tolist()
                st.caption(f"발견된 지표: {', '.join(str(m) for m in _available_metrics)}")

                _rows = []
                for _metric in ["트래픽", "거래액"]:
                    _match = _odf[_odf[_metric_col] == _metric]
                    if _match.empty:
                        continue
                    _row = _match.iloc[0]
                    for _c in _date_cols:
                        _v = _row[_c]
                        try:
                            _v = float(_v)
                        except (TypeError, ValueError):
                            _v = 0.0
                        _rows.append({"날짜": pd.Timestamp(_c), "지표": _metric, "값": _v})

                _long = pd.DataFrame(_rows)
                _long["날짜"] = pd.to_datetime(_long["날짜"])
                _pivot = _long.pivot_table(index="날짜", columns="지표", values="값", aggfunc="sum").fillna(0)
                _pivot = _pivot.sort_index()

                st.success(f"기타부서 데이터 인식 완료 · 기간: {_pivot.index.min().date()} ~ {_pivot.index.max().date()}")

                # 월별 합계 요약
                _monthly = _pivot.resample("ME").sum()
                _monthly.index = _monthly.index.strftime("%Y-%m")
                st.markdown("**월별 합계 (Total에 추가로 더해야 할 금액)**")
                st.dataframe(
                    _monthly.rename(columns={"트래픽": "UV", "거래액": "거래액"})
                    .style.format("{:,.0f}"),
                    use_container_width=True,
                )

                st.caption(
                    "⚠️ 구매객수/CR/객단가는 이 원본에 없어 계산할 수 없어요. "
                    "UV·거래액만 Total에 더하면 실제 사내 리포트와 맞아떨어질 거예요."
                )

                _daily_csv = _pivot.reset_index()
                _daily_csv["날짜"] = _daily_csv["날짜"].dt.strftime("%Y-%m-%d")
                st.download_button(
                    "⬇️ 기타부서_일별.csv 다운로드",
                    _daily_csv.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                    file_name="기타부서_일별.csv", mime="text/csv",
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"파일을 읽는 중 문제가 발생했습니다: {e}")
