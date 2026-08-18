"""
쇼핑검색광고 대시보드 - 데이터 변환기

원본 리포트(엑셀/CSV)를 대시보드용 CSV(tableau_daily.csv, category_daily.csv)로 변환합니다.
모든 원본이 "매번 전체 누적 기간을 다시 추출"하는 형태이므로, 새 원본을 받을 때마다
그냥 다시 변환해서 기존 CSV를 통째로 교체하면 됩니다. (증분 병합 불필요)
"""

import io
import os
import streamlit as st
import pandas as pd

st.set_page_config(page_title="쇼핑검색광고 데이터 변환기", layout="wide")
st.title("🔄 쇼핑검색광고 데이터 변환기")
st.caption("원본 리포트를 업로드하면 대시보드용 CSV를 만들어 다운로드할 수 있습니다.")

# 이 변환기가 대시보드와 같은 리포에서 실행 중이면(로컬 클론 또는 Streamlit Cloud 배포),
# data/tableau_daily.csv가 이미 디스크에 있으므로 매번 다시 업로드하지 않아도 자동으로 찾아 씁니다.
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_AUTO_TABLEAU_PATH = os.path.join(_BASE_DIR, "data", "tableau_daily.csv")

tab1, tab2 = st.tabs(["① 일일리포트[태블로] (tableau_daily.csv)", "② 카테고리별 실적 (category_daily.csv)"])


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def check_missing_dates(dates: pd.Series) -> list:
    full_range = pd.date_range(dates.min(), dates.max(), freq="D")
    missing = sorted(set(full_range) - set(dates))
    return missing


# ════════════════════════════════════════════════════════════════
# TAB 1: 일일리포트[태블로] 일자별 RAW → 01(실적 흐름)/02(전년비교) 페이지의 기준 데이터
# ════════════════════════════════════════════════════════════════
with tab1:
    auto_found = os.path.exists(_AUTO_TABLEAU_PATH)

    st.markdown("""
    **원본 파일**: 일일리포트[태블로] 일자별 RAW 엑셀 (거래액·구매건수 등 전체 큰 흐름 기준)
    맨 아래 "총합계" 같은 합계 행이 있어도 자동으로 제외합니다.

    **고정 구간 + 증분 업데이트**: 특정 날짜까지는 확정된(고정) 실적이라 값이 바뀌지 않는다면,
    고정 기준일 **이전** 데이터는 기존 파일 값을 그대로 유지하고, 기준일 **이후** 데이터만
    새로 올린 원본 값으로 채우거나 덮어씁니다.
    """)

    if auto_found:
        st.success(f"✅ 이 리포에 있는 `data/tableau_daily.csv`를 기존 파일로 자동 인식했습니다. 매번 다시 올리지 않아도 됩니다.")
    else:
        st.info("이 환경에서는 `data/tableau_daily.csv`를 찾지 못했습니다 (로컬에서 리포 폴더 밖에서 실행했거나, 별도 환경일 수 있어요). 아래에 기존 파일을 직접 올려주세요.")

    c1, c2 = st.columns(2)
    with c1:
        daily_file = st.file_uploader("① 원본 엑셀 업로드 (.xlsx)", type=["xlsx"], key="daily_upload")
    with c2:
        existing_file = st.file_uploader(
            "② 기존 tableau_daily.csv 업로드 (자동 인식 안 될 때만, 또는 다른 파일로 덮어쓰고 싶을 때)",
            type=["csv"], key="existing_upload",
        )

    fixed_cutoff = st.date_input(
        "🔒 고정 기준일 (이 날짜까지는 기존 파일 값을 그대로 유지)",
        value=pd.Timestamp("2026-07-31"), key="fixed_cutoff",
    )

    if daily_file is not None:
        try:
            raw = pd.read_excel(daily_file, sheet_name=0)
        except Exception as e:
            st.error(f"엑셀 읽기 실패: {e}")
            st.stop()

        if "기간_일자+요일" not in raw.columns:
            st.error(
                "'기간_일자+요일' 컬럼을 찾을 수 없습니다. 원본 리포트 형식이 바뀌지 않았는지 확인해주세요.\n\n"
                f"현재 컬럼: {list(raw.columns)}"
            )
            st.stop()

        def _parse_date(s):
            date_part = str(s).split(" ")[0]
            try:
                return pd.to_datetime(date_part, format="%y-%m-%d")
            except (ValueError, TypeError):
                return pd.NaT

        raw["date"] = raw["기간_일자+요일"].apply(_parse_date)
        n_dropped = raw["date"].isna().sum()
        raw = raw.dropna(subset=["date"])
        result = raw.sort_values("date").reset_index(drop=True)
        if n_dropped:
            st.caption(f"※ 원본에서 날짜로 해석되지 않는 행 {n_dropped}개(합계 행 등)는 제외했습니다.")

        cutoff_ts = pd.Timestamp(fixed_cutoff)

        # 우선순위: 수동 업로드(②) > 자동 인식(리포의 data/tableau_daily.csv)
        existing = None
        existing_source = None
        if existing_file is not None:
            try:
                existing = pd.read_csv(existing_file, parse_dates=["date"])
                existing_source = "업로드한 파일"
            except Exception as e:
                st.error(f"기존 CSV 읽기 실패: {e}")
                st.stop()
        elif auto_found:
            existing = pd.read_csv(_AUTO_TABLEAU_PATH, parse_dates=["date"])
            existing_source = "리포의 data/tableau_daily.csv (자동 인식)"

        if existing is not None:
            missing_cols = set(existing.columns) - set(result.columns)
            if missing_cols:
                st.warning(f"⚠️ 기존 파일에는 있지만 새 원본엔 없는 컬럼: {sorted(missing_cols)} — 새 원본 기준 컬럼으로 맞춥니다.")

            fixed_part = existing[existing["date"] <= cutoff_ts]
            new_part = result[result["date"] > cutoff_ts]
            # 컬럼을 새 원본 기준으로 통일 (기존 파일에 없는 컬럼은 NaN)
            fixed_part = fixed_part.reindex(columns=result.columns)
            result = pd.concat([fixed_part, new_part], ignore_index=True).sort_values("date").reset_index(drop=True)

            st.success(
                f"병합 완료 (기존 파일 출처: {existing_source}): 고정구간 {len(fixed_part):,}행"
                f"({fixed_part['date'].min().date() if len(fixed_part) else '-'}~{fixed_part['date'].max().date() if len(fixed_part) else '-'}) "
                f"+ 신규구간 {len(new_part):,}행"
                f"({new_part['date'].min().date() if len(new_part) else '-'}~{new_part['date'].max().date() if len(new_part) else '-'}) "
                f"= 총 {len(result):,}행"
            )
        else:
            st.success(f"변환 완료: {len(result):,}행 · {result['date'].min().date()} ~ {result['date'].max().date()}")
            st.caption("※ 기존 파일이 없어 전체를 새로 변환했습니다.")

        missing = check_missing_dates(result["date"])
        if missing:
            st.warning(f"⚠️ 누락된 날짜 {len(missing)}개 발견: {[d.date().isoformat() for d in missing[:10]]}"
                       + (" ..." if len(missing) > 10 else ""))
        else:
            st.caption("✅ 날짜 누락 없음 (연속된 일자 확인 완료)")

        st.dataframe(result.head(10), use_container_width=True)

        st.download_button(
            "📥 tableau_daily.csv 다운로드",
            data=to_csv_bytes(result),
            file_name="tableau_daily.csv",
            mime="text/csv",
            key="dl_daily",
        )
        st.info("다운로드한 파일로 GitHub 리포의 `data/tableau_daily.csv`를 교체(덮어쓰기)하면 됩니다.")


# ════════════════════════════════════════════════════════════════
# TAB 2: 카테고리별 실적 (네이버_쇼검_ep거래액_상세_XXXX.csv)
# ════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("""
    **원본 파일**: `네이버_쇼검_ep거래액_상세_YYYYMMDD.csv` (카테고리별 쇼핑검색광고/EP채널 순결제거래액 상세, UTF-16 탭구분)
    매번 전체 누적 기간이 다시 추출되는 형태이므로, 새로 받은 파일을 그대로 올리면 됩니다.
    """)

    cat_file = st.file_uploader("원본 CSV 업로드 (.csv)", type=["csv"], key="cat_upload")

    if cat_file is not None:
        raw_bytes = cat_file.read()
        text = None
        for enc in ("utf-16", "utf-8-sig", "cp949"):
            try:
                text = raw_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue

        if text is None:
            st.error("파일 인코딩을 인식하지 못했습니다 (utf-16 / utf-8 / cp949 시도 실패). 원본 형식을 확인해주세요.")
            st.stop()

        try:
            raw = pd.read_csv(io.StringIO(text), sep="\t", skiprows=2)
        except Exception as e:
            st.error(f"CSV 파싱 실패: {e}")
            st.stop()

        expected_cols = ["ym", "ymd", "category", "광고_이월", "광고_입점", "광고_정상",
                         "EP_이월", "EP_입점", "EP_정상"]
        if raw.shape[1] != len(expected_cols):
            st.error(
                f"컬럼 개수가 예상과 다릅니다 (예상 {len(expected_cols)}개, 실제 {raw.shape[1]}개). "
                "원본 리포트 형식이 바뀌지 않았는지 확인해주세요."
            )
            st.stop()

        raw.columns = expected_cols

        for c in ["광고_이월", "광고_입점", "광고_정상", "EP_이월", "EP_입점", "EP_정상"]:
            raw[c] = raw[c].astype(str).str.replace(",", "").str.replace("nan", "0")
            raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0)

        raw["date"] = pd.to_datetime(raw["ymd"], format="%Y%m%d")
        result = raw[["date", "category"] + [c for c in expected_cols if c not in ("ym", "ymd", "category")]]
        result = result.sort_values(["date", "category"]).reset_index(drop=True)

        st.success(
            f"변환 완료: {len(result):,}행 · {result['date'].min().date()} ~ {result['date'].max().date()} "
            f"· 카테고리 {result['category'].nunique()}개"
        )
        st.caption(f"카테고리 목록: {', '.join(sorted(result['category'].unique()))}")

        unique_dates = result["date"].drop_duplicates()
        missing = check_missing_dates(unique_dates)
        if missing:
            st.warning(f"⚠️ 누락된 날짜 {len(missing)}개 발견: {[d.date().isoformat() for d in missing[:10]]}"
                       + (" ..." if len(missing) > 10 else ""))
        else:
            st.caption("✅ 날짜 누락 없음 (연속된 일자 확인 완료)")

        st.dataframe(result.head(10), use_container_width=True)

        st.download_button(
            "📥 category_daily.csv 다운로드",
            data=to_csv_bytes(result),
            file_name="category_daily.csv",
            mime="text/csv",
            key="dl_category",
        )
        st.info("다운로드한 파일로 GitHub 리포의 `data/category_daily.csv`를 교체(덮어쓰기)하면 됩니다.")
