import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st


# =========================================================
# 기본 설정
# =========================================================
st.set_page_config(
    page_title="입출고·불량 데이터 조회",
    page_icon="📦",
    layout="wide",
)

st.title("입출고·불량 데이터 조회")
st.caption("SKU별 수량을 확인하고 원하는 기준으로 필터링하여 Excel로 다운로드할 수 있습니다.")


# =========================================================
# 파일 읽기
# =========================================================
DEFAULT_FILE = Path("분석raw.xlsx")


@st.cache_data
def load_excel(file_source):
    xls = pd.ExcelFile(file_source)

    required_sheets = ["03_입고", "04_출고", "05_불량관리"]
    missing = [s for s in required_sheets if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"필수 시트가 없습니다: {', '.join(missing)}")

    inbound = pd.read_excel(xls, sheet_name="03_입고")
    outbound = pd.read_excel(xls, sheet_name="04_출고")
    defect = pd.read_excel(xls, sheet_name="05_불량관리")

    return inbound, outbound, defect


# =========================================================
# 공통 정리 함수
# =========================================================
def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_factory(value):
    """
    공장 규칙
    - C2-S는 그대로 C2-S
    - C2 / c2 계열은 C2
    - C5는 C5
    - 그 외 / 공란은 미상
    """
    text = clean_text(value)

    if not text:
        return "미상"

    upper = text.upper().replace(" ", "")

    # C2-S는 가장 먼저 처리해야 C2로 잘못 바뀌지 않음
    if "C2-S" in upper:
        return "C2-S"
    if "C2" in upper:
        return "C2"
    if "C5" in upper:
        return "C5"

    return "미상"


def normalize_defect_type(value):
    """
    불량타입 규칙
    - 전체라는 단어가 있으면 전체
    - 렌즈라는 단어가 있으면 렌즈
    - 테라는 표현/테 단어가 있으면 테
    - 그 외 기타

    '전체'를 먼저 확인하는 이유:
    전체 불량 데이터에 다른 문자가 포함되어도 전체로 우선 분류하기 위함.
    """
    text = clean_text(value)

    if "전체" in text:
        return "전체"
    if "렌즈" in text:
        return "렌즈"

    # 실제 데이터에는 '(테)', '테2' 등의 형태가 있으므로 '테' 포함으로 처리
    if "테" in text:
        return "테"

    return "기타"


def classify_outbound(value):
    """
    출고 전표명으로 매장명/파트를 분류한다.

    우선순위
    1. 오프라인 반품
    2. 일본
    3. 미국
    4. 시딩/협찬/본부장 -> 협찬
    5. 해외파트 -> B2B
    6. 출고요청 -> 첫 번째와 두 번째 '_' 사이 값을 매장명으로 사용
    7. 그 외 -> 전표명 그대로 유지
    """
    text = clean_text(value)

    # 공백을 제거한 비교용 문자열
    compact = re.sub(r"\s+", "", text)

    # 오프라인 반품
    if "오프라인반품" in compact or "오프라인매장반품" in compact:
        return "오프라인 반품", "오프라인 반품"

    # 일본 / 미국
    if "일본" in text:
        return "일본", "일본"

    if "미국" in text:
        return "미국", "미국"

    # 협찬
    if any(word in text for word in ["시딩", "협찬", "본부장"]):
        return "협찬", "협찬"

    # 해외파트
    if "해외파트" in text:
        return "B2B", "B2B"

    # 출고요청 / 출고 요청
    if "출고요청" in compact:
        parts = text.split("_")

        # 일반적인 형태: 날짜_매장명_출고 요청
        if len(parts) >= 3:
            store_name = parts[1].strip()
            if store_name:
                return store_name, "매장"

        # 예외적으로 '_'가 부족한 경우 원문 유지
        return text, "매장"

    # 그 외 전표명은 원문 그대로 유지
    return text if text else "미상", "기타"


def to_excel_bytes(df, sheet_name="데이터"):
    """
    필터링된 데이터를 Excel로 다운로드하기 위한 bytes 생성
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        ws = writer.book[sheet_name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # 열 너비 자동 조정
        for column_cells in ws.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 40)

    output.seek(0)
    return output.getvalue()


def format_number(value):
    try:
        return f"{int(value):,}"
    except Exception:
        return value


# =========================================================
# 데이터 로드
# =========================================================
uploaded_file = st.sidebar.file_uploader(
    "분석raw.xlsx 업로드",
    type=["xlsx"],
    help="GitHub에 올린 기본 파일 대신 다른 Excel 파일을 분석하려면 업로드하세요.",
)

source = uploaded_file if uploaded_file is not None else DEFAULT_FILE

if not DEFAULT_FILE.exists() and uploaded_file is None:
    st.error(
        "분석raw.xlsx 파일을 찾을 수 없습니다. "
        "GitHub 저장소에 app.py와 같은 위치에 분석raw.xlsx를 올리거나 "
        "왼쪽에서 Excel 파일을 업로드하세요."
    )
    st.stop()

try:
    inbound, outbound, defect = load_excel(source)
except Exception as e:
    st.error(f"Excel을 읽는 중 오류가 발생했습니다: {e}")
    st.stop()


# =========================================================
# 데이터 전처리
# =========================================================
# 입고
inbound = inbound.copy()
inbound["상품코드"] = inbound["상품코드"].astype(str).str.strip()
inbound["입고수량"] = pd.to_numeric(inbound["입고수량"], errors="coerce").fillna(0)
inbound["공장"] = inbound["중국공장"].apply(normalize_factory)

# 출고
outbound = outbound.copy()
outbound["상품코드"] = outbound["상품코드"].astype(str).str.strip()
outbound["수량"] = pd.to_numeric(outbound["수량"], errors="coerce").fillna(0)

classified = outbound["전표명"].apply(classify_outbound)
outbound["매장명"] = classified.apply(lambda x: x[0])
outbound["파트"] = classified.apply(lambda x: x[1])

# 불량
defect = defect.copy()
defect["상품코드"] = defect["상품코드"].astype(str).str.strip()
defect["불량수량"] = pd.to_numeric(defect["불량수량"], errors="coerce").fillna(0)
defect["공장"] = defect["중국공장"].apply(normalize_factory)
defect["불량타입"] = defect["불량유형"].apply(normalize_defect_type)


# =========================================================
# 탭
# =========================================================
tab_in, tab_out, tab_def = st.tabs(["📥 입고", "📤 출고", "⚠️ 불량"])


# =========================================================
# 입고
# =========================================================
with tab_in:
    st.subheader("입고")

    c1, c2 = st.columns(2)

    with c1:
        factories = ["전체"] + sorted(inbound["공장"].dropna().unique().tolist())
        selected_factory = st.selectbox("공장", factories, key="in_factory")

    with c2:
        sku_keyword = st.text_input("SKU 검색", key="in_sku", placeholder="상품코드 입력")

    filtered_in = inbound.copy()

    if selected_factory != "전체":
        filtered_in = filtered_in[filtered_in["공장"] == selected_factory]

    if sku_keyword:
        filtered_in = filtered_in[
            filtered_in["상품코드"].str.contains(sku_keyword, case=False, na=False)
        ]

    # SKU별 수량
    summary_in = (
        filtered_in.groupby(["상품코드", "상품명", "공장"], as_index=False)["입고수량"]
        .sum()
        .sort_values("입고수량", ascending=False)
    )

    total_in = summary_in["입고수량"].sum()

    m1, m2 = st.columns(2)
    m1.metric("SKU 수", f"{summary_in['상품코드'].nunique():,}")
    m2.metric("입고수량", format_number(total_in))

    st.dataframe(
        summary_in,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "📥 입고 SKU별 수량 Excel 다운로드",
        data=to_excel_bytes(summary_in, "입고_SKU별"),
        file_name="입고_SKU별_수량.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("원본 입고 데이터 보기"):
        st.dataframe(filtered_in, use_container_width=True, hide_index=True)


# =========================================================
# 출고
# =========================================================
with tab_out:
    st.subheader("출고")

    c1, c2, c3 = st.columns(3)

    with c1:
        parts = ["전체"] + sorted(outbound["파트"].dropna().unique().tolist())
        selected_part = st.selectbox("파트", parts, key="out_part")

    with c2:
        stores = ["전체"] + sorted(outbound["매장명"].dropna().unique().tolist())
        selected_store = st.selectbox("매장명", stores, key="out_store")

    with c3:
        out_sku_keyword = st.text_input(
            "SKU 검색",
            key="out_sku",
            placeholder="상품코드 입력",
        )

    filtered_out = outbound.copy()

    if selected_part != "전체":
        filtered_out = filtered_out[filtered_out["파트"] == selected_part]

    if selected_store != "전체":
        filtered_out = filtered_out[filtered_out["매장명"] == selected_store]

    if out_sku_keyword:
        filtered_out = filtered_out[
            filtered_out["상품코드"].str.contains(
                out_sku_keyword,
                case=False,
                na=False,
            )
        ]

    # 매장/파트별 SKU 수량
    summary_out = (
        filtered_out.groupby(
            ["파트", "매장명", "상품코드", "상품명"],
            as_index=False,
        )["수량"]
        .sum()
        .sort_values("수량", ascending=False)
    )

    total_out = summary_out["수량"].sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("SKU 수", f"{summary_out['상품코드'].nunique():,}")
    m2.metric("출고수량", format_number(total_out))
    m3.metric("매장/파트 조합", f"{summary_out[['파트', '매장명']].drop_duplicates().shape[0]:,}")

    st.dataframe(
        summary_out,
        use_container_width=True,
        hide_index=True,
    )

    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            "📥 출고 SKU별 수량 Excel 다운로드",
            data=to_excel_bytes(summary_out, "출고_SKU별"),
            file_name="출고_SKU별_수량.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with d2:
        # 매장/파트별 SKU 집계
        store_part_summary = (
            filtered_out.groupby(
                ["파트", "매장명", "상품코드", "상품명"],
                as_index=False,
            )["수량"]
            .sum()
            .sort_values(["파트", "매장명", "수량"], ascending=[True, True, False])
        )

        st.download_button(
            "📥 매장·파트별 SKU 수량 Excel 다운로드",
            data=to_excel_bytes(store_part_summary, "매장파트별"),
            file_name="출고_매장·파트별_SKU_수량.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("원본 출고 데이터 보기"):
        st.dataframe(filtered_out, use_container_width=True, hide_index=True)


# =========================================================
# 불량
# =========================================================
with tab_def:
    st.subheader("불량")

    c1, c2, c3 = st.columns(3)

    with c1:
        factories = ["전체"] + sorted(defect["공장"].dropna().unique().tolist())
        selected_def_factory = st.selectbox("공장", factories, key="def_factory")

    with c2:
        defect_types = ["전체"] + sorted(defect["불량타입"].dropna().unique().tolist())
        selected_def_type = st.selectbox("불량타입", defect_types, key="def_type")

    with c3:
        def_sku_keyword = st.text_input(
            "SKU 검색",
            key="def_sku",
            placeholder="상품코드 입력",
        )

    filtered_def = defect.copy()

    if selected_def_factory != "전체":
        filtered_def = filtered_def[
            filtered_def["공장"] == selected_def_factory
        ]

    if selected_def_type != "전체":
        filtered_def = filtered_def[
            filtered_def["불량타입"] == selected_def_type
        ]

    if def_sku_keyword:
        filtered_def = filtered_def[
            filtered_def["상품코드"].str.contains(
                def_sku_keyword,
                case=False,
                na=False,
            )
        ]

    # 공장별 SKU별 불량수량
    summary_def = (
        filtered_def.groupby(
            ["공장", "불량타입", "상품코드", "상품명"],
            as_index=False,
        )["불량수량"]
        .sum()
        .sort_values("불량수량", ascending=False)
    )

    total_def = summary_def["불량수량"].sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("SKU 수", f"{summary_def['상품코드'].nunique():,}")
    m2.metric("불량수량", format_number(total_def))
    m3.metric("공장 수", f"{summary_def['공장'].nunique():,}")

    st.dataframe(
        summary_def,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "📥 공장별·SKU별 불량 Excel 다운로드",
        data=to_excel_bytes(summary_def, "공장SKU별불량"),
        file_name="불량_공장별_SKU별_수량.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("원본 불량 데이터 보기"):
        st.dataframe(filtered_def, use_container_width=True, hide_index=True)


# =========================================================
# 하단 안내
# =========================================================
st.divider()
st.caption(
    "분류 규칙: C2-S → C2-S / C2·c2 → C2 / C5 → C5 / 기타·공란 → 미상. "
    "불량타입은 전체·렌즈·테 외에는 기타로 분류합니다."
)
