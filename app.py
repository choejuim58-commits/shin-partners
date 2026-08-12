import io
import pandas as pd
import requests
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# ---------------------------------------------------------
# 📌 구글 시트 커넥션 설정
# ---------------------------------------------------------
conn_gs = st.connection("gsheets", type=GSheetsConnection)

def get_gs_data():
    """구글 시트 최신 데이터 읽기"""
    df = conn_gs.read(ttl=0)
    return df.fillna("")

def save_gs_data(df):
    """구글 시트에 데이터 덮어쓰기 저장"""
    conn_gs.update(data=df)
    st.cache_data.clear()

# ---------------------------------------------------------
# 1. 원드라이브 설정 & 변환 함수
# ---------------------------------------------------------
URL_PANEL = "https://1drv.ms/x/c/e159fcf2d96e4b60/IQD9uLWkE5c7QYgaIwIL_YOgAfq_SOnucQ6LYqj58PDDqPg?e=HP4KZH"
URL_BUILDING = "https://1drv.ms/x/c/e159fcf2d96e4b60/IQBMeXt0XDmHT4f-A042xKnQAVl576fJOGz2wN8KgN3hd5Q?e=6X3Y1o"

@st.cache_data(ttl=43200)  # 12시간 캐시
def load_onedrive_excel(url, category_name, target_sheet="2026.8"):
    try:
        download_url = url + "&download=1" if "?" in url else url + "?download=1"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(download_url, headers=headers, allow_redirects=True)
        response.raise_for_status()
        
        excel_data = io.BytesIO(response.content)
        
        raw_df = pd.read_excel(excel_data, engine='openpyxl', sheet_name=target_sheet, header=None)
        
        header_idx = None
        for idx, row in raw_df.iterrows():
            row_str = row.astype(str).str.replace(" ", "")
            if row_str.str.contains("제품|품명|규격").any():
                header_idx = idx
                break
        
        if header_idx is not None:
            df = raw_df.iloc[header_idx + 1:].copy()
            df.columns = raw_df.iloc[header_idx].values
        else:
            df = raw_df.copy()

        df = df.dropna(how='all')
        
        df.columns = [str(col).strip() if pd.notna(col) else "이름없음" for col in df.columns]
        
        rename_dict = {}
        for col in df.columns:
            cleaned_col = col.replace(" ", "")
            if "밴들당" in cleaned_col or "밴들수량" in cleaned_col:
                rename_dict[col] = "밴들당 수량"
            elif cleaned_col == "밴들":
                rename_dict[col] = "밴들수"
        df = df.rename(columns=rename_dict)

        df.insert(0, "분류", category_name)
        
        if "제품" in df.columns:
            df["제품"] = df["제품"].ffill()
        if "입고 날짜" in df.columns:
            df["입고 날짜"] = df["입고 날짜"].ffill()
            
        num_cols = ["밴들당 수량", "밴들수", "낱매", "현재고", "이월재고", "수입입고", "매입입고", "수정/불량", "단가"]
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        if all(k in df.columns for k in ["밴들당 수량", "밴들수", "낱매"]):
            df["현재고"] = (df["밴들당 수량"] * df["밴들수"]) + df["낱매"]

        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
        df = df.loc[:, df.columns != "이름없음"]
        df = df.fillna("") 
        
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%Y-%m-%d').fillna("")

        return df

    except Exception as e:
        st.error(f"[{category_name}] 원드라이브 읽기 실패: {e}")
        return pd.DataFrame()

# ---------------------------------------------------------
# 2. Streamlit 설정 및 메인 타이틀
# ---------------------------------------------------------
st.set_page_config(
    page_title="신앤파트너스 단가 & 원드라이브 수량 관리", page_icon="📦", layout="wide"
)
st.title("📦 수입상 vs 합판상 단가 & 원드라이브 수량 통합 시스템")

tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "🌐 원드라이브 실시간 수량",
    "🔍 단가 검색 및 비교", 
    "📁 엑셀 일괄 등록", 
    "✏️ 개별 수동 등록", 
    "🛠️ 데이터 직접 수정/삭제"
])

# 🎨 엑셀 서식 스타일 적용 함수
def apply_excel_style(styler):
    num_cols = ["밴들당 수량", "밴들수", "낱매", "현재고", "이월재고", "수입입고", "매입입고", "수정/불량", "단가"]
    valid_num_cols = [col for col in num_cols if col in styler.data.columns]
    
    styler = styler.format("{:.0f}", subset=valid_num_cols)
    
    return styler.map(
        lambda x: 'color: #000000; font-weight: bold;', subset=['밴들수'] if '밴들수' in styler.data.columns else []
    ).map(
        lambda x: 'color: #1e7e34;', subset=['낱매'] if '낱매' in styler.data.columns else []
    ).map(
        lambda x: 'color: #007bff;', subset=['현재고'] if '현재고' in styler.data.columns else []
    ).map(
        lambda x: 'color: #e83e8c; font-weight: bold; font-style: italic;', subset=['이월재고'] if '이월재고' in styler.data.columns else []
    )

# --- [탭 0] 원드라이브 실시간 수량 ---
with tab0:
    col_title, col_btn = st.columns([4, 1])
    with col_title:
        st.subheader("🌐 원드라이브(OneDrive) 수량 현황")
        st.caption("💡 데이터는 빠른 속도를 위해 메모리에 저장됩니다. 아침 정산 후 최신화가 필요할 때 오른쪽 버튼을 누르세요.")
    with col_btn:
        st.write("")
        if st.button("🔄 원드라이브 데이터 최신화", type="primary"):
            st.cache_data.clear()
            st.rerun()

    with st.spinner("원드라이브 8월 데이터 로딩 중..."):
        df_panel = load_onedrive_excel(URL_PANEL, "판재류", target_sheet="2026.8")
        df_building = load_onedrive_excel(URL_BUILDING, "건축자재", target_sheet="2026.8")
        df_onedrive_all = pd.concat([df_panel, df_building], ignore_index=True).fillna("")

    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["전체 보기", "🪵 판재류", "🏗️ 건축자재"])
    
    with sub_tab1:
        search_od = st.text_input("원드라이브 자재 검색 (제품명, 규격, 회사 등)", key="search_od")
        if not df_onedrive_all.empty:
            df_display = df_onedrive_all.copy()
            if search_od:
                mask = df_display.astype(str).apply(lambda row: row.str.contains(search_od, case=False).any(), axis=1)
                df_display = df_display[mask]
            
            styled_df = df_display.style.pipe(apply_excel_style)
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
        else:
            st.info("원드라이브 데이터가 없거나 불러오는 데 실패했습니다.")

    with sub_tab2:
        if not df_panel.empty:
            styled_panel = df_panel.style.pipe(apply_excel_style)
            st.dataframe(styled_panel, use_container_width=True, hide_index=True)
        else:
            st.info("판재류 데이터가 없습니다.")

    with sub_tab3:
        if not df_building.empty:
            styled_building = df_building.style.pipe(apply_excel_style)
            st.dataframe(styled_building, use_container_width=True, hide_index=True)
        else:
            st.info("건축자재 데이터가 없습니다.")

# --- [탭 1] 단가 검색 및 가격 비교 ---
with tab1:
    st.subheader("🔎 자재 단가 조회 & 비교 (구글 시트 기반)")

    col1, col2 = st.columns([1, 2])
    with col1:
        category_filter = st.selectbox(
            "조회할 단가표 선택",
            ["전체 (비교해서 보기)", "수입상 단가만", "합판상 단가만"],
        )
    with col2:
        search_kw = st.text_input(
            "자재명 또는 규격 검색 (예: 멀바우, MDF, OSB, 1220x2440)",
            key="search_kw",
        )

    df_gs = get_gs_data()

    if not df_gs.empty:
        # 필터링 적용
        if category_filter == "수입상 단가만":
            df_gs = df_gs[df_gs["category"] == "수입상"]
        elif category_filter == "합판상 단가만":
            df_gs = df_gs[df_gs["category"] == "합판상"]

        if search_kw:
            mask = (
                df_gs["name"].astype(str).str.contains(search_kw, case=False) |
                df_gs["item_type"].astype(str).str.contains(search_kw, case=False)
            )
            df_gs = df_gs[mask]

        if not df_gs.empty:
            df_show = df_gs.copy()
            if "price" in df_show.columns:
                df_show["price"] = pd.to_numeric(df_show["price"], errors="coerce").fillna(0).astype(int)
                df_show["price"] = df_show["price"].apply(lambda x: f"{x:,} 원")
            
            rename_map = {"category": "구분", "item_type": "품목군", "name": "규격/자재명", "price": "단가", "remark": "비고"}
            df_show = df_show.rename(columns=rename_map)
            st.dataframe(df_show, use_container_width=True, hide_index=True)
        else:
            st.info("검색된 데이터가 없습니다.")
    else:
        st.info("구글 시트에 등록된 데이터가 없습니다.")

# --- [탭 2] 엑셀 일괄 등록 ---
with tab2:
    st.subheader("📁 엑셀(.xlsx) 파일 업로드")

    target_category = st.radio(
        "어떤 단가표 데이터인가요?",
        ["수입상", "합판상"],
        horizontal=True,
    )
    uploaded_file = st.file_uploader(
        "엑셀 파일을 등록해 주세요", type=["xlsx", "csv"]
    )

    st.markdown("""
    > **💡 엑셀 작성 양식 (첫번째 줄 제목)**  
    > `품목군` | `규격` | `단가` | `비고`
    """)

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df_up = pd.read_csv(uploaded_file)
            else:
                df_up = pd.read_excel(uploaded_file)

            st.write("▼ 업로드 파일 미리보기")
            st.dataframe(df_up.head(5))

            if st.button(f"💾 [{target_category}] 구글 시트에 저장"):
                df_current = get_gs_data()
                
                new_rows = []
                for _, row in df_up.iterrows():
                    new_rows.append({
                        "category": target_category,
                        "item_type": str(row.get("품목군", "")),
                        "name": str(row.get("규격", "")),
                        "price": int(pd.to_numeric(row.get("단가", 0), errors='coerce') or 0),
                        "remark": str(row.get("비고", ""))
                    })
                
                df_new = pd.DataFrame(new_rows)
                df_updated = pd.concat([df_current, df_new], ignore_index=True)
                
                save_gs_data(df_updated)
                st.success(f"총 {len(new_rows)}개의 [{target_category}] 단가가 성공적으로 구글 시트에 등록되었습니다!")
                st.rerun()

        except Exception as e:
            st.error(f"등록 실패: 엑셀 파일 열 이름을 확인해 주세요. ({e})")

# --- [탭 3] 개별 수동 등록 ---
with tab3:
    st.subheader("✏️ 자재 하나씩 직접 입력")
    with st.form("manual_form"):
        f_cat = st.selectbox("단가 구분", ["수입상", "합판상"])
        f_type = st.text_input("품목군 (예: 멀바우 집성목, OSB, MDF)")
        f_name = st.text_input("규격 (예: 12x910x2400)")
        f_price = st.number_input("단가 (원)", min_value=0, step=1000)
        f_remark = st.text_input("비고 (예: ANEKA, 3*8제품 등)")
        submitted = st.form_submit_button("저장하기")

    if submitted:
        if f_name and f_price > 0:
            df_current = get_gs_data()
            
            new_data = pd.DataFrame([{
                "category": f_cat,
                "item_type": f_type,
                "name": f_name,
                "price": int(f_price),
                "remark": f_remark
            }])
            
            df_updated = pd.concat([df_current, new_data], ignore_index=True)
            save_gs_data(df_updated)
            
            st.success(f"[{f_cat}] '{f_name}' 단가가 구글 시트에 성공적으로 저장되었습니다!")
            st.rerun()
        else:
            st.error("규격과 단가를 꼭 입력해 주세요.")

# --- [탭 4] 데이터 직접 수정 및 삭제 ---
with tab4:
    st.subheader("🛠️ 구글 시트 데이터 직접 수정 & 삭제")
    st.caption("💡 셀을 더블클릭하여 금액/텍스트를 수정하거나 행을 선택하여 삭제할 수 있습니다.")

    edit_cat = st.selectbox(
        "수정/관리할 단가표 선택",
        ["전체 데이터", "수입상", "합판상"],
        key="edit_cat_select"
    )

    df_edit = get_gs_data()

    if not df_edit.empty:
        if edit_cat == "수입상":
            df_edit_filtered = df_edit[df_edit["category"] == "수입상"]
        elif edit_cat == "합판상":
            df_edit_filtered = df_edit[df_edit["category"] == "합판상"]
        else:
            df_edit_filtered = df_edit

        edited_df = st.data_editor(
            df_edit_filtered,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "category": st.column_config.SelectboxColumn("구분", options=["수입상", "합판상"], required=True),
                "item_type": st.column_config.TextColumn("품목군"),
                "name": st.column_config.TextColumn("규격/자재명", required=True),
                "price": st.column_config.NumberColumn("단가 (원)", format="%d 원", step=500, required=True),
                "remark": st.column_config.TextColumn("비고"),
            },
            hide_index=True,
            key="gs_editor"
        )

        btn_col1, btn_col2 = st.columns([1, 4])

        with btn_col1:
            if st.button("💾 구글 시트에 수정사항 반영하기", type="primary"):
                # 전체 데이터를 관리하기 위해 선택된 카테고리 외 기존 데이터와 병합 후 저장
                if edit_cat == "전체 데이터":
                    final_df = edited_df
                else:
                    other_df = df_edit[df_edit["category"] != edit_cat]
                    final_df = pd.concat([other_df, edited_df], ignore_index=True)

                save_gs_data(final_df)
                st.success("✅ 구글 시트에 수정사항이 성공적으로 반영되었습니다!")
                st.rerun()

        with btn_col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 현재 수정본 엑셀로 내려받기",
                data=buffer.getvalue(),
                file_name="단가표_최신수정본.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("수정/관리할 데이터가 없습니다. 먼저 [📁 엑셀 일괄 등록] 탭에서 데이터를 등록해 주세요.")
