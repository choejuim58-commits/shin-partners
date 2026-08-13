import io
import pandas as pd
import requests
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# ---------------------------------------------------------
# 📌 1. 장바구니 세션 상태 초기화
# ---------------------------------------------------------
if "cart" not in st.session_state:
    st.session_state.cart = []

# ---------------------------------------------------------
# 📌 구글 시트 커넥션 & 데이터 로드 함수 (2번, 3번 적용)
# ---------------------------------------------------------
conn_gs = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data(ttl=300)
def get_gs_data():
    """구글 시트 최신 데이터 읽기 & 사전 정제 (속도 최적화)"""
    df = conn_gs.read(ttl=0).fillna("")
    if not df.empty and "price" in df.columns:
        df["price_num"] = pd.to_numeric(
            df["price"].astype(str).str.replace("원", "").str.replace(",", "").str.strip(),
            errors="coerce"
        ).fillna(0).astype(int)
    else:
        df["price_num"] = 0
    return df

def save_gs_data(df):
    """구글 시트에 데이터 덮어쓰기 저장"""
    # 저장 전 '구분' -> '품목군' -> '규격/자재명' 순 자동 정렬
    if not df.empty:
        sort_cols = [c for c in ["category", "item_type", "name"] if c in df.columns]
        if sort_cols:
            df = df.sort_values(by=sort_cols).reset_index(drop=True)
            
    conn_gs.update(data=df)
    st.cache_data.clear()

# ---------------------------------------------------------
# 원드라이브 설정 & 변환 함수
# ---------------------------------------------------------
URL_PANEL = "https://1drv.ms/x/c/e159fcf2d96e4b60/IQD9uLWkE5c7QYgaIwIL_YOgAfq_SOnucQ6LYqj58PDDqPg?e=HP4KZH"
URL_BUILDING = "https://1drv.ms/x/c/e159fcf2d96e4b60/IQBMeXt0XDmHT4f-A042xKnQAVl576fJOGz2wN8KgN3hd5Q?e=6X3Y1o"

@st.cache_data(ttl=43200)
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
# Streamlit 기본 레이아웃 설정
# ---------------------------------------------------------
st.set_page_config(page_title="신앤파트너스 단가 & 수량 관리", page_icon="📦", layout="wide")
st.title("📦 수입상 vs 합판상 단가 & 원드라이브 수량 통합 시스템")

tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🌐 원드라이브 실시간 수량",
    "🔎 단가 검색 및 담기",
    f"🧺 장바구니 ({len(st.session_state.cart)}개)",
    "📁 엑셀 일괄 등록", 
    "✏️ 개별 수동 등록", 
    "🛠️ 데이터 직접 수정/삭제"
])

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
    with col_btn:
        if st.button("🔄 원드라이브 데이터 최신화", type="primary"):
            st.cache_data.clear()
            st.rerun()

    with st.spinner("원드라이브 데이터 로딩 중..."):
        df_panel = load_onedrive_excel(URL_PANEL, "판재류", target_sheet="2026.8")
        df_building = load_onedrive_excel(URL_BUILDING, "건축자재", target_sheet="2026.8")
        df_onedrive_all = pd.concat([df_panel, df_building], ignore_index=True).fillna("")

    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["전체 보기", "🪵 판재류", "🏗️ 건축자재"])
    with sub_tab1:
        search_od = st.text_input("원드라이브 자재 검색", key="search_od")
        if not df_onedrive_all.empty:
            df_display = df_onedrive_all.copy()
            if search_od:
                mask = df_display.astype(str).apply(lambda row: row.str.contains(search_od, case=False).any(), axis=1)
                df_display = df_display[mask]
            st.dataframe(df_display.style.pipe(apply_excel_style), use_container_width=True, hide_index=True)
    with sub_tab2:
        if not df_panel.empty:
            st.dataframe(df_panel.style.pipe(apply_excel_style), use_container_width=True, hide_index=True)
    with sub_tab3:
        if not df_building.empty:
            st.dataframe(df_building.style.pipe(apply_excel_style), use_container_width=True, hide_index=True)

# --- [탭 1] 단가 검색 및 담기 (버그 완전 수정!) ---
with tab1:
    col_t1, col_t2 = st.columns([3, 1])
    with col_t1:
        st.subheader("🔎 자재 단가 조회 & 장바구니 담기")
    with col_t2:
        if st.button("🔄 단가표 즉시 새로고침", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_gs = get_gs_data()

    if not df_gs.empty:
        # 검색 필터 영역
        with st.form("search_form"):
            col1, col2, col3 = st.columns([2, 3, 1])
            with col1:
                category_filter = st.selectbox("조회할 단가표 선택", ["전체 (비교해서 보기)", "수입상 단가만", "합판상 단가만"], key="tab1_cat_select")
            with col2:
                search_kw = st.text_input("자재명/규격 입력 후 Enter", placeholder="예: 멀바우, MDF, OSB...", key="tab1_search_kw")
            with col3:
                st.write("")
                st.write("")
                search_submitted = st.form_submit_button("🔍 검색", use_container_width=True)

        df_filtered = df_gs.copy()
        if category_filter == "수입상 단가만":
            df_filtered = df_filtered[df_filtered["category"] == "수입상"]
        elif category_filter == "합판상 단가만":
            df_filtered = df_filtered[df_filtered["category"] == "합판상"]

        if search_kw:
            mask = (
                df_filtered["name"].astype(str).str.contains(search_kw, case=False) |
                df_filtered["item_type"].astype(str).str.contains(search_kw, case=False)
            )
            df_filtered = df_filtered[mask]

        if not df_filtered.empty:
            # 💡 [핵심 해결책] 고유 식별자(ID) 부여해서 인덱스/객체 꼬임 완벽 방지
            df_filtered = df_filtered.reset_index(drop=True)
            df_filtered['item_id'] = df_filtered.index

            with st.expander("🛒 선택한 품목 장바구니에 바로 담기", expanded=True):
                c_sel, c_qty, c_btn = st.columns([3, 1, 1])

                with c_sel:
                    # ID를 오퍼션으로 전달
                    id_list = df_filtered['item_id'].tolist()
                    
                    def format_by_id(item_id):
                        row = df_filtered.loc[item_id]
                        return f"[{row['category']}] {row['item_type']} | {row['name']} ({row['price_num']:,}원)"

                    selected_id = st.selectbox(
                        "담을 품목 선택",
                        options=id_list,
                        format_func=format_by_id,
                        label_visibility="collapsed"
                    )

                with c_qty:
                    add_qty = st.number_input("수량", min_value=1, value=1, step=1, label_visibility="collapsed")

                with c_btn:
                    if st.button("🛒 담기", type="primary", use_container_width=True):
                        selected_row = df_filtered.loc[selected_id]
                        
                        exists = False
                        for c in st.session_state.cart:
                            if c["category"] == selected_row["category"] and c["name"] == selected_row["name"]:
                                c["qty"] += add_qty
                                exists = True
                                break
                        if not exists:
                            st.session_state.cart.append({
                                "category": selected_row["category"],
                                "item_type": selected_row["item_type"],
                                "name": selected_row["name"],
                                "price": int(selected_row["price_num"]),
                                "qty": int(add_qty)
                            })
                        st.toast(f"✅ '{selected_row['name']}' {add_qty}개가 담겼습니다!", icon="🛒")
                        st.rerun()

            # 단가표 출력
            df_show = df_filtered.copy()
            df_show["price"] = df_show["price_num"].apply(lambda x: f"{x:,} 원")
            rename_map = {"category": "구분", "item_type": "품목군", "name": "규격/자재명", "price": "단가", "remark": "비고"}
            df_show = df_show.rename(columns=rename_map)
            
            st.dataframe(df_show[["구분", "품목군", "규격/자재명", "단가", "비고"]], use_container_width=True, hide_index=True)
        else:
            st.info("검색된 데이터가 없습니다.")
    else:
        st.info("구글 시트에 등록된 데이터가 없습니다.")

# --- [탭 2] 장바구니 및 견적서 확인 ---
with tab2:
    st.subheader("🧺 내가 담은 자재 목록 & 예상 견적")

    if st.session_state.cart:
        cart_df = pd.DataFrame(st.session_state.cart)
        cart_df["total_price"] = cart_df["price"] * cart_df["qty"]

        total_sum = cart_df["total_price"].sum()
        total_count = cart_df["qty"].sum()
        
        m1, m2, m3 = st.columns(3)
        m1.metric("선택한 자재 종류", f"{len(cart_df)}종")
        m2.metric("총 선택 수량", f"{total_count:,} 개")
        m3.metric("💳 총 예상 견적 금액", f"{total_sum:,} 원")

        st.markdown("---")

        disp_cart = cart_df[["category", "item_type", "name", "price", "qty", "total_price"]].copy()
        disp_cart["price"] = disp_cart["price"].apply(lambda x: f"{x:,}원")
        disp_cart["total_price"] = disp_cart["total_price"].apply(lambda x: f"{x:,}원")
        disp_cart.columns = ["구분", "품목군", "자재명/규격", "단가", "수량", "합계금액"]

        st.dataframe(disp_cart, use_container_width=True, hide_index=True)

        col_b1, col_b2, _ = st.columns([1, 1, 2])
        with col_b1:
            if st.button("🗑️ 장바구니 전체 비우기", use_container_width=True):
                st.session_state.cart = []
                st.rerun()

        with col_b2:
            excel_cart = cart_df[["category", "item_type", "name", "price", "qty", "total_price"]].copy()
            excel_cart.columns = ["구분", "품목군", "규격/자재명", "단가(원)", "수량", "합계금액(원)"]
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                excel_cart.to_excel(writer, index=False, sheet_name="선택자재_견적서")
            
            st.download_button(
                label="📥 견적 내역 엑셀 받기",
                data=buffer.getvalue(),
                file_name="선택자재_견적목록.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("💡 장바구니가 비어 있습니다.\n[🔎 단가 검색 및 담기] 탭에서 원하는 품목을 담아보세요!")

# --- [탭 3] 엑셀 일괄 등록 ---
with tab3:
    st.subheader("📁 엑셀(.xlsx) 파일 업로드")
    target_category = st.radio("어떤 단가표 데이터인가요?", ["수입상", "합판상"], horizontal=True)
    uploaded_file = st.file_uploader("엑셀 파일을 등록해 주세요", type=["xlsx", "csv"])

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
            st.error(f"등록 실패: {e}")

# --- [탭 4] 개별 수동 등록 ---
with tab4:
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
            st.success(f"[{f_cat}] '{f_name}' 단가가 구글 시트에 저장되었습니다!")
            st.rerun()
        else:
            st.error("규격과 단가를 입력해 주세요.")

# --- [탭 5] 데이터 직접 수정/삭제 ---
with tab5:
    st.subheader("🛠️ 구글 시트 데이터 직접 수정 & 삭제")
    edit_cat = st.selectbox("수정/관리할 단가표 선택", ["전체 데이터", "수입상", "합판상"], key="edit_cat_select")
    df_edit = get_gs_data()

    if not df_edit.empty:
        df_edit = df_edit.rename(columns={"비고": "remark", "구분": "category", "품목군": "item_type", "규격/자재명": "name", "단가": "price"})
        if "remark" in df_edit.columns:
            df_edit["remark"] = df_edit["remark"].fillna("").astype(str)

        if edit_cat == "수입상":
            df_edit_filtered = df_edit[df_edit["category"] == "수입상"].copy().reset_index(drop=True)
        elif edit_cat == "합판상":
            df_edit_filtered = df_edit[df_edit["category"] == "합판상"].copy().reset_index(drop=True)
        else:
            df_edit_filtered = df_edit.copy().reset_index(drop=True)

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
            key=f"gs_editor_{edit_cat}"
        )

        if st.button("💾 구글 시트에 수정사항 반영하기", type="primary"):
            if edit_cat == "전체 데이터":
                final_df = edited_df
            else:
                other_df = df_edit[df_edit["category"] != edit_cat]
                final_df = pd.concat([other_df, edited_df], ignore_index=True)

            save_gs_data(final_df)
            st.success("✅ 구글 시트에 성공적으로 반영되었습니다!")
            st.rerun()
