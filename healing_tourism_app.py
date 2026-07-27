
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import io
import folium
from streamlit_folium import st_folium
from geopy.geocoders import Nominatim
import time

st.set_page_config(page_title="전북 치유관광자원 등록 시스템", layout="wide")

DB_PATH = "healing_tourism.db"
ADMIN_PASSWORD_DEFAULT = "0000"

THEMES = ["자연/치유", "힐링/명상", "전통/생활문화", "뷰티/스파", "치유음식", "인문치유"]
OPERATORS = ["공공", "민간", "기타"]

REG_TYPE_SERVICE = "치유관광서비스 제공형"
REG_TYPE_FACILITY = "치유관광시설 이용형"
FACILITY_BUSINESS_TYPES = [
    "호텔업(관광호텔업 등)",
    "전문휴양업(숙박시설 갖춘 업)",
    "종합휴양업(숙박시설 갖춘 업)",
    "한옥체험업(숙박 체험 시설)",
    "관광펜션업",
    "해당없음"
]


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS resources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT NOT NULL,
            contact TEXT,
            operator_type TEXT,
            operating_hours TEXT,
            theme TEXT,
            has_specialist TEXT,
            program_count INTEGER,
            weekly_freq REAL,
            registration_type TEXT,
            facility_business_type TEXT,
            lat REAL,
            lon REAL,
            status TEXT DEFAULT '검토중',
            memo TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()


init_db()


@st.cache_data(show_spinner=False)
def geocode_address(address: str):
    try:
        geolocator = Nominatim(user_agent="jeonbuk_healing_tourism_app")
        loc = geolocator.geocode(address, timeout=10)
        if loc:
            return loc.latitude, loc.longitude
    except Exception:
        pass
    return None, None


def eligibility_check(theme, has_specialist, program_count, weekly_freq, registration_type, facility_business_type):
    """치유관광사업자 등록제도 실무 가이드라인 기반 자가 체크리스트 로직"""
    issues = []
    ok = True

    if registration_type == REG_TYPE_SERVICE:
        if program_count < 1:
            issues.append("치유관광 프로그램이 1개 이상 필요합니다.")
            ok = False
        if weekly_freq < 1:
            issues.append("연평균 주 1회(연 52회) 이상 상시 운영 기준 미충족 가능성이 있습니다.")
            ok = False
    else:
        if program_count < 2:
            issues.append("치유관광시설 이용형은 프로그램 2개 이상이 필요합니다.")
            ok = False
        if weekly_freq < 1:
            issues.append("2개 프로그램 기준 연 104회(프로그램당 주1회) 이상 운영 기준 확인이 필요합니다.")
            ok = False
        if facility_business_type == "해당없음":
            issues.append("호텔업/전문(종합)휴양업(숙박)/한옥체험업/관광펜션업 중 하나의 등록·지정이 필요합니다.")
            ok = False

    if has_specialist != "있음":
        issues.append("치유관광 관련 교육과정 이수 전문인력 1명 이상 확보가 필요합니다.")
        ok = False

    if not issues:
        issues.append("등록기준(안) 충족 가능성이 높습니다. 관할 시·군·구 등록신청서 제출을 준비하세요.")

    return ok, issues


def add_resource(data: dict):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO resources
        (name, address, contact, operator_type, operating_hours, theme, has_specialist,
         program_count, weekly_freq, registration_type, facility_business_type,
         lat, lon, status, memo, created_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        data["name"], data["address"], data["contact"], data["operator_type"],
        data["operating_hours"], data["theme"], data["has_specialist"],
        data["program_count"], data["weekly_freq"], data["registration_type"],
        data["facility_business_type"], data["lat"], data["lon"], "검토중", "",
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))
    conn.commit()
    conn.close()


def load_resources():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM resources ORDER BY id DESC", conn)
    conn.close()
    return df


def update_status(res_id, status, memo):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE resources SET status=?, memo=? WHERE id=?", (status, memo, res_id))
    conn.commit()
    conn.close()


def delete_resource(res_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM resources WHERE id=?", (res_id,))
    conn.commit()
    conn.close()


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="치유관광자원목록")
        ws = writer.sheets["치유관광자원목록"]
        for i, col in enumerate(df.columns, 1):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str)])
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len + 4, 40)
    return output.getvalue()


def render_map(df: pd.DataFrame, height=520):
    center = [35.8242, 127.1480]  # 전북 대략 중심
    m = folium.Map(location=center, zoom_start=9)
    theme_colors = {
        "자연/치유": "green", "힐링/명상": "purple", "전통/생활문화": "orange",
        "뷰티/스파": "pink", "치유음식": "red", "인문치유": "blue"
    }
    for _, row in df.iterrows():
        if pd.notna(row.get("lat")) and pd.notna(row.get("lon")):
            popup_html = f"""
            <b>{row['name']}</b><br>
            주소: {row['address']}<br>
            테마: {row['theme']}<br>
            운영주체: {row['operator_type']}<br>
            전문인력: {row['has_specialist']}<br>
            상태: {row['status']}
            """
            folium.Marker(
                location=[row["lat"], row["lon"]],
                popup=folium.Popup(popup_html, max_width=250),
                tooltip=row["name"],
                icon=folium.Icon(color=theme_colors.get(row["theme"], "gray"))
            ).add_to(m)
    st_folium(m, width=None, height=height)


def registrant_page():
    st.header("🌿 예비 치유관광자원 등록")
    st.caption("전북특별자치도문화관광재단 · 치유관광산업 육성에 관한 법률 기반 예비 등록 시스템")

    with st.form("register_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("등록지명 *")
            address = st.text_input("주소 *", placeholder="예: 전북특별자치도 전주시 완산구 ...")
            contact = st.text_input("연락처")
            operator_type = st.selectbox("운영주체", OPERATORS)
            operating_hours = st.text_input("운영시간", placeholder="예: 09:00~18:00, 연중무휴")
        with col2:
            theme = st.selectbox("테마유형", THEMES)
            has_specialist = st.radio("치유관광 전문인력 보유 여부", ["있음", "없음"], horizontal=True)
            registration_type = st.radio("등록 희망 유형", [REG_TYPE_SERVICE, REG_TYPE_FACILITY])
            facility_business_type = st.selectbox(
                "치유관광시설 관련 업종(시설 이용형만 해당)", FACILITY_BUSINESS_TYPES
            )
            program_count = st.number_input("운영 중(예정) 치유관광 프로그램 수", min_value=0, step=1)
            weekly_freq = st.number_input("1개 프로그램당 평균 주간 운영 횟수", min_value=0.0, step=0.5)

        submitted = st.form_submit_button("등록하기")

    if submitted:
        if not name or not address:
            st.error("등록지명과 주소는 필수 입력 항목입니다.")
        else:
            with st.spinner("주소를 좌표로 변환하는 중..."):
                lat, lon = geocode_address(address)
            if lat is None:
                st.warning("주소를 좌표로 변환하지 못했습니다. 지도에는 표시되지 않으며, 주소를 더 구체적으로 입력해 다시 시도해보세요.")
            data = {
                "name": name, "address": address, "contact": contact,
                "operator_type": operator_type, "operating_hours": operating_hours,
                "theme": theme, "has_specialist": has_specialist,
                "program_count": int(program_count), "weekly_freq": float(weekly_freq),
                "registration_type": registration_type,
                "facility_business_type": facility_business_type,
                "lat": lat, "lon": lon
            }
            add_resource(data)
            st.success(f"'{name}' 이(가) 등록되었습니다. (검토중 상태)")

            ok, issues = eligibility_check(theme, has_specialist, int(program_count), float(weekly_freq),
                                            registration_type, facility_business_type)
            st.subheader("📋 치유관광사업자 등록기준 자가 체크(참고용)")
            if ok:
                st.success("등록기준(안)을 충족할 가능성이 높습니다.")
            else:
                st.info("아래 항목을 보완하면 정식 등록 신청이 수월합니다.")
            for i in issues:
                st.write(f"- {i}")
            st.caption("※ 본 체크는 실무 가이드라인 요약 기준의 참고용 자가진단이며, 최종 등록 여부는 관할 시·군·구 심사에 따릅니다.")

    st.divider()
    st.subheader("🗺️ 등록 현황 지도")
    df = load_resources()
    if df.empty:
        st.info("아직 등록된 자원이 없습니다.")
    else:
        theme_filter = st.multiselect("테마 필터", THEMES, default=THEMES)
        filtered = df[df["theme"].isin(theme_filter)] if theme_filter else df
        render_map(filtered)
        st.caption(f"총 등록 건수: {len(df)}건")


def admin_login():
    st.header("🔐 관리자 로그인")
    pw = st.text_input("비밀번호를 입력하세요 (임시 비밀번호: 0000)", type="password")
    if st.button("로그인"):
        if pw == st.session_state.get("admin_password", ADMIN_PASSWORD_DEFAULT):
            st.session_state["is_admin"] = True
            st.success("로그인 성공")
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")


def admin_page():
    st.header("🛠️ 관리자 페이지")
    st.caption("등록된 예비 치유관광자원을 검토·관리합니다.")

    if st.button("로그아웃"):
        st.session_state["is_admin"] = False
        st.rerun()

    with st.expander("비밀번호 변경"):
        new_pw = st.text_input("새 비밀번호 설정", type="password", key="new_pw")
        if st.button("비밀번호 변경 적용"):
            if new_pw:
                st.session_state["admin_password"] = new_pw
                st.success("비밀번호가 변경되었습니다.")
            else:
                st.warning("새 비밀번호를 입력하세요.")

    df = load_resources()
    st.subheader(f"📊 등록 현황 (총 {len(df)}건)")

    if df.empty:
        st.info("등록된 데이터가 없습니다.")
        return

    col1, col2, col3 = st.columns(3)
    col1.metric("전체 등록 수", len(df))
    col2.metric("전문인력 보유", int((df["has_specialist"] == "있음").sum()))
    col3.metric("승인 완료", int((df["status"] == "승인").sum()))

    status_filter = st.multiselect("상태 필터", ["검토중", "승인", "보완요청", "반려"],
                                    default=["검토중", "승인", "보완요청", "반려"])
    view_df = df[df["status"].isin(status_filter)]
    st.dataframe(view_df, use_container_width=True, height=300)

    st.subheader("✏️ 개별 항목 관리")
    options = view_df["id"].tolist()
    if options:
        sel_id = st.selectbox("관리할 항목 선택 (ID)", options,
                               format_func=lambda x: f"{x} - {df[df['id']==x]['name'].values[0]}")
        row = df[df["id"] == sel_id].iloc[0]
        st.write(row.to_frame().T)
        new_status = st.selectbox("상태 변경", ["검토중", "승인", "보완요청", "반려"],
                                   index=["검토중", "승인", "보완요청", "반려"].index(row["status"]))
        memo = st.text_area("관리자 메모", value=row["memo"] if row["memo"] else "")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("저장"):
                update_status(sel_id, new_status, memo)
                st.success("저장되었습니다.")
                st.rerun()
        with c2:
            if st.button("삭제", type="secondary"):
                delete_resource(sel_id)
                st.warning("삭제되었습니다.")
                st.rerun()

    st.divider()
    st.subheader("📥 엑셀 다운로드")
    excel_data = to_excel_bytes(df)
    st.download_button(
        label="전체 등록 목록 엑셀로 다운로드",
        data=excel_data,
        file_name=f"전북_치유관광자원_목록_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()
    st.subheader("🗺️ 전체 등록 지도")
    render_map(df)


def main():
    st.sidebar.title("🌿 전북 치유관광자원 등록시스템")
    st.sidebar.caption("전북특별자치도문화관광재단\n치유관광산업 육성에 관한 법률 기반")
    menu = st.sidebar.radio("메뉴", ["등록자 화면", "관리자 화면"])

    if "is_admin" not in st.session_state:
        st.session_state["is_admin"] = False

    if menu == "등록자 화면":
        registrant_page()
    else:
        if st.session_state["is_admin"]:
            admin_page()
        else:
            admin_login()

    st.sidebar.divider()
    st.sidebar.markdown("""
    **참고: 치유관광 등록기준 요약**
    - 서비스 제공형: 프로그램 1개↑ + 전문인력 1명↑
    - 시설 이용형: 프로그램 2개↑ + 전문인력 1명↑ + 숙박 관련 업종
    - 상시운영: 연평균 주1회(연 52회)↑ 기준
    - 등록은 의무가 아닌 임의 등록제입니다.
    """)


if __name__ == "__main__":
    main()
