"""
學校資料管理與情境式查詢系統 v2
── Excel 驅動版 ──
使用方式：streamlit run app.py
Excel 欄位（固定）：
  學校名稱 | 學年度學期 | 學科 | 老師姓名 | 兼任職務
  任教年段 | 任教班級 | 教科書版本 | 特殊身分備註 | 班級人數
"""

import streamlit as st
import pandas as pd
import io

# ─────────────────────────────────────────────────────────────────────
#  常數定義
# ─────────────────────────────────────────────────────────────────────
# 十科固定順序（用於版本對照表與班級陣容）
SUBJECTS = ["國文", "英文", "數學", "物理", "化學",
            "生物", "地球科學", "歷史", "地理", "公民與社會"]

# 年段關鍵字對應（模糊比對用）
GRADE_KEYWORDS = {"高一": ["高一", "1年", "一年"], "高二": ["高二", "2年", "二年"], "高三": ["高三", "3年", "三年"]}

# Badge 顏色對應（特殊身分備註）
BADGE_COLORS = {
    "學科召集人": "#1976D2",
    "年級召集人": "#388E3C",
    "行政組長":   "#F57C00",
    "代理":       "#7B1FA2",
    "兼課":       "#00838F",
    "留職停薪":   "#D32F2F",
}

# ─────────────────────────────────────────────────────────────────────
#  頁面設定
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="學校資料管理系統",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────
#  全域 CSS（Mobile-First RWD）
# ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* 字型 */
html, body, [class*="css"] {
    font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif;
}
/* 頂部 Header */
.header-bar {
    background: linear-gradient(135deg, #1565C0 0%, #0D47A1 100%);
    color: white;
    padding: 0.9rem 1.2rem 0.7rem;
    border-radius: 12px;
    margin-bottom: 1rem;
}
.header-bar h1 { font-size: 1.25rem; margin: 0; }
.header-bar p  { font-size: 0.82rem; margin: 0.2rem 0 0; opacity: 0.85; }

/* 老師卡片 */
.teacher-card {
    background: white;
    border: 1px solid #E0E0E0;
    border-radius: 10px;
    padding: 0.7rem 1rem;
    margin-bottom: 0.45rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}
.teacher-card .name  { font-size: 1rem; font-weight: 600; color: #1A237E; }
.teacher-card .meta  { font-size: 0.78rem; color: #555; margin-top: 3px; }

/* 摘要欄 */
.summary-box {
    background: #E3F2FD;
    border-left: 4px solid #1565C0;
    border-radius: 6px;
    padding: 0.55rem 1rem;
    margin-top: 0.7rem;
    font-size: 0.9rem;
}
.summary-warn {
    background: #FFF3E0;
    border-left: 4px solid #E65100;
    border-radius: 6px;
    padding: 0.55rem 1rem;
    margin-top: 0.5rem;
    font-size: 0.88rem;
    color: #BF360C;
}
/* 手機按鈕更大 */
@media (max-width: 768px) {
    .stSelectbox > div > div { min-height: 44px; font-size: 1rem; }
    .stNumberInput input     { min-height: 44px; font-size: 1rem; }
    .stButton > button       { min-height: 44px; font-size: 1rem; width: 100%; }
    .header-bar h1           { font-size: 1.05rem; }
}
/* Sidebar 上傳區塊 */
.upload-hint {
    font-size: 0.8rem;
    color: #666;
    margin-top: 0.4rem;
    line-height: 1.5;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────
#  工具函式
# ─────────────────────────────────────────────────────────────────────
def badge_html(label: str) -> str:
    """將特殊身分備註轉為彩色 HTML Badge"""
    if not label or label == "-":
        return ""
    color = BADGE_COLORS.get(label, "#607D8B")
    return (
        f'<span style="background:{color};color:#fff;padding:2px 7px;'
        f'border-radius:10px;font-size:0.72rem;margin-left:5px;'
        f'white-space:nowrap;display:inline-block;">{label}</span>'
    )


def safe_str(val) -> str:
    """將 NaN 或 None 轉為空字串，其餘轉 str，防止整個頁面崩潰"""
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    return str(val).strip()


def load_excel(file_obj) -> pd.DataFrame | None:
    """
    讀取上傳的 Excel，做基本欄位檢查與 NaN 清理。
    回傳已清理的 DataFrame，若欄位不符回傳 None。
    """
    required_cols = {
        "學校名稱", "學年度學期", "學科", "老師姓名",
        "兼任職務", "任教年段", "任教班級", "教科書版本",
        "特殊身分備註", "班級人數"
    }
    try:
        df = pd.read_excel(file_obj, dtype=str)  # 全部先讀為字串，避免數字欄位被轉 float
    except Exception as e:
        st.sidebar.error(f"❌ 讀取 Excel 失敗：{e}")
        return None

    # 欄位名稱去除前後空白
    df.columns = [c.strip() for c in df.columns]

    # 檢查必要欄位
    missing = required_cols - set(df.columns)
    if missing:
        st.sidebar.error(f"❌ Excel 缺少欄位：{', '.join(sorted(missing))}")
        return None

    # 全欄 NaN → 空字串，確保後續 .str 操作不報錯
    df = df.fillna("")

    # 班級人數轉數值（空字串→0，非數字→0）
    def to_int(v):
        try:
            return int(float(str(v))) if str(v).strip() not in ("", "-") else 0
        except Exception:
            return 0

    df["班級人數_int"] = df["班級人數"].apply(to_int)

    return df


def filter_df(df: pd.DataFrame, school: str, semester: str) -> pd.DataFrame:
    """依學校 + 學期過濾 DataFrame"""
    return df[(df["學校名稱"] == school) & (df["學年度學期"] == semester)].copy()


def grade_match(grade_cell: str, target_grade: str) -> bool:
    """
    判斷 [任教年段] 欄位是否包含目標年段。
    欄位值可能是 '高一'、'高一,高二'、'1' 等多種格式。
    """
    cell = safe_str(grade_cell)
    # 直接包含
    if target_grade in cell:
        return True
    # 關鍵字比對
    for kw in GRADE_KEYWORDS.get(target_grade, []):
        if kw in cell:
            return True
    return False


def class_match(class_cell: str, target_class: str) -> bool:
    """
    判斷 [任教班級] 欄位是否包含目標班級。
    支援 contains 模糊比對，例如 '101' in '101, 102, 103'。
    """
    return target_class.strip() in safe_str(class_cell)


# ─────────────────────────────────────────────────────────────────────
#  Sidebar：上傳 Excel + 全局控制項
# ─────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 資料來源")
    uploaded_file = st.file_uploader(
        "上傳 Excel 總表（.xlsx）",
        type=["xlsx"],
        help="欄位：學校名稱、學年度學期、學科、老師姓名、兼任職務、任教年段、任教班級、教科書版本、特殊身分備註、班級人數",
    )
    st.markdown("""
    <div class="upload-hint">
    📋 <b>必要欄位（固定 10 欄）：</b><br>
    學校名稱、學年度學期、學科、老師姓名、兼任職務、任教年段、任教班級、教科書版本、特殊身分備註、班級人數
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # 讀取 Excel
    df_raw = None
    if uploaded_file:
        df_raw = load_excel(uploaded_file)
        if df_raw is not None:
            st.success(f"✅ 已載入 {len(df_raw)} 筆資料")

    # 從 Excel 動態產生下拉選單
    if df_raw is not None:
        semesters = sorted(df_raw["學年度學期"].unique().tolist())
        schools   = sorted(df_raw["學校名稱"].unique().tolist())
    else:
        semesters = ["（請先上傳 Excel）"]
        schools   = ["（請先上傳 Excel）"]

    st.markdown("## ⚙️ 全局切換")
    sel_semester = st.selectbox("📅 學年度 / 學期", semesters, key="g_semester")
    sel_school   = st.selectbox("🏫 學校",          schools,   key="g_school")

    st.divider()
    st.markdown("## 🖼️ 座位表圖片路徑")
    seat_img_path = st.text_input(
        "輸入圖片相對或絕對路徑",
        value="",
        placeholder="例如：seating_101.jpg",
        key="seat_img_path",
    )


# ─────────────────────────────────────────────────────────────────────
#  Header
# ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header-bar">
  <h1>🏫 學校資料管理與情境式查詢系統</h1>
  <p>Mobile-First · Excel 驅動 · {sel_semester} ｜ {sel_school}</p>
</div>
""", unsafe_allow_html=True)

# 尚未上傳時顯示引導頁
if df_raw is None:
    st.info("👈 請先在左側側邊欄上傳 Excel 總表（.xlsx），系統即自動載入所有功能。")
    st.markdown("""
    ### 📌 Excel 總表格式說明

    請確保 Excel 第一列為以下欄位名稱（順序不限）：

    | 欄位名稱 | 說明 | 範例 |
    |---|---|---|
    | 學校名稱 | 學校全名 | 陽明高中 |
    | 學年度學期 | 學期代碼 | 114上 |
    | 學科 | 十科之一 | 數學 |
    | 老師姓名 | 老師全名 | 王大明 |
    | 兼任職務 | 職稱 | 教師 / 組長 / 主任 |
    | 任教年段 | 年段文字 | 高一 / 高一,高二 |
    | 任教班級 | 班級號碼（逗號分隔） | 101, 102, 201 |
    | 教科書版本 | 出版社名稱 | 龍騰 |
    | 特殊身分備註 | 特殊身分 | 學科召集人 / 代理 / 留職停薪 |
    | 班級人數 | 整數，可空白 | 32 |

    > 💡 空白欄位（NaN）系統會自動以 `-` 顯示，不會造成程式錯誤。
    """)
    st.stop()  # 未上傳前停止渲染後續內容

# 依全局選單過濾出當前 DataFrame
df = filter_df(df_raw, sel_school, sel_semester)

if df.empty:
    st.warning(f"⚠️ 目前 Excel 中找不到「{sel_school}」×「{sel_semester}」的資料，請確認選單或 Excel 內容。")
    st.stop()


# ─────────────────────────────────────────────────────────────────────
#  五個功能分頁
# ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🟢 發放情境", "🟡 打包計算", "📚 用書版本", "🖼️ 座位表", "🔍 配課總表"
])


# ══════════════════════════════════════════════════════════════════════
#  Tab 1：發放情境（工具 A：按年段找老師 ｜ 工具 B：按班級看陣容）
# ══════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("🟢 發放情境：年段與班級速查")

    subtab_a, subtab_b = st.tabs(["工具 A｜按年段找老師", "工具 B｜按班級看陣容"])

    # ── 工具 A：選學科 + 年段，列出符合老師（附 Badge）──
    with subtab_a:
        col_s, col_g = st.columns(2)
        with col_s:
            # 動態從 Excel 取得學科清單（保留原始順序 + 固定十科補全）
            avail_subjects = [s for s in SUBJECTS if s in df["學科"].values]
            if not avail_subjects:
                avail_subjects = SUBJECTS
            sel_subj = st.selectbox("學科", avail_subjects, key="ta_subj")

        with col_g:
            sel_grade = st.selectbox("年段", ["高一", "高二", "高三"], key="ta_grade")

        # 過濾：學科 match + 年段 match（模糊）
        matched = df[
            (df["學科"] == sel_subj) &
            df["任教年段"].apply(lambda v: grade_match(v, sel_grade))
        ]

        st.caption(f"符合「{sel_grade}｜{sel_subj}」共 **{len(matched)}** 位老師")

        if matched.empty:
            st.info("查無符合的老師，請確認 Excel 資料或切換條件。")
        else:
            for _, row in matched.iterrows():
                sp    = safe_str(row.get("特殊身分備註", ""))
                duty  = safe_str(row.get("兼任職務", ""))
                bh    = badge_html(sp) if sp else ""
                ver   = safe_str(row.get("教科書版本", ""))
                classes = safe_str(row.get("任教班級", ""))
                st.markdown(f"""
                <div class="teacher-card">
                  <div class="name">{safe_str(row['老師姓名'])}{bh}</div>
                  <div class="meta">
                    兼任職務：{duty or '—'} ｜ 任教班級：{classes or '—'} ｜ 版本：{ver or '—'}
                  </div>
                </div>
                """, unsafe_allow_html=True)

    # ── 工具 B：輸入班級號碼，看 10 科陣容 ──
    with subtab_b:
        st.caption("輸入班級號碼（例如：101），系統以「包含比對」搜尋任教班級欄位。")

        # 自動從 Excel 產生班級建議清單（拆解所有任教班級值）
        all_class_vals = set()
        for cell in df["任教班級"]:
            for c in str(cell).replace("，", ",").split(","):
                c = c.strip()
                if c:
                    all_class_vals.add(c)
        all_classes_sorted = sorted(all_class_vals)

        # 提供下拉 + 自由輸入兩種方式
        col_cls1, col_cls2 = st.columns([1, 1])
        with col_cls1:
            dropdown_class = st.selectbox(
                "從下拉選擇班級",
                ["（手動輸入）"] + all_classes_sorted,
                key="tb_dropdown",
            )
        with col_cls2:
            manual_class = st.text_input("或手動輸入班級號碼", value="", placeholder="例如 101", key="tb_manual")

        # 決定最終使用的班級
        target_class = manual_class.strip() if manual_class.strip() else (
            dropdown_class if dropdown_class != "（手動輸入）" else ""
        )

        if not target_class:
            st.info("請選擇或輸入班級號碼。")
        else:
            # 找出這個班所有學科的老師（contains 比對）
            cls_df = df[df["任教班級"].apply(lambda v: class_match(v, target_class))]

            st.caption(f"班級「{target_class}」— 找到 {len(cls_df)} 筆任教記錄")

            if cls_df.empty:
                st.warning("查無此班級的任教資料，請確認 Excel 中 [任教班級] 欄位的格式。")
            else:
                # 以學科為軸，整理成一頁清單
                for subj in SUBJECTS:
                    subj_rows = cls_df[cls_df["學科"] == subj]
                    if subj_rows.empty:
                        st.markdown(f"""
                        <div class="teacher-card" style="opacity:0.45;">
                          <div class="name" style="font-size:0.88rem;">【{subj}】</div>
                          <div class="meta">— 無資料 —</div>
                        </div>""", unsafe_allow_html=True)
                    else:
                        for _, row in subj_rows.iterrows():
                            sp   = safe_str(row.get("特殊身分備註", ""))
                            duty = safe_str(row.get("兼任職務", ""))
                            ver  = safe_str(row.get("教科書版本", ""))
                            bh   = badge_html(sp) if sp else ""
                            st.markdown(f"""
                            <div class="teacher-card">
                              <div class="name" style="font-size:0.88rem;">【{subj}】{safe_str(row['老師姓名'])}{bh}</div>
                              <div class="meta">兼任：{duty or '—'} ｜ 版本：{ver or '—'}</div>
                            </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
#  Tab 2：打包計算器
# ══════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("🟡 打包點收計算器")
    st.caption("班級人數來源：Excel [班級人數] 欄位。若空白系統自動標示警告，並允許現場手動輸入。")

    # 上方控制項
    col_e1, col_e2 = st.columns([1, 1])
    with col_e1:
        extra_qty = st.number_input("每班額外備用量（份）", min_value=0, value=0, step=1, key="extra_qty")
    with col_e2:
        sel_pack_grades = st.multiselect("計入計算的年段", ["高一", "高二", "高三"],
                                          default=["高一", "高二", "高三"], key="pack_grades")

    st.divider()

    # ── 從 Excel 抓取班級人數資料 ──
    # 策略：以「任教班級」欄位拆解出所有唯一班級，再對應「班級人數」
    # 因為一個老師會帶多班，我們只取 班級+人數 的唯一組合
    # 先建立 班級 → 人數 字典
    class_count_map: dict[str, int] = {}
    # 只取班級人數有值的列（避免同班不同老師產生重複）
    for _, row in df.iterrows():
        for cls in str(row.get("任教班級", "")).replace("，", ",").split(","):
            cls = cls.strip()
            if not cls:
                continue
            if cls not in class_count_map:
                cnt = int(row["班級人數_int"]) if row["班級人數_int"] else 0
                class_count_map[cls] = cnt
            else:
                # 若已有且為 0，嘗試用非零值覆蓋
                if class_count_map[cls] == 0 and row["班級人數_int"]:
                    class_count_map[cls] = int(row["班級人數_int"])

    # 判斷年段（班級前一碼：1xx=高一, 2xx=高二, 3xx=高三）
    def class_to_grade(cls_code: str) -> str:
        c = cls_code.strip()
        if c.startswith("1"):
            return "高一"
        elif c.startswith("2"):
            return "高二"
        elif c.startswith("3"):
            return "高三"
        return "其他"

    # 臨時人數 session_state 初始化
    if "temp_counts" not in st.session_state:
        st.session_state.temp_counts = {}

    # 整理成顯示資料
    grand_students = 0
    grand_pack = 0
    warn_classes = []

    for grade in ["高一", "高二", "高三"]:
        grade_classes = {k: v for k, v in class_count_map.items() if class_to_grade(k) == grade}
        if not grade_classes:
            continue

        in_calc = grade in sel_pack_grades
        st.markdown(f"#### {grade}{'（計入計算）' if in_calc else '（未選入）'}")

        grade_students = 0
        grade_pack = 0

        cols_header = st.columns([2, 2, 2, 2, 3])
        cols_header[0].markdown("**班級**")
        cols_header[1].markdown("**原始人數**")
        cols_header[2].markdown("**暫時人數**")
        cols_header[3].markdown("**備用量**")
        cols_header[4].markdown("**本次打包**")

        for cls in sorted(grade_classes.keys()):
            base_count = grade_classes[cls]
            temp_key = f"temp_{cls}"

            # 臨時人數輸入（現場手動輸入）
            temp_val = st.session_state.temp_counts.get(cls, base_count)

            c0, c1, c2, c3, c4 = st.columns([2, 2, 2, 2, 3])
            with c0:
                st.write(cls)
            with c1:
                if base_count == 0:
                    st.markdown('<span style="color:#D32F2F;">⚠️ 0（空白）</span>', unsafe_allow_html=True)
                    warn_classes.append(cls)
                else:
                    st.write(base_count)
            with c2:
                # 臨時輸入框
                new_temp = st.number_input(
                    f"暫時人數_{cls}", min_value=0,
                    value=temp_val, step=1,
                    label_visibility="collapsed",
                    key=temp_key,
                )
                st.session_state.temp_counts[cls] = new_temp
            with c3:
                st.write(extra_qty if in_calc else 0)
            with c4:
                pack = (new_temp + extra_qty) if in_calc else new_temp
                st.markdown(f"**{pack}**")
                grade_pack += pack
                grade_students += new_temp

        if in_calc:
            st.markdown(f"""
            <div class="summary-box">
              <b>{grade} 小計</b>：{len(grade_classes)} 班 ｜ 學生合計 {grade_students} 人
              ｜ <span style="color:#1565C0;font-weight:700;">本次打包 {grade_pack} 份</span>
            </div>
            """, unsafe_allow_html=True)
            grand_students += grade_students
            grand_pack += grade_pack

        st.write("")  # 間距

    # 人數空白警示
    if warn_classes:
        st.markdown(f"""
        <div class="summary-warn">
          ⚠️ 下列班級人數在 Excel 中為空白，已以暫時人數欄取代，請確認輸入：<br>
          <b>{", ".join(warn_classes)}</b>
        </div>
        """, unsafe_allow_html=True)

    # 全校合計
    st.markdown(f"""
    <div class="summary-box" style="background:#FFF8E1;border-color:#F9A825;margin-top:1rem;">
      🏫 <b>全校合計（已選年段）</b>
      ｜ 學生 {grand_students} 人
      ｜ <span style="color:#E65100;font-weight:700;font-size:1.05rem;">本次打包 {grand_pack} 份</span>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
#  Tab 3：用書版本對照表
# ══════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("📚 十科用書版本對照表")
    st.caption("自動從 Excel 交叉統計各學科 × 各年段的教科書版本；同一格複數版本自動以「/」合併。")

    # 交叉統計：學科 × 年段 → 版本集合
    # 年段判斷：優先用 [任教年段] 欄位，模糊比對三個年段
    version_matrix: dict[str, dict[str, set]] = {
        subj: {"高一": set(), "高二": set(), "高三": set()} for subj in SUBJECTS
    }

    for _, row in df.iterrows():
        subj = safe_str(row.get("學科", ""))
        ver  = safe_str(row.get("教科書版本", ""))
        if subj not in SUBJECTS or not ver or ver == "-":
            continue
        grade_cell = safe_str(row.get("任教年段", ""))
        for grade in ["高一", "高二", "高三"]:
            if grade_match(grade_cell, grade):
                version_matrix[subj][grade].add(ver)

    # 整理成 DataFrame 顯示
    rows_v = []
    for subj in SUBJECTS:
        rows_v.append({
            "學科":  subj,
            "高一":  "/".join(sorted(version_matrix[subj]["高一"])) or "—",
            "高二":  "/".join(sorted(version_matrix[subj]["高二"])) or "—",
            "高三":  "/".join(sorted(version_matrix[subj]["高三"])) or "—",
        })

    df_ver = pd.DataFrame(rows_v)
    st.dataframe(df_ver, use_container_width=True, hide_index=True)
    st.info("💡 同一欄位「/」分隔代表該年段同時使用多個版本（例如：龍騰/三民）。")


# ══════════════════════════════════════════════════════════════════════
#  Tab 4：座位表
# ══════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("🖼️ 座位表檢視")

    seat_subtab1, seat_subtab2 = st.tabs(["文字 / 表格座位表", "圖片座位表"])

    with seat_subtab1:
        st.markdown("#### 📋 文字 / 表格式座位表")
        default_seat = (
            "（請在此貼上或輸入文字版座位表）\n\n"
            "範例：\n"
            "╔═══╦═══╦═══╦═══╦═══╗\n"
            "║ 1 ║ 2 ║ 3 ║ 4 ║ 5 ║\n"
            "╠═══╬═══╬═══╬═══╬═══╣\n"
            "║ 6 ║ 7 ║ 8 ║ 9 ║10 ║\n"
            "╚═══╩═══╩═══╩═══╩═══╝"
        )
        st.text_area("座位表內容（可直接編輯）", value=default_seat, height=260, key="seat_text")

    with seat_subtab2:
        st.markdown("#### 🖼️ 圖片座位表")
        # 路徑由左側 sidebar 輸入，此處顯示
        img_path = seat_img_path.strip()
        if img_path:
            try:
                st.image(img_path, use_container_width=True, caption=f"座位表：{img_path}")
            except Exception as e:
                st.warning(f"⚠️ 無法載入圖片（{img_path}）：{e}")
        else:
            st.info("請在左側側邊欄輸入圖片路徑（例如：seating_101.jpg），圖片將自動縮放以適應螢幕。")


# ══════════════════════════════════════════════════════════════════════
#  Tab 5：配課總表（完整 Excel 原始資料檢視 + 關鍵字搜尋）
# ══════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("🔍 教師配課總表")
    st.caption(f"顯示 {sel_semester}｜{sel_school} 的所有配課記錄，支援關鍵字搜尋。")

    kw = st.text_input("🔍 關鍵字搜尋（姓名 / 學科 / 特殊身分 / 班級）", key="kw_all")

    # 顯示用 DataFrame（去除 _int 衍生欄位）
    display_cols = ["學科", "老師姓名", "兼任職務", "任教年段", "任教班級", "教科書版本", "特殊身分備註", "班級人數"]
    df_show = df[[c for c in display_cols if c in df.columns]].copy()
    # 補上缺少的欄位
    for col in display_cols:
        if col not in df_show.columns:
            df_show[col] = "—"
    # 空字串顯示為 —
    df_show = df_show.replace({"": "—"})

    if kw:
        mask = df_show.apply(lambda r: kw in r.to_string(), axis=1)
        df_show = df_show[mask]

    st.dataframe(df_show, use_container_width=True, hide_index=True)
    st.caption(f"共 {len(df_show)} 筆記錄")

    # 匯出目前篩選結果為 CSV
    csv_bytes = df_show.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="⬇️ 匯出目前結果（CSV）",
        data=csv_bytes,
        file_name=f"{sel_school}_{sel_semester}_配課總表.csv",
        mime="text/csv",
    )