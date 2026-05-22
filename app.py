"""
學校資料管理與情境式查詢系統 v3
── Excel 驅動 · 完全重寫版 ──

執行方式：
    pip install streamlit pandas openpyxl
    streamlit run app.py

Excel 欄位（固定 9 欄，順序不限但名稱須完全一致）：
    學校名稱 | 學年度學期 | 學科 | 老師姓名 | 任教班級
    任教年段 | 教科書版本 | 特殊身分備註 | 班級人數
"""

import streamlit as st
import pandas as pd

# ──────────────────────────────────────────────
#  常數
# ──────────────────────────────────────────────
REQUIRED_COLS = [
    "學校名稱", "學年度學期", "學科", "老師姓名",
    "任教班級", "任教年段", "教科書版本", "特殊身分備註", "班級人數",
]

SUBJECTS = [
    "國文", "英文", "數學", "物理", "化學",
    "生物", "地球科學", "歷史", "地理", "公民與社會",
]

BADGE_COLORS = {
    "學科召集人": "#1565C0",
    "年級召集人": "#2E7D32",
    "行政組長":   "#E65100",
    "代理":       "#6A1B9A",
    "兼課":       "#00695C",
    "留職停薪":   "#C62828",
}

# ──────────────────────────────────────────────
#  頁面設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="學校資料管理系統",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
#  全域 CSS（Mobile-First）
# ──────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif;
}
.header-bar {
    background: linear-gradient(135deg, #1565C0, #0D47A1);
    color: white;
    padding: 0.85rem 1.1rem 0.7rem;
    border-radius: 12px;
    margin-bottom: 1rem;
}
.header-bar h1 { font-size: 1.2rem; margin: 0; }
.header-bar p  { font-size: 0.8rem; margin: 0.15rem 0 0; opacity: 0.85; }

/* 老師卡片 */
.t-card {
    background: #fff;
    border: 1px solid #E0E7FF;
    border-left: 4px solid #3F51B5;
    border-radius: 10px;
    padding: 0.75rem 1rem;
    margin-bottom: 0.5rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
.t-card .t-name { font-size: 1.05rem; font-weight: 700; color: #1A237E; }
.t-card .t-meta { font-size: 0.82rem; color: #555; margin-top: 4px; line-height: 1.6; }

/* Badge */
.badge {
    display: inline-block;
    padding: 1px 8px;
    border-radius: 10px;
    font-size: 0.72rem;
    color: #fff;
    margin-left: 6px;
    vertical-align: middle;
}

/* 摘要塊 */
.sum-box {
    background: #E3F2FD;
    border-left: 4px solid #1565C0;
    border-radius: 6px;
    padding: 0.5rem 1rem;
    margin: 0.6rem 0;
    font-size: 0.9rem;
}
.warn-box {
    background: #FFF3E0;
    border-left: 4px solid #E65100;
    border-radius: 6px;
    padding: 0.5rem 1rem;
    margin: 0.4rem 0;
    font-size: 0.88rem;
    color: #BF360C;
}

/* 手機友善 */
@media (max-width: 768px) {
    .stSelectbox > div > div { min-height: 44px; font-size: 1rem; }
    .stNumberInput input     { min-height: 44px; font-size: 1rem; }
    .stButton > button       { min-height: 44px; width: 100%; }
    .header-bar h1           { font-size: 1rem; }
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
#  工具函式
# ──────────────────────────────────────────────
def safe(val) -> str:
    """NaN / None / 空值 → 空字串，其餘轉 str 並去除首尾空白"""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val).strip()


def make_badge(label: str) -> str:
    """產生彩色 HTML badge"""
    if not label:
        return ""
    color = BADGE_COLORS.get(label, "#607D8B")
    return f'<span class="badge" style="background:{color};">{label}</span>'


def parse_classes(cell: str) -> list[str]:
    """
    將任教班級儲存格字串拆解為班級清單。
    支援全形/半形逗號、頓號、空格等分隔符。
    """
    if not cell:
        return []
    import re
    parts = re.split(r"[,，、\s]+", cell)
    return [p.strip() for p in parts if p.strip()]


def sort_classes(cls_list: list[str]) -> list[str]:
    """
    班級排序：純數字者按數字大小，其餘按字串排序。
    例如：['312','101','資三甲'] → ['101','312','資三甲']
    """
    def key(s):
        try:
            return (0, int(s))
        except ValueError:
            return (1, s)
    return sorted(cls_list, key=key)


def grade_contains(cell: str, target: str) -> bool:
    """判斷任教年段欄位是否包含目標年段（模糊比對）"""
    mapping = {
        "高一": ["高一", "一年", "Grade1", "grade1", "1年"],
        "高二": ["高二", "二年", "Grade2", "grade2", "2年"],
        "高三": ["高三", "三年", "Grade3", "grade3", "3年"],
    }
    cell_s = safe(cell)
    if target in cell_s:
        return True
    for kw in mapping.get(target, []):
        if kw in cell_s:
            return True
    return False


@st.cache_data(show_spinner="📂 讀取 Excel 中…")
def load_excel(file_bytes: bytes, file_name: str) -> pd.DataFrame | None:
    """
    讀取 Excel 並做防呆處理：
    - 全部欄位讀為字串
    - NaN → 空字串
    - 班級人數額外產生整數版 _班級人數_int
    - 驗證必要欄位是否存在
    """
    import io
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, engine="openpyxl")
    except Exception as e:
        st.sidebar.error(f"❌ Excel 讀取失敗：{e}")
        return None

    # 去除欄位名稱多餘空白
    df.columns = [c.strip() for c in df.columns]

    # 檢查必要欄位
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.sidebar.error(f"❌ Excel 缺少欄位：{', '.join(missing)}")
        return None

    # 全欄 NaN → 空字串
    df = df.fillna("").astype(str)
    # astype(str) 會把 nan 轉成 'nan' 字串，再清掉
    df = df.replace({"nan": "", "NaN": "", "None": ""})
    # 去除所有儲存格首尾空白（雙保險：先全欄，再對關鍵比對欄位單獨處理）
    df = df.apply(lambda col: col.str.strip())
    # 明確對「學校名稱」「學年度學期」做額外 strip，防止隱藏空白字元
    df["學校名稱"]   = df["學校名稱"].str.strip()
    df["學年度學期"] = df["學年度學期"].str.strip()

    # 班級人數整數欄（供計算用）
    def to_int(v: str) -> int:
        try:
            return int(float(v)) if v else 0
        except Exception:
            return 0
    df["_班級人數_int"] = df["班級人數"].apply(to_int)

    return df


# ──────────────────────────────────────────────
#  Sidebar：上傳 + 全域切換
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 上傳 Excel 總表")
    uploaded = st.file_uploader(
        "選擇 .xlsx 檔案",
        type=["xlsx"],
        help="必要欄位：學校名稱、學年度學期、學科、老師姓名、任教班級、任教年段、教科書版本、特殊身分備註、班級人數",
    )

    df_raw: pd.DataFrame | None = None
    if uploaded is not None:
        # 用 file_id（name+size）做快取 key，避免同名不同檔無法重新載入
        file_bytes = uploaded.read()
        df_raw = load_excel(file_bytes, uploaded.name)
        if df_raw is not None:
            st.success(f"✅ 已載入 **{len(df_raw)}** 筆資料")

    st.divider()

    # ── 全域切換選單（從 Excel 動態抓取） ──
    st.markdown("## ⚙️ 全域切換")

    if df_raw is not None:
        school_options   = sorted(set(df_raw["學校名稱"].str.strip().unique().tolist()))
        semester_options = sorted(set(df_raw["學年度學期"].str.strip().unique().tolist()))
    else:
        school_options   = ["（請先上傳 Excel）"]
        semester_options = ["（請先上傳 Excel）"]

    sel_school   = st.selectbox("🏫 選擇學校",    school_options,   key="sel_school")
    sel_semester = st.selectbox("📅 選擇學年度/學期", semester_options, key="sel_semester")

    st.divider()
    st.markdown("## 🖼️ 座位表圖片路徑")
    seat_img_path = st.text_input(
        "圖片相對/絕對路徑",
        placeholder="例：seating_101.jpg",
        key="seat_img_path",
    )


# ──────────────────────────────────────────────
#  Header
# ──────────────────────────────────────────────
st.markdown(f"""
<div class="header-bar">
  <h1>🏫 學校資料管理與情境式查詢系統</h1>
  <p>{sel_semester} ｜ {sel_school} ── Mobile-First · Excel 驅動</p>
</div>
""", unsafe_allow_html=True)

# ── 未上傳時顯示說明並停止 ──
if df_raw is None:
    st.info("👈 請先在左側側邊欄上傳 Excel 總表（.xlsx），系統即自動載入所有功能。")
    st.markdown("""
### 📋 Excel 欄位格式說明
| 欄位名稱 | 說明 | 範例 |
|---|---|---|
| 學校名稱 | 學校全名 | 國立中大壢中 |
| 學年度學期 | 學期代碼 | 114上 |
| 學科 | 科目名稱 | 數學 |
| 老師姓名 | 老師全名 | 王大明 |
| 任教班級 | 逗號分隔班級 | 101, 102, 201 |
| 任教年段 | 年段文字 | 高一 |
| 教科書版本 | 出版社 | 龍騰 |
| 特殊身分備註 | 特殊身分 | 學科召集人 |
| 班級人數 | 整數，可空白 | 32 |

> 💡 空白欄位系統自動以 `—` 顯示，不會造成程式錯誤。
    """)
    st.stop()

# ── 依全域選單過濾出當前學校 + 學期的資料 ──
# 過濾時兩邊都 strip，徹底防止隱藏空白導致比對失敗
_school_clean   = sel_school.strip()
_semester_clean = sel_semester.strip()
df: pd.DataFrame = df_raw[
    (df_raw["學校名稱"].str.strip() == _school_clean) &
    (df_raw["學年度學期"].str.strip() == _semester_clean)
].copy()

if df.empty:
    st.warning(
        f"⚠️ 在 Excel 中找不到「**{sel_school}**」×「**{sel_semester}**」的資料，"
        "請確認學校名稱與學期是否正確。"
    )
    st.stop()

# ──────────────────────────────────────────────
#  三個功能標籤頁
# ──────────────────────────────────────────────
tab_table, tab_find, tab_pack = st.tabs([
    "📄 配課總表",
    "🔍 按年段找老師",
    "📦 打包點收計算器",
])


# ══════════════════════════════════════════════
#  頁籤一：配課總表
# ══════════════════════════════════════════════
with tab_table:
    st.subheader("📄 配課總表")
    st.caption(f"資料來源：{sel_school}｜{sel_semester}，共 {len(df)} 筆")

    # ── 篩選控制項 ──
    col_kw, col_subj, col_grade = st.columns([2, 1, 1])

    with col_kw:
        kw = st.text_input("🔍 關鍵字搜尋（老師姓名 / 班級 / 備註）", key="t1_kw")

    with col_subj:
        subj_opts = ["全部"] + [s for s in SUBJECTS if s in df["學科"].values]
        sel_subj_t1 = st.selectbox("學科篩選", subj_opts, key="t1_subj")

    with col_grade:
        grade_opts = ["全部", "高一", "高二", "高三"]
        sel_grade_t1 = st.selectbox("年段篩選", grade_opts, key="t1_grade")

    # ── 套用篩選 ──
    df_show = df[REQUIRED_COLS].copy()          # 只取 9 個必要欄位，順序固定
    df_show = df_show.replace({"": "—"})       # 空字串顯示為 —

    if kw:
        mask = df_show.apply(lambda r: kw.lower() in r.to_string().lower(), axis=1)
        df_show = df_show[mask]

    if sel_subj_t1 != "全部":
        df_show = df_show[df_show["學科"] == sel_subj_t1]

    if sel_grade_t1 != "全部":
        # 年段篩選：對原始 df 的任教年段欄做模糊比對
        grade_mask = df.loc[df_show.index, "任教年段"].apply(
            lambda v: grade_contains(v, sel_grade_t1)
        )
        df_show = df_show[grade_mask]

    st.dataframe(df_show, use_container_width=True, hide_index=True)
    st.caption(f"篩選後共 **{len(df_show)}** 筆")

    # 匯出 CSV
    csv_bytes = df_show.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        "⬇️ 匯出目前結果（CSV）",
        data=csv_bytes,
        file_name=f"{sel_school}_{sel_semester}_配課總表.csv",
        mime="text/csv",
    )


# ══════════════════════════════════════════════
#  頁籤二：按年段找老師（發放情境）
# ══════════════════════════════════════════════
with tab_find:
    st.subheader("🔍 按年段找老師")

    col_s, col_g = st.columns(2)
    with col_s:
        avail_subj = [s for s in SUBJECTS if s in df["學科"].values]
        sel_subj_t2 = st.selectbox("學科", avail_subj or SUBJECTS, key="t2_subj")
    with col_g:
        sel_grade_t2 = st.selectbox("年段", ["高一", "高二", "高三"], key="t2_grade")

    # ── 過濾：學科 + 年段（模糊） ──
    matched = df[
        (df["學科"] == sel_subj_t2) &
        df["任教年段"].apply(lambda v: grade_contains(v, sel_grade_t2))
    ]

    if matched.empty:
        st.info("查無符合資料，請確認 Excel 中的學科與任教年段欄位內容。")
    else:
        # ── groupby 合併：同名老師 → 一張卡片 ──
        def merge_teacher_rows(grp: pd.DataFrame) -> pd.Series:
            def first_nonempty(series):
                for v in series:
                    s = safe(v)
                    if s and s != "—":
                        return s
                return "—"

            # 拆解所有班級儲存格 → 合併去重 → 排序
            all_cls: list[str] = []
            for cell in grp["任教班級"]:
                all_cls.extend(parse_classes(safe(cell)))
            merged_cls = ", ".join(sort_classes(list(set(all_cls)))) or "—"

            return pd.Series({
                "特殊身分備註": first_nonempty(grp["特殊身分備註"]),
                "教科書版本":   first_nonempty(grp["教科書版本"]),
                "任教年段":     first_nonempty(grp["任教年段"]),
                "任教班級_合併": merged_cls,
            })

        merged = (
            matched
            .groupby("老師姓名", sort=True)
            .apply(merge_teacher_rows)
            .reset_index()
        )

        st.caption(f"符合「{sel_grade_t2}｜{sel_subj_t2}」共 **{len(merged)}** 位老師")

        for _, row in merged.iterrows():
            sp      = safe(row.get("特殊身分備註", ""))
            ver     = safe(row.get("教科書版本", ""))
            classes = safe(row.get("任教班級_合併", ""))
            bh      = make_badge(sp) if sp and sp != "—" else ""

            # 備註行：只有有值才顯示
            badge_line = f"身分備註：{sp}" if (sp and sp != "—") else ""

            st.markdown(f"""
            <div class="t-card">
              <div class="t-name">{safe(row['老師姓名'])}{bh}</div>
              <div class="t-meta">
                任教班級：{classes}<br>
                用書版本：{ver if ver and ver != "—" else "（未填）"}<br>
                {"身分備註：" + sp if (sp and sp != "—") else ""}
              </div>
            </div>
            """, unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  頁籤三：打包點收計算器
# ══════════════════════════════════════════════
with tab_pack:
    st.subheader("📦 打包點收計算器")

    # 全域：每人發放本數
    col_p1, col_p2 = st.columns([1, 2])
    with col_p1:
        copies_per_person = st.number_input(
            "每人發放本數",
            min_value=1, value=1, step=1,
            key="copies_per_person",
            help="預設 1，可改為 2（例如同時發兩份）",
        )
    with col_p2:
        st.markdown(
            '<div class="sum-box" style="margin-top:1.6rem;">'
            '調整左方「每人發放本數」後，下方總計即時更新。</div>',
            unsafe_allow_html=True,
        )

    st.divider()

    # ── 從 df 拆解出所有唯一班級，並對應人數 ──
    # 每個班級取「班級人數_int 最大非零值」作為預設，避免同班不同老師重複列
    class_count_raw: dict[str, int] = {}  # 班級 → Excel 原始人數
    for _, row in df.iterrows():
        for cls in parse_classes(safe(row["任教班級"])):
            if cls not in class_count_raw:
                class_count_raw[cls] = int(row["_班級人數_int"])
            else:
                # 若已有值且為 0，嘗試用非零覆蓋
                if class_count_raw[cls] == 0 and int(row["_班級人數_int"]) > 0:
                    class_count_raw[cls] = int(row["_班級人數_int"])

    sorted_classes = sort_classes(list(class_count_raw.keys()))

    if not sorted_classes:
        st.info("目前篩選條件下沒有任何班級資料，請確認 Excel 的任教班級欄位。")
        st.stop()

    # ── session_state 初始化（保存手動輸入的暫時人數） ──
    # key 格式：pack_{學校}_{學期}_{班級}，避免切換學校時數值錯亂
    ss_prefix = f"pack_{sel_school}_{sel_semester}_"

    # 統計警告班級數
    warn_classes: list[str] = []

    # ── 表頭 ──
    hc = st.columns([2, 2, 3, 3])
    hc[0].markdown("**班級**")
    hc[1].markdown("**人數**")
    hc[2].markdown("**本次發放本數**")
    hc[3].markdown("**狀態**")

    grand_total = 0

    for cls in sorted_classes:
        base_count = class_count_raw[cls]
        ss_key = f"{ss_prefix}{cls}"

        # session_state 初始化（第一次才設）
        if ss_key not in st.session_state:
            st.session_state[ss_key] = base_count

        c0, c1, c2, c3 = st.columns([2, 2, 3, 3])

        with c0:
            st.write(f"**{cls}**")

        with c1:
            if base_count == 0:
                # 人數空白：提供手動輸入框
                warn_classes.append(cls)
                manual_count = st.number_input(
                    f"輸入人數_{cls}",
                    min_value=0,
                    value=st.session_state[ss_key],
                    step=1,
                    key=ss_key,
                    label_visibility="collapsed",
                )
                effective_count = manual_count
            else:
                # Excel 有值：顯示數字（仍允許手動覆蓋）
                manual_count = st.number_input(
                    f"人數_{cls}",
                    min_value=0,
                    value=base_count,
                    step=1,
                    key=ss_key,
                    label_visibility="collapsed",
                )
                effective_count = manual_count

        with c2:
            pack_qty = effective_count * copies_per_person
            st.write(f"**{pack_qty}** 本")

        with c3:
            if base_count == 0:
                st.markdown(
                    '<span style="color:#C62828;font-weight:600;">'
                    '⚠️ 尚未輸入人數</span>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<span style="color:#2E7D32;">✅ 已有人數</span>',
                    unsafe_allow_html=True,
                )

        grand_total += pack_qty

    st.divider()

    # ── 警告彙整 ──
    if warn_classes:
        st.markdown(
            f'<div class="warn-box">⚠️ 下列班級在 Excel 中人數為空白，'
            f'請在上方手動輸入：<b>{", ".join(warn_classes)}</b></div>',
            unsafe_allow_html=True,
        )

    # ── 大字體總計 ──
    st.metric(
        label=f"📦 {sel_school}｜{sel_semester} 總打包本數",
        value=f"{grand_total} 本",
        delta=f"共 {len(sorted_classes)} 個班級 × {copies_per_person} 本/人",
    )
