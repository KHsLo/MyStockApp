# ============================================================
# app.py：股票篩選器（Streamlit）
#   - Tab 1：日線模式（讀 simply_report_YYYYMMDD_daily.xlsx）
#   - Tab 2：週線模式（讀 simply_report_YYYYMMDD_weekly.xlsx）
#   - Tab 3：日週交集（Tab 1 ∩ Tab 2，只顯示基本 4 欄）
#   - Tab 4：日線3線糾結 + 紅K棒
#   - Tab 5：週線3線糾結 + 紅K棒
#   - Tab 6：日週3線糾結交集
#   - Tab 7：日線均線色帶（滿足最後交易日均線色帶條件 == 'O'）
#   - Tab 8：週線均線色帶
#   - Tab 9：日週均線色帶交集
#   - 側邊欄：全域日期選擇
#   - 篩選條件：勾選欄位 → 展開該欄位可選值 → 多選/範圍
#   - 日期欄位 6 個互斥，只能選一個；預設日期起迄為「當天與當天」
# ============================================================
import os
import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
# 直接把根目錄當作讀取區，不再往內尋找 scanner_output 資料夾
OUTPUT_DIR = CURRENT_DIR


# =========================
# 欄位定義
# =========================

# 文字欄位（關鍵字搜尋，預設不勾選）
TEXT_FIELDS = ['代號', '公司名稱', '產業類別']

# 日期欄位（互斥，只能選一個）
DATE_FIELDS = [
    '最近MJ交點', 'stochRSI最近交叉日期',
    'ut bot買入訊號日期', 'ut bot賣出訊號日期',
    'HSI金叉日期', 'HSI死叉日期',
]

# 類別型欄位：(欄位名, 可選值清單, 預設勾選, 預設選中值)
CATEGORICAL_FIELDS = [
    ('KAMA10>20>30', ['O', 'X'], True, ['O']),
    ('收盤價>KAMA10', ['O', 'X'], True, ['O']),
    ('趨勢線突破', ['O', '△', 'X'], True, ['O']),
    ('EHMA趨勢', ['綠帶(多)', '紅帶(空)'], True, ['綠帶(多)']),
    ('StochRSI_交點等級',
     ['極度超賣', '超賣', '弱勢反彈／超賣修復', '偏弱', '中性',
      '偏強', '強勢', '超買／強勢', '極度超買', '資料不足'],
     True,
     ['極度超賣', '超賣', '弱勢反彈／超賣修復', '偏弱', '中性', '偏強']),
    ('stochRSI最近交叉方向', ['金叉', '死叉'], True, ['金叉']),
    ('StochRSI_K_趨勢',
     ['明顯上升', '緩步上升', '持平', '緩步下降', '明顯下降'],
     True, ['明顯上升', '緩步上升', '持平']),
    ('MACD紅綠柱', ['紅柱', '綠柱'], True, ['紅柱']),
    ('J線位置', ['>0', '<0'], True, ['>0']),
    ('J線方向', ['上升', '下降', '持平'], True, ['上升', '持平']),
    ('MFI方向', ['上升', '下降', '持平'], True, ['上升', '持平']),
    ('OBV方向', ['上升', '下降', '持平'], True, ['上升', '持平']),
    ('型態等級', ['強', '中', '弱'], False, ['強']),
    ('底底高', ['O', 'X'], False, ['O']),
    ('突破狀態',
     ['已突破', '接近', '未突破', '已跌破', '未接近', '未跌破'],
     False, ['已突破', '接近']),
     # ★ 新增：均線色帶條件（可篩選）
    ('滿足最後交易日均線色帶條件', ['O', 'X'], False, ['O']),
    # ★ 新增：CCI 穿越 100（可篩選）
    ('滿足最後交易日CCI穿越100', ['O', 'X'], False, ['O']),
]

# 數值型欄位（共通）：(欄位名, 最小值, 最大值, 預設值, 步進, 預設勾選)
NUMERIC_FIELDS_COMMON = [
    ('實收資本額（億）', 0.0, 10000.0, (0.0, 10000.0), 0.5, False),
    ('Hurst Exponent', 0.0, 100.0, (50.0, 100.0), 1.0, True),
    ('STC數值', -100.0, 200.0, (0.0, 80.0), 1.0, True),
    ('StochRSI_K', -100.0, 200.0, (0.0, 80.0), 1.0, True),
    ('RSI（14）', 0.0, 100.0, (0.0, 80.0), 1.0, True),
    ('當時J值', -100.0, 200.0, (0.0, 80.0), 1.0, True),
]

# 數值型欄位（換手率，日線/週線不同）
NUMERIC_FIELDS_TURNOVER_DAILY = [
    ('單日換手率％', 0.0, 100.0, (0.0, 100.0), 0.1, False),
    ('5期均換手率％', 0.0, 100.0, (0.0, 100.0), 0.1, False),
    ('20期均換手率％', 0.0, 100.0, (0.0, 100.0), 0.1, False),
    ('換手率5期比', 0.0, 10.0, (0.0, 10.0), 0.1, False),
    ('換手率20期比', 0.0, 10.0, (0.0, 10.0), 0.1, False),
]
NUMERIC_FIELDS_TURNOVER_WEEKLY = [
    ('單週換手率％', 0.0, 100.0, (0.0, 100.0), 0.1, False),
    ('5期均換手率％', 0.0, 100.0, (0.0, 100.0), 0.1, False),
    ('20期均換手率％', 0.0, 100.0, (0.0, 100.0), 0.1, False),
    ('換手率5期比', 0.0, 10.0, (0.0, 10.0), 0.1, False),
    ('換手率20期比', 0.0, 10.0, (0.0, 10.0), 0.1, False),
]


def get_numeric_fields(period_label):
    if period_label == 'daily':
        return NUMERIC_FIELDS_COMMON + NUMERIC_FIELDS_TURNOVER_DAILY
    return NUMERIC_FIELDS_COMMON + NUMERIC_FIELDS_TURNOVER_WEEKLY


# =========================
# 資料讀取
# =========================
def list_simply_dates():
    """列出所有 available 的日期（YYYYMMDD）"""
    if not os.path.exists(OUTPUT_DIR):
        return []
    pat = re.compile(r'^simply_report_(\d{8})_daily\.xlsx$')
    dates = set()
    for f in os.listdir(OUTPUT_DIR):
        m = pat.match(f)
        if m:
            dates.add(m.group(1))
    return sorted(dates, reverse=True)


@st.cache_data(show_spinner=False)
def load_simply(date_str, period_label):
    path = os.path.join(OUTPUT_DIR, f"simply_report_{date_str}_{period_label}.xlsx")
    if not os.path.exists(path):
        return None
    try:
        return pd.read_excel(path, sheet_name=0)
    except Exception:
        return None


# =========================
# Arrow 相容處理
# =========================
def _arrow_safe(df):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            non_null = df[c].dropna()
            if len(non_null) == 0:
                continue
            types = non_null.map(type).nunique()
            if types > 1:
                df[c] = df[c].apply(lambda x: '' if pd.isna(x) else str(x))
    return df


# =========================
# 取消全選 回呼
# =========================
def _make_deselect_all_cb(keys):
    """建立「取消全部勾選」的回呼函式。"""
    def cb():
        for k in keys:
            st.session_state[k] = False
    return cb


# =========================
# 篩選 UI 構建（在主區域，不在 sidebar）
# =========================
def build_filter_ui(df, period_label, key_prefix):
    """
    在主區域構建篩選 UI，回傳 filters dict。
    每個欄位前有 checkbox，勾選才展開可選值。
    日期欄位互斥，共用一個 radio。
    頂部提供「取消全部勾選」按鈕。
    日期預設起迄為「當天與當天」。
    """
    filters = {}

    # ---- 收集所有 checkbox 的 key（用於「取消全部勾選」）----
    all_cb_keys = []
    for field in TEXT_FIELDS:
        if field in df.columns:
            all_cb_keys.append(f'{key_prefix}_cb_text_{field}')
    for field, _opts, _dc, _dv in CATEGORICAL_FIELDS:
        if field in df.columns:
            all_cb_keys.append(f'{key_prefix}_cb_cat_{field}')
    for field, _lo, _hi, _dv, _st, _dc in get_numeric_fields(period_label):
        if field in df.columns:
            all_cb_keys.append(f'{key_prefix}_cb_num_{field}')

    # ---- 初始化：若 key 不存在，依預設值寫入 ----
    for field, _opts, default_checked, _dv in CATEGORICAL_FIELDS:
        if field in df.columns:
            k = f'{key_prefix}_cb_cat_{field}'
            if k not in st.session_state:
                st.session_state[k] = default_checked
    for field, _lo, _hi, _dv, _st, default_checked in get_numeric_fields(period_label):
        if field in df.columns:
            k = f'{key_prefix}_cb_num_{field}'
            if k not in st.session_state:
                st.session_state[k] = default_checked
    for field in TEXT_FIELDS:
        if field in df.columns:
            k = f'{key_prefix}_cb_text_{field}'
            if k not in st.session_state:
                st.session_state[k] = False

    # ---- 全域操作按鈕 ----
    btn_col1, btn_col2 = st.columns([1, 3])
    with btn_col1:
        st.button(
            "🧹 取消全部勾選",
            key=f'{key_prefix}_deselect_all',
            on_click=_make_deselect_all_cb(all_cb_keys),
            use_container_width=True,
        )
    with btn_col2:
        st.caption("💡 點擊後會取消所有已勾選的篩選欄位，再按下方「開始執行」重新篩選")

    # ---- 文字欄位 ----
    with st.expander("📝 文字欄位（關鍵字搜尋）", expanded=False):
        for field in TEXT_FIELDS:
            if field not in df.columns:
                continue
            cb_key = f'{key_prefix}_cb_text_{field}'
            is_checked = st.checkbox(field, key=cb_key)
            if is_checked:
                val = st.text_input(
                    f"{field} 關鍵字",
                    key=f'{key_prefix}_text_{field}',
                    placeholder='留空則不篩選',
                )
                if val and val.strip():
                    filters[field] = val.strip()

    # ---- 日期欄位（互斥）----
    date_fields_present = [f for f in DATE_FIELDS if f in df.columns]
    if date_fields_present:
        with st.expander("📅 日期欄位（僅能選一個）", expanded=False):
            selected = st.radio(
                "選擇要篩選的日期欄位",
                options=['不啟用'] + date_fields_present,
                index=0,
                horizontal=True,
                key=f'{key_prefix}_date_field_radio',
            )
            if selected != '不啟用':
                # ★ 預設日期起迄改為「當天與當天」
                today = datetime.today().date()
                default_range = (today, today)
                d = st.date_input(
                    f"{selected} 日期區間",
                    value=default_range,
                    key=f'{key_prefix}_date_range_{selected}',
                )
                if isinstance(d, tuple) and len(d) == 2:
                    filters['_date_field'] = selected
                    filters['_date_range'] = d

    # ---- 類別型欄位 ----
    with st.expander("🏷️ 類別型欄位", expanded=True):
        for field, options, default_checked, default_val in CATEGORICAL_FIELDS:
            if field not in df.columns:
                continue
            cb_key = f'{key_prefix}_cb_cat_{field}'
            is_checked = st.checkbox(field, key=cb_key)
            if is_checked:
                safe_default = [v for v in default_val if v in options]
                val = st.multiselect(
                    f"{field} 選項",
                    options=options,
                    default=safe_default,
                    key=f'{key_prefix}_cat_{field}',
                    label_visibility='collapsed',
                )
                if val:
                    filters[field] = val

    # ---- 數值型欄位 ----
    with st.expander("🔢 數值型欄位", expanded=True):
        for field, lo, hi, default_val, step, default_checked in get_numeric_fields(period_label):
            if field not in df.columns:
                continue
            cb_key = f'{key_prefix}_cb_num_{field}'
            is_checked = st.checkbox(field, key=cb_key)
            if is_checked:
                val = st.slider(
                    f"{field} 範圍",
                    min_value=float(lo),
                    max_value=float(hi),
                    value=(float(default_val[0]), float(default_val[1])),
                    step=float(step),
                    key=f'{key_prefix}_num_{field}',
                    label_visibility='collapsed',
                )
                filters[field] = val

    return filters


# =========================
# 篩選邏輯
# =========================
def apply_filters(df, filters):
    if df is None or df.empty:
        return df
    mask = pd.Series([True] * len(df), index=df.index)

    # 日期欄位
    if '_date_field' in filters and '_date_range' in filters:
        date_col = filters['_date_field']
        start, end = filters['_date_range']
        if date_col in df.columns:
            s = pd.to_datetime(df[date_col], errors='coerce')
            mask = mask & s.notna() & \
                   (s >= pd.to_datetime(start)) & \
                   (s <= pd.to_datetime(end))

    # 其他欄位
    for field, val in filters.items():
        if field.startswith('_'):
            continue
        if field not in df.columns:
            continue
        if isinstance(val, str):
            mask = mask & df[field].astype(str).str.contains(
                val, case=False, na=False, regex=False)
        elif isinstance(val, list):
            mask = mask & df[field].astype(str).isin([str(v) for v in val])
        elif isinstance(val, tuple) and len(val) == 2:
            s = pd.to_numeric(df[field], errors='coerce')
            mask = mask & s.notna() & (s >= val[0]) & (s <= val[1])

    return df[mask].copy()


# =========================
# 結果顯示
# =========================
def _calc_table_height(n_rows, row_px=35, header_px=40, min_h=80, max_h=800):
    """依列數動態計算 dataframe 高度，避免大量空白列。"""
    if n_rows <= 0:
        return min_h
    h = header_px + row_px * n_rows
    return int(min(max(h, min_h), max_h))


def _build_pinned_config(df, extra_pinned=None):
    """建立凍結欄位 config；只保留 df 實際存在的欄位。"""
    pinned_cfg = {
        '代號':            st.column_config.TextColumn('代號', pinned=True),
        '公司名稱':        st.column_config.TextColumn('公司名稱', pinned=True),
        '產業類別':        st.column_config.TextColumn('產業類別', pinned=True),
        '實收資本額（億）': st.column_config.NumberColumn('實收資本額（億）', pinned=True),
    }
    if extra_pinned:
        for k, v in extra_pinned.items():
            pinned_cfg[k] = v
    return {k: v for k, v in pinned_cfg.items() if k in df.columns}


def render_result(df, key_prefix, extra_pinned=None, sort_by=None, sort_ascending=True):
    if df is None or df.empty:
        st.warning("🔍 查無符合條件的股票")
        return

    # ★ 排序（若有指定）
    if sort_by and sort_by in df.columns:
        df = df.sort_values(sort_by, ascending=sort_ascending,
                            na_position='last').reset_index(drop=True)

    # ============================================================
    # ★ 關鍵：把所有「要凍結的欄位」推到 DataFrame 最左邊，
    #   且順序與 column_config 完全一致，Streamlit 才會全部凍結。
    # ============================================================
    pinned_order = ['代號', '公司名稱', '產業類別', '實收資本額（億）']
    if extra_pinned:
        for k in extra_pinned.keys():
            if k not in pinned_order:
                pinned_order.append(k)
    pinned_order = [c for c in pinned_order if c in df.columns]

    if pinned_order:
        other_cols = [c for c in df.columns if c not in pinned_order]
        df = df[pinned_order + other_cols]

    st.success(f"✅ 共 {len(df)} 檔符合")

    column_config = _build_pinned_config(df, extra_pinned=extra_pinned)

    st.data_editor(
        _arrow_safe(df),
        width='stretch',
        height=_calc_table_height(len(df)),
        column_config=column_config,
        column_order=list(df.columns),
        hide_index=True,
        disabled=True,  # 设为只读
        key=f'{key_prefix}_editor_{datetime.now().timestamp()}',  # 确保 key 唯一
    )

    csv = df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "📥 匯出篩選結果（CSV）",
        data=csv,
        file_name=f"filtered_{key_prefix}_{datetime.today().strftime('%Y%m%d')}.csv",
        mime='text/csv',
        key=f'{key_prefix}_download_csv',
    )


# =========================
# Tab 1：日線
# =========================
def render_tab_daily():
    st.header("🔵 日線模式")
    period_label = 'daily'
    key_prefix = 'daily'

    date_str = st.session_state.get('selected_date')
    if not date_str:
        st.info("👈 請先在左側選擇日期")
        return

    df = load_simply(date_str, period_label)
    if df is None:
        st.error(f"❌ 讀不到 simply_report_{date_str}_{period_label}.xlsx")
        return

    st.caption(f"📂 simply_report_{date_str}_{period_label}.xlsx（共 {len(df)} 檔）")

    filters = build_filter_ui(df, period_label, key_prefix)

    run_clicked = st.button(
        "▶️ 開始執行（日線）",
        type='primary',
        key=f'{key_prefix}_run',
        use_container_width=True,
    )

    if run_clicked:
        result = apply_filters(df, filters)
        st.session_state['daily_result'] = result
        st.session_state['daily_date'] = date_str

    stored_date = st.session_state.get('daily_date')
    if stored_date and stored_date != date_str:
        st.warning(f"⚠️ 你儲存的結果日期是 {stored_date}，"
                   f"目前側邊欄選擇的日期是 {date_str}，請重新執行篩選")

    if 'daily_result' in st.session_state:
        st.markdown("---")
        render_result(st.session_state['daily_result'], key_prefix)
    else:
        st.info("👆 設定完篩選條件後，點選「開始執行」")


# =========================
# Tab 2：週線
# =========================
def render_tab_weekly():
    st.header("🟢 週線模式")
    period_label = 'weekly'
    key_prefix = 'weekly'

    date_str = st.session_state.get('selected_date')
    if not date_str:
        st.info("👈 請先在左側選擇日期")
        return

    df = load_simply(date_str, period_label)
    if df is None:
        st.error(f"❌ 讀不到 simply_report_{date_str}_{period_label}.xlsx")
        return

    st.caption(f"📂 simply_report_{date_str}_{period_label}.xlsx（共 {len(df)} 檔）")

    filters = build_filter_ui(df, period_label, key_prefix)

    run_clicked = st.button(
        "▶️ 開始執行（週線）",
        type='primary',
        key=f'{key_prefix}_run',
        use_container_width=True,
    )

    if run_clicked:
        result = apply_filters(df, filters)
        st.session_state['weekly_result'] = result
        st.session_state['weekly_date'] = date_str

    stored_date = st.session_state.get('weekly_date')
    if stored_date and stored_date != date_str:
        st.warning(f"⚠️ 你儲存的結果日期是 {stored_date}，"
                   f"目前側邊欄選擇的日期是 {date_str}，請重新執行篩選")

    if 'weekly_result' in st.session_state:
        st.markdown("---")
        render_result(st.session_state['weekly_result'], key_prefix)
    else:
        st.info("👆 設定完篩選條件後，點選「開始執行」")


# =========================
# Tab 4/5：3線糾結 + 紅K棒
#   - 日期鎖定為最新交易日
#   - 硬過濾：最後交易日標準3線糾結 非空 且 標準3線糾結K棒型態 == '紅K棒'
#   - 其餘提供與 Tab1/Tab2 相同的完整篩選 UI
# =========================
def render_tab_3line(period_label, key_prefix, title):
    st.header(title)

    # ---- 日期鎖定為最新交易日 ----
    dates = list_simply_dates()
    if not dates:
        st.error("❌ 找不到任何 simply_report 檔案")
        return
    date_str = dates[0]  # 已 reverse=True，取最新
    st.caption(f"📅 日期鎖定為最新交易日：**{date_str}**（此頁面不受側邊欄日期影響）")

    df = load_simply(date_str, period_label)
    if df is None:
        st.error(f"❌ 讀不到 simply_report_{date_str}_{period_label}.xlsx")
        return

    # ---- 檢查欄位是否存在 ----
    col_3line = '最後交易日標準3線糾結'
    col_type  = '標準3線糾結K棒型態'
    missing = [c for c in [col_3line, col_type] if c not in df.columns]
    if missing:
        st.error(
            f"❌ 此日期檔案缺少欄位：{'、'.join(missing)}\n\n"
            f"請用新版（V68 以上）重新產出 simply_report_{date_str}_{period_label}.xlsx"
        )
        return

    # ---- 硬過濾：非空 + 紅K棒 ----
    mask = df[col_3line].notna() & (df[col_type].astype(str).str.strip() == '紅K棒')
    df_filtered = df[mask].copy()

    if df_filtered.empty:
        st.warning("🔍 此日期無符合『3線糾結 + 紅K棒』的個股")
        return

    st.caption(
        f"📂 simply_report_{date_str}_{period_label}.xlsx "
        f"（原始 {len(df)} 檔 → 3線糾結+紅K棒 **{len(df_filtered)} 檔**）"
    )

    # ---- 沿用 Tab1/Tab2 的篩選 UI ----
    filters = build_filter_ui(df_filtered, period_label, key_prefix)

    run_clicked = st.button(
        "▶️ 開始執行",
        type='primary',
        key=f'{key_prefix}_run',
        use_container_width=True,
    )

    if run_clicked:
        result = apply_filters(df_filtered, filters)

        # 把 3line 欄位與基本欄位重排到前面（與釘選順序一致）
        front_cols_order = ['代號', '公司名稱', '產業類別', '實收資本額（億）',
                            col_3line, col_type]
        front_cols_order = [c for c in front_cols_order if c in result.columns]
        rest_cols = [c for c in result.columns if c not in front_cols_order]
        result = result[front_cols_order + rest_cols]

        st.session_state[f'{key_prefix}_result'] = result
        st.session_state[f'{key_prefix}_date'] = date_str

    # ---- 若儲存的結果日期與目前最新日期不同，提示 ----
    stored_date = st.session_state.get(f'{key_prefix}_date')
    if stored_date and stored_date != date_str:
        st.warning(
            f"⚠️ 你儲存的結果日期是 {stored_date}，目前最新交易日是 {date_str}，"
            f"請重新執行篩選"
        )

    # ---- 顯示結果 ----
    if f'{key_prefix}_result' in st.session_state:
        st.markdown("---")
        extra_pinned = {
            col_3line: st.column_config.NumberColumn(
                col_3line, pinned=True, format='%.5f'),
            col_type: st.column_config.TextColumn(
                col_type, pinned=True),
        }
        render_result(
            st.session_state[f'{key_prefix}_result'],
            key_prefix,
            extra_pinned=extra_pinned,
            sort_by=col_3line,
            sort_ascending=True,
        )
    else:
        st.info("👆 設定完篩選條件後，點選「開始執行」")


# =========================
# Tab 7/8：均線色帶（日線 / 週線）
#   - 日期鎖定為最新交易日
#   - 硬過濾：滿足最後交易日均線色帶條件 == 'O'
#   - 其餘提供與 Tab1/Tab2 相同的完整篩選 UI
# =========================
def render_tab_ma_ribbon(period_label, key_prefix, title):
    st.header(title)

    # ---- 日期鎖定為最新交易日 ----
    dates = list_simply_dates()
    if not dates:
        st.error("❌ 找不到任何 simply_report 檔案")
        return
    date_str = dates[0]  # 已 reverse=True，取最新
    st.caption(f"📅 日期鎖定為最新交易日：**{date_str}**（此頁面不受側邊欄日期影響）")

    df = load_simply(date_str, period_label)
    if df is None:
        st.error(f"❌ 讀不到 simply_report_{date_str}_{period_label}.xlsx")
        return

    # ---- 檢查欄位是否存在 ----
    col_ribbon = '滿足最後交易日均線色帶條件'
    if col_ribbon not in df.columns:
        st.error(
            f"❌ 此日期檔案缺少欄位：{col_ribbon}\n\n"
            f"請用新版重新產出 simply_report_{date_str}_{period_label}.xlsx"
        )
        return

    # ---- 硬過濾：== 'O' ----
    mask = df[col_ribbon].astype(str).str.strip() == 'O'
    df_filtered = df[mask].copy()

    if df_filtered.empty:
        st.warning("🔍 此日期無符合『最後交易日均線色帶條件』的個股")
        return

    st.caption(
        f"📂 simply_report_{date_str}_{period_label}.xlsx "
        f"（原始 {len(df)} 檔 → 均線色帶 **{len(df_filtered)} 檔**）"
    )

    # ---- 沿用 Tab1/Tab2 的篩選 UI ----
    filters = build_filter_ui(df_filtered, period_label, key_prefix)

    run_clicked = st.button(
        "▶️ 開始執行",
        type='primary',
        key=f'{key_prefix}_run',
        use_container_width=True,
    )

    if run_clicked:
        result = apply_filters(df_filtered, filters)

        # 把色帶欄位與基本欄位重排到前面（與釘選順序一致）
        front_cols_order = ['代號', '公司名稱', '產業類別', '實收資本額（億）',
                            col_ribbon]
        front_cols_order = [c for c in front_cols_order if c in result.columns]
        rest_cols = [c for c in result.columns if c not in front_cols_order]
        result = result[front_cols_order + rest_cols]

        st.session_state[f'{key_prefix}_result'] = result
        st.session_state[f'{key_prefix}_date'] = date_str

    # ---- 若儲存的結果日期與目前最新日期不同，提示 ----
    stored_date = st.session_state.get(f'{key_prefix}_date')
    if stored_date and stored_date != date_str:
        st.warning(
            f"⚠️ 你儲存的結果日期是 {stored_date}，目前最新交易日是 {date_str}，"
            f"請重新執行篩選"
        )

    # ---- 顯示結果 ----
    if f'{key_prefix}_result' in st.session_state:
        st.markdown("---")
        extra_pinned = {
            col_ribbon: st.column_config.TextColumn(col_ribbon, pinned=True),
        }
        render_result(
            st.session_state[f'{key_prefix}_result'],
            key_prefix,
            extra_pinned=extra_pinned,
        )
    else:
        st.info("👆 設定完篩選條件後，點選「開始執行」")

# =========================
# Tab 10/11：CCI 穿越 100（日線 / 週線）
#   - 日期鎖定為最新交易日
#   - 硬過濾：滿足最後交易日CCI穿越100 == 'O'
#   - 其餘提供與 Tab1/Tab2 相同的完整篩選 UI
# =========================
def render_tab_cci(period_label, key_prefix, title):
    st.header(title)

    # ---- 日期鎖定為最新交易日 ----
    dates = list_simply_dates()
    if not dates:
        st.error("❌ 找不到任何 simply_report 檔案")
        return
    date_str = dates[0]  # 已 reverse=True，取最新
    st.caption(f"📅 日期鎖定為最新交易日：**{date_str}**（此頁面不受側邊欄日期影響）")

    df = load_simply(date_str, period_label)
    if df is None:
        st.error(f"❌ 讀不到 simply_report_{date_str}_{period_label}.xlsx")
        return

    # ---- 檢查欄位是否存在 ----
    col_cci = '滿足最後交易日CCI穿越100'
    if col_cci not in df.columns:
        st.error(
            f"❌ 此日期檔案缺少欄位：{col_cci}\n\n"
            f"請用新版重新產出 simply_report_{date_str}_{period_label}.xlsx"
        )
        return

    # ---- 硬過濾：== 'O' ----
    mask = df[col_cci].astype(str).str.strip() == 'O'
    df_filtered = df[mask].copy()

    if df_filtered.empty:
        st.warning("🔍 此日期無符合『最後交易日CCI穿越100』的個股")
        return

    st.caption(
        f"📂 simply_report_{date_str}_{period_label}.xlsx "
        f"（原始 {len(df)} 檔 → CCI穿越100 **{len(df_filtered)} 檔**）"
    )

    # ---- 沿用 Tab1/Tab2 的篩選 UI ----
    filters = build_filter_ui(df_filtered, period_label, key_prefix)

    run_clicked = st.button(
        "▶️ 開始執行",
        type='primary',
        key=f'{key_prefix}_run',
        use_container_width=True,
    )

    if run_clicked:
        result = apply_filters(df_filtered, filters)

        # 把 CCI 欄位與基本欄位重排到前面（與釘選順序一致）
        front_cols_order = ['代號', '公司名稱', '產業類別', '實收資本額（億）',
                            col_cci]
        front_cols_order = [c for c in front_cols_order if c in result.columns]
        rest_cols = [c for c in result.columns if c not in front_cols_order]
        result = result[front_cols_order + rest_cols]

        st.session_state[f'{key_prefix}_result'] = result
        st.session_state[f'{key_prefix}_date'] = date_str

    # ---- 若儲存的結果日期與目前最新日期不同，提示 ----
    stored_date = st.session_state.get(f'{key_prefix}_date')
    if stored_date and stored_date != date_str:
        st.warning(
            f"⚠️ 你儲存的結果日期是 {stored_date}，目前最新交易日是 {date_str}，"
            f"請重新執行篩選"
        )

    # ---- 顯示結果 ----
    if f'{key_prefix}_result' in st.session_state:
        st.markdown("---")
        extra_pinned = {
            col_cci: st.column_config.TextColumn(col_cci, pinned=True),
        }
        render_result(
            st.session_state[f'{key_prefix}_result'],
            key_prefix,
            extra_pinned=extra_pinned,
        )
    else:
        st.info("👆 設定完篩選條件後，點選「開始執行」")

# =========================
# Tab 3：日週交集
# =========================
def render_tab_intersection():
    st.header("🔀 日週交集")

    df_daily = st.session_state.get('daily_result')
    df_weekly = st.session_state.get('weekly_result')

    if df_daily is None or df_weekly is None:
        st.info("👈 請先在「日線模式」與「週線模式」分別執行篩選，"
                "交集結果會自動顯示在這裡")
        return

    daily_date = st.session_state.get('daily_date', '?')
    weekly_date = st.session_state.get('weekly_date', '?')

    st.caption(f"📂 日線結果：{daily_date}（{len(df_daily)} 檔）"
               f"　|　週線結果：{weekly_date}（{len(df_weekly)} 檔）")

    codes_d = set(df_daily['代號'].astype(str))
    codes_w = set(df_weekly['代號'].astype(str))
    common = codes_d & codes_w

    if not common:
        st.warning(f"🔍 查無交集"
                   f"（日線 {len(codes_d)} 檔、週線 {len(codes_w)} 檔，"
                   f"無相同個股）")
        return

    st.success(f"✅ 日線符合 {len(codes_d)} 檔、週線符合 {len(codes_w)} 檔，"
               f"交集 {len(common)} 檔")

    # 只顯示基本 4 欄
    base_cols = ['代號', '公司名稱', '產業類別', '實收資本額（億）']
    base_cols = [c for c in base_cols if c in df_daily.columns]
    result = df_daily[df_daily['代號'].astype(str).isin(common)][base_cols].copy()
    result = result.reset_index(drop=True)

    column_config = _build_pinned_config(result)

    st.dataframe(
        _arrow_safe(result),
        width='stretch',
        height=_calc_table_height(len(result)),
        column_config=column_config,
    )

    csv = result.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "📥 匯出交集結果（CSV）",
        data=csv,
        file_name=f"intersection_{datetime.today().strftime('%Y%m%d')}.csv",
        mime='text/csv',
        key='intersection_download_csv',
    )


# =========================
# Tab 6：日週3線糾結交集
# =========================
def render_tab_intersection_3line():
    st.header("🔀 日週3線糾結交集")

    df_d = st.session_state.get('daily_3line_result')
    df_w = st.session_state.get('weekly_3line_result')

    if df_d is None or df_w is None:
        st.info("👈 請先在「日線3線糾結」與「週線3線糾結」分別執行篩選，"
                "交集結果會自動顯示在這裡")
        return

    daily_date  = st.session_state.get('daily_3line_date', '?')
    weekly_date = st.session_state.get('weekly_3line_date', '?')

    st.caption(
        f"📂 日線3線結果：{daily_date}（{len(df_d)} 檔）"
        f"　|　週線3線結果：{weekly_date}（{len(df_w)} 檔）"
    )

    codes_d = set(df_d['代號'].astype(str))
    codes_w = set(df_w['代號'].astype(str))
    common = codes_d & codes_w

    if not common:
        st.warning(
            f"🔍 查無交集（日線 {len(codes_d)} 檔、週線 {len(codes_w)} 檔，無相同個股）"
        )
        return

    st.success(
        f"✅ 日線3線符合 {len(codes_d)} 檔、週線3線符合 {len(codes_w)} 檔，"
        f"日週共振 **{len(common)} 檔**"
    )

    # 只顯示基本 4 欄
    base_cols = ['代號', '公司名稱', '產業類別', '實收資本額（億）']
    base_cols = [c for c in base_cols if c in df_d.columns]
    result = df_d[df_d['代號'].astype(str).isin(common)][base_cols].copy()
    result = result.reset_index(drop=True)

    column_config = _build_pinned_config(result)

    st.dataframe(
        _arrow_safe(result),
        width='stretch',
        height=_calc_table_height(len(result)),
        column_config=column_config,
    )

    csv = result.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "📥 匯出交集結果（CSV）",
        data=csv,
        file_name=f"intersection_3line_{datetime.today().strftime('%Y%m%d')}.csv",
        mime='text/csv',
        key='intersection_3line_download_csv',
    )


# =========================
# Tab 9：日週均線色帶交集
# =========================
def render_tab_intersection_ribbon():
    st.header("🔀 日週均線色帶交集")

    df_d = st.session_state.get('daily_ribbon_result')
    df_w = st.session_state.get('weekly_ribbon_result')

    if df_d is None or df_w is None:
        st.info("👈 請先在「日線均線色帶」與「週線均線色帶」分別執行篩選，"
                "交集結果會自動顯示在這裡")
        return

    daily_date  = st.session_state.get('daily_ribbon_date', '?')
    weekly_date = st.session_state.get('weekly_ribbon_date', '?')

    st.caption(
        f"📂 日線結果：{daily_date}（{len(df_d)} 檔）"
        f"　|　週線結果：{weekly_date}（{len(df_w)} 檔）"
    )

    codes_d = set(df_d['代號'].astype(str))
    codes_w = set(df_w['代號'].astype(str))
    common = codes_d & codes_w

    if not common:
        st.warning(
            f"🔍 查無交集（日線 {len(codes_d)} 檔、週線 {len(codes_w)} 檔，無相同個股）"
        )
        return

    st.success(
        f"✅ 日線符合 {len(codes_d)} 檔、週線符合 {len(codes_w)} 檔，"
        f"日週共振 **{len(common)} 檔**"
    )

    # 只顯示基本 4 欄
    base_cols = ['代號', '公司名稱', '產業類別', '實收資本額（億）']
    base_cols = [c for c in base_cols if c in df_d.columns]
    result = df_d[df_d['代號'].astype(str).isin(common)][base_cols].copy()
    result = result.reset_index(drop=True)

    column_config = _build_pinned_config(result)

    st.dataframe(
        _arrow_safe(result),
        width='stretch',
        height=_calc_table_height(len(result)),
        column_config=column_config,
    )

    csv = result.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "📥 匯出交集結果（CSV）",
        data=csv,
        file_name=f"intersection_ribbon_{datetime.today().strftime('%Y%m%d')}.csv",
        mime='text/csv',
        key='intersection_ribbon_download_csv',
    )

# =========================
# Tab 12：日週 CCI 穿越 100 交集
# =========================
def render_tab_intersection_cci():
    st.header("🔀 日週CCI穿越100交集")

    df_d = st.session_state.get('daily_cci_result')
    df_w = st.session_state.get('weekly_cci_result')

    if df_d is None or df_w is None:
        st.info("👈 請先在「日線CCI穿越100」與「週線CCI穿越100」分別執行篩選，"
                "交集結果會自動顯示在這裡")
        return

    daily_date  = st.session_state.get('daily_cci_date', '?')
    weekly_date = st.session_state.get('weekly_cci_date', '?')

    st.caption(
        f"📂 日線結果：{daily_date}（{len(df_d)} 檔）"
        f"　|　週線結果：{weekly_date}（{len(df_w)} 檔）"
    )

    codes_d = set(df_d['代號'].astype(str))
    codes_w = set(df_w['代號'].astype(str))
    common = codes_d & codes_w

    if not common:
        st.warning(
            f"🔍 查無交集（日線 {len(codes_d)} 檔、週線 {len(codes_w)} 檔，無相同個股）"
        )
        return

    st.success(
        f"✅ 日線符合 {len(codes_d)} 檔、週線符合 {len(codes_w)} 檔，"
        f"日週共振 **{len(common)} 檔**"
    )

    # 只顯示基本 4 欄
    base_cols = ['代號', '公司名稱', '產業類別', '實收資本額（億）']
    base_cols = [c for c in base_cols if c in df_d.columns]
    result = df_d[df_d['代號'].astype(str).isin(common)][base_cols].copy()
    result = result.reset_index(drop=True)

    column_config = _build_pinned_config(result)

    st.dataframe(
        _arrow_safe(result),
        width='stretch',
        height=_calc_table_height(len(result)),
        column_config=column_config,
    )

    csv = result.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "📥 匯出交集結果（CSV）",
        data=csv,
        file_name=f"intersection_cci_{datetime.today().strftime('%Y%m%d')}.csv",
        mime='text/csv',
        key='intersection_cci_download_csv',
    )

# =========================
# 側邊欄（只放全域設定）
# =========================
def build_sidebar():
    st.sidebar.markdown("### ⚙️ 全域設定")

    dates = list_simply_dates()
    if not dates:
        st.sidebar.error("❌ 找不到任何 simply_report 檔案")
        st.sidebar.caption(f"請確認 {OUTPUT_DIR} 內有 "
                           f"simply_report_YYYYMMDD_daily.xlsx")
        return None

    date_str = st.sidebar.selectbox(
        "📅 選擇日期",
        dates,
        key='selected_date',
    )
    st.sidebar.caption(
        f"將讀取：\n"
        f"- simply_report_{date_str}_daily.xlsx\n"
        f"- simply_report_{date_str}_weekly.xlsx"
    )
    st.sidebar.markdown("---")
    st.sidebar.caption(
        "💡 操作流程：\n"
        "1. 在「日線」Tab 設定篩選條件 → 按開始執行\n"
        "2. 在「週線」Tab 設定篩選條件 → 按開始執行\n"
        "3. 切到「日週交集」Tab 自動顯示交集結果\n"
        "4. 3線糾結 / 均線色帶 Tab 會鎖定最新交易日"
    )
    return date_str


# =========================
# 主程式
# =========================
def main():
    st.set_page_config(page_title="股票篩選器", layout="wide")
    st.title("📈 simply_report 股票篩選器")

    build_sidebar()

    (tab1, tab2, tab3, tab4, tab5, tab6,
     tab7, tab8, tab9, tab10, tab11, tab12) = st.tabs([
        "🔵 日線", "🟢 週線", "🔀 日週交集",
        "🎯 日線3線糾結", "🎯 週線3線糾結", "🎯 日週3線交集",
        "🎯 日線均線色帶", "🎯 週線均線色帶", "🎯 日週均線色帶交集",
        "🎯 日線CCI穿越100", "🎯 週線CCI穿越100", "🎯 日週CCI交集",
    ])
    with tab1:
        render_tab_daily()
    with tab2:
        render_tab_weekly()
    with tab3:
        render_tab_intersection()
    with tab4:
        render_tab_3line('daily', 'daily_3line', "🎯 日線 3線糾結 + 紅K棒")
    with tab5:
        render_tab_3line('weekly', 'weekly_3line', "🎯 週線 3線糾結 + 紅K棒")
    with tab6:
        render_tab_intersection_3line()
    with tab7:
        render_tab_ma_ribbon('daily', 'daily_ribbon', "🎯 日線 均線色帶")
    with tab8:
        render_tab_ma_ribbon('weekly', 'weekly_ribbon', "🎯 週線 均線色帶")
    with tab9:
        render_tab_intersection_ribbon()
    with tab10:
        render_tab_cci('daily', 'daily_cci', "🎯 日線 CCI穿越100")
    with tab11:
        render_tab_cci('weekly', 'weekly_cci', "🎯 週線 CCI穿越100")
    with tab12:
        render_tab_intersection_cci()


if __name__ == "__main__":
    main()
