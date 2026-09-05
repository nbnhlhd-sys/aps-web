# -*- coding: utf-8 -*-
"""
APS 排产系统 V2.0 — Streamlit 网页版
=====================================
基于原 Tkinter 桌面版业务逻辑改写，核心改进：
1. st.set_page_config(layout="wide")  强制宽屏，杜绝两边留白
2. st.columns() + st.container(border=True)  行列工整布局，杜绝"一字长蛇阵"
3. 输入框下方直接写计算公式，利用 Streamlit 自动重跑机制实现"无按钮实时计算"

运行方式：
    pip install streamlit pandas numpy plotly openpyxl
    streamlit run aps_web.py
"""

import streamlit as st
import pandas as pd
# 加大 Styler 渲染上限（默认262144单元格，大数据量会报错）
pd.set_option("styler.render.max_elements", 5000000)
import numpy as np
from datetime import datetime, timedelta, date
import plotly.graph_objects as go
import io
import re
import os

# ============================================================
# 【要求1】最顶部强制开启宽屏模式
# ============================================================
st.set_page_config(
    page_title="APS 排产系统 V2.0",
    page_icon="🏭",
    layout="wide",                          # ← 关键：全屏宽屏
    initial_sidebar_state="expanded",
)

# ============================================================
# 全局常量 & 业务计算函数（从原 Tkinter 版移植）
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "aps_data.xlsx")

COL_MACHINE = ["机台状态", "机台编号", "机台名称", "最大吨位", "机台订单统计"]
COL_INJECTION = [
    "产品品号", "品名", "规格", "小时产能 (个)", "每日总工时 (小时)",
    "班产 (个)", "推荐机台", "成品周期 (秒)", "模具穴数 (个)", "换模时间比例 (%)"
]
COL_PACK = [
    "产品品号", "品名", "规格", "小时产能 (个)", "每日总工时 (小时)",
    "班产 (个)", "推荐机台"
]
COL_WORKORDER = [
    "订单类型", "计划批号", "工单号", "产品品号", "品名", "规格",
    "预计产量", "已生产量", "需生产数", "生产机台号", "小时产能 (个)",
    "班产 (个)", "生产周期 (天)", "预计开工日期", "预计完工日期",
    "计划开始日期", "计划结束日期", "订单状态", "班次", "换模时间比例 (%)", "备注"
]

SHIFT_HOURS = {"白班": 8.0, "晚班": 8.0}  # 可在侧边栏修改
_last_blocked_count = 0  # 排产时被"预计开工日期=当天"阻止的订单数
_schedule_days = 60      # 排产范围：预计开工日期在(当天, 当天+N天]内才参与排产（可自定义）
_display_days = 7        # 列表显示范围：预计开工日期在(当天-N天, 当天+∞]内才显示（可自定义）

# 用户管理
COL_USERS = ["用户名", "密码", "角色", "备注"]
DEFAULT_ADMIN = {"用户名": "admin", "密码": "123456", "角色": "管理员", "备注": "系统默认管理员"}


def save_all():
    """统一保存：4张数据表 + 系统设置（班次时长、排产/显示范围）到 Excel
    采用原子写入：先写临时文件，成功后 os.replace 替换原文件，防止写入中断损坏文件"""
    import tempfile
    tmp_path = DATA_FILE + ".tmp.xlsx"
    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as w:
            st.session_state.df_machine.to_excel(w, sheet_name="机台配置", index=False)
            st.session_state.df_injection.to_excel(w, sheet_name="注塑配置", index=False)
            st.session_state.df_pack.to_excel(w, sheet_name="包装配置", index=False)
            st.session_state.df_workorder.to_excel(w, sheet_name="工单排产", index=False)
            sys_df = pd.DataFrame([
                {"参数": "白班时长", "值": st.session_state.shift_white},
                {"参数": "晚班时长", "值": st.session_state.shift_night},
                {"参数": "排产范围天数", "值": _schedule_days},
                {"参数": "列表显示天数", "值": _display_days},
            ])
            sys_df.to_excel(w, sheet_name="系统设置", index=False)
            st.session_state.df_users.to_excel(w, sheet_name="用户管理", index=False)
        # 原子替换：同分区下 os.replace 是原子操作
        os.replace(tmp_path, DATA_FILE)
        return True, ""
    except Exception as e:
        # 清理临时文件
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        return False, str(e)


def load_all_data():
    """从 Excel 加载全部数据，文件损坏时自动备份并重建空表
    返回 (loaded_ok, message)"""
    if not os.path.exists(DATA_FILE):
        return False, "数据文件不存在，将使用空表"
    try:
        st.session_state.df_machine = pd.read_excel(DATA_FILE, sheet_name="机台配置", dtype=object).reindex(columns=COL_MACHINE).fillna("")
        st.session_state.df_injection = pd.read_excel(DATA_FILE, sheet_name="注塑配置", dtype=object).reindex(columns=COL_INJECTION).fillna("")
        st.session_state.df_pack = pd.read_excel(DATA_FILE, sheet_name="包装配置", dtype=object).reindex(columns=COL_PACK).fillna("")
        st.session_state.df_workorder = pd.read_excel(DATA_FILE, sheet_name="工单排产", dtype=object).reindex(columns=COL_WORKORDER).fillna("")
        return True, "数据加载成功"
    except Exception as e:
        err_msg = str(e)
        # 文件损坏（CRC错误 / zip错误 / openpyxl读取错误）→ 备份并重建
        if any(kw in err_msg.lower() for kw in ["crc", "zip", "badzip", "corrupt", "error reading", "file is not"]):
            try:
                backup_path = DATA_FILE + f".corrupted_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                os.rename(DATA_FILE, backup_path)
                # 重建空文件
                with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
                    pd.DataFrame(columns=COL_MACHINE).to_excel(w, sheet_name="机台配置", index=False)
                    pd.DataFrame(columns=COL_INJECTION).to_excel(w, sheet_name="注塑配置", index=False)
                    pd.DataFrame(columns=COL_PACK).to_excel(w, sheet_name="包装配置", index=False)
                    pd.DataFrame(columns=COL_WORKORDER).to_excel(w, sheet_name="工单排产", index=False)
                st.session_state.df_machine = pd.DataFrame(columns=COL_MACHINE, dtype=object)
                st.session_state.df_injection = pd.DataFrame(columns=COL_INJECTION, dtype=object)
                st.session_state.df_pack = pd.DataFrame(columns=COL_PACK, dtype=object)
                st.session_state.df_workorder = pd.DataFrame(columns=COL_WORKORDER, dtype=object)
                return True, f"⚠ 数据文件损坏已自动修复（原文件备份为 {os.path.basename(backup_path)}），数据已重置为空表"
            except Exception as e2:
                return False, f"数据文件损坏且重建失败: {e2}"
        return False, f"数据加载异常: {err_msg}"


def _normalize_col(s):
    """列名归一化：全角转半角、去空格、去括号内容、小写，用于模糊匹配"""
    import unicodedata
    s = unicodedata.normalize("NFKC", str(s))  # 全角→半角
    s = s.replace(" ", "").replace("　", "")
    # 去掉括号及内容（个）(天)(%)等
    import re
    s = re.sub(r"[（(].*?[)）]", "", s)
    return s.strip().lower()


def _import_excel_sheet(xls, sheet_name, target_cols, col_aliases):
    """灵活导入单个sheet：自动匹配sheet名和列名别名；找不到匹配sheet时用第一个sheet"""
    actual_sheet = None
    available = xls.sheet_names
    for sn in available:
        if sn.strip() == sheet_name or any(alias in sn for alias in col_aliases.get("_sheet", [])):
            actual_sheet = sn
            break
    if actual_sheet is None:
        if available:
            actual_sheet = available[0]
        else:
            return pd.DataFrame(columns=target_cols)
    df = pd.read_excel(xls, sheet_name=actual_sheet, dtype=object)
    # 列名自动映射（精确匹配优先，然后归一化模糊匹配）
    rename_map = {}
    # 预计算目标列的归一化名
    target_norm = {t: _normalize_col(t) for t in target_cols}
    for col in df.columns:
        col_str = str(col).strip()
        col_norm = _normalize_col(col_str)
        matched = None
        # 1. 精确匹配
        for target, aliases in col_aliases.items():
            if target == "_sheet":
                continue
            if col_str == target or col_str in aliases:
                matched = target
                break
        # 2. 归一化模糊匹配（包含关系）
        if matched is None:
            for target, aliases in col_aliases.items():
                if target == "_sheet":
                    continue
                tn = target_norm.get(target, _normalize_col(target))
                if col_norm == tn or col_norm in tn or tn in col_norm:
                    matched = target
                    break
                for alias in aliases:
                    an = _normalize_col(alias)
                    if col_norm == an or col_norm in an or an in col_norm:
                        matched = target
                        break
                if matched:
                    break
        if matched:
            rename_map[col] = matched
    df = df.rename(columns=rename_map)
    df = df.reindex(columns=target_cols).fillna("")
    return df


def render_config_import(title, sheet_name, target_cols, aliases, df_key, key_prefix):
    """通用配置表导入UI：选文件→预览→确认导入→进度条。和工单导入方式一致。"""
    with st.container(border=True):
        st.markdown(f"**📥 {title}**（选文件→预览→点确认导入）")
        _mode = st.radio("导入方式", ["覆盖当前", "追加到现有"],
                         horizontal=True, key=f"{key_prefix}_imp_mode", index=0)
        _cfg_up_key = f"{key_prefix}_import_file_{st.session_state.get(f'_{key_prefix}_import_counter', 0)}"
        _up = st.file_uploader(f"选择{title}Excel", type=["xlsx"],
                               key=_cfg_up_key, label_visibility="collapsed")
        if _up is not None:
            try:
                _xls = pd.ExcelFile(_up)
                _df_prev = _import_excel_sheet(_xls, sheet_name, target_cols, aliases)
                # 列名诊断
                _raw_cols = [str(c) for c in _xls.parse(_xls.sheet_names[0], nrows=0).columns]
                _matched = [c for c in target_cols if c in _df_prev.columns and _df_prev[c].astype(str).str.strip().ne("").any()]
                st.info(f"📋 预览：{len(_df_prev)}条（匹配{len(_matched)}/{len(target_cols)}列，sheet：{', '.join(_xls.sheet_names)}）")
                if len(_matched) < len(target_cols) // 2:
                    st.warning(f"⚠ 列名匹配较少。文件原始列名：{', '.join(_raw_cols[:12])}")
                st.session_state[f"{key_prefix}_import_data"] = (_df_prev, _mode)
            except Exception as e:
                st.error(f"文件解析失败：{e}")
                st.session_state[f"{key_prefix}_import_data"] = None

            if st.button("✅ 确认导入", type="primary", use_container_width=True,
                         disabled=st.session_state.get(f"{key_prefix}_import_data") is None,
                         key=f"{key_prefix}_confirm_btn"):
                _data = st.session_state.get(f"{key_prefix}_import_data")
                if _data:
                    _df_imp, _mode = _data
                    _bar = st.progress(0, text="正在导入...")
                    try:
                        _bar.progress(30, text="读取数据...")
                        if _mode == "覆盖当前":
                            st.session_state[df_key] = _df_imp
                        else:
                            st.session_state[df_key] = pd.concat(
                                [st.session_state[df_key], _df_imp],
                                ignore_index=True).astype(object)
                        _bar.progress(60, text="保存到Excel...")
                        _ok, _err = save_all()
                        if not _ok:
                            _bar.empty()
                            st.error(f"❌ 保存失败：{_err}")
                            return
                        _bar.progress(90, text="刷新...")
                        try:
                            st.session_state._data_file_mtime = os.path.getmtime(DATA_FILE)
                        except Exception:
                            pass
                        _bar.progress(100, text="导入完成！")
                        st.success(f"✅ 已导入 {len(_df_imp)} 条{title.replace('导入','')}")
                        st.session_state[f"{key_prefix}_import_data"] = None
                        st.session_state[f"_{key_prefix}_import_counter"] = st.session_state.get(f"_{key_prefix}_import_counter", 0) + 1
                        st.rerun()
                    except Exception as e:
                        _bar.empty()
                        st.error(f"导入失败：{e}")


def render_config_batch_edit(df_key, editable_cols, key_prefix, recalc_func=None):
    """通用配置表批量编辑：选择产品→选择字段→输入新值→应用→自动收起刷新"""
    _exp_key = f"_{key_prefix}_be_expanded"
    if _exp_key not in st.session_state:
        st.session_state[_exp_key] = False
    with st.expander("✏️ 批量编辑（选择产品，批量修改字段）", expanded=st.session_state[_exp_key]):
        df_cfg = st.session_state[df_key].copy()
        if df_cfg.shape[0] == 0:
            st.info("暂无数据，请先添加或导入。")
            return
        # 构建可选标签
        _labels = []
        _label_to_idx = {}
        for _pos in range(min(len(df_cfg), 5000)):
            _idx = df_cfg.index[_pos]
            _r = df_cfg.loc[_idx]
            _lbl = f"{_r.get('产品品号','')} | {_r.get('品名','')}"
            _labels.append(_lbl)
            _label_to_idx[_lbl] = _idx
        _sel = st.multiselect(f"选择产品（当前共{len(df_cfg)}条，留空=全部）",
                              options=_labels, key=f"{key_prefix}_be_select")
        if not _sel:
            _target_idx = list(df_cfg.index)
        else:
            _target_idx = [_label_to_idx[l] for l in _sel if l in _label_to_idx]
        st.caption(f"将修改 {len(_target_idx)} 条产品")

        _field = st.selectbox("选择要修改的字段", editable_cols, key=f"{key_prefix}_be_field")
        _new_val = st.text_input(f"输入新的「{_field}」值（留空不修改）",
                                 key=f"{key_prefix}_be_value", placeholder="留空=不修改")

        if st.button("✅ 应用批量修改", type="primary", use_container_width=True,
                     key=f"{key_prefix}_be_apply"):
            if _new_val.strip():
                _cnt = 0
                for _idx in _target_idx:
                    st.session_state[df_key].at[_idx, _field] = _new_val.strip()
                    _cnt += 1
                # 重算产能（如果有重算函数）
                if recalc_func:
                    recalc_func()
                _ok, _err = save_all()
                if _ok:
                    st.success(f"✅ 已批量修改 {_cnt} 条产品的「{_field}」")
                    st.session_state[_exp_key] = False  # 自动收起
                    st.rerun()
                else:
                    st.error(f"❌ 保存失败：{_err}")
            else:
                st.warning("请输入新值")


# 各表的列名别名映射
WO_ALIASES = {
    "_sheet": ["工单", "workorder", "work order"],
    "订单类型": ["类型", "订单类别", "type"],
    "计划批号": ["批号", "计划号", "批次号"],
    "工单号": ["工单编号", "单号", "work order no", "wo no"],
    "产品品号": ["品号", "产品编号", "料号", "product no"],
    "品名": ["产品名称", "产品名", "product name"],
    "规格": ["规格型号", "型号"],
    "预计产量": ["订单数量", "数量", "产量", "order qty"],
    "已生产量": ["已产量", "完成数量", "produced qty"],
    "需生产数": ["未产量", "剩余数量"],
    "生产机台号": ["机台号", "机台", "设备号", "machine"],
    "小时产能 (个)": ["小时产能", "产能"],
    "班产 (个)": ["班产"],
    "生产周期 (天)": ["生产周期", "周期"],
    "预计开工日期": ["开工日期", "开始日期", "预计开始"],
    "预计完工日期": ["完工日期", "结束日期", "预计结束", "交期"],
    "计划开始日期": ["计划开始", "排产开始"],
    "计划结束日期": ["计划结束", "排产结束"],
    "订单状态": ["状态"],
    "班次": ["班别"],
    "换模时间比例 (%)": ["换模比例", "换模时间"],
    "备注": ["remark", "note"],
}
MACHINE_ALIASES = {
    "_sheet": ["机台", "machine"],
    "机台状态": ["状态"],
    "机台编号": ["机台号", "编号", "machine no"],
    "机台名称": ["名称", "machine name"],
    "最大吨位": ["吨位"],
    "机台订单统计": ["订单统计"],
}
INJECTION_ALIASES = {
    "_sheet": ["注塑", "injection"],
    "产品品号": ["品号", "产品编号"],
    "品名": ["产品名称"],
    "规格": ["规格型号"],
    "小时产能 (个)": ["小时产能", "产能"],
    "每日总工时 (小时)": ["每日工时", "工时"],
    "班产 (个)": ["班产"],
    "推荐机台": ["推荐机台号"],
    "成品周期 (秒)": ["周期", "成型周期"],
    "模具穴数 (个)": ["穴数", "模穴数"],
    "换模时间比例 (%)": ["换模比例"],
}
PACK_ALIASES = {
    "_sheet": ["包装", "pack"],
    "产品品号": ["品号", "产品编号"],
    "品名": ["产品名称"],
    "规格": ["规格型号"],
    "小时产能 (个)": ["小时产能", "产能"],
    "每日总工时 (小时)": ["每日工时", "工时"],
    "班产 (个)": ["班产"],
    "推荐机台": ["推荐机台号"],
}


def load_system_settings():
    """从 Excel 的'系统设置'sheet 加载班次时长和范围参数，不存在则用默认值"""
    global _schedule_days, _display_days
    if not os.path.exists(DATA_FILE):
        return
    try:
        sys_df = pd.read_excel(DATA_FILE, sheet_name="系统设置", dtype=object)
        for _, r in sys_df.iterrows():
            key = str(r.get("参数", "")).strip()
            val = r.get("值")
            if key == "白班时长" and val is not None:
                st.session_state.shift_white = float(val)
            elif key == "晚班时长" and val is not None:
                st.session_state.shift_night = float(val)
            elif key == "排产范围天数" and val is not None:
                _schedule_days = int(float(val))
            elif key == "列表显示天数" and val is not None:
                _display_days = int(float(val))
    except Exception:
        pass


def load_users():
    """从Excel加载用户列表，不存在则创建默认admin"""
    if not os.path.exists(DATA_FILE):
        st.session_state.df_users = pd.DataFrame([DEFAULT_ADMIN], columns=COL_USERS, dtype=object)
        return
    try:
        df = pd.read_excel(DATA_FILE, sheet_name="用户管理", dtype=object)
        if df.shape[0] == 0:
            df = pd.DataFrame([DEFAULT_ADMIN], columns=COL_USERS, dtype=object)
        st.session_state.df_users = df.reindex(columns=COL_USERS).fillna("").astype(object)
    except Exception:
        st.session_state.df_users = pd.DataFrame([DEFAULT_ADMIN], columns=COL_USERS, dtype=object)


def verify_login(username, password):
    """验证用户名密码，返回(成功, 角色)"""
    df = st.session_state.df_users
    match = df[(df["用户名"].astype(str).str.strip() == username.strip()) &
               (df["密码"].astype(str).str.strip() == password)]
    if len(match) > 0:
        return True, str(match.iloc[0]["角色"]).strip()
    return False, ""


def _on_wo_select(event):
    """工单列表行选中回调（模块级定义，避免条件块内定义导致重跑时报错）"""
    try:
        if event is not None and hasattr(event, "selection") and event.selection is not None:
            rows = getattr(event.selection, "rows", None)
            if rows is not None:
                st.session_state.wo_selected = list(rows)
                return
        st.session_state.wo_selected = []
    except Exception:
        st.session_state.wo_selected = []


def safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        if isinstance(v, float) and pd.isna(v):
            return default
        s = str(v).strip()
        if s in ("", "nan", "<NA>", "None", "NaT"):
            return default
        return float(s)
    except (TypeError, ValueError):
        return default


def parse_machines(mach_str):
    """'13.21.22' / '13,21,22' / '13 21 22' → 编号列表"""
    if pd.isna(mach_str):
        return []
    s = str(mach_str).strip()
    if s == "" or s.lower() in ("nan", "<na>", "none"):
        return []
    s = s.replace("，", ",").replace("、", ",").replace(".", ",").replace(" ", ",")
    return [m.strip() for m in s.split(",") if m.strip() != ""]


def calc_injection(cycle_sec, cavity, daily_hour):
    """注塑：小时产能 = 3600/周期×穴数；班产 = 小时产能×每日工时"""
    try:
        c, cav, dh = float(cycle_sec), float(cavity), float(daily_hour)
        if c <= 0 or cav <= 0:
            return 0.0, 0.0
        hour_cap = 3600.0 / c * cav
        batch = hour_cap * dh
        return round(hour_cap, 1), round(batch, 2)
    except Exception:
        return 0.0, 0.0


def calc_pack(hour_cap, daily_hour):
    """包装：班产 = 小时产能×每日工时"""
    try:
        h, dh = float(hour_cap), float(daily_hour)
        return round(h, 1), round(h * dh, 2)
    except Exception:
        return 0.0, 0.0


def get_shift_count(shift):
    return {"白班": 1, "晚班": 1, "白晚班": 2}.get(str(shift).strip(), 1)


def calc_daily_hours(shift):
    s = str(shift).strip()
    wh = SHIFT_HOURS.get("白班", 8.0)
    nh = SHIFT_HOURS.get("晚班", 8.0)
    if s == "晚班":
        return nh
    if s == "白晚班":
        return wh + nh
    return wh


def calc_work_cycle(produce_qty, hour_cap, shift="白班", mold_ratio=0.0):
    """
    生产周期(天) = 需生产数 / (小时产能 × 每日班次工时) × (1 + 换模比例%)
    """
    try:
        pq, hc = float(produce_qty), float(hour_cap)
        if hc <= 0 or pq < 0:
            return 0.0
        daily = calc_daily_hours(shift)
        if daily <= 0:
            return 0.0
        try:
            ratio = float(str(mold_ratio).strip().replace("%", "")) \
                if str(mold_ratio).strip() not in ("", "nan", "<NA>", "None") else 0.0
        except (ValueError, TypeError):
            ratio = 0.0
        days = pq / (hc * daily) * (1 + ratio / 100.0)
        return round(days, 2)
    except Exception:
        return 0.0


def match_product_to_workorder(df_work, df_inj, df_pack):
    """工单产品匹配：从注塑/包装配置表带出品名、规格、产能、换模比例"""
    df = df_work.copy()
    inj_map = {str(r["产品品号"]).strip(): r for _, r in df_inj.iterrows()
               if pd.notna(r["产品品号"])}
    pack_map = {str(r["产品品号"]).strip(): r for _, r in df_pack.iterrows()
                if pd.notna(r["产品品号"])}
    for idx, row in df.iterrows():
        pn = str(row["产品品号"]).strip() if pd.notna(row["产品品号"]) else ""
        if pn in inj_map:
            src = inj_map[pn]
            df.at[idx, "品名"] = src.get("品名", "")
            df.at[idx, "规格"] = src.get("规格", "")
            df.at[idx, "小时产能 (个)"] = round(safe_float(src.get("小时产能 (个)")), 1)
            df.at[idx, "班产 (个)"] = round(safe_float(src.get("班产 (个)")), 2)
            df.at[idx, "订单类型"] = "注塑"
            df.at[idx, "换模时间比例 (%)"] = src.get("换模时间比例 (%)", 0.0)
        elif pn in pack_map:
            src = pack_map[pn]
            df.at[idx, "品名"] = src.get("品名", "")
            df.at[idx, "规格"] = src.get("规格", "")
            df.at[idx, "小时产能 (个)"] = round(safe_float(src.get("小时产能 (个)")), 1)
            df.at[idx, "班产 (个)"] = round(safe_float(src.get("班产 (个)")), 2)
            df.at[idx, "订单类型"] = "包装"
    return df


def sync_workorder_to_config(df_work, df_inj, df_pack):
    """以工单表为准，把产品信息反向同步到注塑/包装配置表（更新或新增）"""
    inj = df_inj.copy()
    pack = df_pack.copy()
    for _, row in df_work.iterrows():
        pn = str(row.get("产品品号", "")).strip()
        if not pn or pn in ("nan", "<NA>", "None"):
            continue
        otype = str(row.get("订单类型", "")).strip()
        name = str(row.get("品名", "")).strip()
        spec = str(row.get("规格", "")).strip()
        hcap = safe_float(row.get("小时产能 (个)"))
        bqty = safe_float(row.get("班产 (个)"))
        mold = safe_float(row.get("换模时间比例 (%)"))
        daily = safe_float(row.get("每日总工时 (小时)", 16))
        if otype == "注塑":
            target = inj
            target_cols = COL_INJECTION
        elif otype == "包装":
            target = pack
            target_cols = COL_PACK
        else:
            continue
        # 查找是否已存在
        existing = target[target["产品品号"].astype(str).str.strip() == pn]
        if len(existing) > 0:
            idx = existing.index[0]
            if name:
                target.at[idx, "品名"] = name
            if spec:
                target.at[idx, "规格"] = spec
            if hcap > 0:
                target.at[idx, "小时产能 (个)"] = hcap
            if bqty > 0:
                target.at[idx, "班产 (个)"] = bqty
            if otype == "注塑" and mold >= 0:
                target.at[idx, "换模时间比例 (%)"] = mold
        else:
            # 新增
            new_row = {c: "" for c in target_cols}
            new_row["产品品号"] = pn
            new_row["品名"] = name
            new_row["规格"] = spec
            new_row["小时产能 (个)"] = hcap
            new_row["班产 (个)"] = bqty
            new_row["每日总工时 (小时)"] = daily if daily > 0 else 16
            if otype == "注塑":
                new_row["换模时间比例 (%)"] = mold
                new_row["成品周期 (秒)"] = 0
                new_row["模具穴数 (个)"] = 1
                new_row["推荐机台"] = ""
            target = pd.concat([target, pd.DataFrame([new_row])], ignore_index=True).astype(object)
        if otype == "注塑":
            inj = target
        else:
            pack = target
    return inj, pack


def auto_schedule(df_work, df_machine):
    """
    按机台分组独立排产：
    - 每台机台第一个工单从当天开始，后续接续
    - 维修状态机台不参与排产
    - 白班/晚班独立追踪机台占用
    - 排产范围可自定义：预计开工日期在(当天, 当天+N天]内才参与排产（默认60天）
    - 预计开工日期<当天 或 >当天+N天 的订单不参与排产
    """
    df = df_work.copy()
    today_str = datetime.now().strftime("%Y-%m-%d")
    today_dt = pd.to_datetime(today_str)
    repair_machs = set()
    for _, r in df_machine.iterrows():
        if str(r.get("机台状态", "")).strip() == "维修":
            no = str(r.get("机台编号", "")).strip()
            if no:
                repair_machs.add(no)

    # 只排"生产中"且有有效机台的工单
    work_mask = (df["订单状态"] == "生产中")
    # 【排产范围】预计开工日期在(当天, 当天+N天]内才参与排产（N可自定义，默认60天）
    # 预计开工日期为空 → 参与排产；<当天 或 >当天+N天 → 跳过（当天也参与排产）
    global _schedule_days
    range_end = (datetime.now() + timedelta(days=_schedule_days)).strftime("%Y-%m-%d")
    start_col = df["预计开工日期"].astype(str).str.strip()
    # 可排产 = 空日期 OR (当天 <= 日期 <= 当天+N天)
    schedulable_mask = (start_col == "") | (
        (start_col >= today_str) & (start_col <= range_end)
    )
    blocked_mask = work_mask & (~schedulable_mask)
    blocked_count = int(blocked_mask.sum())
    global _last_blocked_count
    _last_blocked_count = blocked_count
    if blocked_count > 0:
        df.loc[blocked_mask, "计划开始日期"] = ""
        df.loc[blocked_mask, "计划结束日期"] = ""
    work_mask = work_mask & schedulable_mask
    work_df = df[work_mask].copy()
    if work_df.shape[0] == 0:
        return df

    machine_ends = {}  # {机台号: {"white": datetime, "night": datetime}}

    # 分离单机台 & 多机台
    single = []
    multi = []
    for idx in work_df.index:
        machs = [m for m in parse_machines(df.loc[idx, "生产机台号"]) if m not in repair_machs]
        if len(machs) == 1:
            single.append((idx, machs[0]))
        elif len(machs) > 1:
            multi.append((idx, machs))
        else:
            df.at[idx, "计划开始日期"] = ""
            df.at[idx, "计划结束日期"] = ""

    # 单机台分组排产
    groups = {}
    for idx, mach in single:
        groups.setdefault(mach, []).append(idx)

    def _mach_sort(m):
        try:
            return (0, int(m))
        except (ValueError, TypeError):
            return (1, m)

    for mach in sorted(groups.keys(), key=_mach_sort):
        for idx in groups[mach]:
            row = df.loc[idx]
            cycle = safe_float(row["生产周期 (天)"])
            if cycle <= 0:
                df.at[idx, "计划开始日期"] = ""
                df.at[idx, "计划结束日期"] = ""
                continue
            shift = str(row["班次"]).strip() if pd.notna(row["班次"]) else "白班"
            if shift not in ("白班", "晚班", "白晚班"):
                shift = "白班"
            if mach not in machine_ends:
                machine_ends[mach] = {"white": today_dt, "night": today_dt}
            ends = machine_ends[mach]
            if shift == "白班":
                s_dt = ends["white"]
            elif shift == "晚班":
                s_dt = ends["night"]
            else:
                s_dt = max(ends["white"], ends["night"])
            e_dt = s_dt + timedelta(days=cycle)
            df.at[idx, "计划开始日期"] = s_dt.strftime("%Y-%m-%d")
            df.at[idx, "计划结束日期"] = e_dt.strftime("%Y-%m-%d")
            if shift == "白班":
                ends["white"] = e_dt
            elif shift == "晚班":
                ends["night"] = e_dt
            else:
                ends["white"] = e_dt
                ends["night"] = e_dt

    # 多机台排产
    for idx, machs in multi:
        row = df.loc[idx]
        cycle = safe_float(row["生产周期 (天)"])
        if cycle <= 0:
            df.at[idx, "计划开始日期"] = ""
            df.at[idx, "计划结束日期"] = ""
            continue
        shift = str(row["班次"]).strip() if pd.notna(row["班次"]) else "白班"
        if shift not in ("白班", "晚班", "白晚班"):
            shift = "白班"
        effective = cycle / len(machs)
        for m in machs:
            if m not in machine_ends:
                machine_ends[m] = {"white": today_dt, "night": today_dt}
        s_dt = today_dt
        for m in machs:
            ends = machine_ends[m]
            if shift == "白班":
                avail = ends["white"]
            elif shift == "晚班":
                avail = ends["night"]
            else:
                avail = max(ends["white"], ends["night"])
            if avail > s_dt:
                s_dt = avail
        e_dt = s_dt + timedelta(days=effective)
        df.at[idx, "计划开始日期"] = s_dt.strftime("%Y-%m-%d")
        df.at[idx, "计划结束日期"] = e_dt.strftime("%Y-%m-%d")
        for m in machs:
            ends = machine_ends[m]
            if shift == "白班":
                ends["white"] = e_dt
            elif shift == "晚班":
                ends["night"] = e_dt
            else:
                ends["white"] = e_dt
                ends["night"] = e_dt

    return df


def update_machine_status(df_work, df_machine):
    """有生产中工单→生产；无工单且非维修→闲置；维修保持"""
    active = set()
    for _, row in df_work.iterrows():
        if str(row.get("订单状态", "")).strip() == "生产中":
            for m in parse_machines(row.get("生产机台号", "")):
                if m:
                    active.add(m)
    df = df_machine.copy()
    for idx, row in df.iterrows():
        cur = str(row.get("机台状态", "闲置")).strip()
        if cur not in ("生产", "维修", "闲置"):
            cur = "闲置"
        no = str(row.get("机台编号", "")).strip()
        if cur == "维修":
            continue
        df.at[idx, "机台状态"] = "生产" if no in active else "闲置"
    return df


# ============================================================
# Session State 初始化（数据持久化到内存 + Excel）
# ============================================================
def init_state():
    if "df_machine" not in st.session_state:
        st.session_state.df_machine = pd.DataFrame(columns=COL_MACHINE, dtype=object)
    if "df_injection" not in st.session_state:
        st.session_state.df_injection = pd.DataFrame(columns=COL_INJECTION, dtype=object)
    if "df_pack" not in st.session_state:
        st.session_state.df_pack = pd.DataFrame(columns=COL_PACK, dtype=object)
    if "df_workorder" not in st.session_state:
        st.session_state.df_workorder = pd.DataFrame(columns=COL_WORKORDER, dtype=object)
    if "df_users" not in st.session_state:
        st.session_state.df_users = pd.DataFrame([DEFAULT_ADMIN], columns=COL_USERS, dtype=object)
    if "current_user" not in st.session_state:
        st.session_state.current_user = None
    if "current_role" not in st.session_state:
        st.session_state.current_role = None
    if "shift_white" not in st.session_state:
        st.session_state.shift_white = 8.0
    if "shift_night" not in st.session_state:
        st.session_state.shift_night = 8.0
    if "page" not in st.session_state:
        st.session_state.page = "生产看板"
    if "wo_selected" not in st.session_state:
        st.session_state.wo_selected = []
    # 尝试加载已有 Excel（含损坏自动恢复）
    if not st.session_state.get("_loaded", False):
        ok, msg = load_all_data()
        if not ok and os.path.exists(DATA_FILE):
            st.warning(msg)
        elif ok and "损坏" in msg:
            st.warning(msg)
        st.session_state._loaded = True
        # 系统设置单独加载，失败不影响数据表
        load_system_settings()
        load_users()


init_state()
SHIFT_HOURS["白班"] = st.session_state.shift_white
SHIFT_HOURS["晚班"] = st.session_state.shift_night


# ============================================================
# 侧边栏导航
# ============================================================
with st.sidebar:
    st.markdown("## 🏭 APS 排产系统")
    st.caption("V2.0 网页版 | 宽屏布局")
    st.divider()

    if st.session_state.current_user:
        # 当前用户信息
        role_icon = "👑" if st.session_state.current_role == "管理员" else "👤"
        st.markdown(f"{role_icon} **{st.session_state.current_user}**（{st.session_state.current_role}）")
        if st.button("🚪 退出登录", use_container_width=True):
            st.session_state.current_user = None
            st.session_state.current_role = None
            st.session_state.page = "生产看板"
            st.rerun()
        st.divider()

        pages = ["生产看板", "工单排产", "机台配置", "注塑配置", "包装配置", "甘特图", "系统设置"]
        icons = ["🏠", "📋", "⚙️", "🔧", "📦", "📊", "🛠️"]
        if st.session_state.current_role == "管理员":
            pages.append("用户管理")
            icons.append("👥")
        for pg, ic in zip(pages, icons):
            if st.button(f"{ic} {pg}", use_container_width=True,
                         type="primary" if st.session_state.page == pg else "secondary"):
                st.session_state.page = pg
                st.rerun()

        st.divider()
        st.markdown("### 📅 排产/显示范围")
        sd = st.number_input("排产范围(天)", 1, 365, _schedule_days, 1,
                             help="预计开工日期在(当天, 当天+N天]内的订单才参与排产")
        dd = st.number_input("列表显示范围(天)", 1, 90, _display_days, 1,
                             help="预计开工日期在(当天-N天, 以后)的订单才显示在列表中")
        if sd != _schedule_days:
            _schedule_days = sd
            save_all()
            st.rerun()
        if dd != _display_days:
            _display_days = dd
            save_all()
            st.rerun()

        st.divider()
        # 数据持久化按钮
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 保存", use_container_width=True):
                ok, err = save_all()
                if ok:
                    st.success("已保存到 aps_data.xlsx（含班次/范围设置）")
                else:
                    st.error(f"保存失败: {err}")
        with c2:
            if st.button("📤 导出", use_container_width=True):
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    st.session_state.df_machine.to_excel(w, sheet_name="机台配置", index=False)
                    st.session_state.df_injection.to_excel(w, sheet_name="注塑配置", index=False)
                    st.session_state.df_pack.to_excel(w, sheet_name="包装配置", index=False)
                    st.session_state.df_workorder.to_excel(w, sheet_name="工单排产", index=False)
                st.download_button("⬇ 下载 Excel", buf.getvalue(),
                                   file_name="aps_export.xlsx", use_container_width=True)


# ============================================================
# 页面：工单排产（核心页 — 自动计算演示）
# ============================================================
def page_workorder():
    st.title("📋 工单排产管理")
    st.caption("修改任意输入框 → 页面自动重跑 → 需生产数 / 生产周期 / 预计完工日期 实时刷新，无需点击计算按钮")
    st.info("📌 排产规则：**预计开工日期在[当天, 当天+N天]范围内的订单参与排产**（N可在左侧栏自定义，默认60天）；早于当天或超出范围的订单不参与排产。列表仅显示近N天内的订单（默认7天）。")

    # 【文件变化自动检测】外部导入/修改文件后自动重载
    if os.path.exists(DATA_FILE):
        try:
            _cur_mtime = os.path.getmtime(DATA_FILE)
            _last_mtime = st.session_state.get("_data_file_mtime", 0)
            if _cur_mtime > _last_mtime + 0.5:
                if _last_mtime > 0:
                    ok, msg = load_all_data()
                    load_system_settings()
                    load_users()
                    if ok:
                        st.success(f"🔄 检测到数据文件已更新，已自动重新加载（{len(st.session_state.df_workorder)}条工单）。")
                st.session_state._data_file_mtime = _cur_mtime
        except Exception:
            pass

    # 手动重载 + 内存/文件对比
    _rl1, _rl2, _rl3 = st.columns([1, 1, 1])
    with _rl1:
        if st.button("🔄 重新加载", use_container_width=True):
            st.session_state._loaded = False
            st.session_state._data_file_mtime = 0
            st.rerun()
    with _rl2:
        st.metric("内存工单", len(st.session_state.df_workorder))
    with _rl3:
        _fc = 0
        if os.path.exists(DATA_FILE):
            try:
                _fc = len(pd.read_excel(DATA_FILE, sheet_name="工单排产", dtype=object))
            except Exception:
                pass
        st.metric("文件工单", _fc)

    df_work = st.session_state.df_workorder
    df_inj = st.session_state.df_injection
    df_pack = st.session_state.df_pack
    df_mach = st.session_state.df_machine

    # ----------------------------------------------------------
    # 区域1：新增/编辑工单表单（用 fieldset 风格的 border container）
    # ----------------------------------------------------------
    with st.container(border=True):
        st.subheader("✏️ 新增 / 编辑工单")

        # 【要求2】第一行：基本信息 4 列
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            order_type = st.selectbox("订单类型", ["注塑", "包装"], index=0)
        with col2:
            plan_no = st.text_input("计划批号", value="", placeholder="如 P20260901")
        with col3:
            work_no = st.text_input("工单号", value="", placeholder="如 WO-001")
        with col4:
            product_no = st.text_input("产品品号", value="", placeholder="输入品号自动匹配")

        # 第二行：数量信息 4 列
        col5, col6, col7, col8 = st.columns(4)
        with col5:
            order_qty = st.number_input("预计产量", min_value=0, value=10000, step=100)
        with col6:
            produced_qty = st.number_input("已生产量", min_value=0, value=0, step=100)
        with col7:
            shift = st.selectbox("班次", ["白班", "晚班", "白晚班"], index=0)
        with col8:
            start_date = st.date_input("预计开工日期", value=date.today())

        # 第三行：机台 & 状态 3 列 + 备注
        col9, col10, col11 = st.columns(3)
        with col9:
            # 从机台配置表取可用机台（排除维修）
            available_machs = [str(r["机台编号"]).strip() for _, r in df_mach.iterrows()
                               if str(r.get("机台状态", "")).strip() != "维修"
                               and pd.notna(r.get("机台编号")) and str(r["机台编号"]).strip() != ""]
            selected_machs = st.multiselect("生产机台号（可多选）",
                                            options=available_machs if available_machs else ["请先在机台配置添加"],
                                            default=[])
        with col10:
            order_status = st.selectbox("订单状态", ["生产中", "暂停", "维修"], index=0)
        with col11:
            mold_ratio = st.number_input("换模时间比例 (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.5)

        # ================================================================
        # 【要求3】无按钮自动计算 —— 公式直接写在输入框下方
        # 只要上面任何输入框变化，Streamlit 自动重跑，这里立刻重新计算
        # ================================================================
        st.divider()
        # ① 需生产数 = 预计产量 - 已生产量
        need_produce = max(int(order_qty) - int(produced_qty), 0)

        # ② 从配置表匹配产品信息
        pn = product_no.strip()
        matched_name = ""
        matched_spec = ""
        hour_cap = 0.0
        batch_qty = 0.0
        matched_mold = mold_ratio
        if pn:
            for _, r in df_inj.iterrows():
                if str(r["产品品号"]).strip() == pn:
                    matched_name = str(r.get("品名", ""))
                    matched_spec = str(r.get("规格", ""))
                    hour_cap = safe_float(r.get("小时产能 (个)"))
                    batch_qty = safe_float(r.get("班产 (个)"))
                    matched_mold = safe_float(r.get("换模时间比例 (%)"), mold_ratio)
                    break
            else:
                for _, r in df_pack.iterrows():
                    if str(r["产品品号"]).strip() == pn:
                        matched_name = str(r.get("品名", ""))
                        matched_spec = str(r.get("规格", ""))
                        hour_cap = safe_float(r.get("小时产能 (个)"))
                        batch_qty = safe_float(r.get("班产 (个)"))
                        matched_mold = safe_float(r.get("换模时间比例 (%)"), mold_ratio)
                        break

        # ③ 生产周期(天)
        work_cycle = calc_work_cycle(need_produce, hour_cap, shift, matched_mold)

        # ④ 预计完工日期
        if work_cycle > 0:
            finish_date = start_date + timedelta(days=work_cycle)
        else:
            finish_date = start_date

        # 紧凑计算结果条（单行显示）
        _cap_txt = f"{hour_cap:.1f}个/时" if hour_cap > 0 else "未匹配"
        _match_txt = f"✅ {matched_name}" if (pn and hour_cap > 0) else (f"⚠ 品号「{pn}」未配置" if pn else "请输入品号")
        st.markdown(
            f"<div style='background:#f0f7ff;padding:6px 14px;border-radius:6px;font-size:13px;'>"
            f"🧮 <b>需生产 {need_produce:,}</b> 个 &nbsp;|&nbsp; "
            f"<b>小时产能 {_cap_txt}</b> &nbsp;|&nbsp; "
            f"<b>周期 {work_cycle:.2f}天</b>（{shift}{calc_daily_hours(shift):.0f}h/天）&nbsp;|&nbsp; "
            f"<b>完工 {finish_date.strftime('%Y-%m-%d')}</b> &nbsp;|&nbsp; "
            f"<span style='color:{'#2e7d32' if hour_cap>0 else '#d84315'}'>{_match_txt}</span>"
            f"</div>", unsafe_allow_html=True)

        # 底部按钮行 —— 【要求2】水平并排 + use_container_width
        st.divider()
        btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)
        with btn_col1:
            if st.button("➕ 添加工单", use_container_width=True, type="primary"):
                new_row = {
                    "订单类型": order_type, "计划批号": plan_no, "工单号": work_no,
                    "产品品号": pn, "品名": matched_name, "规格": matched_spec,
                    "预计产量": order_qty, "已生产量": produced_qty, "需生产数": need_produce,
                    "生产机台号": ",".join(selected_machs),
                    "小时产能 (个)": hour_cap, "班产 (个)": batch_qty,
                    "生产周期 (天)": work_cycle,
                    "预计开工日期": start_date.strftime("%Y-%m-%d"),
                    "预计完工日期": finish_date.strftime("%Y-%m-%d"),
                    "计划开始日期": "", "计划结束日期": "",
                    "订单状态": order_status, "班次": shift,
                    "换模时间比例 (%)": matched_mold, "备注": "",
                }
                st.session_state.df_workorder = pd.concat(
                    [st.session_state.df_workorder, pd.DataFrame([new_row])],
                    ignore_index=True).astype(object)
                st.success(f"工单 {work_no or plan_no} 已添加！")
                st.rerun()
        with btn_col2:
            if st.button("🔄 全量自动排产", use_container_width=True):
                df = match_product_to_workorder(st.session_state.df_workorder,
                                                st.session_state.df_injection,
                                                st.session_state.df_pack)
                # 重算需生产数 & 周期
                for idx, row in df.iterrows():
                    oq = safe_float(row["预计产量"])
                    pq = safe_float(row["已生产量"])
                    np_ = int(max(oq - pq, 0))
                    df.at[idx, "需生产数"] = np_
                    hc = safe_float(row["小时产能 (个)"])
                    if hc > 0:
                        df.at[idx, "生产周期 (天)"] = calc_work_cycle(
                            np_, hc, row.get("班次", "白班"),
                            row.get("换模时间比例 (%)", 0.0))
                df = auto_schedule(df, st.session_state.df_machine)
                st.session_state.df_machine = update_machine_status(df, st.session_state.df_machine)
                st.session_state.df_workorder = df
                msg = "全量自动排产完成！"
                if _last_blocked_count > 0:
                    msg += f"\n\n⚠ 有 {_last_blocked_count} 个订单未参与排产（预计开工日期不在排产范围：当天~当天+{_schedule_days}天内，已清空其排产日期）。"
                st.success(msg)
                st.rerun()
        with btn_col3:
            if st.button("🗑 清除全部工单", use_container_width=True):
                st.session_state.df_workorder = pd.DataFrame(columns=COL_WORKORDER, dtype=object)
                save_all()
                try:
                    st.session_state._data_file_mtime = os.path.getmtime(DATA_FILE)
                except Exception:
                    pass
                st.warning("已清除全部工单数据并保存")
                st.rerun()
        with btn_col4:
            if st.button("🔄 同步配置", use_container_width=True):
                st.session_state.df_workorder = match_product_to_workorder(
                    st.session_state.df_workorder,
                    st.session_state.df_injection,
                    st.session_state.df_pack)
                for _idx, _row in st.session_state.df_workorder.iterrows():
                    _oq = safe_float(_row["预计产量"])
                    _pq = safe_float(_row["已生产量"])
                    st.session_state.df_workorder.at[_idx, "需生产数"] = int(max(_oq - _pq, 0))
                    _hc = safe_float(_row["小时产能 (个)"])
                    if _hc > 0:
                        st.session_state.df_workorder.at[_idx, "生产周期 (天)"] = calc_work_cycle(
                            st.session_state.df_workorder.at[_idx, "需生产数"],
                            _hc, _row.get("班次", "白班"),
                            _row.get("换模时间比例 (%)", 0.0))
                save_all()
                st.success("已同步配置表数据到全部工单")
                st.rerun()

    # ----------------------------------------------------------
    # 区域2：工单列表（颜色标记 + 表头排序 + 显示范围过滤 + 批量导入）
    # ----------------------------------------------------------
    with st.container(border=True):
        st.subheader("📊 工单列表")
        st.caption("👆 点击表头可排序；🔴 近7天内开工  🟡 当天开工  🔵 预计完工晚于计划结束")

        # 筛选行
        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1:
            f_plan = st.text_input("筛选-计划批号", key="f_plan")
        with fc2:
            f_work = st.text_input("筛选-工单号", key="f_work")
        with fc3:
            f_mach = st.text_input("筛选-机台号", key="f_mach")
        with fc4:
            f_prod = st.text_input("筛选-品号", key="f_prod")

        # 【修改1】订单类型勾选筛选 + 显示全部开关
        ft1, ft2, ft3, ft4 = st.columns([1, 1, 1, 3])
        with ft1:
            show_inj = st.checkbox("☑ 显示注塑", value=True, key="f_type_inj")
        with ft2:
            show_pack = st.checkbox("☑ 显示包装", value=True, key="f_type_pack")
        with ft3:
            show_all_dates = st.checkbox("📅 显示全部", value=False, key="f_show_all",
                                         help="勾选后忽略日期范围，显示所有工单")
        with ft4:
            st.caption("（导入数据后建议勾选「显示全部」查看历史工单）")

        # 预计完工日期快捷筛选
        fc_fin1, fc_fin2, fc_fin3, fc_fin4 = st.columns(4)
        with fc_fin1:
            fin_before = st.checkbox("⏰ 七天前", value=False, key="f_fin_before",
                                     help="预计完工日期在7天以前")
        with fc_fin2:
            fin_today = st.checkbox("📌 当天", value=False, key="f_fin_today",
                                    help="预计完工日期是今天")
        with fc_fin3:
            fin_after = st.checkbox("🔮 七天后", value=False, key="f_fin_after",
                                    help="预计完工日期在7天以后")
        with fc_fin4:
            st.caption("（预计完工日期筛选，不勾选=不限制）")

        # 颜色筛选（红/黄/蓝/白）
        fc_clr1, fc_clr2, fc_clr3, fc_clr4 = st.columns(4)
        with fc_clr1:
            clr_red = st.checkbox("🔴 近7天开工", value=False, key="f_clr_red",
                                  help="预计开工日期在近7天内（不含今天）")
        with fc_clr2:
            clr_yellow = st.checkbox("🟡 当天开工", value=False, key="f_clr_yellow",
                                     help="预计开工日期是今天")
        with fc_clr3:
            clr_blue = st.checkbox("🔵 完工超期", value=False, key="f_clr_blue",
                                   help="预计完工日期晚于计划结束日期")
        with fc_clr4:
            clr_white = st.checkbox("⚪ 正常", value=False, key="f_clr_white",
                                    help="无颜色标记的正常工单")

        # 【修改3】日期范围自定义筛选（紧凑布局：从/到左右并排）
        with st.expander("📅 日期筛选（留空=不筛选）", expanded=False):
            dcol1, dcol2, dcol3, dcol4 = st.columns(4)
            def _compact_date(col, label, key_from, key_to):
                with col:
                    st.caption(label)
                    _sf, _st = st.columns(2)
                    with _sf:
                        v_from = st.date_input("从", value=None, key=key_from,
                                               format="YYYY-MM-DD", label_visibility="collapsed")
                    with _st:
                        v_to = st.date_input("到", value=None, key=key_to,
                                             format="YYYY-MM-DD", label_visibility="collapsed")
                return v_from, v_to
            f_exp_start_from, f_exp_start_to = _compact_date(dcol1, "预计开工", "f_exp_start_from", "f_exp_start_to")
            f_exp_finish_from, f_exp_finish_to = _compact_date(dcol2, "预计完工", "f_exp_finish_from", "f_exp_finish_to")
            f_plan_start_from, f_plan_start_to = _compact_date(dcol3, "计划开始", "f_plan_start_from", "f_plan_start_to")
            f_plan_end_from, f_plan_end_to = _compact_date(dcol4, "计划结束", "f_plan_end_from", "f_plan_end_to")

        # 编辑模式开关
        edit_mode = st.checkbox("✏️ 编辑模式（开启后可直接修改表格单元格，关闭后显示颜色标记）",
                                value=False, key="wo_edit_mode")

        df_show = st.session_state.df_workorder.copy()

        # 订单类型过滤
        _sel_types = []
        if show_inj:
            _sel_types.append("注塑")
        if show_pack:
            _sel_types.append("包装")
        # 两个都勾选=不过滤类型（避免导入数据中订单类型为空时被全部过滤）
        if len(_sel_types) == 2:
            pass
        elif _sel_types:
            df_show = df_show[df_show["订单类型"].astype(str).str.strip().isin(_sel_types)]
        else:
            df_show = df_show.iloc[0:0]  # 都不勾选 → 空表

        # 【要求4】显示范围过滤：预计开工日期在(当天-N天, 以后]才显示，N天前的不显示
        # 勾选"显示全部"则跳过此过滤
        if not show_all_dates:
            global _display_days
            _today_str = datetime.now().strftime("%Y-%m-%d")
            _cutoff = (datetime.now() - timedelta(days=_display_days)).strftime("%Y-%m-%d")
            _start_col = df_show["预计开工日期"].astype(str).str.strip()
            df_show = df_show[(_start_col == "") | (_start_col >= _cutoff)]

        # 预计完工日期快捷筛选（七天前/当天/七天后）
        if fin_before or fin_today or fin_after:
            _fin_col = pd.to_datetime(df_show["预计完工日期"], errors="coerce")
            _today_d = datetime.now().date()
            _mask = pd.Series([False] * len(df_show), index=df_show.index)
            if fin_before:
                _mask = _mask | ((_fin_col.notna()) & (_fin_col.dt.date <= _today_d - timedelta(days=7)))
            if fin_today:
                _mask = _mask | ((_fin_col.notna()) & (_fin_col.dt.date == _today_d))
            if fin_after:
                _mask = _mask | ((_fin_col.notna()) & (_fin_col.dt.date >= _today_d + timedelta(days=7)))
            df_show = df_show[_mask]

        if f_plan.strip():
            df_show = df_show[df_show["计划批号"].astype(str).str.contains(f_plan.strip(), na=False)]
        if f_work.strip():
            df_show = df_show[df_show["工单号"].astype(str).str.contains(f_work.strip(), na=False)]
        if f_mach.strip():
            df_show = df_show[df_show["生产机台号"].astype(str).str.contains(f_mach.strip(), na=False)]
        if f_prod.strip():
            df_show = df_show[df_show["产品品号"].astype(str).str.contains(f_prod.strip(), na=False)]

        # 【修改3】日期范围筛选：4个日期字段各自的从~到范围
        def _filter_date_range(df, col_name, d_from, d_to):
            """对指定日期列做范围筛选，d_from/d_to为None时不筛选该端"""
            if d_from is None and d_to is None:
                return df
            col = df[col_name].astype(str).str.strip()
            mask = pd.Series([True] * len(df), index=df.index)
            if d_from is not None:
                mask = mask & (col >= d_from.strftime("%Y-%m-%d"))
            if d_to is not None:
                mask = mask & (col <= d_to.strftime("%Y-%m-%d"))
            return df[mask]

        df_show = _filter_date_range(df_show, "预计开工日期", f_exp_start_from, f_exp_start_to)
        df_show = _filter_date_range(df_show, "预计完工日期", f_exp_finish_from, f_exp_finish_to)
        df_show = _filter_date_range(df_show, "计划开始日期", f_plan_start_from, f_plan_start_to)
        df_show = _filter_date_range(df_show, "计划结束日期", f_plan_end_from, f_plan_end_to)

        # 颜色筛选（红/黄/蓝/白）
        if clr_red or clr_yellow or clr_blue or clr_white:
            _today_s = datetime.now().strftime("%Y-%m-%d")
            _today_d = datetime.strptime(_today_s, "%Y-%m-%d").date()
            _cutoff_d = _today_d - timedelta(days=_display_days)
            def _get_color(r):
                start_s = str(r.get("预计开工日期", "")).strip()
                finish_s = str(r.get("预计完工日期", "")).strip()
                plan_end_s = str(r.get("计划结束日期", "")).strip()
                # 红：近7天开工（不含今天）
                if start_s and start_s not in ("nan", "<NA>", "None"):
                    try:
                        sd = datetime.strptime(start_s, "%Y-%m-%d").date()
                        if _cutoff_d <= sd < _today_d:
                            return "red"
                    except Exception:
                        pass
                # 黄：当天开工
                if start_s == _today_s:
                    return "yellow"
                # 蓝：完工超期
                if finish_s and plan_end_s and finish_s not in ("nan", "<NA>", "None") \
                        and plan_end_s not in ("nan", "<NA>", "None"):
                    try:
                        if finish_s > plan_end_s:
                            return "blue"
                    except Exception:
                        pass
                return "white"
            _colors = df_show.apply(_get_color, axis=1)
            _sel_colors = []
            if clr_red: _sel_colors.append("red")
            if clr_yellow: _sel_colors.append("yellow")
            if clr_blue: _sel_colors.append("blue")
            if clr_white: _sel_colors.append("white")
            df_show = df_show[_colors.isin(_sel_colors)]

        if df_show.shape[0] == 0:
            _total = len(st.session_state.df_workorder)
            if _total > 0:
                _hint = f"（内存中共{_total}条工单，被筛选过滤）"
                if not show_all_dates:
                    _hint += " → 试试勾选「📅 显示全部」"
            else:
                # 检查文件里有没有数据
                _file_count = 0
                if os.path.exists(DATA_FILE):
                    try:
                        _fc = pd.read_excel(DATA_FILE, sheet_name="工单排产", dtype=object)
                        _file_count = len(_fc)
                    except Exception:
                        pass
                if _file_count > 0:
                    _hint = f"（文件中有{_file_count}条工单但未加载 → 点上方「🔄 从文件重新加载」）"
                else:
                    _hint = ""
            st.info(f"暂无工单数据{_hint}，请在上方添加或导入。")
        else:
            # ===== 【要求5】行颜色标记函数 =====
            def _row_colors(row):
                """返回整行的背景色样式列表
                优先级：大红(近7天开工) > 黄色(当天开工) > 蓝色(预计完工>计划结束)
                """
                today_s = datetime.now().strftime("%Y-%m-%d")
                today_d = datetime.strptime(today_s, "%Y-%m-%d").date()
                cutoff_d = today_d - timedelta(days=_display_days)
                start_s = str(row.get("预计开工日期", "")).strip()
                finish_s = str(row.get("预计完工日期", "")).strip()
                plan_end_s = str(row.get("计划结束日期", "")).strip()
                bg = ""
                # 蓝色：预计完工日期 > 计划结束日期
                if finish_s and plan_end_s and finish_s not in ("nan", "<NA>", "None") \
                        and plan_end_s not in ("nan", "<NA>", "None"):
                    try:
                        if finish_s > plan_end_s:
                            bg = "background-color: #4A90D9; color: white;"
                    except Exception:
                        pass
                # 黄色：预计开工日期 == 当天
                if start_s == today_s:
                    bg = "background-color: #FFD700; color: black;"
                # 大红：预计开工日期在[当天-N天, 当天)
                if start_s and start_s not in ("nan", "<NA>", "None"):
                    try:
                        start_d = datetime.strptime(start_s, "%Y-%m-%d").date()
                        if cutoff_d <= start_d < today_d:
                            bg = "background-color: #DC143C; color: white;"
                    except (ValueError, TypeError):
                        pass
                return [bg] * len(row)

            # ===== 【要求1】颜色标记 + 可排序的 DataFrame（点击表头自动排序）=====
            # 编辑模式：可直接修改单元格（保留原始索引，修改后只更新被编辑的行）
            if edit_mode:
                _date_cols = ["预计开工日期", "预计完工日期", "计划开始日期", "计划结束日期"]
                df_edit = df_show.copy()
                for _dc in _date_cols:
                    df_edit[_dc] = pd.to_datetime(df_edit[_dc], errors="coerce")
                edited = st.data_editor(
                    df_edit,
                    use_container_width=True,
                    num_rows="dynamic",
                    key="wo_editor_main",
                    column_config={
                        "订单类型": st.column_config.SelectboxColumn(options=["注塑", "包装"]),
                        "订单状态": st.column_config.SelectboxColumn(options=["生产中", "暂停", "维修"]),
                        "班次": st.column_config.SelectboxColumn(options=["白班", "晚班", "白晚班"]),
                        "预计开工日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "预计完工日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "计划开始日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                        "计划结束日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    },
                    hide_index=False,
                    height=840,
                )
                # 检测变化：逐行逐列比较
                _changed = False
                _df_full = st.session_state.df_workorder.copy()
                for _idx in edited.index:
                    if _idx in _df_full.index:
                        for _col in COL_WORKORDER:
                            if _col in _date_cols:
                                _new_val = edited.at[_idx, _col]
                                if pd.notna(_new_val):
                                    _new_str = _new_val.strftime("%Y-%m-%d") if hasattr(_new_val, "strftime") else str(_new_val)[:10]
                                else:
                                    _new_str = ""
                                _old_str = str(_df_full.at[_idx, _col]).strip()
                                if _new_str != _old_str:
                                    _df_full.at[_idx, _col] = _new_str
                                    _changed = True
                            else:
                                _new_val = str(edited.at[_idx, _col]).strip() if pd.notna(edited.at[_idx, _col]) else ""
                                _old_val = str(_df_full.at[_idx, _col]).strip()
                                if _new_val != _old_val:
                                    _df_full.at[_idx, _col] = _new_val
                                    _changed = True
                if _changed:
                    # 重算需生产数和周期
                    for _idx in _df_full.index:
                        _oq = safe_float(_df_full.at[_idx, "预计产量"])
                        _pq = safe_float(_df_full.at[_idx, "已生产量"])
                        _df_full.at[_idx, "需生产数"] = int(max(_oq - _pq, 0))
                        _hc = safe_float(_df_full.at[_idx, "小时产能 (个)"])
                        if _hc > 0:
                            _df_full.at[_idx, "生产周期 (天)"] = calc_work_cycle(
                                _df_full.at[_idx, "需生产数"], _hc,
                                _df_full.at[_idx, "班次"], _df_full.at[_idx, "换模时间比例 (%)"])
                    st.session_state.df_workorder = _df_full.astype(object)
                    save_all()
                    st.success("已保存修改")
                    st.rerun()
            else:
                # 非编辑模式：颜色标记显示
                _cell_count = df_show.shape[0] * df_show.shape[1]
                if _cell_count > 200000:
                    st.caption(f"⚠ 数据量较大（{df_show.shape[0]}行×{df_show.shape[1]}列={_cell_count:,}单元格），已关闭行颜色标记以提升性能。")
                    st.dataframe(
                        df_show,
                        use_container_width=True,
                        hide_index=True,
                        height=840,
                    )
                else:
                    styled = df_show.style.apply(_row_colors, axis=1)
                    st.dataframe(
                        styled,
                        use_container_width=True,
                        hide_index=True,
                        height=840,
                    )
            # 保存显示行的原始索引，用于选中行上移/下移映射
            st.session_state._df_show_index = list(df_show.index)
            # 构建多选标签（超大数据只取前5000行，避免下拉框卡顿）
            _row_labels = []
            _label_to_orig = {}
            _max_select = min(len(df_show), 5000)
            for pos in range(_max_select):
                orig_idx = df_show.index[pos]
                r = df_show.loc[orig_idx]
                _lbl = f"{r.get('工单号','')} | {r.get('品名','')} | {r.get('产品品号','')}"
                _row_labels.append(_lbl)
                _label_to_orig[_lbl] = orig_idx
            st.session_state._label_to_orig = _label_to_orig
            if len(df_show) > _max_select:
                st.caption(f"（共{len(df_show)}行，顺序调整仅显示前{_max_select}行）")

            # 【可靠选行】用 multiselect 替代不稳定的 on_select 回调
            _sel_labels = st.multiselect(
                "🔽 选择要调整顺序的工单（可多选，选中后点上移/下移）",
                options=_row_labels,
                default=[],
                key="wo_multiselect",
            )
            st.session_state.wo_selected_labels = _sel_labels
            _sel_n = len(_sel_labels)
            if _sel_n > 0:
                st.caption(f"已选中 {_sel_n} 行，可点击下方上移/下移调整排产优先级顺序")

            # 批量操作按钮
            st.divider()
            bc1, bc2, bc3, bc4 = st.columns(4)
            with bc1:
                if st.button("🔄 对当前列表重新排产", use_container_width=True):
                    df = auto_schedule(st.session_state.df_workorder, st.session_state.df_machine)
                    st.session_state.df_machine = update_machine_status(df, st.session_state.df_machine)
                    st.session_state.df_workorder = df
                    msg = "重新排产完成！"
                    if _last_blocked_count > 0:
                        msg += f"\n\n⚠ 有 {_last_blocked_count} 个订单未参与排产（预计开工日期不在排产范围：当天~当天+{_schedule_days}天内）。"
                    st.success(msg)
                    st.rerun()
            with bc2:
                _sel_labels = st.session_state.get("wo_selected_labels", [])
                _has_sel = len(_sel_labels) > 0
                if st.button("⬆ 选中行上移", use_container_width=True, disabled=not _has_sel):
                    _label_map = st.session_state.get("_label_to_orig", {})
                    _orig = sorted([_label_map[lbl] for lbl in _sel_labels if lbl in _label_map])
                    if _orig:
                        df_mv = st.session_state.df_workorder.copy()
                        for idx in _orig:
                            if idx > 0 and (idx - 1) not in _orig:
                                _tmp = df_mv.iloc[idx - 1].copy()
                                df_mv.iloc[idx - 1] = df_mv.iloc[idx].copy()
                                df_mv.iloc[idx] = _tmp
                        st.session_state.df_workorder = df_mv
                        st.session_state.wo_selected_labels = []
                        st.rerun()
            with bc3:
                if st.button("⬇ 选中行下移", use_container_width=True, disabled=not _has_sel):
                    _label_map = st.session_state.get("_label_to_orig", {})
                    _orig = sorted([_label_map[lbl] for lbl in _sel_labels if lbl in _label_map], reverse=True)
                    if _orig:
                        df_mv = st.session_state.df_workorder.copy()
                        _max = len(df_mv) - 1
                        for idx in _orig:
                            if idx < _max and (idx + 1) not in _orig:
                                _tmp = df_mv.iloc[idx + 1].copy()
                                df_mv.iloc[idx + 1] = df_mv.iloc[idx].copy()
                                df_mv.iloc[idx] = _tmp
                        st.session_state.df_workorder = df_mv
                        st.session_state.wo_selected_labels = []
                        st.rerun()
            with bc4:
                if st.button("📤 导出当前筛选", use_container_width=True):
                    buf = io.BytesIO()
                    df_show.to_excel(buf, index=False)
                    st.download_button("下载", buf.getvalue(), file_name="工单筛选结果.xlsx",
                                       use_container_width=True)

            # ===== 批量编辑（筛选后选中行批量修改字段）=====
            _wo_be_key = "_wo_be_expanded"
            if _wo_be_key not in st.session_state:
                st.session_state[_wo_be_key] = False
            with st.expander("✏️ 批量编辑（筛选后选中工单，批量修改字段）", expanded=st.session_state[_wo_be_key]):
                if df_show.shape[0] == 0:
                    st.info("当前筛选结果为空，请先调整筛选条件。")
                else:
                    # 构建可选标签
                    _be_labels = []
                    _be_label_to_idx = {}
                    for _pos in range(min(len(df_show), 5000)):
                        _orig_idx = df_show.index[_pos]
                        _r = df_show.loc[_orig_idx]
                        _lbl = f"{_r.get('工单号','')} | {_r.get('品名','')} | {_r.get('产品品号','')}"
                        _be_labels.append(_lbl)
                        _be_label_to_idx[_lbl] = _orig_idx
                    _be_sel = st.multiselect(f"选择要编辑的工单（当前筛选共{len(df_show)}条，留空=全部）",
                                             options=_be_labels, key="be_select")
                    if not _be_sel:
                        _be_target_idx = list(df_show.index)
                    else:
                        _be_target_idx = [_be_label_to_idx[l] for l in _be_sel if l in _be_label_to_idx]
                    st.caption(f"将修改 {len(_be_target_idx)} 条工单（只修改填写了的字段，留空的字段不变）")

                    be1, be2, be3 = st.columns(3)
                    with be1:
                        _be_type = st.selectbox("订单类型", ["不修改", "注塑", "包装"], key="be_type")
                    with be2:
                        _be_status = st.selectbox("订单状态", ["不修改", "生产中", "暂停", "维修"], key="be_status")
                    with be3:
                        _be_shift = st.selectbox("班次", ["不修改", "白班", "晚班", "白晚班"], key="be_shift")
                    be4, be5, be6 = st.columns(3)
                    with be4:
                        _be_mach = st.text_input("生产机台号（留空不修改）", key="be_mach", placeholder="如 13,21")
                    with be5:
                        _be_date = st.date_input("预计开工日期（不修改请留空）", value=None, key="be_date",
                                                 format="YYYY-MM-DD")
                    with be6:
                        _be_mold = st.number_input("换模时间比例%（-1不修改）", min_value=-1.0, max_value=100.0,
                                                   value=-1.0, step=0.5, key="be_mold")

                    if st.button("✅ 应用批量修改", type="primary", use_container_width=True):
                        # 先同步配置表数据（品名/产能等）
                        st.session_state.df_workorder = match_product_to_workorder(
                            st.session_state.df_workorder,
                            st.session_state.df_injection,
                            st.session_state.df_pack)
                        # 再应用批量修改（不被配置表覆盖）
                        _cnt = 0
                        for _idx in _be_target_idx:
                            if _be_type != "不修改":
                                st.session_state.df_workorder.at[_idx, "订单类型"] = _be_type
                            if _be_status != "不修改":
                                st.session_state.df_workorder.at[_idx, "订单状态"] = _be_status
                            if _be_shift != "不修改":
                                st.session_state.df_workorder.at[_idx, "班次"] = _be_shift
                            if _be_mach.strip():
                                st.session_state.df_workorder.at[_idx, "生产机台号"] = _be_mach.strip()
                            if _be_date is not None:
                                st.session_state.df_workorder.at[_idx, "预计开工日期"] = _be_date.strftime("%Y-%m-%d")
                            if _be_mold >= 0:
                                st.session_state.df_workorder.at[_idx, "换模时间比例 (%)"] = _be_mold
                            _cnt += 1
                        # 重算需生产数和生产周期（换模比例变了周期也要变）
                        for _idx in _be_target_idx:
                            _oq = safe_float(st.session_state.df_workorder.at[_idx, "预计产量"])
                            _pq = safe_float(st.session_state.df_workorder.at[_idx, "已生产量"])
                            st.session_state.df_workorder.at[_idx, "需生产数"] = int(max(_oq - _pq, 0))
                            _hc = safe_float(st.session_state.df_workorder.at[_idx, "小时产能 (个)"])
                            if _hc > 0:
                                st.session_state.df_workorder.at[_idx, "生产周期 (天)"] = calc_work_cycle(
                                    st.session_state.df_workorder.at[_idx, "需生产数"], _hc,
                                    st.session_state.df_workorder.at[_idx, "班次"],
                                    st.session_state.df_workorder.at[_idx, "换模时间比例 (%)"])
                        # 以工单表为准，反向同步产品数据到配置表
                        st.session_state.df_injection, st.session_state.df_pack = sync_workorder_to_config(
                            st.session_state.df_workorder,
                            st.session_state.df_injection,
                            st.session_state.df_pack)
                        _save_ok, _save_err = save_all()
                        if not _save_ok:
                            st.error(f"❌ 保存失败：{_save_err}")
                            return
                        st.success(f"✅ 已批量修改 {_cnt} 条工单，并同步产品数据到配置表")
                        st.session_state[_wo_be_key] = False  # 自动收起
                        st.rerun()

        # ===== 批量导入（放在批量编辑下面）=====
        with st.container(border=True):
            st.markdown("**📥 批量导入工单 Excel**（选文件→预览→点确认导入）")
            _imp_mode = st.radio("导入方式", ["覆盖当前工单", "追加到现有工单"],
                                 horizontal=True, key="wo_imp_mode", index=0)
            _wo_up_key = f"wo_import_file_{st.session_state.get('_wo_import_counter', 0)}"
            _uploaded = st.file_uploader("选择工单 Excel", type=["xlsx"], key=_wo_up_key,
                                         label_visibility="collapsed")
            if _uploaded is not None:
                try:
                    _xls = pd.ExcelFile(_uploaded)
                    _prev_wo = _import_excel_sheet(_xls, "工单排产", COL_WORKORDER, WO_ALIASES)
                    _prev_mach = _import_excel_sheet(_xls, "机台配置", COL_MACHINE, MACHINE_ALIASES)
                    _prev_inj = _import_excel_sheet(_xls, "注塑配置", COL_INJECTION, INJECTION_ALIASES)
                    _prev_pack = _import_excel_sheet(_xls, "包装配置", COL_PACK, PACK_ALIASES)
                    st.info(f"📋 预览：工单{len(_prev_wo)}条 / 机台{len(_prev_mach)}条 / "
                            f"注塑{len(_prev_inj)}条 / 包装{len(_prev_pack)}条"
                            f"（sheet：{', '.join(_xls.sheet_names)}）")
                    _raw_cols = [str(c) for c in _xls.parse(_xls.sheet_names[0], nrows=0).columns]
                    _matched = [c for c in COL_WORKORDER if c in _prev_wo.columns and _prev_wo[c].astype(str).str.strip().ne("").any()]
                    if len(_matched) < 5:
                        st.warning(f"⚠ 列名匹配较少（仅{len(_matched)}/{len(COL_WORKORDER)}列有数据）。"
                                   f"文件原始列名：{', '.join(_raw_cols[:15])}。")
                    st.session_state._wo_import_data = (_prev_wo, _prev_mach, _prev_inj, _prev_pack, _imp_mode)
                except Exception as e:
                    st.error(f"文件解析失败：{e}")
                    st.session_state._wo_import_data = None

                if st.button("✅ 确认导入", type="primary", use_container_width=True,
                             disabled=st.session_state.get("_wo_import_data") is None):
                    _data = st.session_state.get("_wo_import_data")
                    if _data:
                        _prev_wo, _prev_mach, _prev_inj, _prev_pack, _mode = _data
                        _bar = st.progress(0, text="正在导入...")
                        try:
                            _bar.progress(20, text="读取工单数据...")
                            if _mode == "覆盖当前工单":
                                if len(_prev_wo) > 0:
                                    st.session_state.df_workorder = _prev_wo
                                if len(_prev_mach) > 0:
                                    st.session_state.df_machine = _prev_mach
                                if len(_prev_inj) > 0:
                                    st.session_state.df_injection = _prev_inj
                                if len(_prev_pack) > 0:
                                    st.session_state.df_pack = _prev_pack
                            else:
                                if len(_prev_wo) > 0:
                                    st.session_state.df_workorder = pd.concat(
                                        [st.session_state.df_workorder, _prev_wo],
                                        ignore_index=True).astype(object)
                            _bar.progress(40, text="同步配置表到工单...")
                            st.session_state.df_workorder = match_product_to_workorder(
                                st.session_state.df_workorder,
                                st.session_state.df_injection,
                                st.session_state.df_pack)
                            for _idx, _row in st.session_state.df_workorder.iterrows():
                                _oq = safe_float(_row["预计产量"])
                                _pq = safe_float(_row["已生产量"])
                                st.session_state.df_workorder.at[_idx, "需生产数"] = int(max(_oq - _pq, 0))
                                _hc = safe_float(_row["小时产能 (个)"])
                                if _hc > 0:
                                    st.session_state.df_workorder.at[_idx, "生产周期 (天)"] = calc_work_cycle(
                                        st.session_state.df_workorder.at[_idx, "需生产数"],
                                        _hc, _row.get("班次", "白班"),
                                        _row.get("换模时间比例 (%)", 0.0))
                            _bar.progress(60, text="保存到Excel...")
                            _ok, _err = save_all()
                            if not _ok:
                                _bar.empty()
                                st.error(f"❌ 保存到Excel失败：{_err}\n\n可能原因：文件正被Excel打开，请关闭后重试。")
                                return
                            _bar.progress(80, text="刷新数据...")
                            try:
                                st.session_state._data_file_mtime = os.path.getmtime(DATA_FILE)
                            except Exception:
                                pass
                            _bar.progress(100, text="导入完成！")
                            try:
                                _verify = pd.read_excel(DATA_FILE, sheet_name="工单排产", dtype=object)
                                _verify_count = len(_verify)
                            except Exception:
                                _verify_count = -1
                            st.success(f"✅ 已导入 {len(_prev_wo)} 条工单，文件验证：{_verify_count}条。如列表为空请勾选「显示全部」。")
                            st.session_state._wo_import_data = None
                            st.session_state["_wo_import_counter"] = st.session_state.get("_wo_import_counter", 0) + 1
                            st.rerun()
                        except Exception as e:
                            _bar.empty()
                            st.error(f"导入失败：{e}")


# ============================================================
# 页面：机台配置
# ============================================================
def page_machine():
    st.title("⚙️ 机台配置表")
    st.caption("机台状态：生产 / 维修 / 闲置。维修机台不参与排产。")

    with st.container(border=True):
        st.subheader("新增机台")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            m_status = st.selectbox("机台状态", ["闲置", "生产", "维修"], key="m_status")
        with c2:
            m_no = st.text_input("机台编号", key="m_no", placeholder="如 13")
        with c3:
            m_name = st.text_input("机台名称", key="m_name", placeholder="如 注塑机13号")
        with c4:
            m_ton = st.number_input("最大吨位", min_value=0, value=160, step=10, key="m_ton")

        bc1, bc2, bc3 = st.columns(3)
        with bc1:
            if st.button("➕ 添加机台", use_container_width=True, type="primary"):
                if not m_no.strip():
                    st.error("机台编号不能为空")
                else:
                    new = {"机台状态": m_status, "机台编号": m_no.strip(),
                           "机台名称": m_name, "最大吨位": m_ton, "机台订单统计": 0}
                    st.session_state.df_machine = pd.concat(
                        [st.session_state.df_machine, pd.DataFrame([new])],
                        ignore_index=True).astype(object)
                    st.success(f"机台 {m_no} 已添加")
                    st.rerun()
        with bc2:
            if st.button("🔄 同步工单机台号", use_container_width=True):
                st.success("已同步（维修机台自动排除）")
        with bc3:
            if st.button("🗑 清空", use_container_width=True):
                st.session_state.df_machine = pd.DataFrame(columns=COL_MACHINE, dtype=object)
                save_all()
                st.rerun()

    with st.container(border=True):
        st.subheader("机台列表")
        df = st.session_state.df_machine.copy()
        # 自动计算订单统计
        if st.session_state.df_workorder.shape[0] > 0:
            counts = {}
            for _, row in st.session_state.df_workorder.iterrows():
                if str(row.get("订单状态", "")).strip() == "生产中":
                    for m in parse_machines(row.get("生产机台号", "")):
                        counts[m] = counts.get(m, 0) + 1
            df["机台订单统计"] = df["机台编号"].apply(
                lambda x: counts.get(str(x).strip(), 0) if pd.notna(x) else 0)
            st.session_state.df_machine = df

        edited = st.data_editor(
            df, use_container_width=True, num_rows="dynamic", key="mach_editor",
            column_config={
                "机台状态": st.column_config.SelectboxColumn(options=["生产", "维修", "闲置"]),
            },
            hide_index=True, height=350,
        )
        if not edited.equals(df):
            st.session_state.df_machine = edited.reindex(columns=COL_MACHINE).fillna("").astype(object)
            st.rerun()

    # 导入区域（放在列表后面）
    render_config_import("导入机台Excel", "机台配置", COL_MACHINE, MACHINE_ALIASES,
                         "df_machine", "mach")


# ============================================================
# 页面：注塑产品配置
# ============================================================
def page_injection():
    st.title("🔧 注塑产品配置表")
    st.caption("输入成品周期、模具穴数、每日工时 → 自动计算小时产能和班产（无按钮实时刷新）")

    with st.container(border=True):
        st.subheader("新增注塑产品")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            pn = st.text_input("品号", key="inj_pn", placeholder="如 INJ-001")
        with c2:
            pname = st.text_input("品名", key="inj_name")
        with c3:
            pspec = st.text_input("规格", key="inj_spec")
        with c4:
            prec = st.text_input("推荐机台", key="inj_rec", placeholder="如 13,21,22")

        c5, c6, c7, c8 = st.columns(4)
        with c5:
            cycle = st.number_input("成品周期(秒)", min_value=0.0, value=30.0, step=1.0, key="inj_cycle")
        with c6:
            cavity = st.number_input("模具穴数", min_value=1, value=1, step=1, key="inj_cavity")
        with c7:
            daily = st.number_input("每日工时(h)", min_value=0.0, value=16.0, step=1.0, key="inj_daily")
        with c8:
            ratio = st.number_input("换模比例(%)", min_value=0.0, value=0.0, step=0.5, key="inj_ratio")

        # 自动计算（紧凑显示）
        h_cap, b_qty = calc_injection(cycle, cavity, daily)
        calc_bar = st.container()
        with calc_bar:
            st.markdown(
                f"<div style='background:#f0f7ff;padding:8px 14px;border-radius:6px;font-size:14px;'>"
                f"⚡ 自动计算：<b>小时产能 {h_cap:.1f}</b> 个/时 &nbsp;|&nbsp; "
                f"<b>班产 {b_qty:.0f}</b> 个/班 &nbsp;|&nbsp; "
                f"<span style='color:#888'>3600÷{cycle:.0f}s×{cavity}穴 × {daily:.0f}h</span>"
                f"</div>", unsafe_allow_html=True)

        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button("➕ 添加产品", use_container_width=True, type="primary"):
                if not pn.strip():
                    st.error("产品品号不能为空")
                else:
                    new = {"产品品号": pn.strip(), "品名": pname, "规格": pspec,
                           "小时产能 (个)": h_cap, "每日总工时 (小时)": daily,
                           "班产 (个)": b_qty, "推荐机台": prec,
                           "成品周期 (秒)": cycle, "模具穴数 (个)": cavity,
                           "换模时间比例 (%)": ratio}
                    st.session_state.df_injection = pd.concat(
                        [st.session_state.df_injection, pd.DataFrame([new])],
                        ignore_index=True).astype(object)
                    st.success(f"产品 {pn} 已添加")
                    st.rerun()
        with bc2:
            if st.button("🗑 清空", use_container_width=True):
                st.session_state.df_injection = pd.DataFrame(columns=COL_INJECTION, dtype=object)
                save_all()
                st.rerun()

    with st.container(border=True):
        st.subheader("注塑产品列表")
        # 品号筛选
        filt_col1, filt_col2 = st.columns([1, 4])
        with filt_col1:
            f_pn_inj = st.text_input("品号筛选", key="f_pn_inj", placeholder="输入品号模糊匹配")
        df = st.session_state.df_injection.copy()
        if f_pn_inj.strip():
            df = df[df["产品品号"].astype(str).str.contains(f_pn_inj.strip(), na=False)]
        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic",
                                key="inj_editor", hide_index=True, height=350)
        if not edited.equals(df):
            # 先强制 object 类型，否则 Arrow string 列不接受写入 float
            edited = edited.astype(object)
            # 编辑后自动重算产能
            for idx, row in edited.iterrows():
                h, b = calc_injection(row.get("成品周期 (秒)", 0),
                                      row.get("模具穴数 (个)", 1),
                                      row.get("每日总工时 (小时)", 0))
                edited.at[idx, "小时产能 (个)"] = h
                edited.at[idx, "班产 (个)"] = b
            st.session_state.df_injection = edited.reindex(columns=COL_INJECTION).fillna("").astype(object)
            st.rerun()

    # 批量编辑
    def _recalc_inj():
        for _idx, _row in st.session_state.df_injection.iterrows():
            _h, _b = calc_injection(_row.get("成品周期 (秒)", 0),
                                    _row.get("模具穴数 (个)", 1),
                                    _row.get("每日总工时 (小时)", 0))
            st.session_state.df_injection.at[_idx, "小时产能 (个)"] = _h
            st.session_state.df_injection.at[_idx, "班产 (个)"] = _b
    render_config_batch_edit("df_injection",
                             ["推荐机台", "每日总工时 (小时)", "换模时间比例 (%)", "成品周期 (秒)", "模具穴数 (个)"],
                             "inj", _recalc_inj)

    # 导入区域（放在列表后面）
    render_config_import("导入注塑Excel", "注塑配置", COL_INJECTION, INJECTION_ALIASES,
                         "df_injection", "inj")


# ============================================================
# 页面：包装产品配置
# ============================================================
def page_pack():
    st.title("📦 包装产品配置表")
    st.caption("输入小时产能、每日工时 → 自动计算班产（无按钮实时刷新）")

    with st.container(border=True):
        st.subheader("新增包装产品")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            pn = st.text_input("品号", key="pack_pn", placeholder="如 PK-001")
        with c2:
            pname = st.text_input("品名", key="pack_name")
        with c3:
            pspec = st.text_input("规格", key="pack_spec")
        with c4:
            prec = st.text_input("推荐机台", key="pack_rec", placeholder="如 1,2")

        c5, c6, c7 = st.columns(3)
        with c5:
            hcap = st.number_input("小时产能(个)", min_value=0.0, value=500.0, step=10.0, key="pack_hcap")
        with c6:
            daily = st.number_input("每日工时(h)", min_value=0.0, value=16.0, step=1.0, key="pack_daily")
        with c7:
            ratio = st.number_input("换模比例(%)", min_value=0.0, value=0.0, step=0.5, key="pack_ratio")

        # 自动计算（紧凑显示）
        _, b_qty = calc_pack(hcap, daily)
        st.markdown(
            f"<div style='background:#f0fff4;padding:8px 14px;border-radius:6px;font-size:14px;'>"
            f"⚡ 自动计算：<b>班产 {b_qty:.0f}</b> 个/班 &nbsp;|&nbsp; "
            f"<span style='color:#888'>{hcap:.0f}个/时 × {daily:.0f}h</span>"
            f"</div>", unsafe_allow_html=True)

        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button("➕ 添加产品", use_container_width=True, type="primary"):
                if not pn.strip():
                    st.error("产品品号不能为空")
                else:
                    new = {"产品品号": pn.strip(), "品名": pname, "规格": pspec,
                           "小时产能 (个)": hcap, "每日总工时 (小时)": daily,
                           "班产 (个)": b_qty, "推荐机台": prec,
                           "换模时间比例 (%)": ratio}
                    st.session_state.df_pack = pd.concat(
                        [st.session_state.df_pack, pd.DataFrame([new])],
                        ignore_index=True).astype(object)
                    st.success(f"产品 {pn} 已添加")
                    st.rerun()
        with bc2:
            if st.button("🗑 清空", use_container_width=True):
                st.session_state.df_pack = pd.DataFrame(columns=COL_PACK, dtype=object)
                save_all()
                st.rerun()

    with st.container(border=True):
        st.subheader("包装产品列表")
        # 品号筛选
        filt_col1, filt_col2 = st.columns([1, 4])
        with filt_col1:
            f_pn_pack = st.text_input("品号筛选", key="f_pn_pack", placeholder="输入品号模糊匹配")
        df = st.session_state.df_pack.copy()
        if f_pn_pack.strip():
            df = df[df["产品品号"].astype(str).str.contains(f_pn_pack.strip(), na=False)]
        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic",
                                key="pack_editor", hide_index=True, height=350)
        if not edited.equals(df):
            edited = edited.astype(object)
            for idx, row in edited.iterrows():
                _, b = calc_pack(row.get("小时产能 (个)", 0),
                                 row.get("每日总工时 (小时)", 0))
                edited.at[idx, "班产 (个)"] = b
            st.session_state.df_pack = edited.reindex(columns=COL_PACK).fillna("").astype(object)
            st.rerun()

    # 批量编辑
    def _recalc_pack():
        for _idx, _row in st.session_state.df_pack.iterrows():
            _, _b = calc_pack(_row.get("小时产能 (个)", 0),
                              _row.get("每日总工时 (小时)", 0))
            st.session_state.df_pack.at[_idx, "班产 (个)"] = _b
    render_config_batch_edit("df_pack",
                             ["推荐机台", "每日总工时 (小时)"],
                             "pack", _recalc_pack)

    # 导入区域（放在列表后面）
    render_config_import("导入包装Excel", "包装配置", COL_PACK, PACK_ALIASES,
                         "df_pack", "pack")


# ============================================================
# 页面：甘特图
# ============================================================
def page_gantt():
    st.title("📊 排产甘特图")
    st.caption("按机台×班次展示工单排产时间轴，注塑=蓝色，包装=绿色，晚班=深色；数据与工单列表实时同步")

    # 【需求2】甘特图同步工单列表：顶部操作栏
    top1, top2, top3, top4, top5 = st.columns(5)
    with top1:
        if st.button("🔄 重新排产", use_container_width=True, type="primary"):
            df_sched = match_product_to_workorder(st.session_state.df_workorder,
                                                  st.session_state.df_injection,
                                                  st.session_state.df_pack)
            for idx, row in df_sched.iterrows():
                oq = safe_float(row["预计产量"])
                pq = safe_float(row["已生产量"])
                np_ = int(max(oq - pq, 0))
                df_sched.at[idx, "需生产数"] = np_
                hc = safe_float(row["小时产能 (个)"])
                if hc > 0:
                    df_sched.at[idx, "生产周期 (天)"] = calc_work_cycle(
                        np_, hc, row.get("班次", "白班"),
                        row.get("换模时间比例 (%)", 0.0))
            df_sched = auto_schedule(df_sched, st.session_state.df_machine)
            st.session_state.df_machine = update_machine_status(df_sched, st.session_state.df_machine)
            st.session_state.df_workorder = df_sched
            st.success("已重新排产，甘特图已同步更新")
            st.rerun()
    with top2:
        show_g_inj = st.checkbox("显示注塑", value=True, key="g_inj")
    with top3:
        show_g_pack = st.checkbox("显示包装", value=True, key="g_pack")
    with top4:
        st.metric("排产中工单数", len(st.session_state.df_workorder[
            st.session_state.df_workorder["计划开始日期"].astype(str).str.strip() != ""]))
    with top5:
        g_show_all = st.checkbox("📅 显示全部", value=False, key="g_show_all")

    df = st.session_state.df_workorder.copy()

    # 订单类型筛选（与工单列表同步）
    _g_types = []
    if show_g_inj:
        _g_types.append("注塑")
    if show_g_pack:
        _g_types.append("包装")
    if _g_types:
        df = df[df["订单类型"].astype(str).str.strip().isin(_g_types)]

    # 显示范围过滤（与工单列表同步：预计开工日期 >= 当天-N天，或为空）
    if not g_show_all:
        _g_cutoff = (datetime.now() - timedelta(days=_display_days)).strftime("%Y-%m-%d")
        _g_start = df["预计开工日期"].astype(str).str.strip()
        df = df[(_g_start == "") | (_g_start >= _g_cutoff)]

    df = df[df["计划开始日期"].notna() & (df["计划开始日期"].astype(str).str.strip() != "")
            & df["计划结束日期"].notna() & (df["计划结束日期"].astype(str).str.strip() != "")]

    if df.shape[0] == 0:
        st.warning("暂无排产数据，请先在「工单排产」页点击「全量自动排产」。")
        return

    try:
        df["s_dt"] = pd.to_datetime(df["计划开始日期"], errors="coerce")
        df["e_dt"] = pd.to_datetime(df["计划结束日期"], errors="coerce")
        df = df.dropna(subset=["s_dt", "e_dt"])
    except Exception as e:
        st.error(f"日期解析错误: {e}")
        return

    if df.shape[0] == 0:
        st.warning("排产日期格式有误，请检查工单数据。")
        return

    # 展开多机台 & 班次，构建甘特图 DataFrame
    gantt_rows = []
    for _, row in df.iterrows():
        machs = parse_machines(row["生产机台号"]) or ["未分配"]
        shift = str(row.get("班次", "白班")).strip()
        if shift not in ("白班", "晚班", "白晚班"):
            shift = "白班"
        ord_type = str(row.get("订单类型", "")).strip()
        wno = str(row.get("工单号", ""))
        pname = str(row.get("品名", ""))
        for m in machs:
            if shift in ("白班", "白晚班"):
                gantt_rows.append({"Task": f"机台{m}·白班", "Start": row["s_dt"],
                                   "Finish": row["e_dt"], "类型": ord_type, "班次": "白班",
                                   "工单号": wno, "品名": pname})
            if shift in ("晚班", "白晚班"):
                gantt_rows.append({"Task": f"机台{m}·晚班", "Start": row["s_dt"],
                                   "Finish": row["e_dt"], "类型": ord_type, "班次": "晚班",
                                   "工单号": wno, "品名": pname})

    if not gantt_rows:
        st.info("当前筛选范围内无有效排产数据。")
        return

    gdf = pd.DataFrame(gantt_rows)

    # 时间范围选择
    c1, c2 = st.columns(2)
    with c1:
        view_weeks = st.slider("显示周数", 1, 8, 2, key="g_weeks")
    with c2:
        week_offset = st.slider("周偏移（0=本周）", -4, 8, 0, key="g_offset")

    today = datetime.now().date()
    monday = today - timedelta(days=today.weekday())
    ws = monday + timedelta(days=week_offset * 7)
    we = ws + timedelta(days=7 * view_weeks - 1)

    # 按可见时间范围过滤
    gdf = gdf[(gdf["Start"].dt.date <= we) & (gdf["Finish"].dt.date >= ws)]
    if gdf.shape[0] == 0:
        st.info(f"时间范围 {ws.strftime('%Y-%m-%d')} ~ {we.strftime('%Y-%m-%d')} 内无排产工单，可调整「显示周数」或「周偏移」。")
        return

    # 颜色映射
    def _gantt_color(r):
        cmap = {
            ("注塑", "白班"): "#4A90D9", ("注塑", "晚班"): "#2E5F8A",
            ("包装", "白班"): "#50B86A", ("包装", "晚班"): "#2E7D4F",
        }
        return cmap.get((r["类型"], r["班次"]), "#999999")
    gdf["颜色"] = gdf.apply(_gantt_color, axis=1)
    gdf["标签"] = gdf["工单号"] + " " + gdf["品名"]

    # 【修复】用 px.timeline 绘制甘特图，确保条形正常显示
    import plotly.express as px
    fig = px.timeline(
        gdf,
        x_start="Start",
        x_end="Finish",
        y="Task",
        color="颜色",
        color_discrete_map="identity",
        hover_data={"工单号": True, "品名": True, "类型": True, "班次": True, "颜色": False},
        text="标签",
    )
    fig.update_yaxes(autorange="reversed", title="", gridcolor="#eee")
    fig.update_xaxes(
        range=[ws, we + timedelta(days=1)],
        tickformat="%m-%d",
        gridcolor="#eee",
        title="日期",
    )
    fig.update_traces(textposition="inside", textfont_size=9, textfont_color="white")
    fig.update_layout(
        height=max(400, gdf["Task"].nunique() * 32),
        margin=dict(l=10, r=10, t=40, b=10),
        plot_bgcolor="white",
        title=f"APS 工单甘特图（{ws.strftime('%Y-%m-%d')} ~ {we.strftime('%Y-%m-%d')} · 共{gdf.shape[0]}条作业）",
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)

    # 图例
    lc1, lc2, lc3, lc4 = st.columns(4)
    lc1.markdown("<div style='background:#4A90D9;color:white;padding:6px;text-align:center;border-radius:4px'>注塑·白班</div>", unsafe_allow_html=True)
    lc2.markdown("<div style='background:#2E5F8A;color:white;padding:6px;text-align:center;border-radius:4px'>注塑·晚班</div>", unsafe_allow_html=True)
    lc3.markdown("<div style='background:#50B86A;color:white;padding:6px;text-align:center;border-radius:4px'>包装·白班</div>", unsafe_allow_html=True)
    lc4.markdown("<div style='background:#2E7D4F;color:white;padding:6px;text-align:center;border-radius:4px'>包装·晚班</div>", unsafe_allow_html=True)


# ============================================================
# 页面：系统设置
# ============================================================
def page_settings():
    st.title("🛠️ 系统设置")
    st.caption("班次时长、数据管理、系统信息")

    with st.container(border=True):
        st.subheader("⏱ 班次时长设置")
        c1, c2 = st.columns(2)
        with c1:
            sw = st.number_input("白班时长 (小时)", 0.0, 24.0, st.session_state.shift_white, 0.5)
        with c2:
            sn = st.number_input("晚班时长 (小时)", 0.0, 24.0, st.session_state.shift_night, 0.5)
        if st.button("保存班次设置", use_container_width=True, type="primary"):
            st.session_state.shift_white = sw
            st.session_state.shift_night = sn
            SHIFT_HOURS["白班"] = sw
            SHIFT_HOURS["晚班"] = sn
            save_all()
            st.success(f"已保存：白班 {sw}h / 晚班 {sn}h（已写入Excel，重启不丢失）")
            st.rerun()

    with st.container(border=True):
        st.subheader("💾 数据管理")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("机台数", len(st.session_state.df_machine))
        with c2:
            st.metric("注塑产品数", len(st.session_state.df_injection))
        with c3:
            st.metric("包装产品数", len(st.session_state.df_pack))
        with c4:
            st.metric("工单数", len(st.session_state.df_workorder))

        bc1, bc2, bc3, bc4 = st.columns(4)
        with bc1:
            if st.button("💾 保存", use_container_width=True):
                save_all()
                st.success(f"已保存到 {DATA_FILE}")
        with bc2:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                st.session_state.df_machine.to_excel(w, sheet_name="机台配置", index=False)
                st.session_state.df_injection.to_excel(w, sheet_name="注塑配置", index=False)
                st.session_state.df_pack.to_excel(w, sheet_name="包装配置", index=False)
                st.session_state.df_workorder.to_excel(w, sheet_name="工单排产", index=False)
            st.download_button("📤 导出", buf.getvalue(),
                               file_name="aps_full_export.xlsx", use_container_width=True)
        with bc3:
            if st.button("🔄 重载文件", use_container_width=True):
                st.session_state._loaded = False
                st.rerun()
        with bc4:
            if st.button("🗑 清空数据", use_container_width=True):
                st.session_state.df_machine = pd.DataFrame(columns=COL_MACHINE, dtype=object)
                st.session_state.df_injection = pd.DataFrame(columns=COL_INJECTION, dtype=object)
                st.session_state.df_pack = pd.DataFrame(columns=COL_PACK, dtype=object)
                st.session_state.df_workorder = pd.DataFrame(columns=COL_WORKORDER, dtype=object)
                save_all()
                st.warning("已清空全部数据并保存")
                st.rerun()

    with st.container(border=True):
        st.subheader("📖 计算公式说明")
        st.markdown("""
        | 计算项 | 公式 |
        |--------|------|
        | 注塑小时产能 | 3600 ÷ 成品周期(秒) × 模具穴数 |
        | 班产 | 小时产能 × 每日总工时 |
        | 需生产数 | 预计产量 − 已生产量 |
        | 生产周期(天) | 需生产数 ÷ (小时产能 × 每日班次工时) × (1 + 换模比例%) |
        | 预计完工日期 | 预计开工日期 + 生产周期(天) |
        | 每日班次工时 | 白班=白班时长，晚班=晚班时长，白晚班=白班+晚班 |
        """)


# ============================================================
# 页面：生产看板（首页）
# ============================================================
def page_dashboard():
    st.title("🏠 生产看板")
    st.caption(f"实时生产数据总览 | {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    df_wo = st.session_state.df_workorder.copy()
    df_mach = st.session_state.df_machine.copy()
    today_str = datetime.now().strftime("%Y-%m-%d")

    # ===== 第一行：核心指标（紧凑单行）=====
    total_orders = len(df_wo)
    in_production = len(df_wo[df_wo["订单状态"].astype(str).str.strip() == "生产中"])
    paused = len(df_wo[df_wo["订单状态"].astype(str).str.strip().isin(["暂停", "维修"])])
    scheduled = len(df_wo[df_wo["计划开始日期"].astype(str).str.strip() != ""])
    overdue = 0
    for _, r in df_wo.iterrows():
        pe = str(r.get("计划结束日期", "")).strip()
        fc = str(r.get("预计完工日期", "")).strip()
        if pe and fc and pe not in ("nan", "<NA>") and fc not in ("nan", "<NA>"):
            if pe > fc:
                overdue += 1
    today_start = len(df_wo[df_wo["计划开始日期"].astype(str).str.strip() == today_str])
    _prod_pct = f"{in_production/total_orders*100:.0f}%" if total_orders else "0%"
    st.markdown(
        f"<div style='background:#f0f7ff;padding:8px 16px;border-radius:6px;font-size:14px;'>"
        f"📋 <b>工单总数 {total_orders}</b> &nbsp;|&nbsp; "
        f"<b>生产中 {in_production}</b>（{_prod_pct}）&nbsp;|&nbsp; "
        f"<b>暂停/维修 {paused}</b> &nbsp;|&nbsp; "
        f"<b>已排产 {scheduled}</b> &nbsp;|&nbsp; "
        f"<b>今日开工 {today_start}</b> &nbsp;|&nbsp; "
        f"<b style='color:{'#d84315' if overdue>0 else '#2e7d32'}'>超期 {overdue}</b>"
        f"</div>", unsafe_allow_html=True)

    st.divider()

    # ===== 第二行：机台状态 + 订单类型分布 =====
    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.subheader("⚙️ 机台状态")
        mach_prod = len(df_mach[df_mach["机台状态"].astype(str).str.strip() == "生产"])
        mach_repair = len(df_mach[df_mach["机台状态"].astype(str).str.strip() == "维修"])
        mach_idle = len(df_mach[df_mach["机台状态"].astype(str).str.strip() == "闲置"])
        total_mach = len(df_mach)
        util = mach_prod / total_mach * 100 if total_mach > 0 else 0
        st.markdown(
            f"<div style='background:#f5f5f5;padding:6px 12px;border-radius:6px;font-size:13px;'>"
            f"<b>生产中 {mach_prod}</b> &nbsp;|&nbsp; <b>维修中 {mach_repair}</b> &nbsp;|&nbsp; <b>闲置 {mach_idle}</b> &nbsp;|&nbsp; "
            f"利用率 <b>{util:.0f}%</b>（{mach_prod}/{total_mach}台）"
            f"</div>", unsafe_allow_html=True)
        if total_mach > 0:
            st.progress(util / 100)
        if mach_repair > 0:
            repair_list = df_mach[df_mach["机台状态"].astype(str).str.strip() == "维修"][["机台编号", "机台名称"]]
            st.warning("🔧 维修中机台：" + "、".join(
                [f"{r['机台编号']}({r['机台名称']})" for _, r in repair_list.iterrows()]))

    with col_right:
        st.subheader("📊 订单类型分布")
        inj_count = len(df_wo[df_wo["订单类型"].astype(str).str.strip() == "注塑"])
        pack_count = len(df_wo[df_wo["订单类型"].astype(str).str.strip() == "包装"])
        st.markdown(
            f"<div style='background:#f5f5f5;padding:6px 12px;border-radius:6px;font-size:13px;'>"
            f"<b style='color:#4A90D9'>注塑 {inj_count}</b> &nbsp;|&nbsp; "
            f"<b style='color:#50B86A'>包装 {pack_count}</b>"
            f"</div>", unsafe_allow_html=True)
        if total_orders > 0:
            chart_df = pd.DataFrame({
                "类型": ["注塑", "包装"],
                "数量": [inj_count, pack_count],
            })
            import plotly.express as px
            fig = px.bar(chart_df, x="类型", y="数量", color="类型",
                         color_discrete_map={"注塑": "#4A90D9", "包装": "#50B86A"},
                         text="数量", height=180)
            fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ===== 第三行：今日排产 + 即将到期 =====
    col_today, col_soon = st.columns([1, 1])

    with col_today:
        st.subheader("📅 今日排产工单")
        today_df = df_wo[df_wo["计划开始日期"].astype(str).str.strip() == today_str].copy()
        if today_df.shape[0] == 0:
            st.info("今日无排产工单")
        else:
            show_cols = ["工单号", "产品品号", "品名", "生产机台号", "班次",
                         "需生产数", "计划结束日期"]
            avail_cols = [c for c in show_cols if c in today_df.columns]
            st.dataframe(today_df[avail_cols], use_container_width=True, hide_index=True, height=250)

    with col_soon:
        st.subheader("⚠️ 即将到期/超期工单")
        soon_list = []
        for _, r in df_wo.iterrows():
            fc = str(r.get("预计完工日期", "")).strip()
            pe = str(r.get("计划结束日期", "")).strip()
            if not fc or fc in ("nan", "<NA>"):
                continue
            try:
                fc_dt = datetime.strptime(fc, "%Y-%m-%d").date()
                today_d = datetime.now().date()
                days_left = (fc_dt - today_d).days
                if days_left <= 3:  # 3天内到期或已超期
                    soon_list.append({
                        "工单号": str(r.get("工单号", "")),
                        "品号": str(r.get("产品品号", "")),
                        "预计完工": fc,
                        "计划结束": pe if pe else "未排产",
                        "状态": "已超期" if (pe and pe > fc) else f"剩余{days_left}天",
                    })
            except (ValueError, TypeError):
                continue
        if not soon_list:
            st.success("近期无到期工单")
        else:
            soon_df = pd.DataFrame(soon_list)
            st.dataframe(soon_df, use_container_width=True, hide_index=True, height=250)

    st.divider()

    # ===== 快捷操作 =====
    st.subheader("⚡ 快捷操作")
    q1, q2, q3, q4 = st.columns(4)
    with q1:
        if st.button("📋 前往工单排产", use_container_width=True):
            st.session_state.page = "工单排产"
            st.rerun()
    with q2:
        if st.button("📊 查看甘特图", use_container_width=True):
            st.session_state.page = "甘特图"
            st.rerun()
    with q3:
        if st.button("🔄 一键全量排产", use_container_width=True, type="primary"):
            df_s = match_product_to_workorder(df_wo, st.session_state.df_injection, st.session_state.df_pack)
            for idx, row in df_s.iterrows():
                oq = safe_float(row["预计产量"])
                pq = safe_float(row["已生产量"])
                np_ = int(max(oq - pq, 0))
                df_s.at[idx, "需生产数"] = np_
                hc = safe_float(row["小时产能 (个)"])
                if hc > 0:
                    df_s.at[idx, "生产周期 (天)"] = calc_work_cycle(
                        np_, hc, row.get("班次", "白班"), row.get("换模时间比例 (%)", 0.0))
            df_s = auto_schedule(df_s, st.session_state.df_machine)
            st.session_state.df_machine = update_machine_status(df_s, st.session_state.df_machine)
            st.session_state.df_workorder = df_s
            st.success("全量排产完成，看板已更新")
            st.rerun()
    with q4:
        if st.button("💾 保存数据", use_container_width=True):
            save_all()
            st.success("数据已保存（含班次时长、排产范围等系统设置）")


# ============================================================
# 页面：登录
# ============================================================
def page_login():
    st.markdown("<h1 style='text-align:center;'>🏭 APS 排产系统</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;color:#888;'>V2.0 网页版 | 请登录</p>", unsafe_allow_html=True)
    st.markdown("---")
    _, col, _ = st.columns([1, 2, 1])
    with col:
        with st.container(border=True):
            st.markdown("### 🔐 用户登录")
            username = st.text_input("用户名", key="login_username", placeholder="请输入用户名")
            password = st.text_input("密码", type="password", key="login_password", placeholder="请输入密码")
            if st.button("登 录", type="primary", use_container_width=True):
                if not username.strip():
                    st.error("请输入用户名")
                elif not password:
                    st.error("请输入密码")
                else:
                    ok, role = verify_login(username, password)
                    if ok:
                        st.session_state.current_user = username.strip()
                        st.session_state.current_role = role
                        st.session_state.page = "生产看板"
                        st.success(f"登录成功！欢迎 {username.strip()}（{role}）")
                        st.rerun()
                    else:
                        st.error("用户名或密码错误")
            st.caption("默认管理员：admin / 123456（登录后可在用户管理中修改）")


# ============================================================
# 页面：用户管理（仅管理员）
# ============================================================
def page_user_management():
    st.subheader("👥 用户管理")
    st.caption("管理员可添加、编辑、删除用户。所有用户共享同一套生产数据。")

    df_users = st.session_state.df_users.copy()

    # 添加新用户
    with st.expander("➕ 添加新用户", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            new_user = st.text_input("用户名", key="new_user_name", placeholder="必填")
        with c2:
            new_pass = st.text_input("密码", key="new_user_pass", type="password", placeholder="必填")
        with c3:
            new_role = st.selectbox("角色", ["操作员", "管理员"], key="new_user_role")
        with c4:
            new_note = st.text_input("备注", key="new_user_note", placeholder="选填")
        if st.button("✅ 添加用户", type="primary", use_container_width=True):
            if not new_user.strip():
                st.error("用户名不能为空")
            elif not new_pass:
                st.error("密码不能为空")
            elif new_user.strip() in df_users["用户名"].astype(str).str.strip().values:
                st.error("用户名已存在")
            else:
                new_row = pd.DataFrame([{
                    "用户名": new_user.strip(),
                    "密码": new_pass,
                    "角色": new_role,
                    "备注": new_note.strip()
                }], columns=COL_USERS)
                st.session_state.df_users = pd.concat([df_users, new_row], ignore_index=True).astype(object)
                save_all()
                st.success(f"已添加用户：{new_user.strip()}")
                st.rerun()

    st.markdown("---")
    st.markdown("**现有用户列表**（可直接编辑用户名、密码、角色、备注）")

    # 可编辑的用户列表
    edited = st.data_editor(
        df_users,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        key="user_editor",
        height=300,
        column_config={
            "用户名": st.column_config.TextColumn("用户名", required=True),
            "密码": st.column_config.TextColumn("密码", required=True),
            "角色": st.column_config.SelectboxColumn("角色", options=["操作员", "管理员"], required=True),
            "备注": st.column_config.TextColumn("备注"),
        }
    )

    if not edited.equals(df_users):
        # 检查用户名重复
        user_names = edited["用户名"].astype(str).str.strip()
        if user_names.duplicated().any():
            st.error("⚠ 用户名重复，未保存。请修改重复的用户名。")
        elif (user_names == "").any():
            st.error("⚠ 用户名不能为空，未保存。")
        else:
            st.session_state.df_users = edited.reindex(columns=COL_USERS).fillna("").astype(object)
            save_all()
            st.success("✅ 用户信息已保存")
            st.rerun()

    # 删除用户（单独按钮，避免误删）
    st.markdown("---")
    with st.expander("🗑 删除用户", expanded=False):
        del_names = [u for u in df_users["用户名"].astype(str).str.strip() if u != "admin"]
        if not del_names:
            st.info("没有可删除的用户（admin不可删除）")
        else:
            del_sel = st.selectbox("选择要删除的用户", del_names, key="del_user_sel")
            if st.button("⚠ 删除该用户", type="secondary", use_container_width=True):
                if del_sel == st.session_state.current_user:
                    st.error("不能删除当前登录用户")
                else:
                    st.session_state.df_users = st.session_state.df_users[
                        st.session_state.df_users["用户名"].astype(str).str.strip() != del_sel
                    ].reset_index(drop=True)
                    save_all()
                    st.success(f"已删除用户：{del_sel}")
                    st.rerun()


# ============================================================
# 路由
# ============================================================
page_map = {
    "生产看板": page_dashboard,
    "工单排产": page_workorder,
    "机台配置": page_machine,
    "注塑配置": page_injection,
    "包装配置": page_pack,
    "甘特图": page_gantt,
    "系统设置": page_settings,
    "用户管理": page_user_management,
}

# 未登录 → 登录页；已登录 → 对应页面
if st.session_state.current_user is None:
    page_login()
else:
    page_map.get(st.session_state.page, page_dashboard)()
