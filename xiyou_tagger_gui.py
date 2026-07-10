# -*- coding: utf-8 -*-
"""
西柚数据清洗打标工具 - 综合版
支持：西柚数据清洗+打标、单独关键词打标、两者同时进行
"""

import os
import re
import sys
import traceback
import threading
import tkinter as tk
import ctypes
from tkinter import ttk, filedialog, messagebox
from typing import List, Optional, Tuple, Dict


def _enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

_enable_dpi_awareness()


def lazy_import_pandas():
    global pd
    import pandas as pd
    return pd


def lazy_import_openpyxl():
    import openpyxl
    return openpyxl


class Rule:
    __slots__ = ("rule_id", "tag_name", "mode", "keywords", "neg_keywords",
                 "ignore_case", "priority", "output_value",
                 "category", "category_priority", "group_values")

    MODES = {
        "包含任一": ["or", "OR", "或", "任一", "any", "包含", "包含或"],
        "包含全部": ["and", "AND", "且", "全部", "all", "必须全部包含", "同时包含"],
        "精确等于": ["equals", "=", "等于", "exact", "完全等于"],
        "开头是": ["starts_with", "startswith", "prefix", "开头", "前缀", "以…开头", "以...开头"],
        "结尾是": ["ends_with", "endswith", "suffix", "结尾", "后缀", "以…结尾", "以...结尾"],
        "正则": ["regex", "Regex", "REGEX", "正则表达式", "pattern"],
    }

    def __init__(self, rule_id, tag_name, mode, keywords, neg_keywords,
                 ignore_case, priority, output_value="",
                 category="", category_priority=None, group_values=None):
        self.rule_id = str(rule_id).strip() if rule_id is not None else ""
        self.tag_name = str(tag_name).strip() if tag_name is not None else ""
        self.mode = self._normalize_mode(mode)
        self.keywords = self._split_kws(keywords, self.mode)
        self.neg_keywords = self._split_kws(neg_keywords, "包含任一")
        self.ignore_case = self._parse_bool(ignore_case)
        out = str(output_value).strip() if output_value is not None else ""
        if out.lower() in ("nan", "none", "null"):
            out = ""
        self.output_value = out
        cat = str(category).strip() if category is not None else ""
        if cat.lower() in ("nan", "none", "null", "-"):
            cat = ""
        self.category = cat
        try:
            cp_raw = str(category_priority).strip() if category_priority is not None else ""
            self.category_priority = int(float(cp_raw)) if cp_raw and cp_raw.lower() not in ("nan", "none") else 9999
        except (ValueError, TypeError):
            self.category_priority = 9999
        self.group_values = self._normalize_group_values(group_values)
        try:
            self.priority = int(float(priority)) if priority is not None and str(priority).strip() != "" else 9999
        except (ValueError, TypeError):
            self.priority = 9999
        self.keywords.sort(key=lambda k: -len(k))

    @staticmethod
    def _normalize_group_values(gv) -> List[str]:
        if gv is None:
            return []
        if isinstance(gv, str):
            s = gv.strip()
            if not s or s.lower() in ("nan", "none", "null", "-"):
                return []
            return [s]
        result = []
        for item in gv:
            if item is None:
                continue
            s = str(item).strip()
            if not s or s.lower() in ("nan", "none", "null", "-"):
                continue
            result.append(s)
        return result

    @staticmethod
    def _normalize_mode(mode) -> str:
        s = str(mode).strip() if mode is not None else "包含任一"
        if not s:
            return "包含任一"
        for canonical, aliases in Rule.MODES.items():
            if s == canonical:
                return canonical
            if s in aliases or s.lower() in [a.lower() for a in aliases]:
                return canonical
        return "包含任一"

    @staticmethod
    def _split_kws(raw, mode: str = "包含任一") -> List[str]:
        if raw is None:
            return []
        try:
            if 'pd' in globals() and isinstance(raw, pd.DataFrame):
                return []
            if hasattr(raw, '__class__') and raw.__class__.__name__ == 'Series':
                if pd.isna(raw):
                    return []
        except Exception:
            pass
        text = str(raw).strip()
        if not text or text.lower() in ("nan", "none", "null"):
            return []
        if mode == "正则":
            for sep in ["\r\n", "\n", ";", "；"]:
                text = text.replace(sep, "\n")
            return [k.strip() for k in text.split("\n") if k.strip()]
        for sep in ["，", "、", "/", ";", "；", "|"]:
            text = text.replace(sep, ",")
        return [k.strip() for k in text.split(",") if k.strip()]

    @staticmethod
    def _parse_bool(val) -> bool:
        if isinstance(val, bool):
            return val
        s = str(val).strip().lower() if val is not None else ""
        return s in ("是", "true", "yes", "y", "1", "忽略", "ignore", "忽略大小写")

    def match(self, text: str) -> Optional[Tuple[str, str]]:
        if not self.keywords or not self.tag_name:
            return None
        haystack = str(text) if text is not None else ""

        if self.ignore_case:
            hay_check = haystack.lower()
            kws_for_match = [k.lower() for k in self.keywords]
            neg_kws = [k.lower() for k in self.neg_keywords]
        else:
            hay_check = haystack
            kws_for_match = list(self.keywords)
            neg_kws = list(self.neg_keywords)

        for nk in neg_kws:
            if nk and nk in hay_check:
                return None

        mode = self.mode
        matched_raw = None
        if mode == "包含任一":
            for kw in kws_for_match:
                if kw in hay_check:
                    matched_raw = self._orig_kw(kw); break
        elif mode == "包含全部":
            if all(kw in hay_check for kw in kws_for_match):
                matched_raw = self._orig_kw(self.keywords[0])
        elif mode == "精确等于":
            for kw in kws_for_match:
                if hay_check == kw or hay_check.strip() == kw:
                    matched_raw = self._orig_kw(kw); break
        elif mode == "开头是":
            for kw in kws_for_match:
                if hay_check.startswith(kw):
                    matched_raw = self._orig_kw(kw); break
        elif mode == "结尾是":
            for kw in kws_for_match:
                if hay_check.endswith(kw):
                    matched_raw = self._orig_kw(kw); break
        elif mode == "正则":
            flags = re.IGNORECASE if self.ignore_case else 0
            for kw in self.keywords:
                try:
                    import re as regex_module
                    m = regex_module.search(kw, haystack, flags=flags)
                    if m:
                        matched_raw = m.group(0) if m.group(0) else kw
                        break
                except Exception:
                    continue
        else:
            for kw in kws_for_match:
                if kw in hay_check:
                    matched_raw = self._orig_kw(kw); break

        if matched_raw is None:
            return None
        cell_value = self.output_value if self.output_value else matched_raw
        return (cell_value, matched_raw)

    def _orig_kw(self, lower_kw: str) -> str:
        if not self.ignore_case:
            return lower_kw
        for orig in self.keywords:
            if orig.lower() == lower_kw:
                return orig
        return lower_kw


def _detect_group_columns(df_cols: List[str]) -> List[str]:
    level_map: Dict[int, str] = {}
    for col in df_cols:
        c = col.strip()
        m = re.match(r'^分组值(\d*)$', c)
        if m:
            suffix = m.group(1)
            level = int(suffix) if suffix else 1
            if level not in level_map:
                level_map[level] = c
            continue
        m = re.match(r'^分组值\.(\d+)$', c)
        if m:
            level = int(m.group(1)) + 1
            if level not in level_map:
                level_map[level] = c
            continue
        m = re.match(r'^分组(\d+)$', c)
        if m:
            level = int(m.group(1))
            if level not in level_map:
                level_map[level] = c
            continue
    return [level_map[k] for k in sorted(level_map.keys())]


def load_rules(filepath: str) -> Tuple[List[Rule], str]:
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    try:
        xl = pd.ExcelFile(filepath)
        sheet_name = '打标规则' if '打标规则' in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        raise ValueError(f"无法读取规则表：{e}")

    df.columns = [str(c).strip() for c in df.columns]
    required = {"标签名称", "关键词"}
    has_label = any(c.strip() == "标签名称" for c in df.columns)
    if not has_label:
        raise ValueError(
            "规则表格式不正确，请使用标准模板。\n\n"
            "规则表必须包含列：标签名称、关键词\n"
            f"当前列名：{list(df.columns)}"
        )
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"规则表缺少必要列：{', '.join(sorted(missing))}\n当前列：{list(df.columns)}")

    group_cols = _detect_group_columns(list(df.columns))

    rules: List[Rule] = []
    for _, row in df.iterrows():
        tag = str(row.get("标签名称", "") or "").strip()
        kw = str(row.get("关键词", "") or "").strip()
        if not tag or not kw:
            continue
        gv_list = []
        for gc in group_cols:
            val = str(row.get(gc, "") or "").strip()
            if val.lower() in ("nan", "none", "null", "-"):
                val = ""
            gv_list.append(val)
        while gv_list and not gv_list[-1]:
            gv_list.pop()
        rules.append(Rule(
            rule_id=row.get("规则ID", ""),
            tag_name=tag,
            mode=row.get("匹配模式", "包含任一"),
            keywords=kw,
            neg_keywords=row.get("否定关键词", "") if "否定关键词" in df.columns else "",
            ignore_case=row.get("忽略大小写", True),
            priority=row.get("优先级", 9999),
            output_value=row.get("输出值", "") if "输出值" in df.columns else "",
            category=row.get("分类", "") if "分类" in df.columns else "",
            category_priority=row.get("分类优先级", None) if "分类优先级" in df.columns else None,
            group_values=gv_list,
        ))

    if not rules:
        raise ValueError("规则表中没有有效的规则。")

    rules.sort(key=lambda r: (r.priority, -len(r.keywords[0]) if r.keywords else 0))
    seen = set()
    out_cols = []
    for r in rules:
        if r.tag_name not in seen:
            seen.add(r.tag_name)
            out_cols.append(r.tag_name)
    max_level = max((len(r.group_values) for r in rules), default=0)
    info = f"共加载 {len(rules)} 条规则，输出 {len(out_cols)} 列"
    if max_level > 0:
        info += f"，{max_level}级分组"
    return rules, info


TAG_OUTPUT_COLUMN = {
    '品牌词': '品牌词',
    '竞品词': '竞品品牌',
    '机型': '机型',
    'case/cover': '否词',
    'screen protector': 'screen protector',
    'tempered glass': 'glass',
    'camera lens': 'camera lens',
    'privacy': 'privacy',
    '磁吸': '磁吸',
    '透明': '透明',
    '硅胶': '硅胶',
    '支架': '支架',
}

FINAL_TAG_COLUMNS = [
    '集合', '分类', 'screen protector', '品牌词', '竞品品牌',
    '机型', 'camera lens', 'privacy', 'glass', '否词'
]


def apply_tags_for_xiyou(df, kw_col, rules):
    result = df.copy()

    tag_order = []
    tag_groups = {}
    for r in rules:
        if r.tag_name not in tag_groups:
            tag_groups[r.tag_name] = []
            tag_order.append(r.tag_name)
        tag_groups[r.tag_name].append(r)

    result['集合'] = "0"
    result['分类'] = "0"

    for tag_name in tag_order:
        out_col = TAG_OUTPUT_COLUMN.get(tag_name)
        if out_col and out_col not in result.columns:
            result[out_col] = "0"

    for idx, row in result.iterrows():
        text = str(row.get(kw_col, "") or "")
        hit_rules = []

        for out_name in tag_order:
            group_rules = tag_groups[out_name]
            for rule in group_rules:
                m = rule.match(text)
                if m is not None:
                    cell_val, matched_kw = m
                    out_col = TAG_OUTPUT_COLUMN.get(out_name)
                    if out_col:
                        result.at[idx, out_col] = str(cell_val)
                    hit_rules.append((rule, cell_val, matched_kw))
                    break

        if hit_rules:
            cat_candidates = [(r, cv, mk) for r, cv, mk in hit_rules
                              if r.category and r.category_priority < 9999]
            if cat_candidates:
                cat_candidates.sort(key=lambda x: (x[0].category_priority, -len(x[2])))
                win_rule, win_val, win_kw = cat_candidates[0]
                cat_name = win_rule.category
                result.at[idx, '集合'] = cat_name
                if win_rule.group_values and win_rule.group_values[0]:
                    result.at[idx, '分类'] = f"{cat_name}-{win_rule.group_values[0]}"
                elif cat_name == '竞品词':
                    result.at[idx, '分类'] = f"{cat_name}-{win_val if win_val else win_kw}"
                else:
                    result.at[idx, '分类'] = cat_name
            else:
                result.at[idx, '集合'] = "否"
                result.at[idx, '分类'] = "无机型词"
        else:
            result.at[idx, '集合'] = "否"
            result.at[idx, '分类'] = "无机型词"

    return result


TARGET_COLUMNS = [
    '站点', 'asin', '品牌', '产品系列', '周期',
    '关键词 (数据来源于西柚找词)', '翻译', '周搜索量去重', '周平均搜索量',
    '流量', '自然流量', '广告流量', '流量占比', '自然流量占比', '广告流量占比',
    '流量获得率', '自然流量获得率', '广告流量获得率',
    '流量分布(自然)', '流量分布(广告)', '展示位置',
    '自然排名', '自然排名(页码-页内排名)', '抓取时间',
    'SP广告排名', 'SP广告排名(页码-页内排名)', '抓取时间1',
    '周平均关键词排名', 'CPC建议竞价($)', '建议竞价范围($)',
    '点击转化率(均值)', '周平均竞争难度', '竞争难度档位',
    '自然位滚动率', 'Top3周平均点击份额', 'Top3周平均转化份额', 'Top3 ASIN',
]


def _extract_info_from_filename(filename):
    pattern = r'关键词反查结果_([A-Za-z]{2})_([A-Z0-9]{10})_(.*)\.xlsx'
    match = re.match(pattern, filename)
    if match:
        return match.group(1), match.group(2), match.group(3)
    return "", "", ""


def load_xiyou_data(folder_path):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    dfs = []
    for filename in os.listdir(folder_path):
        if not filename.endswith('.xlsx'):
            continue
        site, asin, period = _extract_info_from_filename(filename)
        if not asin:
            continue
        filepath = os.path.join(folder_path, filename)
        try:
            df = pd.read_excel(filepath, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            df['站点'] = site
            df['asin'] = asin
            df['周期'] = period
            dfs.append(df)
        except Exception:
            continue
    if not dfs:
        raise ValueError("未找到有效的西柚数据文件")
    combined = pd.concat(dfs, ignore_index=True)
    return combined


def load_mapping(mapping_path):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    df = pd.read_excel(mapping_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    asin_col = None
    for col in df.columns:
        if col.upper() == 'ASIN':
            asin_col = col
            break

    if not asin_col:
        raise ValueError("mapping表必须包含ASIN列")

    df.rename(columns={asin_col: 'ASIN'}, inplace=True)
    df = df.drop_duplicates(subset=['ASIN'], keep='first')

    return df


def clean_xiyou_data(raw_df, mapping_df):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for old_name, new_name in [
        ('抓取时间.1', '抓取时间1'),
        ('CPC建议竞价(£)', 'CPC建议竞价($)'),
        ('建议竞价范围(£)', '建议竞价范围($)'),
    ]:
        if old_name in df.columns:
            df.rename(columns={old_name: new_name}, inplace=True)

    df['asin'] = df['asin'].astype(str).str.strip()
    mapping_df['ASIN'] = mapping_df['ASIN'].astype(str).str.strip()
    df = df.merge(mapping_df, left_on='asin', right_on='ASIN', how='left')

    for col in ['品牌', '产品系列', '机型', '细分']:
        if col not in df.columns:
            df[col] = ''
        else:
            df[col] = df[col].fillna('')

    kw_col = '关键词 (数据来源于西柚找词)'
    if kw_col in df.columns:
        first_occurrence = df.groupby(kw_col).cumcount() == 0
        df['周搜索量去重'] = df['周平均搜索量'].where(first_occurrence, '')

    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ''

    df = df[TARGET_COLUMNS]
    return df


def create_tag_sheet(cleaned_df):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    kw_col = '关键词 (数据来源于西柚找词)'
    unique_df = cleaned_df.drop_duplicates(subset=[kw_col]).copy()
    unique_df['__sort_val'] = pd.to_numeric(unique_df['周平均搜索量'], errors='coerce')
    unique_df = unique_df.sort_values(by='__sort_val', ascending=False)

    result = unique_df[[kw_col, '周搜索量去重'] + FINAL_TAG_COLUMNS].copy()
    result.rename(columns={kw_col: '关键词'}, inplace=True)
    return result


def load_keywords_data(filepath):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    df = pd.read_excel(filepath, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    kw_col = None
    for col in df.columns:
        if col == '关键词' or 'keyword' in col.lower() or '搜索词' in col:
            kw_col = col
            break
    if kw_col is None:
        kw_col = df.columns[0]

    df = df[[kw_col]].copy()
    df.rename(columns={kw_col: '关键词'}, inplace=True)
    df['关键词'] = df['关键词'].astype(str).str.strip()
    df = df.drop_duplicates(subset=['关键词'])
    df = df[df['关键词'] != '']
    return df


def apply_tags_to_keywords(keywords_df, rules):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()

    result = keywords_df.copy()

    has_category = any(r.category and r.category_priority < 9999 for r in rules)

    user_tag_names = []
    user_tag_set = set()
    for r in rules:
        if r.tag_name not in user_tag_set:
            user_tag_set.add(r.tag_name)
            user_tag_names.append(r.tag_name)

    RESERVED_SUMMARY = {"大分类", "KW/ASIN"}
    overridden_summary = set()
    if has_category:
        for tn in user_tag_names:
            if tn in RESERVED_SUMMARY:
                overridden_summary.add(tn)

    tag_order = []
    tag_groups = {}
    for r in rules:
        if r.tag_name not in tag_groups:
            tag_groups[r.tag_name] = []
            tag_order.append(r.tag_name)
        tag_groups[r.tag_name].append(r)

    max_group_level = 0
    if has_category:
        for r in rules:
            if r.category and r.category_priority < 9999:
                max_group_level = max(max_group_level, len(r.group_values))
        if max_group_level == 0:
            max_group_level = 1

    cat_col = "大分类" if "大分类" not in overridden_summary else None
    kw_type_col = "KW/ASIN" if "KW/ASIN" not in overridden_summary else None
    group_cols = [f"分组{i+1}" for i in range(max_group_level)]

    summary_cols = []
    if cat_col:
        summary_cols.append(cat_col)
    summary_cols.extend(group_cols)
    if kw_type_col:
        summary_cols.append(kw_type_col)

    for tc in tag_order:
        result[tc] = ""

    if has_category:
        for c in summary_cols:
            result[c] = ""

    for idx, row in result.iterrows():
        text = str(row.get('关键词', "") or "")
        hit_rules = []

        for out_name in tag_order:
            group_rules = tag_groups[out_name]
            for rule in group_rules:
                m = rule.match(text)
                if m is not None:
                    cell_val, matched_kw = m
                    result.at[idx, out_name] = cell_val
                    hit_rules.append((rule, cell_val, matched_kw))
                    break

        if has_category and hit_rules and summary_cols:
            cat_candidates = [(r, cv, mk) for r, cv, mk in hit_rules
                              if r.category and r.category_priority < 9999]
            if cat_candidates:
                cat_candidates.sort(key=lambda x: (x[0].category_priority, -len(x[2])))
                win_rule, win_val, win_kw = cat_candidates[0]
                cat_name = win_rule.category

                is_asin = cat_name.upper() == "ASIN" or win_rule.tag_name.upper() in ("ASIN", "ASIN_")

                if kw_type_col:
                    result.at[idx, kw_type_col] = "ASIN" if is_asin else "KW"
                if cat_col:
                    result.at[idx, cat_col] = cat_name

                if is_asin:
                    for gc in group_cols:
                        result.at[idx, gc] = "ASIN"
                else:
                    for level, gc in enumerate(group_cols):
                        if level < len(win_rule.group_values) and win_rule.group_values[level]:
                            result.at[idx, gc] = f"{cat_name}-{win_rule.group_values[level]}"
                        elif level == 0:
                            label = win_val if win_val else win_kw
                            result.at[idx, gc] = f"{cat_name}_{label}"

    if has_category:
        orig_cols = [c for c in keywords_df.columns]
        kw_idx = orig_cols.index('关键词')
        ordered_summary = []
        if cat_col:
            ordered_summary.append(cat_col)
        elif "大分类" in overridden_summary:
            ordered_summary.append("大分类")
        ordered_summary.extend(group_cols)
        if kw_type_col:
            ordered_summary.append(kw_type_col)
        elif "KW/ASIN" in overridden_summary:
            ordered_summary.append("KW/ASIN")
        remaining_tags = [t for t in tag_order if t not in overridden_summary]
        remaining_orig = [c for c in orig_cols[kw_idx+1:] if c not in tag_order and c not in set(summary_cols)]
        new_cols = orig_cols[:kw_idx+1] + ordered_summary + remaining_tags + remaining_orig
        result = result[new_cols]

    return result


def save_result(output_path, cleaned_df=None, tag_df=None, keyword_tag_df=None):
    global pd
    if 'pd' not in globals():
        pd = lazy_import_pandas()
    lazy_import_openpyxl()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if cleaned_df is not None:
            cleaned_df.to_excel(writer, index=False, sheet_name="西柚数据源")
        if tag_df is not None:
            tag_df.to_excel(writer, index=False, sheet_name="打标")
        if keyword_tag_df is not None:
            keyword_tag_df.to_excel(writer, index=False, sheet_name="关键词打标")


class XiYouTaggerApp:
    BG_COLOR = "#FFFFFF"
    CARD_BG = "#FFFFFF"
    PRIMARY = "#FF6A00"
    PRIMARY_HOVER = "#FF8533"
    PRIMARY_ACTIVE = "#E55F00"
    TEXT = "#333333"
    MUTED = "#666666"
    BORDER = "#FFE8D6"

    def __init__(self, root):
        self.root = root
        self.root.title("西柚数据清洗打标工具")
        self.root.geometry("1500x1050")
        self.root.minsize(960, 700)
        self.root.configure(bg=self.BG_COLOR)

        self._set_window_icon()

        self.rules_path = None
        self.xiyou_folder = None
        self.mapping_path = None
        self.keywords_path = None
        self.rules = []
        self.cleaned_df = None
        self.tag_df = None
        self.keyword_tag_df = None
        self._is_running = False

        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        FONT_FAMILY = "Microsoft YaHei UI"
        FONT_BOLD = (FONT_FAMILY, 10, "bold")

        style.configure("TFrame", background=self.BG_COLOR)
        style.configure("Card.TFrame", background=self.CARD_BG)
        style.configure("TLabel", background=self.BG_COLOR, foreground=self.TEXT, font=(FONT_FAMILY, 10))
        style.configure("Title.TLabel", background=self.BG_COLOR, foreground=self.TEXT, font=(FONT_FAMILY, 19, "bold"))
        style.configure("CardBold.TLabel", background=self.CARD_BG, foreground=self.TEXT, font=FONT_BOLD)
        style.configure("CardSub.TLabel", background=self.CARD_BG, foreground=self.MUTED, font=(FONT_FAMILY, 9))
        style.configure("Status.TLabel", background=self.CARD_BG, foreground=self.MUTED, font=(FONT_FAMILY, 9))

        style.configure("Primary.TButton",
                        background=self.PRIMARY, foreground="white", font=(FONT_FAMILY, 10, "bold"),
                        padding=(22, 9), borderwidth=0, relief="flat", focuscolor="none")
        style.map("Primary.TButton",
                  background=[("active", self.PRIMARY_HOVER), ("pressed", self.PRIMARY_ACTIVE), ("disabled", "#FFD8B8")],
                  foreground=[("disabled", "#FFFFFF")])

        style.configure("Upload.TButton",
                        background=self.CARD_BG, foreground=self.TEXT, font=(FONT_FAMILY, 9),
                        padding=(14, 6), borderwidth=1, relief="solid", bordercolor=self.BORDER, focuscolor="none")
        style.map("Upload.TButton",
                  background=[("active", "#FFF5EB"), ("pressed", "#FFE8D6")],
                  bordercolor=[("active", self.PRIMARY), ("pressed", self.PRIMARY)],
                  foreground=[("active", self.PRIMARY)])

        style.configure("Export.TButton",
                        background=self.CARD_BG, foreground=self.PRIMARY, font=(FONT_FAMILY, 10, "bold"),
                        padding=(18, 8), borderwidth=1, relief="solid", bordercolor=self.PRIMARY, focuscolor="none")
        style.map("Export.TButton",
                  background=[("active", "#FFF5EB"), ("pressed", "#FFE8D6"), ("disabled", self.CARD_BG)],
                  foreground=[("disabled", "#B2B2B2")], bordercolor=[("disabled", self.BORDER)])

        btn_tpl_pad = (12, 6)
        style.configure("Template.TButton",
                        background=self.CARD_BG, foreground="#555555", font=(FONT_FAMILY, 9),
                        padding=btn_tpl_pad, borderwidth=1, relief="solid", bordercolor=self.BORDER, focuscolor="none")
        style.map("Template.TButton",
                  background=[("active", "#FFF5EB"), ("pressed", "#FFE8D6")],
                  foreground=[("active", self.PRIMARY)])

        style.configure("TNotebook", background=self.CARD_BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab",
                        background=self.CARD_BG, foreground=self.MUTED, font=(FONT_FAMILY, 10),
                        padding=(20, 10), borderwidth=0, focuscolor="none")
        style.map("TNotebook.Tab",
                  background=[("selected", self.CARD_BG), ("active", "#FFF5EB")],
                  foreground=[("selected", self.PRIMARY)],
                  font=[("selected", (FONT_FAMILY, 10, "bold"))])

        style.configure("Treeview",
                        background=self.CARD_BG, fieldbackground=self.CARD_BG, foreground=self.TEXT,
                        rowheight=32, font=(FONT_FAMILY, 9), borderwidth=0, relief="flat")
        style.configure("Treeview.Heading",
                        background="#FFF5EB", foreground=self.MUTED, font=(FONT_FAMILY, 9, "bold"),
                        borderwidth=0, relief="flat", padding=(10, 8))
        style.map("Treeview", background=[("selected", "#FFE8D6")], foreground=[("selected", self.TEXT)])
        style.map("Treeview.Heading", background=[("active", "#FFE8D6")])

        style.configure("Vertical.TScrollbar", background=self.BORDER, troughcolor=self.CARD_BG,
                        borderwidth=0, arrowcolor=self.MUTED, relief="flat")
        style.configure("Horizontal.TScrollbar", background=self.BORDER, troughcolor=self.CARD_BG,
                        borderwidth=0, arrowcolor=self.MUTED, relief="flat")
        style.map("Vertical.TScrollbar",
                  background=[("active", self.PRIMARY), ("pressed", self.PRIMARY_HOVER)])
        style.map("Horizontal.TScrollbar",
                  background=[("active", self.PRIMARY), ("pressed", self.PRIMARY_HOVER)])

    def _build_ui(self):
        header = ttk.Frame(self.root, style="TFrame")
        header.pack(fill="x", padx=24, pady=(22, 4))

        ttk.Label(header, text="西柚数据清洗打标工具", style="Title.TLabel").pack(side="left")

        tpl_box = tk.Frame(header, bg=self.BG_COLOR)
        tpl_box.pack(side="right")
        ttk.Button(tpl_box, text="生成Excel模板", style="Template.TButton",
                   command=self._on_template).pack(side="right")
        ttk.Button(tpl_box, text="保存模板", style="Template.TButton",
                   command=self._on_save_template).pack(side="right", padx=(0, 8))
        ttk.Button(tpl_box, text="加载模板", style="Template.TButton",
                   command=self._on_load_template).pack(side="right", padx=(0, 8))

        file_card = self._make_card(self.root, pady=(6, 6))

        r1 = tk.Frame(file_card, bg=self.CARD_BG)
        r1.pack(fill="x", padx=22, pady=(16, 12))
        self._build_file_row(r1, "西柚数据文件夹", self._upload_xiyou_folder, True)

        tk.Frame(file_card, bg=self.BORDER, height=1).pack(fill="x", padx=22)

        r2 = tk.Frame(file_card, bg=self.CARD_BG)
        r2.pack(fill="x", padx=22, pady=(12, 12))
        self._build_file_row(r2, "ASIN mapping", self._upload_mapping, False)

        tk.Frame(file_card, bg=self.BORDER, height=1).pack(fill="x", padx=22)

        r3 = tk.Frame(file_card, bg=self.CARD_BG)
        r3.pack(fill="x", padx=22, pady=(12, 12))
        self._build_file_row(r3, "关键词数据表", self._upload_keywords, False)

        tk.Frame(file_card, bg=self.BORDER, height=1).pack(fill="x", padx=22)

        r4 = tk.Frame(file_card, bg=self.CARD_BG)
        r4.pack(fill="x", padx=22, pady=(12, 16))
        self._build_file_row(r4, "打标规则表", self._upload_rules, False)

        action_card = self._make_card(self.root, pady=(0, 6))
        ai = tk.Frame(action_card, bg=self.CARD_BG)
        ai.pack(fill="x", padx=22, pady=14)
        self.mode_label = ttk.Label(ai, text="两种操作均需打标规则：【西柚流量词分析】西柚数据文件夹+ASIN mapping  |  【单独打标】关键词数据表", style="CardSub.TLabel")
        self.mode_label.pack(side="left")
        self.start_btn = ttk.Button(ai, text="开始处理", style="Primary.TButton", command=self._on_start)
        self.start_btn.pack(side="right", padx=(10, 0))
        self.export_btn = ttk.Button(ai, text="导出结果", style="Export.TButton",
                                     command=self._on_export, state="disabled")
        self.export_btn.pack(side="right")

        nb_card = self._make_card(self.root, pady=(0, 6), expand=True)
        self.notebook = ttk.Notebook(nb_card)
        self.notebook.pack(fill="both", expand=True, padx=0, pady=0)

        rules_frame = tk.Frame(self.notebook, bg=self.CARD_BG)
        self.notebook.add(rules_frame, text="  规则列表  ")
        rh = tk.Frame(rules_frame, bg=self.CARD_BG)
        rh.pack(fill="x", padx=22, pady=(16, 10))
        ttk.Label(rh, text="当前已加载的规则", style="CardBold.TLabel").pack(side="left")
        self.rules_info_lbl = ttk.Label(rh, text="未加载规则", style="CardSub.TLabel")
        self.rules_info_lbl.pack(side="right")

        rules_tree_frame = tk.Frame(rules_frame, bg=self.CARD_BG)
        rules_tree_frame.pack(fill="both", expand=True, padx=22, pady=(0, 16))
        self.rules_tree = ttk.Treeview(rules_tree_frame, show="headings", selectmode="browse",
                                       height=10, style="Treeview")
        rvsb = ttk.Scrollbar(rules_tree_frame, orient="vertical", command=self.rules_tree.yview)
        rhsb = ttk.Scrollbar(rules_tree_frame, orient="horizontal", command=self.rules_tree.xview)
        self.rules_tree.configure(yscrollcommand=rvsb.set, xscrollcommand=rhsb.set)
        self.rules_tree.grid(row=0, column=0, sticky="nsew")
        rvsb.grid(row=0, column=1, sticky="ns")
        rhsb.grid(row=1, column=0, sticky="ew")
        rules_tree_frame.rowconfigure(0, weight=1)
        rules_tree_frame.columnconfigure(0, weight=1)

        xiyou_frame = tk.Frame(self.notebook, bg=self.CARD_BG)
        self.notebook.add(xiyou_frame, text="  西柚数据预览  ")
        xh = tk.Frame(xiyou_frame, bg=self.CARD_BG)
        xh.pack(fill="x", padx=22, pady=(16, 10))
        ttk.Label(xh, text="西柚数据清洗打标结果预览（前500行）", style="CardBold.TLabel").pack(side="left")
        self.xiyou_info_lbl = ttk.Label(xh, text="尚未处理", style="CardSub.TLabel")
        self.xiyou_info_lbl.pack(side="right")

        xiyou_tree_frame = tk.Frame(xiyou_frame, bg=self.CARD_BG)
        xiyou_tree_frame.pack(fill="both", expand=True, padx=22, pady=(0, 16))
        self.xiyou_tree = ttk.Treeview(xiyou_tree_frame, show="headings", selectmode="browse", style="Treeview")
        xvsb = ttk.Scrollbar(xiyou_tree_frame, orient="vertical", command=self.xiyou_tree.yview)
        xhsb = ttk.Scrollbar(xiyou_tree_frame, orient="horizontal", command=self.xiyou_tree.xview)
        self.xiyou_tree.configure(yscrollcommand=xvsb.set, xscrollcommand=xhsb.set)
        self.xiyou_tree.grid(row=0, column=0, sticky="nsew")
        xvsb.grid(row=0, column=1, sticky="ns")
        xhsb.grid(row=1, column=0, sticky="ew")
        xiyou_tree_frame.rowconfigure(0, weight=1)
        xiyou_tree_frame.columnconfigure(0, weight=1)

        keyword_frame = tk.Frame(self.notebook, bg=self.CARD_BG)
        self.notebook.add(keyword_frame, text="  关键词打标预览  ")
        kh = tk.Frame(keyword_frame, bg=self.CARD_BG)
        kh.pack(fill="x", padx=22, pady=(16, 10))
        ttk.Label(kh, text="关键词打标结果预览（前500行）", style="CardBold.TLabel").pack(side="left")
        self.keyword_info_lbl = ttk.Label(kh, text="尚未处理", style="CardSub.TLabel")
        self.keyword_info_lbl.pack(side="right")

        keyword_tree_frame = tk.Frame(keyword_frame, bg=self.CARD_BG)
        keyword_tree_frame.pack(fill="both", expand=True, padx=22, pady=(0, 16))
        self.keyword_tree = ttk.Treeview(keyword_tree_frame, show="headings", selectmode="browse", style="Treeview")
        kvsb = ttk.Scrollbar(keyword_tree_frame, orient="vertical", command=self.keyword_tree.yview)
        khsb = ttk.Scrollbar(keyword_tree_frame, orient="horizontal", command=self.keyword_tree.xview)
        self.keyword_tree.configure(yscrollcommand=kvsb.set, xscrollcommand=khsb.set)
        self.keyword_tree.grid(row=0, column=0, sticky="nsew")
        kvsb.grid(row=0, column=1, sticky="ns")
        khsb.grid(row=1, column=0, sticky="ew")
        keyword_tree_frame.rowconfigure(0, weight=1)
        keyword_tree_frame.columnconfigure(0, weight=1)

        status_bar = self._make_card(self.root, pady=(0, 12))
        self.status_var = tk.StringVar(value="就绪。")
        self.status_dot = tk.Label(status_bar, text="●", bg=self.CARD_BG, fg=self.PRIMARY, font=("Segoe UI", 11))
        self.status_dot.pack(side="left", padx=(18, 8), pady=10)
        sb = tk.Label(status_bar, textvariable=self.status_var, bg=self.CARD_BG, fg=self.MUTED,
                      anchor="w", font=("Microsoft YaHei UI", 9))
        sb.pack(side="left", fill="x", expand=True, pady=10)

    def _make_card(self, parent, pady=(0, 0), expand=False):
        outer = tk.Frame(parent, bg=self.BORDER)
        outer.pack(fill="both" if expand else "x", expand=expand, padx=24, pady=pady)
        card = tk.Frame(outer, bg=self.CARD_BG, highlightbackground=self.BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=(0, 1), pady=(0, 1))
        return card

    def _build_file_row(self, parent, label, cmd, is_folder=False):
        tk.Label(parent, text=label, bg=self.CARD_BG, fg=self.TEXT,
                 font=("Microsoft YaHei UI", 10, "bold"), width=12, anchor="w").pack(side="left")
        btn_text = "选择文件夹" if is_folder else "选择文件"
        btn = ttk.Button(parent, text=btn_text, style="Upload.TButton", command=cmd)
        btn.pack(side="right")
        
        rm_btn = ttk.Button(parent, text="✕", style="Upload.TButton", width=3,
                           command=lambda l=label: self._clear_file(l))
        rm_btn.pack(side="right", padx=(4, 8))
        
        var = tk.StringVar(value=f"未选择{label}")
        lbl = tk.Label(parent, textvariable=var, bg=self.CARD_BG, fg=self.MUTED,
                       font=("Microsoft YaHei UI", 9), anchor="w")
        lbl.pack(side="left", fill="x", expand=True, padx=14)
        if label == "西柚数据文件夹":
            self.xiyou_btn = btn
            self.xiyou_rm_btn = rm_btn
            self.xiyou_var = var
            self.xiyou_lbl = lbl
        elif label == "ASIN mapping":
            self.mapping_btn = btn
            self.mapping_rm_btn = rm_btn
            self.mapping_var = var
            self.mapping_lbl = lbl
        elif label == "关键词数据表":
            self.keywords_btn = btn
            self.keywords_rm_btn = rm_btn
            self.keywords_var = var
            self.keywords_lbl = lbl
        elif label == "打标规则表":
            self.rule_btn = btn
            self.rule_rm_btn = rm_btn
            self.rule_path_var = var
            self.rule_path_lbl = lbl

    def _clear_file(self, label):
        if label == "西柚数据文件夹":
            self.xiyou_folder = None
            self.xiyou_var.set("未选择西柚数据文件夹")
            self.cleaned_df = None
            self.tag_df = None
        elif label == "ASIN mapping":
            self.mapping_path = None
            self.mapping_var.set("未选择ASIN mapping")
            self.cleaned_df = None
            self.tag_df = None
        elif label == "关键词数据表":
            self.keywords_path = None
            self.keywords_var.set("未选择关键词数据表")
            self.keyword_tag_df = None
        elif label == "打标规则表":
            self.rules_path = None
            self.rule_path_var.set("未选择打标规则表")
            self.rules = []
            self._clear_rules_tree()
        self._clear_xiyou_tree()
        self._clear_keyword_tree()
        self.export_btn.configure(state="disabled")
        self._update_mode()

    def _clear_rules_tree(self):
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)
        self.rules_info_lbl.configure(text="未加载规则")

    def _clear_xiyou_tree(self):
        for item in self.xiyou_tree.get_children():
            self.xiyou_tree.delete(item)
        self.xiyou_info_lbl.configure(text="尚未处理")

    def _clear_keyword_tree(self):
        for item in self.keyword_tree.get_children():
            self.keyword_tree.delete(item)
        self.keyword_info_lbl.configure(text="尚未处理")

    def _set_window_icon(self):
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
        if not os.path.exists(icon_path):
            icon_path = os.path.join(os.getcwd(), "app_icon.ico")
        try:
            self.root.iconbitmap(icon_path)
        except Exception:
            pass
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("xiyou.tagger.1")
        except Exception:
            pass

    def _set_status(self, msg, color=None):
        self.status_var.set(msg)
        if color and hasattr(self, 'status_dot'):
            self.status_dot.configure(fg=color)
        self.root.update_idletasks()

    def _on_template(self):
        global pd
        if 'pd' not in globals():
            pd = lazy_import_pandas()

        path = filedialog.asksaveasfilename(
            title="保存规则模板", defaultextension=".xlsx",
            initialfile="打标规则模板.xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            template = pd.DataFrame([
                {"规则ID": 1, "标签名称": "KW/ASIN", "匹配模式": "开头是",
                 "关键词": "B0", "否定关键词": "", "忽略大小写": "否", "优先级": 1, "输出值": "",
                 "分类": "ASIN", "分类优先级": 1, "分组值": ""},
                {"规则ID": 2, "标签名称": "品牌词", "匹配模式": "包含任一",
                 "关键词": "ESR,halolock", "否定关键词": "", "忽略大小写": "是", "优先级": 2, "输出值": "",
                 "分类": "品牌词", "分类优先级": 2, "分组值": ""},
                {"规则ID": 3, "标签名称": "竞品词", "匹配模式": "包含任一",
                 "关键词": "magic john,miracase,otterbox,belkin,jetech,spigen,tech21,torras,tauri,tocol",
                 "否定关键词": "", "忽略大小写": "是", "优先级": 3, "输出值": "",
                 "分类": "竞品词", "分类优先级": 3, "分组值": ""},
                {"规则ID": 4, "标签名称": "机型", "匹配模式": "包含全部",
                 "关键词": "pro,max,17", "否定关键词": "", "忽略大小写": "是", "优先级": 6, "输出值": "17PM",
                 "分类": "机型词", "分类优先级": 40, "分组值": ""},
                {"规则ID": 5, "标签名称": "case/cover", "匹配模式": "包含全部",
                 "关键词": "case", "否定关键词": "", "忽略大小写": "是", "优先级": 18, "输出值": "case",
                 "分类": "大词", "分类优先级": 19, "分组值": "壳"},
                {"规则ID": 6, "标签名称": "screen protector", "匹配模式": "包含全部",
                 "关键词": "screen,protector", "否定关键词": "camera lens,lens protector",
                 "忽略大小写": "是", "优先级": 19, "输出值": "screen protector",
                 "分类": "大词", "分类优先级": 30, "分组值": "膜"},
                {"规则ID": 7, "标签名称": "tempered glass", "匹配模式": "包含全部",
                 "关键词": "tempered,glass", "否定关键词": "screen protector,camera,lens",
                 "忽略大小写": "是", "优先级": 22, "输出值": "glass",
                 "分类": "小词", "分类优先级": 20, "分组值": "膜"},
                {"规则ID": 8, "标签名称": "camera lens", "匹配模式": "包含全部",
                 "关键词": "camera lens", "否定关键词": "", "忽略大小写": "是", "优先级": 24, "输出值": "camera",
                 "分类": "小词", "分类优先级": 15, "分组值": "镜头膜"},
                {"规则ID": 9, "标签名称": "privacy", "匹配模式": "包含全部",
                 "关键词": "privacy", "否定关键词": "", "忽略大小写": "是", "优先级": 27, "输出值": "privacy",
                 "分类": "小词", "分类优先级": 15, "分组值": "防窥膜"},
                {"规则ID": 10, "标签名称": "磁吸", "匹配模式": "正则",
                 "关键词": "magsafe", "否定关键词": "", "忽略大小写": "是", "优先级": 32, "输出值": "magsafe",
                 "分类": "小词", "分类优先级": 16, "分组值": "磁吸"},
                {"规则ID": 11, "标签名称": "透明", "匹配模式": "包含全部",
                 "关键词": "clear", "否定关键词": "", "忽略大小写": "是", "优先级": 36, "输出值": "",
                 "分类": "小词", "分类优先级": 17, "分组值": "透明"},
                {"规则ID": 12, "标签名称": "硅胶", "匹配模式": "包含全部",
                 "关键词": "silicone", "否定关键词": "", "忽略大小写": "是", "优先级": 40, "输出值": "",
                 "分类": "小词", "分类优先级": 18, "分组值": "硅胶"},
                {"规则ID": 13, "标签名称": "支架", "匹配模式": "包含全部",
                 "关键词": "stand", "否定关键词": "", "忽略大小写": "是", "优先级": 43, "输出值": "",
                 "分类": "小词", "分类优先级": 15, "分组值": "支架"},
            ])
            help_df = pd.DataFrame([
                {"列名": "规则ID", "说明": "序号，不影响结果"},
                {"列名": "标签名称", "说明": "输出列名。相同标签名的规则会输出到同一列（按优先级竞争）"},
                {"列名": "匹配模式", "说明": "包含任一/包含全部/精确等于/开头是/结尾是/正则"},
                {"列名": "关键词", "说明": "要匹配的词，多个用英文逗号分隔；支持短语"},
                {"列名": "否定关键词", "说明": "包含这些词则不命中此规则；可选"},
                {"列名": "忽略大小写", "说明": "是/否；英文一般填是"},
                {"列名": "优先级", "说明": "数字越小越优先；同标签名下决定哪条规则胜出"},
                {"列名": "输出值", "说明": "【可选】命中时填入单元格的自定义值；不填则填入匹配到的关键词"},
                {"列名": "分类", "说明": "【可选】大分类名（如ASIN/品牌词/大词/小词）"},
                {"列名": "分类优先级", "说明": "【可选】数字越小越优先；决定「集合」列取哪个分类"},
                {"列名": "分组值", "说明": "【可选】分组名（如壳/膜/镜头膜/防窥膜）"},
            ])
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                template.to_excel(writer, index=False, sheet_name="打标规则")
                help_df.to_excel(writer, index=False, sheet_name="使用说明")
            messagebox.showinfo("模板已生成",
                                f"规则模板已保存到：\n{path}\n\n"
                                "模板包含两个Sheet：\n"
                                "1.「打标规则」：预置13条标签规则\n"
                                "2.「使用说明」：每列含义")
            self._set_status(f"模板已生成：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("生成失败", str(e))

    def _upload_xiyou_folder(self):
        path = filedialog.askdirectory(title="选择西柚下载数据文件夹")
        if not path:
            return
        xlsx_files = [f for f in os.listdir(path) if f.endswith('.xlsx')]
        if not xlsx_files:
            messagebox.showwarning("提示", "该文件夹中没有找到Excel文件")
            return
        self.xiyou_folder = path
        self.xiyou_var.set(f"✓ {os.path.basename(path)}  （{len(xlsx_files)}个Excel文件）")
        self.xiyou_lbl.configure(fg=self.PRIMARY)
        self._update_mode()

    def _upload_mapping(self):
        path = filedialog.askopenfilename(title="选择ASIN mapping表", filetypes=[("Excel", "*.xlsx *.xls"), ("所有", "*.*")])
        if not path:
            return
        try:
            global pd
            if 'pd' not in globals():
                pd = lazy_import_pandas()
            df = pd.read_excel(path, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            asin_found = any(col.upper() == 'ASIN' for col in df.columns)
            if not asin_found:
                raise ValueError("mapping表必须包含ASIN列")
        except Exception as e:
            messagebox.showerror("mapping表错误", str(e))
            return
        self.mapping_path = path
        self.mapping_var.set(f"✓ {os.path.basename(path)}  （{len(df)}行）")
        self.mapping_lbl.configure(fg=self.PRIMARY)
        self._update_mode()

    def _upload_keywords(self):
        path = filedialog.askopenfilename(title="选择关键词数据表", filetypes=[("Excel", "*.xlsx *.xls"), ("所有", "*.*")])
        if not path:
            return
        try:
            global pd
            if 'pd' not in globals():
                pd = lazy_import_pandas()
            df = pd.read_excel(path, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            kw_col = None
            for col in df.columns:
                if col == '关键词' or 'keyword' in col.lower() or '搜索词' in col:
                    kw_col = col
                    break
            if kw_col is None:
                kw_col = df.columns[0]
        except Exception as e:
            messagebox.showerror("关键词数据表错误", str(e))
            return
        self.keywords_path = path
        self.keywords_var.set(f"✓ {os.path.basename(path)}  （{len(df)}行）")
        self.keywords_lbl.configure(fg=self.PRIMARY)
        self._update_mode()

    def _upload_rules(self):
        path = filedialog.askopenfilename(title="选择打标规则表", filetypes=[("Excel", "*.xlsx *.xls"), ("所有", "*.*")])
        if not path:
            return
        try:
            rules, info = load_rules(path)
        except Exception as e:
            messagebox.showerror("规则表错误", str(e))
            return
        self.rules_path = path
        self.rules = rules
        self.rule_path_var.set(f"✓ {os.path.basename(path)}  （{info}）")
        self.rule_path_lbl.configure(fg=self.PRIMARY)
        self._refresh_rules_tree()
        self.notebook.select(0)
        self._update_mode()

    def _update_mode(self):
        has_xiyou = self.xiyou_folder is not None and self.mapping_path is not None
        has_keywords = self.keywords_path is not None
        has_rules = len(self.rules) > 0

        if has_xiyou and has_keywords and has_rules:
            self.mode_label.configure(text="模式：西柚数据清洗打标 + 关键词单独打标")
            self.start_btn.configure(state="normal")
        elif has_xiyou and has_rules:
            self.mode_label.configure(text="模式：西柚数据清洗打标")
            self.start_btn.configure(state="normal")
        elif has_keywords and has_rules:
            self.mode_label.configure(text="模式：关键词单独打标")
            self.start_btn.configure(state="normal")
        else:
            self.mode_label.configure(text="请选择要执行的操作")
            self.start_btn.configure(state="disabled")

    def _refresh_rules_tree(self):
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)
        if not self.rules:
            self.rules_info_lbl.configure(text="未加载规则")
            return
        max_gv = max((len(r.group_values) for r in self.rules), default=1)
        if max_gv < 1:
            max_gv = 1
        group_col_names = []
        for i in range(max_gv):
            label = "分组值" if i == 0 else f"分组值{i+1}"
            group_col_names.append(label)
        cols = ("优先级", "标签名称(输出列)", "匹配模式", "关键词", "否定关键词",
                "忽略大小写", "输出值", "分类", "分类优先级") + tuple(group_col_names)
        self.rules_tree["columns"] = cols
        widths = {"优先级": 50, "标签名称(输出列)": 110, "匹配模式": 70, "关键词": 220,
                  "否定关键词": 110, "忽略大小写": 65, "输出值": 80,
                  "分类": 65, "分类优先级": 65}
        for gc in group_col_names:
            widths[gc] = 65
        for c in cols:
            self.rules_tree.heading(c, text=c)
            self.rules_tree.column(c, width=widths.get(c, 70), minwidth=40, anchor="w", stretch=True)
        for r in self.rules:
            kw_text = ",".join(r.keywords)
            if len(kw_text) > 60:
                kw_text = kw_text[:57] + "..."
            neg_text = ",".join(r.neg_keywords)
            if len(neg_text) > 30:
                neg_text = neg_text[:27] + "..."
            out_text = r.output_value if r.output_value else "(命中关键词)"
            cat_text = r.category if r.category else "-"
            cp_text = str(r.category_priority) if r.category_priority < 9999 else "-"
            values = [r.priority, r.tag_name, r.mode, kw_text,
                      neg_text or "-", "是" if r.ignore_case else "否", out_text,
                      cat_text, cp_text]
            for i in range(max_gv):
                values.append(r.group_values[i] if i < len(r.group_values) and r.group_values[i] else "-")
            row_tag = "oddrow" if len(self.rules_tree.get_children()) % 2 == 0 else "evenrow"
            self.rules_tree.insert("", "end", values=values, tags=(row_tag,))
        out_cols = []
        for r in self.rules:
            if r.tag_name not in out_cols:
                out_cols.append(r.tag_name)
        group_info = f"{max_gv}级分组" if max_gv > 0 else ""
        self.rules_info_lbl.configure(
            text=f"共 {len(self.rules)} 条规则，输出 {len(out_cols)} 列，{group_info}：{', '.join(out_cols[:6])}" + ("..." if len(out_cols) > 6 else ""))

    def _on_save_template(self):
        if not self.rules:
            messagebox.showwarning("提示", "当前没有可保存的规则。")
            return
        global pd
        if 'pd' not in globals():
            pd = lazy_import_pandas()

        default_name = "打标模板.xlsx"
        if self.rules_path:
            base = os.path.splitext(os.path.basename(self.rules_path))[0]
            default_name = f"{base}_模板.xlsx"
        path = filedialog.asksaveasfilename(
            title="保存规则模板", defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 模板", "*.xlsx"), ("所有", "*.*")])
        if not path:
            return
        try:
            max_gv = max((len(r.group_values) for r in self.rules), default=1)
            rows = []
            for r in self.rules:
                row = {
                    "规则ID": r.rule_id,
                    "标签名称": r.tag_name,
                    "匹配模式": r.mode,
                    "关键词": ",".join(r.keywords),
                    "否定关键词": ",".join(r.neg_keywords),
                    "忽略大小写": "是" if r.ignore_case else "否",
                    "优先级": r.priority,
                    "输出值": r.output_value,
                    "分类": r.category,
                    "分类优先级": "" if r.category_priority >= 9999 else r.category_priority,
                }
                for i in range(max_gv):
                    col_name = "分组值" if i == 0 else f"分组值{i+1}"
                    row[col_name] = r.group_values[i] if i < len(r.group_values) else ""
                rows.append(row)
            rules_df = pd.DataFrame(rows)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                rules_df.to_excel(writer, index=False, sheet_name="打标规则")
            messagebox.showinfo("保存成功",
                                f"模板已保存到：\n{path}\n\n"
                                f"共保存 {len(self.rules)} 条规则。")
            self._set_status(f"模板已保存：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _on_load_template(self):
        path = filedialog.askopenfilename(
            title="加载规则模板", filetypes=[("Excel 模板/规则表", "*.xlsx *.xls"), ("所有", "*.*")])
        if not path:
            return
        try:
            rules, info = load_rules(path)
        except Exception as e:
            messagebox.showerror("加载失败", f"无法加载模板：{e}")
            return
        self.rules = rules
        self.rules_path = path
        out_cols = []
        for r in rules:
            if r.tag_name not in out_cols:
                out_cols.append(r.tag_name)
        max_level = max((len(r.group_values) for r in rules), default=0)
        info = f"共加载 {len(rules)} 条规则，输出 {len(out_cols)} 列"
        if max_level > 0:
            info += f"，{max_level}级分组"
        self.rule_path_var.set(f"✓ [模板] {os.path.basename(path)}  （{info}）")
        self.rule_path_lbl.configure(fg=self.PRIMARY)
        self._refresh_rules_tree()
        self.notebook.select(0)
        self._update_mode()
        messagebox.showinfo("加载成功",
                            f"模板加载成功！\n\n"
                            f"共 {len(rules)} 条规则，输出 {len(out_cols)} 列。")

    def _reset_buttons_after_error(self):
        self._is_running = False
        self.start_btn.configure(state="normal")
        self.export_btn.configure(state="disabled")

    def _on_export(self):
        if self.cleaned_df is None and self.keyword_tag_df is None:
            messagebox.showwarning("提示", "请先点击【开始处理】。")
            return
        
        default_name = "数据清洗打标结果.xlsx"
        if self.xiyou_folder:
            default_name = f"{os.path.basename(self.xiyou_folder)}_清洗打标结果.xlsx"
        elif self.keywords_path:
            default_name = f"{os.path.splitext(os.path.basename(self.keywords_path))[0]}_打标结果.xlsx"
        
        path = filedialog.asksaveasfilename(
            title="导出结果", defaultextension=".xlsx",
            initialfile=default_name, filetypes=[("Excel", "*.xlsx"), ("所有", "*.*")])
        if not path:
            return
        
        try:
            self._set_status("正在导出...", self.PRIMARY)
            self.root.update_idletasks()
            save_result(path, self.cleaned_df, self.tag_df, self.keyword_tag_df)
            self._set_status(f"已导出：{os.path.basename(path)}", self.PRIMARY)
            messagebox.showinfo("导出成功", f"结果已保存到：\n{path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("导出失败", str(e))
            self._set_status("导出失败", "#FA5151")

    def _on_start(self):
        if self._is_running:
            return
        has_xiyou = self.xiyou_folder is not None and self.mapping_path is not None
        has_keywords = self.keywords_path is not None
        has_rules = len(self.rules) > 0

        if not has_rules:
            messagebox.showwarning("提示", "请先上传打标规则表或加载模板。")
            return
        if not has_xiyou and not has_keywords:
            messagebox.showwarning("提示", "请至少选择一种操作模式。")
            return

        self._is_running = True
        self.start_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled")

        if has_xiyou and has_keywords:
            self._set_status("正在处理：西柚数据清洗打标 + 关键词单独打标...", self.PRIMARY)
        elif has_xiyou:
            self._set_status("正在处理：西柚数据清洗打标...", self.PRIMARY)
        else:
            self._set_status("正在处理：关键词单独打标...", self.PRIMARY)

        def worker():
            try:
                cleaned_df = None
                tag_df = None
                keyword_tag_df = None

                if has_xiyou:
                    raw_df = load_xiyou_data(self.xiyou_folder)
                    mapping_df = load_mapping(self.mapping_path)
                    cleaned_df = clean_xiyou_data(raw_df, mapping_df)
                    tagged_df = apply_tags_for_xiyou(cleaned_df, '关键词 (数据来源于西柚找词)', self.rules)
                    tag_df = create_tag_sheet(tagged_df)
                    cleaned_df = tagged_df

                if has_keywords:
                    keywords_df = load_keywords_data(self.keywords_path)
                    keyword_tag_df = apply_tags_to_keywords(keywords_df, self.rules)

                self.cleaned_df = cleaned_df
                self.tag_df = tag_df
                self.keyword_tag_df = keyword_tag_df
                self.root.after(0, self._show_result_safe, cleaned_df, tag_df, keyword_tag_df)
            except Exception as e:
                err = traceback.format_exc()
                self.root.after(0, self._on_error, str(e), err)

        threading.Thread(target=worker, daemon=True).start()

    def _on_error(self, err_msg, err_tb):
        self._reset_buttons_after_error()
        messagebox.showerror("处理失败", f"{err_msg}\n\n{err_tb}")
        self._set_status("处理失败", "#FA5151")

    def _show_result_safe(self, cleaned_df, tag_df, keyword_tag_df):
        try:
            self._show_result(cleaned_df, tag_df, keyword_tag_df)
        except Exception as e:
            err = traceback.format_exc()
            self._reset_buttons_after_error()
            messagebox.showerror("结果显示失败", f"{e}\n\n{err}")
            self._set_status("结果显示失败", "#FA5151")
        finally:
            self._is_running = False
            self.start_btn.configure(state="normal")

    def _show_result(self, cleaned_df, tag_df, keyword_tag_df):
        for item in self.xiyou_tree.get_children():
            self.xiyou_tree.delete(item)
        for item in self.keyword_tree.get_children():
            self.keyword_tree.delete(item)

        has_xiyou = cleaned_df is not None
        has_keywords = keyword_tag_df is not None

        if has_xiyou:
            preview_df = cleaned_df.head(500)
            cols = list(preview_df.columns)
            self.xiyou_tree["columns"] = cols
            for col in cols:
                self.xiyou_tree.heading(col, text=col)
                self.xiyou_tree.column(col, width=100, minwidth=50, anchor="w", stretch=True)
            for idx, row in preview_df.iterrows():
                values = [str(row.get(c, "")) for c in cols]
                row_tag = "oddrow" if idx % 2 == 0 else "evenrow"
                self.xiyou_tree.insert("", "end", values=values, tags=(row_tag,))
            self.xiyou_info_lbl.configure(
                text=f"西柚数据源：{len(cleaned_df)}行 × {len(cleaned_df.columns)}列 | 打标：{len(tag_df)}行 × {len(tag_df.columns)}列")

        if has_keywords:
            preview_df = keyword_tag_df.head(500)
            cols = list(preview_df.columns)
            self.keyword_tree["columns"] = cols
            for col in cols:
                self.keyword_tree.heading(col, text=col)
                self.keyword_tree.column(col, width=100, minwidth=50, anchor="w", stretch=True)
            for idx, row in preview_df.iterrows():
                values = [str(row.get(c, "")) for c in cols]
                row_tag = "oddrow" if idx % 2 == 0 else "evenrow"
                self.keyword_tree.insert("", "end", values=values, tags=(row_tag,))
            self.keyword_info_lbl.configure(
                text=f"关键词打标：{len(keyword_tag_df)}行 × {len(keyword_tag_df.columns)}列")

        if has_xiyou or has_keywords:
            self.export_btn.configure(state="normal")
            mode_text = []
            if has_xiyou:
                mode_text.append("西柚数据清洗打标")
            if has_keywords:
                mode_text.append("关键词单独打标")
            self._set_status(f"处理完成：{', '.join(mode_text)}", self.PRIMARY)
        else:
            self._set_status("处理完成")


def main():
    root = tk.Tk()
    app = XiYouTaggerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()