#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_questions.py — 评课报告附录 J「教师课堂提问梳理与评价」校验脚本（class-observe-eval 配套）

用途
    对生成的评课报告 Markdown 做结构化校验，共 14 项：
      1  附录 J 区域存在          2  提问可梳理度等级（J-0）
      3  J-1 序号从 1 连续        4  J-1 列数与表头一致
      5  J-2 逐问覆盖 J-1        6  J-1/J-2 单元格字数（≤60，61–80 警告，>80 失败）
      7  原问溯源（需 --source）  8  禁用词（分数/等第排名/否定标签/理想比例）
      9  建议动作取值             10 认知层级取值
      11 非认知问句合并（甲级）   12 深评卡数量 2–4 张（甲级）
      13 J-4 整体画像存在         14 内部标记泄露（附录 J 区域内）
    区域界定与表格定位均容错：附录 J 起始标题支持「### J. …」「### J …」「## 附录 J …」等变体；
    J-1/J-2 表格靠表头关键字定位（J-1 需含「序号」「教师原问」，J-2 需含「序号」「评价」「优化建议」），
    不依赖子节标题文字；找不到 J 区域起始标题时退回全文按表头查找。丙级/丁级为降级输出，3–8 项标 SKIP。

用法
    python3 check_questions.py <报告.md> [--source <课堂输入原文.md>] [--whole] [--json] [--selftest]

    --source    提供听课记录/转写原文时，执行「原问溯源」检查
    --whole     禁用词检查扩大到全文（默认只查附录 J 区域）
    --json      输出 JSON 明细：{file, summary{fail,warn}, checks[{id,name,status,message,details[]}]}
    --selftest  运行内置自测（不要求位置参数）

退出码
    0 = 无 FAIL；1 = 有 FAIL（自测失败同为 1）；2 = 用法错误 / 文件不存在 / 无法解析

仅依赖 Python 标准库，兼容 Python 3.8+。
"""

import argparse
import difflib
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from collections import Counter

SCRIPT_VERSION = "1.0"

# ================================================================ 常量与正则

# 每个检查项最多在人类可读输出中展示的明细条数
MAX_DETAILS_SHOWN = 12

# --- 标题 / 表格识别 ---------------------------------------------------------
HEADING_RE = re.compile(r'^(#{1,6})\s*(.*?)\s*#*\s*$')
SEP_RE = re.compile(r'^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$')
FENCE_RE = re.compile(r'^\s*(```|~~~)')

# 附录 J 的子节标题（J-0 … J-9），如「#### J-1 提问梳理表」「### J1」「附录 J-2」
J_SUB_RE = re.compile(r'^(?:附录\s*)?J\s*[-–—_．.]*\s*([0-9０-９])')
# 附录 J 区域起始标题：「J」后不能紧跟字母数字（排除 J-1），允许「附录 J」「J.」「J：」等
J_START_RE = re.compile(r'^(?:附录\s*)?J(?![0-9A-Za-z])')
J_START_BARE_RE = re.compile(r'^(?:附录\s*)?J\s*[\.．、:：]?\s*(?:[（(][^）)]*[）)])?\s*$')

# 表头关键字（表格靠表头定位，不依赖标题文字）
J1_HEADER_KEYS = ('序号', '原问')            # 「教师原问」或「原问」
J2_HEADER_KEYS = ('序号', '评价', '建议')    # 「优化建议」含「建议」

# --- 可梳理度等级 -------------------------------------------------------------
LEVEL_INFO = {
    '甲': ('甲级 逐句可还原', False),
    '乙': ('乙级 选择性记录', False),
    '丙': ('丙级 仅概括描述', True),
    '丁': ('丁级 无提问信息', True),
}
# (等级, 正则, 层级)：0 = 显式等级词；1 = 别称；2 = 弱线索
LEVEL_TERMS = [
    ('甲', re.compile(r'甲\s*级'), 0),
    ('乙', re.compile(r'乙\s*级'), 0),
    ('丙', re.compile(r'丙\s*级'), 0),
    ('丁', re.compile(r'丁\s*级'), 0),
    ('甲', re.compile(r'逐句可还原'), 1),
    ('乙', re.compile(r'选择性记录'), 1),
    ('丙', re.compile(r'仅概括描述'), 1),
    ('丁', re.compile(r'无提问信息'), 1),
    ('丁', re.compile(r'未做提问梳理'), 2),
    ('丙', re.compile(r'教师提问补充清单'), 2),
]
# 等级词前若紧邻这些「断言词」，可信度更高；若紧邻否定词则降低
ASSERT_RE = re.compile(r'(?:判定|定为|评为|等级|为|属|是|：|:|（|\()[\s*_]*$')
NEGATE_RE = re.compile(r'(?:不|非|未达|不到|无法|不是|仅|只|难|尚|够不上)[\u4e00-\u9fff]{0,2}$')

# --- 取值白名单 ---------------------------------------------------------------
ALLOWED_ACTIONS = ('保留', '加追问', '改写问句', '调整对象或留白', '合并或删减')
ALLOWED_COG_LEVELS = ('事实复述', '理解解释', '分析评价', '迁移创造', '组织管理', '确认求同', '征询意愿')
NON_COGNITIVE = ('组织管理', '确认求同', '征询意愿')

# --- 深评卡 / J-4 --------------------------------------------------------------
# 「深评卡」与「第 X 问」之间允许空格、加粗符号、井号、引用符、间隔号、冒号、括号等变体
CARD_RE = re.compile(r'深评卡[\s\*#>·•・:：．.、,，\-—–~～()（）]*第?\s*(\d+)\s*问')
J4_RE = re.compile(r'整体画像|结构描述|问题链')

# --- 禁用词 -------------------------------------------------------------------
BANNED_PATTERNS = [
    # 分数：数字 + 可选小数 + 「分」，排除「5 分钟」「3 组」「分层/分工/分析/分布/级别/段落/分歧」等
    ('分数', re.compile(r'(?<![\d.])\d+(?:\.\d+)?\s*分(?!\s*[钟组工析布层别段歧秒批之])')),
    ('得分/评分/打分', re.compile(r'得分|评分|打分')),
    ('等第与排名', re.compile(r'达标率|优秀率|合格率|及格率|理想比例|排名|第几名|优于平均|不合格|等第')),
    ('否定标签', re.compile(r'无效提问|低质|劣质|差评|糟糕')),
    # 理想比例：仅查「3-2-1 比例」及其连接符/后缀变体
    ('理想比例', re.compile(r'3\s*[-—–‑－~～:：至到]\s*2\s*[-—–‑－~～:：至到]\s*1\s*(?:的)?\s*(?:比例|原则|结构|分布|配比|占比)')),
]
# 否定语境（如 J-0 说明里的「不打分、不设等第、不做排名」）不算命中
NEG_WORDS_1 = ('不', '未', '无', '非', '零', '没', '没有', '避免', '禁止', '不得', '严禁', '拒绝', '反对', '勿')
NEG_VERBS_2 = ('据此', '做', '设', '进行', '给', '列', '评', '宜', '应', '用', '以', '据', '出现', '使用', '提倡', '主张')

# --- 内部标记泄露 -------------------------------------------------------------
LEAK_PATTERNS = [
    ('.md 文件名', re.compile(r'\.md(?![A-Za-z0-9])')),
    ('references/ 路径', re.compile(r'references/')),
    ('§ 符号', re.compile(r'§')),
    ('反引号', re.compile(r'`')),
    ('W-数字 编号', re.compile(r'(?<![A-Za-z])W-\d+')),
    ('C-数字 编号', re.compile(r'(?<![A-Za-z])C-\d+')),
    ('表 数字- 表号', re.compile(r'表\s*\d+\s*[-—–‑－]')),
    ('teacher_questioning', re.compile(r'teacher_questioning')),
    ('摘录卡', re.compile(r'摘录卡')),
]

# --- 溯源 ----------------------------------------------------------------------
FRAG_SPLIT_RE = re.compile(r'[/／]|…+|\.{2,}|[？?]|\n')
SEQ_TOKEN_SPLIT_RE = re.compile(r'(?:[、,，;；/／|＆&\s]+|and|or|和|及|与|兼)', re.I)
SEQ_RANGE_RE = re.compile(r'^(\d+)\s*[-–—~～－至到]\s*(\d+)$')


# ================================================================ 通用文本工具

def read_text(path):
    """读取文本文件：优先 UTF-8（兼容 BOM），失败时依次尝试 GB18030 / Latin-1。"""
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in ('utf-8-sig', 'gb18030'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('latin-1')


BR_RE = re.compile(r'<br\s*/?>', re.I)
MD_LINK_RE = re.compile(r'!?\[([^\]]*)\]\([^)]*\)')
HTML_TAG_RE = re.compile(r'<[^>]+>')
MD_EMPH_RE = re.compile(r'(\*\*|__|~~|\*|_)(.+?)\1')


def strip_md(s):
    """去掉单元格里的 Markdown 标记（加粗/斜体/删除线/行内代码/链接/HTML 标签），<br> 视作空格。"""
    s = str(s)
    s = BR_RE.sub(' ', s)
    s = MD_LINK_RE.sub(r'\1', s)
    s = HTML_TAG_RE.sub('', s)
    s = re.sub(r'`([^`]*)`', r'\1', s)
    s = s.replace('`', '')
    s = MD_EMPH_RE.sub(r'\2', s)
    return s


def count_chars(s):
    """去掉 Markdown 标记与空白后的字符数（中文一字计 1）。"""
    return len(re.sub(r'\s+', '', strip_md(s)))


def norm_text(s):
    """溯源归一化：NFKC（全角数字/字母转半角）→ 去括注内容 → 去空白、标点、符号，只留字母/数字/汉字。"""
    s = unicodedata.normalize('NFKC', str(s))
    s = re.sub(r'\([^()]*\)', '', s)
    s = re.sub(r'\[[^\[\]]*\]', '', s)
    s = re.sub(r'【[^【】]*】', '', s)
    out = []
    for ch in s:
        if ch.isspace():
            continue
        cat = unicodedata.category(ch)
        if cat[0] in ('P', 'S', 'Z', 'C'):   # 标点、符号、分隔符、控制/格式字符
            continue
        out.append(ch)
    return ''.join(out)


def snippet(text, start, end, width=30):
    """截取命中位置前后各 width 字的片段。"""
    a, b = max(0, start - width), min(len(text), end + width)
    return ('…' if a > 0 else '') + text[a:b] + ('…' if b < len(text) else '')


def is_negated(line, start):
    """命中词之前 4 字内是否有否定词（不/未/无/非/零/没有/避免/禁止/不得/严禁/拒绝/反对，或「不设/不做/不据此…」）。"""
    pre = line[max(0, start - 4):start]
    if pre.endswith(NEG_WORDS_1):
        return True
    for w in ('不', '未', '无', '非', '零'):
        for v in NEG_VERBS_2:
            if pre.endswith(w + v):
                return True
    return False


# ================================================================ 结构解析

def fence_mask(lines):
    """标记围栏代码块（``` 或 ~~~）内的行，标题与表格解析时跳过。"""
    mask = [False] * len(lines)
    in_fence = False
    for i, line in enumerate(lines):
        if FENCE_RE.match(line):
            mask[i] = True
            in_fence = not in_fence
            continue
        mask[i] = in_fence
    return mask


def is_j_start(text):
    """判断标题文字是否为附录 J 区域起始标题（子节 J-0…J-9 不算）。"""
    t = text.strip()
    if J_SUB_RE.match(t):
        return False
    if J_START_RE.match(t):
        return ('提问' in t) or J_START_BARE_RE.match(t) is not None
    return '教师课堂提问梳理与评价' in t


def find_region(lines, headings):
    """定位附录 J 区域：返回 (起始行索引, 结束行索引(不含), 是否找到, 起始标题文字)。
    结束 = 下一个同级或更高级标题（跳过 J 子节标题），或文件末尾。"""
    for k, (idx, level, text) in enumerate(headings):
        if is_j_start(text):
            end = len(lines)
            for jdx, lv2, t2 in headings[k + 1:]:
                if lv2 <= level and not J_SUB_RE.match(t2):
                    end = jdx
                    break
            return idx, end, True, text
    return 0, len(lines), False, ''


def split_row(line):
    """把 Markdown 表格行拆成单元格列表（支持转义的 \\|）。"""
    s = line.strip().replace('\\|', '\x00')
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|'):
        s = s[:-1]
    return [c.replace('\x00', '|').strip() for c in s.split('|')]


def classify_table(header):
    """按表头关键字判断表格类型：'J-1' / 'J-2' / None。"""
    hs = [strip_md(h).replace(' ', '') for h in header]

    def has(key):
        return any(key in h for h in hs)

    if has('序号') and has('原问'):
        return 'J-1'
    # J-2 标准表头含「评价」「优化建议」； tolerate 「可取/可提升」拆列写法
    if has('序号') and has('建议') and (has('评价') or has('可取') or has('可提升')):
        return 'J-2'
    return None


def parse_tables(lines, fence, start, end):
    """解析区域内的所有 Markdown 表格（表头行 + 分隔行 + 数据行）。"""
    tables = []
    i = start
    while i < end - 1:
        if fence[i] or not lines[i].lstrip().startswith('|'):
            i += 1
            continue
        if fence[i + 1] or '|' not in lines[i + 1] or not SEP_RE.match(lines[i + 1]):
            i += 1
            continue
        header = split_row(lines[i])
        rows = []
        j = i + 2
        while j < end and not fence[j] and lines[j].lstrip().startswith('|'):
            rows.append((j + 1, split_row(lines[j])))   # 行号从 1 计
            j += 1
        tables.append({'header_no': i + 1, 'header': header, 'rows': rows, 'kind': classify_table(header)})
        i = j
    return tables


def find_col(header, *keys):
    """按关键字（依次）定位列索引，找不到返回 -1。"""
    hs = [strip_md(h) for h in header]
    for key in keys:
        for i, h in enumerate(hs):
            if key in h:
                return i
    return -1


def table_label(kind, table):
    return '%s（表头第 %d 行）' % (kind, table['header_no'])


def parse_seq_cell(cell):
    """解析序号单元格 → 整数列表。兼容 顿号/逗号/空格/斜杠/and/or/和/及/与 等分隔符，以及 8-14 区间。"""
    s = unicodedata.normalize('NFKC', strip_md(cell))
    s = re.sub(r'[（(][^（）()]*[）)]', '', s)   # 去掉括注，避免括注里的数字被当作序号
    nums = []
    for tok in SEQ_TOKEN_SPLIT_RE.split(s):
        tok = tok.strip()
        if not tok:
            continue
        m = SEQ_RANGE_RE.match(tok)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b and b - a <= 300:
                nums.extend(range(a, b + 1))
            else:
                nums.extend([a, b])
            continue
        nums.extend(int(x) for x in re.findall(r'\d+', tok))
    return nums


# ================================================================ 可梳理度识别

def detect_level(lines):
    """从 J-0 说明文字中识别可梳理度等级。返回 (等级或 None, 证据片段, 说明)。
    打分规则：显式等级词 > 别称 > 弱线索；前接「定为/判定/等级/：」加分，前接否定词减分；
    多个等级同分则视为无法确定。"""
    best = {}   # 等级 -> (得分, 出现顺序, 所在行)
    order = 0
    for text in lines:
        for lvl, rx, tier in LEVEL_TERMS:
            for m in rx.finditer(text):
                pre = text[max(0, m.start() - 4):m.start()]
                score = (2 - tier) * 10
                if ASSERT_RE.search(pre):
                    score += 1
                if NEGATE_RE.search(pre):
                    score -= 3
                cur = best.get(lvl)
                if cur is None or score > cur[0]:
                    best[lvl] = (score, order, strip_md(text).strip())
                order += 1
    if not best:
        return None, '', ''
    top = max(v[0] for v in best.values())
    winners = [l for l, v in best.items() if v[0] == top]
    if len(winners) > 1:
        return None, '', '多个等级得分相同，无法确定：' + '、'.join(LEVEL_INFO[w][0] for w in winners)
    lvl = winners[0]
    return lvl, best[lvl][2], ''


# ================================================================ 原问溯源

def split_fragments(cell):
    """把「教师原问」单元格切成问句片段并归一化；按 ／ / …… ... ？ ? 换行切分，丢弃 <4 字的片段。"""
    text = re.sub(r'<br\s*/?>', '\n', str(cell), flags=re.I)
    text = strip_md(text)
    result = []
    for part in FRAG_SPLIT_RE.split(text):
        n = norm_text(part)
        n = re.sub(r'^(?:教师|老师|说话人\d*|师|T\d?)', '', n)   # 去掉可能残留的说话人标签
        if len(n) >= 4:
            result.append(n)
    return result


def trace_fragments(items, source_text):
    """原问溯源。items: [(序号, 归一化片段)]。
    先在归一化原文中做子串查找；未命中者按句子边界（。！？；换行）切句，
    用 difflib.SequenceMatcher 计算「片段被该句按顺序覆盖的比例」，≥ 0.80 视为模糊命中。"""
    src_norm = norm_text(source_text)
    sentences = []
    for raw in re.split(r'(?<=[。！？；!?;])|\n', source_text):
        n = norm_text(raw)
        if len(n) >= 4:
            sentences.append((raw.strip(), n))
    total, exact, fuzzy = len(items), 0, 0
    misses, fuzzy_hits = [], []
    for seq, frag in items:
        if frag in src_norm:
            exact += 1
            continue
        sm = difflib.SequenceMatcher(None, '', frag, autojunk=False)   # 片段固定为 seq2 以复用缓存
        best_score, best_raw = 0.0, ''
        for raw, n in sentences:
            sm.set_seq1(n)
            matched = sum(b.size for b in sm.get_matching_blocks())
            score = matched / float(len(frag))
            if score > best_score:
                best_score, best_raw = score, raw
        if best_score >= 0.80:
            fuzzy += 1
            fuzzy_hits.append((seq, frag, best_raw, best_score))
        else:
            misses.append((seq, frag, best_score, best_raw))
    return {'total': total, 'exact': exact, 'fuzzy': fuzzy, 'misses': misses, 'fuzzy_hits': fuzzy_hits}


# ================================================================ 主检查流程

def analyze(lines, source_text=None, whole=False):
    """对报告行列表执行全部检查，返回结果字典。"""
    checks = []

    def add(cid, name, status, message, details=None):
        checks.append({'id': cid, 'name': name, 'status': status,
                       'message': message, 'details': list(details or [])})

    fence = fence_mask(lines)
    headings = []
    for i, line in enumerate(lines):
        if fence[i]:
            continue
        m = HEADING_RE.match(line)
        if m and m.group(2).strip():
            headings.append((i, len(m.group(1)), m.group(2).strip()))

    start, end, found, title = find_region(lines, headings)
    tables = parse_tables(lines, fence, start, end)
    j1_tables = [t for t in tables if t['kind'] == 'J-1']
    j2_tables = [t for t in tables if t['kind'] == 'J-2']
    fallback = (not found) and bool(j1_tables or j2_tables)

    # 子节标题（J-0…J-9）位置
    sub_heads = []
    for idx in range(start + 1, end):
        if fence[idx]:
            continue
        m = HEADING_RE.match(lines[idx])
        if m and J_SUB_RE.match(m.group(2).strip()):
            sub_heads.append((idx, m.group(2).strip()))

    # ---------- 1 区域存在 ----------
    if found:
        add(1, '区域存在', 'PASS', '第 %d–%d 行：%s' % (start + 1, end, title[:40]))
    elif fallback:
        add(1, '区域存在', 'FAIL',
            '未找到附录 J 起始标题；已在全文按表头定位到 J-1/J-2 表，后续检查按全文范围执行',
            ['提示：起始标题应形如「### J. 教师课堂提问梳理与评价」「## 附录 J …」'])
    else:
        add(1, '区域存在', 'FAIL', '未找到附录 J 区域起始标题，也未在全文找到 J-1/J-2 表',
            ['提示：起始标题应形如「### J. 教师课堂提问梳理与评价」「## 附录 J …」'])

    # ---------- 2 可梳理度（J-0 说明 = 区域内首个表格/首个 J-1+ 子节标题之前） ----------
    limit = end
    if tables:
        limit = min(limit, tables[0]['header_no'] - 1)
    for idx, text in sub_heads:
        m = J_SUB_RE.match(text)
        if m and m.group(1) != '0':
            limit = min(limit, idx)
    j0_lines = [lines[i] for i in range(start, max(start, limit))]
    level, evidence, note = detect_level(j0_lines) if any(l.strip() for l in j0_lines) else (None, '', '')
    if level is None and not note:
        level, evidence, note = detect_level(lines[start:end])
    degraded = LEVEL_INFO.get(level, ('', False))[1]
    if level:
        if degraded:
            msg = '%s：降级输出，无逐问表属正常' % LEVEL_INFO[level][0]
        else:
            msg = '%s：%s' % (LEVEL_INFO[level][0], evidence.strip()[:40])
        add(2, '可梳理度', 'PASS', msg, ['依据：%s' % evidence.strip()[:60]] if evidence.strip() else [])
    elif note:
        add(2, '可梳理度', 'WARN', '无法确定可梳理度等级（%s）' % note)
    elif found:
        add(2, '可梳理度', 'WARN', 'J-0 说明中未识别出可梳理度等级（甲级/乙级/丙级/丁级，或逐句可还原/选择性记录/仅概括描述/无提问信息）')
    else:
        add(2, '可梳理度', 'WARN', '未找到附录 J 区域，且全文未识别出可梳理度等级')

    # ---------- 3 J-1 存在且序号连续 ----------
    j1_rows = []   # (表, 行号, 单元格, 序号列表)
    if degraded:
        add(3, 'J-1 序号连续', 'SKIP', '降级输出（%s），不要求 J-1 逐问表' % LEVEL_INFO[level][0])
    elif not j1_tables:
        add(3, 'J-1 序号连续', 'FAIL', '未找到 J-1 提问梳理表（表头需同时含「序号」「教师原问」）')
    else:
        seqs, bad_rows = [], []
        for t in j1_tables:
            col = find_col(t['header'], '序号')
            if col < 0:
                bad_rows.append('%s 表头无「序号」列' % table_label('J-1', t))
                continue
            for line_no, cells in t['rows']:
                nums = parse_seq_cell(cells[col]) if col < len(cells) else []
                if not nums:
                    bad_rows.append('第 %d 行序号无法解析：%r' % (line_no, cells[col] if col < len(cells) else cells))
                j1_rows.append((t, line_no, cells, nums))
                seqs.extend(nums)
        missing = sorted(set(range(1, max(seqs) + 1)) - set(seqs)) if seqs else []
        dup = sorted(x for x, c in Counter(seqs).items() if c > 1)
        if bad_rows or missing or dup:
            parts = []
            if missing:
                parts.append('缺少序号 ' + ', '.join(str(x) for x in missing[:30]) + ('…' if len(missing) > 30 else ''))
            if dup:
                parts.append('重复序号 ' + ', '.join(str(x) for x in dup[:30]))
            if bad_rows:
                parts.append('%d 行序号无法解析' % len(bad_rows))
            details = bad_rows[:]
            for x in missing[:30]:
                details.append('序号 %d 未出现在 J-1 中' % x)
            for x in dup[:30]:
                details.append('序号 %d 在 J-1 中重复' % x)
            add(3, 'J-1 序号连续', 'FAIL', '；'.join(parts) + '（共读出 %d 个序号，%d 张表）' % (len(seqs), len(j1_tables)), details)
        else:
            add(3, 'J-1 序号连续', 'PASS', '共 %d 问（%d 张表，序号 1–%d 连续）' % (len(seqs), len(j1_tables), len(seqs)))

    # ---------- 4 J-1 列数一致 ----------
    if degraded:
        add(4, 'J-1 列数一致', 'SKIP', '降级输出，不要求 J-1 逐问表')
    elif not j1_tables:
        add(4, 'J-1 列数一致', 'SKIP', '无 J-1 表')
    else:
        bad, nrows = [], 0
        for t in j1_tables:
            k = len(t['header'])
            for line_no, cells in t['rows']:
                nrows += 1
                if len(cells) != k:
                    bad.append('%s 第 %d 行有 %d 列（表头 %d 列）：%s'
                               % (table_label('J-1', t), line_no, len(cells), k, lines[line_no - 1].strip()[:50]))
        if bad:
            add(4, 'J-1 列数一致', 'FAIL', '%d 行列数与表头不一致（共 %d 行数据）' % (len(bad), nrows), bad)
        else:
            add(4, 'J-1 列数一致', 'PASS', '共 %d 行数据，列数均与表头一致' % nrows)

    # ---------- 5 逐问覆盖 ----------
    j2_rows = []   # (表, 行号, 单元格, 序号列表)
    j2_seq_col = -1
    if degraded:
        add(5, '逐问覆盖', 'SKIP', '降级输出，不要求 J-2 逐问表')
    elif not j1_tables:
        add(5, '逐问覆盖', 'SKIP', '无 J-1 表，无法核对覆盖（见检查 3）')
    elif not j2_tables:
        add(5, '逐问覆盖', 'FAIL', '未找到 J-2 逐问评价表（表头需同时含「序号」「评价」「优化建议」）')
    else:
        for t in j2_tables:
            col = find_col(t['header'], '序号')
            if col < 0:
                continue
            j2_seq_col = col
            for line_no, cells in t['rows']:
                nums = parse_seq_cell(cells[col]) if col < len(cells) else []
                j2_rows.append((t, line_no, cells, nums))
        j1_set = set()
        for _, _, _, nums in j1_rows:
            j1_set.update(nums)
        j2_list = [x for _, _, _, nums in j2_rows for x in nums]
        j2_set = set(j2_list)
        missing = sorted(j1_set - j2_set)
        extra = sorted(j2_set - j1_set)
        dup = sorted(x for x, c in Counter(j2_list).items() if c > 1)
        unparsed = ['第 %d 行序号无法解析：%r' % (ln, cells[j2_seq_col] if j2_seq_col < len(cells) else cells)
                    for _, ln, cells, nums in j2_rows if not nums]
        if missing or extra or dup or unparsed:
            parts = []
            if missing:
                parts.append('缺少序号 ' + ', '.join(str(x) for x in missing[:30]) + ('…' if len(missing) > 30 else ''))
            if extra:
                parts.append('多余序号 ' + ', '.join(str(x) for x in extra[:30]))
            if dup:
                parts.append('重复序号 ' + ', '.join(str(x) for x in dup[:30]))
            if unparsed:
                parts.append('%d 行序号无法解析' % len(unparsed))
            details = unparsed[:]
            for x in missing[:30]:
                details.append('第 %d 问未在 J-2 中评价' % x)
            for x in extra[:30]:
                details.append('J-2 序号 %d 在 J-1 中不存在' % x)
            for x in dup[:30]:
                details.append('序号 %d 在 J-2 中出现 %d 次' % (x, Counter(j2_list)[x]))
            add(5, '逐问覆盖', 'FAIL', '；'.join(parts) + '（J-2 共 %d 行）' % len(j2_rows), details)
        else:
            add(5, '逐问覆盖', 'PASS', 'J-2 共 %d 行，覆盖 J-1 全部 %d 个序号，无多余、无重复' % (len(j2_rows), len(j1_set)))

    # ---------- 6 单元格字数 ----------
    if degraded:
        add(6, '单元格字数', 'SKIP', '降级输出，无逐问表')
    elif not j1_tables and not j2_tables:
        add(6, '单元格字数', 'SKIP', '无 J-1/J-2 表')
    else:
        over, warn = [], []
        ncells = 0
        for kind, tlist in (('J-1', j1_tables), ('J-2', j2_tables)):
            for t in tlist:
                col = find_col(t['header'], '序号')
                for line_no, cells in t['rows']:
                    seq_txt = cells[col] if 0 <= col < len(cells) else '?'
                    for ci, cell in enumerate(cells):
                        if ci == col:
                            continue
                        n = count_chars(cell)
                        ncells += 1
                        if n > 60:
                            colname = strip_md(t['header'][ci]) if ci < len(t['header']) else '第 %d 列' % (ci + 1)
                            item = '%s 第 %s 行「%s」%d 字：%s' % (table_label(kind, t), seq_txt or line_no, colname, n, strip_md(cell)[:30] + '…')
                            (over if n > 80 else warn).append(item)
        if over or warn:
            parts = []
            if over:
                parts.append('%d 格超过 80 字' % len(over))
            if warn:
                parts.append('%d 格 61–80 字' % len(warn))
            add(6, '单元格字数', 'FAIL' if over else 'WARN', '；'.join(parts) + '（共检查 %d 格，上限 60 字）' % ncells, over + warn)
        else:
            add(6, '单元格字数', 'PASS', 'J-1/J-2 共 %d 格（除序号列），均不超过 60 字' % ncells)

    # ---------- 7 原问溯源 ----------
    if degraded:
        add(7, '原问溯源', 'SKIP', '降级输出，无逐问表')
    elif source_text is None:
        add(7, '原问溯源', 'SKIP', '未提供 --source，跳过溯源')
    elif not j1_rows:
        add(7, '原问溯源', 'SKIP', '无 J-1 表，无法溯源')
    else:
        items = []
        for t, line_no, cells, nums in j1_rows:
            qcol = find_col(t['header'], '教师原问', '原问')
            cell = cells[qcol] if 0 <= qcol < len(cells) else ''
            seq_label = nums[0] if nums else line_no
            for frag in split_fragments(cell):
                items.append((seq_label, frag))
        if not items:
            add(7, '原问溯源', 'WARN', 'J-1「教师原问」列无可溯源片段（归一化后均不足 4 字）')
        else:
            r = trace_fragments(items, source_text)
            rate = (r['exact'] + r['fuzzy']) / float(r['total'])
            pct = '%.1f%%' % (rate * 100)
            details = ['第 %s 问 未命中：「%s」%s' % (s, f, ('（最接近原句「%s」，相似度 %.2f）' % (raw[:40], sc)) if raw else '')
                       for s, f, sc, raw in r['misses']]
            details += ['第 %s 问 模糊命中（相似度 %.2f）：「%s」 ≈ 原句「%s」' % (s, sc, f, raw[:40])
                        for s, f, raw, sc in r['fuzzy_hits']]
            base = '片段 %d，精确命中 %d，模糊命中 %d，未命中 %d，命中率 %s' % (r['total'], r['exact'], r['fuzzy'], len(r['misses']), pct)
            if rate < 0.70:
                add(7, '原问溯源', 'FAIL', base + '（低于 70%）', details)
            elif rate < 1.0:
                add(7, '原问溯源', 'WARN', base + '；未命中片段可能是转写错字还原或删口语赘词所致，需人工复核', details)
            else:
                add(7, '原问溯源', 'PASS', base, details)

    # ---------- 8 禁用词 ----------
    if whole:
        scope, scope_desc = range(0, len(lines)), '全文'
    elif found:
        scope, scope_desc = range(start, end), '附录 J 区域 第 %d–%d 行' % (start + 1, end)
    else:
        scope, scope_desc = range(0, len(lines)), '全文（未找到附录 J 区域，已回退）'
    hits = []
    for i in scope:
        line = lines[i]
        if not line.strip():
            continue
        seen = set()
        for label, rx in BANNED_PATTERNS:
            for m in rx.finditer(line):
                key = (label, m.group(0))
                if key in seen or is_negated(line, m.start()):
                    continue
                seen.add(key)
                hits.append('第 %d 行「%s」（%s）：%s' % (i + 1, m.group(0), label, snippet(line, m.start(), m.end(), 30)))
    if hits:
        add(8, '禁用词', 'FAIL', '命中 %d 处（范围：%s）' % (len(hits), scope_desc), hits)
    elif degraded:
        add(8, '禁用词', 'SKIP', '降级输出，无逐问表（附录 J 区域内未命中禁用词）')
    else:
        add(8, '禁用词', 'PASS', '未命中（范围：%s）' % scope_desc)

    # ---------- 9 建议动作取值 ----------
    if degraded:
        add(9, '建议动作取值', 'SKIP', '降级输出，不要求 J-2 逐问表')
    elif not j2_tables:
        add(9, '建议动作取值', 'SKIP', '无 J-2 表')
    else:
        bad, n = [], 0
        for t in j2_tables:
            acol = find_col(t['header'], '建议动作', '动作')
            scol = find_col(t['header'], '序号')
            if acol < 0:
                bad.append('%s 表头无「建议动作」列' % table_label('J-2', t))
                continue
            for line_no, cells in t['rows']:
                if acol >= len(cells):
                    continue
                n += 1
                val = strip_md(cells[acol])
                if not any(a in val for a in ALLOWED_ACTIONS):
                    seq_txt = cells[scol] if 0 <= scol < len(cells) else line_no
                    bad.append('%s 第 %s 行「建议动作」=「%s」，不在五种取值之内' % (table_label('J-2', t), seq_txt, val[:30]))
        if bad:
            add(9, '建议动作取值', 'WARN', '%d 行取值不在五种之内（保留/加追问/改写问句/调整对象或留白/合并或删减）' % len(bad), bad)
        else:
            add(9, '建议动作取值', 'PASS', 'J-2 共 %d 行，均在五种取值之内' % n)

    # ---------- 10 认知层级取值 ----------
    if degraded:
        add(10, '认知层级取值', 'SKIP', '降级输出，不要求 J-1 逐问表')
    elif not j1_tables:
        add(10, '认知层级取值', 'SKIP', '无 J-1 表')
    else:
        bad, n = [], 0
        for t in j1_tables:
            lcol = find_col(t['header'], '认知层级', '认知水平', '层级')
            if lcol < 0:
                bad.append('%s 表头无「认知层级」列' % table_label('J-1', t))
                continue
            for line_no, cells in t['rows']:
                if lcol >= len(cells):
                    continue
                n += 1
                val = strip_md(cells[lcol])
                if not any(x in val for x in ALLOWED_COG_LEVELS):
                    bad.append('%s 第 %d 行「认知层级」=「%s」，不在七种取值之内' % (table_label('J-1', t), line_no, val[:30] or '（空）'))
        if bad:
            add(10, '认知层级取值', 'WARN', '%d 行取值不在七种之内（事实复述/理解解释/分析评价/迁移创造/组织管理/确认求同/征询意愿）' % len(bad), bad)
        else:
            add(10, '认知层级取值', 'PASS', 'J-1 共 %d 行，均在七种取值之内' % n)

    # ---------- 11 非认知问句合并（仅甲级） ----------
    if level != '甲':
        reason = '降级输出，不检查' if degraded else ('仅甲级检查' if level else '未识别可梳理度等级，不检查')
        add(11, '非认知问句合并', 'SKIP', reason)
    elif not j1_rows:
        add(11, '非认知问句合并', 'SKIP', '无 J-1 表')
    else:
        warns = []
        non_cog = Counter()
        singles = {}   # (类别, 环节) -> [单独成行的序号…]
        for t, line_no, cells, nums in j1_rows:
            lcol = find_col(t['header'], '认知层级', '认知水平', '层级')
            scol = find_col(t['header'], '环节')
            val = strip_md(cells[lcol]) if 0 <= lcol < len(cells) else ''
            stage = strip_md(cells[scol]) if 0 <= scol < len(cells) else ''
            kinds = [k for k in NON_COGNITIVE if k in val]
            if kinds and nums:
                kind = kinds[0]
                non_cog[kind] += 1
                seq = nums[0]
                row = next((r for r in j2_rows if seq in r[3]), None)
                if row is not None and len(row[3]) == 1:
                    singles.setdefault((kind, stage), []).append(seq)
        # 同一类在同一环节只有一问时，单独成行属正常；两问及以上仍各自单行，才提示合并
        for (kind, stage), seqs in sorted(singles.items(), key=lambda kv: kv[1][0]):
            if len(seqs) >= 2:
                warns.append('%s（%s）有 %d 问各自单独成行（第 %s 问），建议合并为一行'
                             % (kind, stage or '环节未标', len(seqs), '、'.join(str(s) for s in seqs)))
        for t, line_no, cells, nums in j2_rows:
            if len(nums) > 10:
                warns.append('%s 第 %d 行合并了 %d 个序号，超过 10 个，建议拆行' % (table_label('J-2', t), line_no, len(nums)))
        desc = '、'.join('%s %d 问' % (k, v) for k, v in non_cog.items()) if non_cog else 'J-1 中无非认知问句'
        if warns:
            add(11, '非认知问句合并', 'WARN', '%s；%d 处需关注' % (desc, len(warns)), warns)
        else:
            add(11, '非认知问句合并', 'PASS', '%s；没有同类别同环节可合并而未合并的行，合并行均不超过 10 个序号' % desc)

    # ---------- 12 深评卡数量（甲级期望 2–4 张） ----------
    if level != '甲':
        if level == '乙':
            reason = '乙级不出深评卡，不检查'
        elif degraded:
            reason = '降级输出，不检查'
        else:
            reason = '未识别可梳理度等级，不检查'
        add(12, '深评卡数量', 'SKIP', reason)
    else:
        card_nums = []
        for i in range(start, end):
            m = CARD_RE.search(lines[i])
            if m:
                num = int(unicodedata.normalize('NFKC', m.group(1)))
                if num not in card_nums:
                    card_nums.append(num)
        n = len(card_nums)
        msg = '%d 张（第 %s 问；期望 2–4 张）' % (n, '、'.join(str(x) for x in card_nums) if card_nums else '无')
        if n < 2 or n > 4:
            add(12, '深评卡数量', 'WARN', msg)
        else:
            add(12, '深评卡数量', 'PASS', msg)

    # ---------- 13 J-4 整体画像存在 ----------
    if degraded:
        add(13, 'J-4 存在', 'SKIP', '降级输出，不要求 J-4 整体画像')
    else:
        boundary = start
        if j2_rows:
            boundary = max(r[1] - 1 for r in j2_rows)
        elif j1_rows:
            boundary = max(r[1] - 1 for r in j1_rows)
        elif tables:
            boundary = max(t['header_no'] - 1 for t in tables)
        found_line, kw = None, ''
        for i in range(boundary, end):
            if fence[i]:
                continue
            m = J4_RE.search(lines[i])
            if m:
                found_line, kw = i, m.group(0)
                break
        if found_line is None:
            for idx, text in sub_heads:
                m = J_SUB_RE.match(text)
                if idx >= boundary and m and unicodedata.normalize('NFKC', m.group(1)) == '4':
                    found_line, kw = idx, 'J-4 小标题'
                    break
        if found_line is None:
            add(13, 'J-4 存在', 'WARN', 'J-1/J-2 之后未见「整体画像」「结构描述」「问题链」字样或 J-4 小标题')
        else:
            add(13, 'J-4 存在', 'PASS', '第 %d 行含「%s」' % (found_line + 1, kw))

    # ---------- 14 内部标记泄露（附录 J 区域内） ----------
    if found:
        lscope, ldesc = range(start, end), '附录 J 区域 第 %d–%d 行' % (start + 1, end)
    else:
        lscope, ldesc = range(0, len(lines)), '全文（未找到附录 J 区域，已回退）'
    leaks = []
    for i in lscope:
        line = lines[i]
        if not line.strip():
            continue
        for label, rx in LEAK_PATTERNS:
            m = rx.search(line)
            if m:
                leaks.append('第 %d 行「%s」（%s）：%s' % (i + 1, m.group(0), label, snippet(line, m.start(), m.end(), 30)))
    if leaks:
        add(14, '内部标记泄露', 'FAIL', '附录 J 区域内发现 %d 处内部标记（范围：%s）' % (len(leaks), ldesc), leaks)
    else:
        add(14, '内部标记泄露', 'PASS', '范围内未发现内部标记（范围：%s）' % ldesc)

    checks.sort(key=lambda c: c['id'])
    n_fail = sum(1 for c in checks if c['status'] == 'FAIL')
    n_warn = sum(1 for c in checks if c['status'] == 'WARN')
    return {
        'file': '',
        'script_version': SCRIPT_VERSION,
        'region': {'found': found, 'fallback': fallback, 'start': start + 1 if found else None,
                   'end': end if found else None, 'title': title if found else ''},
        'level': LEVEL_INFO[level][0] if level else None,
        'checks': checks,
        'summary': {'fail': n_fail, 'warn': n_warn,
                    'pass': sum(1 for c in checks if c['status'] == 'PASS'),
                    'skip': sum(1 for c in checks if c['status'] == 'SKIP')},
    }


def exit_code_of(result):
    return 1 if result['summary']['fail'] > 0 else 0


def check_file(path, source_path=None, whole=False):
    """读取报告（与可选原文）并执行检查，返回结果字典；报告为空时抛 ValueError。"""
    text = read_text(path)
    if not text.strip():
        raise ValueError('报告文件为空，无法解析')
    lines = text.splitlines()
    source_text = read_text(source_path) if source_path else None
    result = analyze(lines, source_text, whole)
    result['file'] = os.path.abspath(path)
    if source_path:
        result['source'] = os.path.abspath(source_path)
    return result


def print_human(result):
    """默认人类可读输出。"""
    reg = result['region']
    print('文件：%s' % result['file'])
    if reg['found']:
        print('附录 J 区域：第 %d–%d 行  %s' % (reg['start'], reg['end'], reg['title'][:50]))
    elif reg['fallback']:
        print('附录 J 区域：未找到起始标题（已退回全文定位 J-1/J-2 表）')
    else:
        print('附录 J 区域：未找到')
    if result.get('level'):
        print('可梳理度：%s' % result['level'])
    print()
    for c in result['checks']:
        print('[%s] %d  %s：%s' % (c['status'], c['id'], c['name'], c['message']))
        for d in c['details'][:MAX_DETAILS_SHOWN]:
            print('        - %s' % d)
        if len(c['details']) > MAX_DETAILS_SHOWN:
            print('        …（另 %d 条，见 --json）' % (len(c['details']) - MAX_DETAILS_SHOWN))
    print()
    print('汇总：FAIL %d 项，WARN %d 项' % (result['summary']['fail'], result['summary']['warn']))


# ================================================================ 内置自测样例

SAMPLE_A = '''# 课堂观察评课报告

## 一、这节课是怎么走的

略。

### J. 教师课堂提问梳理与评价

#### J-0 说明

提问可梳理度：甲级（逐句可还原）。本附录只梳理教师提问，学生提问见附录 A；不打分、不设等第、不做排名。

#### J-1 提问梳理表

| 序号 | 环节 | 教师原问 | 认知层级 | 提问对象与学生应答 | 教师接应 |
|:----:|------|----------|----------|--------------------|----------|
| 1 | 导入 | 这幅图上画了什么？ | 事实复述 | 齐答："鸡和兔" | 肯定 |
| 2 | 导入 | 准备好了吗？ | 确认求同 | 齐答："好了" | 转入下一环节 |
| 3 | 审题 | 来解释一下这段话？ | 理解解释 | 点名·学生A："就是问鸡和兔各有几只" | 肯定后追问 |
| 4 | 审题 | 明白了吗？ | 确认求同 | 齐答："明白" | 转入下一环节 |
| 5 | 探究 | 怎么加的？ | 理解解释 | 点名·学生B："一只鸡加了两条腿" | 探因追问 |
| 6 | 总结 | 这节课你学到了什么？ | 分析评价 | 自请·学生C："画图法" | 复述采纳 |

#### J-2 逐问评价与优化建议表

| 序号 | 评价 | 优化建议 | 建议动作 |
|:----:|------|----------|:--------:|
| 1 | 可取：激活旧知。可提升：仅齐答。 | 追问："你怎么看出来的？" | 加追问 |
| 2、4 | 审题环节的确认性问句，学生答"好了"即可。 | 核心概念处改为需要说理由的问法。 | 改写问句 |
| 3 | 可取：请学生转述题意。可提升：只有一人应答。 | 转问："他抓住了哪两个信息？" | 加追问 |
| 5 | 可取：探因。可提升：无。 | 建议保留 | 保留 |
| 6 | 可取：开放总结。可提升：答后即下课。 | 追问："为什么用画图法？" | 加追问 |

#### J-3 重点问句深评

> **深评卡 · 第 5 问**（探究）
> **对话链还原**：教师："怎么加的？" → 学生B："一只鸡加了两条腿" → 教师："为什么？"
> **这里发生了什么**：探因追问把操作引向算理。
> **假如我来教**：保留探因问，追加转问。
> **留给你的问题**：你更想让学生说清"步骤"还是"道理"？

> **深评卡 · 第 6 问**（总结）
> **对话链还原**：教师："这节课你学到了什么？" → 学生C："画图法"
> **这里发生了什么**：总结问开放，答后即下课。
> **假如我来教**：追加转问。
> **留给你的问题**：总结环节你希望留几分钟？

#### J-4 提问整体画像

**结构描述**：共 6 问；事实复述 1 问、确认求同 2 问、理解解释 2 问、分析评价 1 问。
**问题链识别**：审题 → 探究形成一条主线。

### K. 其他附录

略。
'''

# 样例 A 对应的课堂输入原文（含口语赘词，用于验证溯源 100% 命中）
SAMPLE_A_SOURCE = '''师：同学们，这幅图上画了什么？
生（齐）：鸡和兔。
师：嗯，那那准备好了吗？
生（齐）：好了。
师：来解释一下这段话？
学生A：就是问鸡和兔各有几只。
师：明白了吗？
师：怎么加的？
学生B：一只鸡加了两条腿。
师：这节课你学到了什么？
学生C：画图法。
'''

_LONG_CELL = ('这一问的评价文字刻意写得很长很长，' * 6)[:90]   # 90 字，超过 80 字上限

SAMPLE_B = '''### J 教师课堂提问梳理与评价

#### J-0 说明

提问可梳理度：甲级。

#### J-1 提问梳理表

| 序号 | 环节 | 教师原问 | 认知层级 | 提问对象与学生应答 | 教师接应 |
|:----:|------|----------|----------|--------------------|----------|
| 1 | 导入 | 这幅图上画了什么？ | 事实复述 | 齐答 | 肯定 |
| 2 | 导入 | 准备好了吗？ | 确认求同 | 齐答 | 转入 |
| 4 | 审题 | 来解释一下这段话？ | 理解解释 | 点名·学生A | 追问 |

#### J-2 逐问评价与优化建议表

| 序号 | 评价 | 优化建议 | 建议动作 |
|:----:|------|----------|:--------:|
| 1 | 本问得 4.5 分，属无效提问。 | 改写 | 优秀 |
| 2 | ''' + _LONG_CELL + ''' | 追问 | 加追问 |
| 9 | 多余的一行 | 追问 | 保留 |
'''

SAMPLE_C = '''## 附录 J 教师课堂提问梳理与评价

**J-0 说明**

提问可梳理度：丙级（仅概括描述）。输入只有"教师提问后学生回答"之类的概述，没有原话，未做逐问梳理。

**教师提问补充清单**

1. 导入环节教师提出了什么问题？
2. ……
'''

# 找不到 J 起始标题、但全文有 J-1/J-2 表 → 回退模式
SAMPLE_D = '''# 报告

| 序号 | 环节 | 教师原问 | 认知层级 | 提问对象与学生应答 | 教师接应 |
|---|---|---|---|---|---|
| 1 | 导入 | 图上画了什么？ | 事实复述 | 齐答 | 肯定 |
| 2 | 导入 | 准备好了吗？ | 确认求同 | 齐答 | 转入 |

| 序号 | 评价 | 优化建议 | 建议动作 |
|---|---|---|---|
| 1 | 可取：激活旧知。 | 追问 | 加追问 |
| 2 | 确认性问句。 | 改写 | 改写问句 |
'''

# 溯源小样例：原问含「（推定）」并删了口语赘词；原文有转写错字
TRACE_SOURCE = '''师：嗯，那那大家看一下这个图，上面画了些什么东西？
生：鸡和兔。
师：一只鸡加了两条腿，为十么呢，说说你的想法？
'''
TRACE_QUESTION = '大家看一下这个图，上面画了些什么东西？（推定）／一只鸡加了两条腿，为什么？（推定）'


def selftest():
    """内置自测：样例 A/B/C、回退模式、溯源小样例。全部通过返回 0，否则打印失败详情并返回 1。"""
    failures = []

    def expect(cond, msg):
        if not cond:
            failures.append(msg)

    def status_of(result):
        return {c['id']: c['status'] for c in result['checks']}

    # ---- 样例 A：甲级，完全合规（附带原文，验证溯源 100%）
    res_a = analyze(SAMPLE_A.splitlines(), source_text=SAMPLE_A_SOURCE)
    st = status_of(res_a)
    bad = [(i, s) for i, s in st.items() if s in ('FAIL', 'WARN')]
    expect(not bad, '样例 A 应无 FAIL/WARN，实际：%s' % bad)
    expect(all(st.get(i) == 'PASS' for i in range(1, 15)), '样例 A 全部 14 项应为 PASS，实际：%s' % st)
    expect(st.get(7) == 'PASS' and '命中率 100.0%' in res_a['checks'][6]['message'], '样例 A 溯源应 100% 命中')
    expect(res_a['level'] == '甲级 逐句可还原', '样例 A 等级应为甲级，实际：%s' % res_a['level'])
    expect(exit_code_of(res_a) == 0, '样例 A 退出码应为 0')

    # ---- 样例 B：甲级，多处问题
    res_b = analyze(SAMPLE_B.splitlines())
    st = status_of(res_b)
    fail_set = sorted(i for i, s in st.items() if s == 'FAIL')
    expect(fail_set == [3, 5, 6, 8], '样例 B 的 FAIL 集合应为 [3, 5, 6, 8]，实际：%s' % fail_set)
    expect(st.get(9) == 'WARN', '样例 B 检查 9 应为 WARN，实际：%s' % st.get(9))
    expect(exit_code_of(res_b) == 1, '样例 B 退出码应为 1')

    # ---- 样例 C：丙级降级（经文件路径执行，覆盖读文件链路）
    tmpdir = tempfile.mkdtemp(prefix='check_questions_selftest_')
    path_c = os.path.join(tmpdir, 'report_c.md')
    with open(path_c, 'w', encoding='utf-8') as f:
        f.write(SAMPLE_C.replace('\n', '\r\n'))   # 顺便验证 CRLF 兼容
    res_c = check_file(path_c)
    shutil.rmtree(tmpdir, ignore_errors=True)
    st = status_of(res_c)
    expect(all(s != 'FAIL' for s in st.values()), '样例 C 应无 FAIL，实际：%s' % st)
    expect(all(st.get(i) == 'SKIP' for i in range(3, 9)), '样例 C 检查 3–8 应为 SKIP，实际：%s' % st)
    expect(res_c['level'] == '丙级 仅概括描述', '样例 C 等级应为丙级，实际：%s' % res_c['level'])
    expect(exit_code_of(res_c) == 0, '样例 C 退出码应为 0')

    # ---- 样例 D：无 J 起始标题 → 回退全文定位表格
    res_d = analyze(SAMPLE_D.splitlines())
    st = status_of(res_d)
    expect(st.get(1) == 'FAIL' and st.get(3) == 'PASS' and st.get(5) == 'PASS',
            '样例 D 应为 区域 FAIL、序号 PASS、覆盖 PASS，实际：%s' % st)

    # ---- 溯源小样例：删赘词 + （推定）+ 转写错字
    frags = split_fragments(TRACE_QUESTION)
    expect(len(frags) == 2, '溯源小样例应切出 2 个片段，实际：%s' % frags)
    r = trace_fragments([(1, f) for f in frags], TRACE_SOURCE)
    expect(r['exact'] == 1 and r['fuzzy'] == 1 and not r['misses'],
            '溯源小样例应 1 精确 + 1 模糊命中，实际：%s' % r)

    # ---- 输出
    print('自测：样例 A（甲级合规）      FAIL=%d WARN=%d 退出码=%d' % (res_a['summary']['fail'], res_a['summary']['warn'], exit_code_of(res_a)))
    print('自测：样例 B（甲级多处问题）  FAIL=%d WARN=%d 退出码=%d  FAIL 项=%s' % (res_b['summary']['fail'], res_b['summary']['warn'], exit_code_of(res_b), fail_set))
    print('自测：样例 C（丙级降级）      FAIL=%d WARN=%d 退出码=%d  等级=%s' % (res_c['summary']['fail'], res_c['summary']['warn'], exit_code_of(res_c), res_c['level']))
    print('自测：样例 D（无标题回退）    检查1=%s 检查3=%s 检查5=%s' % (status_of(res_d).get(1), status_of(res_d).get(3), status_of(res_d).get(5)))
    print('自测：溯源小样例              片段=%d 精确=%d 模糊=%d 未命中=%d' % (r['total'], r['exact'], r['fuzzy'], len(r['misses'])))
    if failures:
        print()
        for msg in failures:
            print('失败：' + msg)
        print('自测失败')
        return 1
    print()
    print('自测通过')
    return 0


# ================================================================ 命令行入口

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog='check_questions.py',
        description='评课报告附录 J「教师课堂提问梳理与评价」校验脚本（退出码：0 无 FAIL；1 有 FAIL；2 用法错误/文件不存在/无法解析）')
    parser.add_argument('report', nargs='?', help='评课报告 Markdown 路径')
    parser.add_argument('--source', help='课堂输入原文（听课记录/转写），用于原问溯源检查')
    parser.add_argument('--whole', action='store_true', help='禁用词检查扩大到全文（默认只查附录 J 区域）')
    parser.add_argument('--json', action='store_true', help='输出 JSON 明细')
    parser.add_argument('--selftest', action='store_true', help='运行内置自测，不要求位置参数')
    args = parser.parse_args(argv)

    if args.selftest:
        return selftest()
    if not args.report:
        parser.error('缺少报告路径（或使用 --selftest 运行内置自测）')

    for p in (args.report, args.source):
        if p and not os.path.isfile(p):
            print('错误：文件不存在或不是普通文件：%s' % p, file=sys.stderr)
            return 2
    try:
        result = check_file(args.report, args.source, args.whole)
    except ValueError as e:
        print('错误：%s' % e, file=sys.stderr)
        return 2
    except (OSError, UnicodeDecodeError) as e:
        print('错误：无法读取/解析文件：%s' % e, file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_human(result)
    return exit_code_of(result)


if __name__ == '__main__':
    sys.exit(main())
