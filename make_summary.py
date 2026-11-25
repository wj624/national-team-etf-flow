# -*- coding: utf-8 -*-
"""
从 _etf_data.json（refresh.py 每天生成的完整数据）提炼出「给机器/给 AI/给人」的三份摘要。

产出（都放在仓库根目录，随每日构建一起提交）：
  latest.json   —— 结构化结论数，约 15~25KB，供 Apps Script 邮件脚本读取
  llms.txt      —— 人话版摘要，也用于喂大模型
  snapshot.md   —— 与 llms.txt 同内容，GitHub 上可直接渲染阅读

设计原则：**所有计算都放在这里，下游（GAS）只负责展示**。
这个仓库的指标（连续天数、历史分位、倍量、越跌越买）比单纯展示复杂，
放在 Python 里算便于调试，也让 GAS 脚本保持简单。

用法：
  python make_summary.py              # 读 _etf_data.json，写三份摘要
  python make_summary.py --check      # 只打印 llms.txt，不写文件
  python make_summary.py --src X.json # 指定输入
"""
import json
import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = 'https://wj624.github.io/national-team-etf-flow'

# 分类展示顺序（与页面一致）
CAT_ORDER = ['沪深300', '上证50', '中证500', '中证1000', '创业板', '创业板50',
             '科创板50', '上证180', '上证180金融', '深证100', '中证800', '中证A500', 'MSCI中国A50']

WEEKDAY_CN = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']


# ---------------------------------------------------------------- 工具
def nd(x, dflt=None):
    """安全转 float"""
    try:
        v = float(x)
        return v if v == v else dflt          # 过滤 NaN
    except (TypeError, ValueError):
        return dflt


def fmt_date(k):
    """把 20261030 / 2026-10-30 / 20261030 各种写法统一成 2026-10-30"""
    s = str(k).strip()
    if len(s) == 8 and s.isdigit():
        return f'{s[:4]}-{s[4:6]}-{s[6:]}'
    return s


def weekday_cn(d):
    try:
        return WEEKDAY_CN[datetime.strptime(d, '%Y-%m-%d').weekday()]
    except Exception:
        return ''


def sgn(x):
    return 1 if x > 0 else (-1 if x < 0 else 0)


def pct_rank(series, value):
    """value 在 series 中的百分位（0-100），表示「历史上有多少比例的日子不如今天」"""
    arr = [x for x in series if x is not None]
    if not arr:
        return None
    below = sum(1 for x in arr if x <= value)
    return round(below * 100.0 / len(arr), 1)


def yi(x, d=1):
    """亿元格式化，带正负号"""
    if x is None:
        return '—'
    return f'{x:+,.{d}f}'


def pct(x, d=1):
    if x is None:
        return '—'
    return f'{x:+.{d}f}%'


def ratio(x, d=1):
    if x is None:
        return '—'
    return f'{x:.{d}f}%'


# ---------------------------------------------------------------- 核心计算
def build(p):
    dates = [fmt_date(d) for d in (p.get('dates') or [])]
    asof = fmt_date(p.get('data_asof') or (dates[-1] if dates else ''))
    n_days = len(dates)

    # ---- 每日全表合计 ----
    dt = {}
    for k, v in (p.get('date_total') or {}).items():
        vv = nd(v)
        if vv is not None:
            dt[fmt_date(k)] = vv
    dt_seq = [(d, dt[d]) for d in dates if d in dt]          # 按时间顺序
    total_series = [v for _, v in dt_seq]

    # ---- 分类逐日 ----
    ct = {}
    for k, v in (p.get('cat_total') or {}).items():
        ct[fmt_date(k)] = {str(c): nd(x, 0.0) for c, x in (v or {}).items()}

    # ---- ETF 明细 ----
    etfs_raw = p.get('etfs') or []
    s2025 = {e['c']: nd(e.get('s2025'), 0.0) or 0.0 for e in etfs_raw}
    n2025 = {e['c']: nd(e.get('n2025'), 0.0) or 0.0 for e in etfs_raw}

    today_net = dt.get(asof)
    if today_net is None and dt_seq:
        asof, today_net = dt_seq[-1]

    # ---- 今日净买卖的分配：买入家数 / 卖出家数 ----
    buy_cnt = sell_cnt = flat_cnt = 0
    for e in etfs_raw:
        v = nd((e.get('series') or {}).get(asof), nd(e.get('today'), 0.0)) or 0.0
        if v > 0.01:
            buy_cnt += 1
        elif v < -0.01:
            sell_cnt += 1
        else:
            flat_cnt += 1

    # ---- 历史分位与排名 ----
    pct_today = pct_rank(total_series, today_net) if today_net is not None else None
    rank_txt = ''
    if today_net is not None and total_series:
        if today_net >= 0:
            pos = sum(1 for x in total_series if x > today_net) + 1
            rank_txt = f'史上第 {pos} 大单日净买入（共 {len(total_series)} 个交易日）'
        else:
            pos = sum(1 for x in total_series if x < today_net) + 1
            rank_txt = f'史上第 {pos} 大单日净卖出（共 {len(total_series)} 个交易日）'

    # ---- 连续同向天数 ----
    streak_days, streak_sum, streak_dir = 0, 0.0, ''
    if dt_seq:
        last_dir = sgn(dt_seq[-1][1])
        if last_dir != 0:
            for d, v in reversed(dt_seq):
                if sgn(v) == last_dir:
                    streak_days += 1
                    streak_sum += v
                else:
                    break
        streak_dir = 'buy' if last_dir > 0 else ('sell' if last_dir < 0 else 'flat')

    # ---- 多窗口累计 ----
    def win(n):
        tail = total_series[-n:]
        return round(sum(tail), 1) if tail else None

    w = {'d5': win(5), 'd20': win(20), 'd60': win(60), 'd250': win(250)}

    # ---- 长短期方向是否相反（这是最容易误读的地方，单独点出来）----
    reversal = None
    if w['d20'] is not None and w['d250'] is not None:
        s_short, s_long = sgn(w['d20']), sgn(w['d250'])
        if s_short != 0 and s_long != 0 and s_short != s_long:
            reversal = {
                'short': 'buy' if s_short > 0 else 'sell',
                'long': 'buy' if s_long > 0 else 'sell',
                'text': (f'近 20 日净{"买入" if s_short>0 else "卖出"} {abs(w["d20"]):,.1f} 亿元，'
                         f'而近 250 日累计净{"买入" if s_long>0 else "卖出"} {abs(w["d250"]):,.1f} 亿元。'
                         f'短期方向与一年维度相反，说明这波是{"回补" if s_short>0 else "撤退"}，'
                         f'不是趋势的延续，别把短期动向直接外推成长期方向。')
            }

    # ---- 指数涨跌（中证全指 kline） ----
    idx_chg = None
    kl = [x for x in (p.get('kline') or []) if x]
    kl_map = {fmt_date(x.get('d')): x for x in kl}
    if asof in kl_map and len(kl) >= 2:
        cur = kl_map[asof]
        try:
            i = [fmt_date(x.get('d')) for x in kl].index(asof)
            if i > 0:
                prev = kl[i - 1]
                c0, c1 = nd(prev.get('c')), nd(cur.get('c'))
                if c0 and c1:
                    idx_chg = round((c1 / c0 - 1) * 100, 2)
        except ValueError:
            pass

    # ---- 越跌越买判定 ----
    def tag_of(chg, net):
        if chg is None or net is None:
            return '—'
        if abs(chg) < 0.05:
            return '横盘'
        if chg < 0 and net > 0:
            return '越跌越买'
        if chg < 0 and net < 0:
            return '下跌也在卖'
        if chg > 0 and net > 0:
            return '顺势加仓'
        return '涨了就卖'

    today_tag = tag_of(idx_chg, today_net)

    # 近 10 个交易日的「指数涨跌 × 净买卖」对照
    diverge = []
    kl_dates = [fmt_date(x.get('d')) for x in kl]
    for d, v in dt_seq[-10:]:
        chg = None
        if d in kl_map:
            i = kl_dates.index(d)
            if i > 0:
                c0, c1 = nd(kl[i - 1].get('c')), nd(kl[i].get('c'))
                if c0 and c1:
                    chg = round((c1 / c0 - 1) * 100, 2)
        diverge.append({'d': d, 'idx': chg, 'net': round(v, 1), 'tag': tag_of(chg, v)})

    # 连续「越跌越买」天数
    diverge_streak = 0
    for row in reversed(diverge):
        if row['tag'] == '越跌越买':
            diverge_streak += 1
        else:
            break

    # ---- 持仓市值 ----
    mv = {}
    for k, v in (p.get('mv_total') or {}).items():
        vv = nd(v)
        if vv is not None:
            mv[fmt_date(k)] = vv
    mv_seq = [mv[d] for d in dates if d in mv]
    mv_now = mv.get(asof)
    mv_pct = pct_rank(mv_seq, mv_now) if (mv_now is not None and mv_seq) else None
    mv_hi = mv_lo = mv_hi_d = mv_lo_d = None
    if mv_seq:
        tail = [(d, mv[d]) for d in dates if d in mv][-250:]
        if tail:
            mv_hi_d, mv_hi = max(tail, key=lambda x: x[1])
            mv_lo_d, mv_lo = min(tail, key=lambda x: x[1])
    mv_chg = None
    if mv_now is not None and len(mv_seq) >= 2:
        prev = mv_seq[-2]
        if prev:
            mv_chg = round((mv_now / prev - 1) * 100, 2)

    # ---- 分类汇总 ----
    cats = []
    for cat in CAT_ORDER:
        t = nd((ct.get(asof) or {}).get(cat), 0.0) or 0.0
        w5 = round(sum(nd((ct.get(d) or {}).get(cat), 0.0) or 0.0 for d, _ in dt_seq[-5:]), 1)
        m20 = round(sum(nd((ct.get(d) or {}).get(cat), 0.0) or 0.0 for d, _ in dt_seq[-20:]), 1)
        # 该类 2025-12-31 市值（亿元），用于倍量金额
        base_mv = 0.0
        members = [e for e in etfs_raw if e.get('cat') == cat]
        for e in members:
            base_mv += (s2025.get(e['c'], 0.0) * n2025.get(e['c'], 0.0))
        base_mv = round(base_mv, 1)
        bei = round(t / (base_mv * 0.001), 2) if base_mv > 0 else None
        cats.append({
            'name': cat, 'today': round(t, 1), 'w5': w5, 'm20': m20,
            'base_mv': base_mv, 'bei': bei, 'n_etf': len(members),
        })
    cats.sort(key=lambda x: -abs(x['today']))

    # ---- ETF 明细 ----
    etfs = []
    for e in etfs_raw:
        c = e['c']
        dsh = (e.get('dsh') or {}).get(asof)
        dsh = nd(dsh)
        base = s2025.get(c, 0.0) * 0.001
        bei = round(dsh / base, 2) if (dsh is not None and base > 0) else None
        etfs.append({
            'c': c, 'n': e.get('n', ''), 'cat': e.get('cat', '其他'), 'm': e.get('m', ''),
            'today': round(nd(e.get('today'), 0.0) or 0.0, 1),
            'week': round(nd(e.get('week'), 0.0) or 0.0, 1),
            'month': round(nd(e.get('month'), 0.0) or 0.0, 1),
            'dsh': round(dsh, 4) if dsh is not None else None,
            'bei': bei,
            'share': round(nd(e.get('last_share'), 0.0) or 0.0, 1),
            'nav': nd(e.get('last_nav')),
        })
    etfs.sort(key=lambda x: -abs(x['today']))

    # ---- 最近 20 日序列（供 GAS 画趋势/判断） ----
    recent = [{'d': d, 'net': round(v, 1)} for d, v in dt_seq[-20:]]

    return {
        # updated 是「数据本身的时间戳」（数据未推进时 refresh.py 会沿用旧值以跳过提交），
        # generated 才是「本文件生成时间」，两者不是一回事，邮件里要分开显示。
        'updated': p.get('updated') or '',
        'generated': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'asof': asof,
        'asof_weekday': weekday_cn(asof),
        'scope': {'etfs': len(etfs_raw), 'cats': len(CAT_ORDER), 'days': n_days,
                  'start': dates[0] if dates else ''},
        'today': {
            'net': round(today_net, 1) if today_net is not None else None,
            'buy_count': buy_cnt, 'sell_count': sell_cnt, 'flat_count': flat_cnt,
            'pct': pct_today,
            'rank_text': rank_txt,
            'streak_days': streak_days, 'streak_dir': streak_dir,
            'streak_sum': round(streak_sum, 1),
            'idx_chg': idx_chg, 'tag': today_tag,
        },
        'windows': {'d5': w['d5'], 'd20': w['d20'], 'd60': w['d60'], 'd250': w['d250']},
        'reversal': reversal,
        'mv': {'now': round(mv_now, 1) if mv_now is not None else None,
               'chg_1d': mv_chg, 'pct': mv_pct,
               'hi': round(mv_hi, 1) if mv_hi is not None else None, 'hi_date': mv_hi_d,
               'lo': round(mv_lo, 1) if mv_lo is not None else None, 'lo_date': mv_lo_d},
        'cats': cats,
        'etfs': etfs,
        'diverge': {'streak': diverge_streak, 'rows': diverge},
        'recent': recent,
    }


# ---------------------------------------------------------------- 文字摘要
def render(d):
    t = d['today']
    L = []
    A = L.append

    A('# 国家队 ETF 净买卖 · 数据摘要')
    A('')
    A(f"> 数据日：{d['asof']}（{d['asof_weekday']}）　数据时间戳 {d['updated'] or '—'}　本摘要生成于 {d['generated']}")
    A(f'> 在线看板：{BASE}/')
    A(f'> 机器可读：{BASE}/latest.json')
    A('')
    A('## 这是什么')
    A('')
    A(f"跟踪 **{d['scope']['etfs']} 只** 宽基 ETF（归入 {d['scope']['cats']} 个指数类别）每日的**份额变化**，")
    A('换算成净买入/净卖出的金额。这批 ETF 是国家队（中央汇金、证金等）托盘的主要工具，')
    A('所以它们的申赎方向常被当作「国家队动向」的代理指标。')
    A('')
    A('**核心算法**：')
    A('```')
    A('净买卖金额 =（当日基金份额 − 前一日基金份额）× 当日单位净值')
    A('       正值 = 净申购（资金进场，红色）')
    A('       负值 = 净赎回（资金离场，绿色）')
    A('倍量份额   = 当日净申购份额 ÷ 该 ETF 在 2025-12-31 的份额 × 0.1%')
    A('```')
    A('「倍量」是相对自身规模归一化后的强度，便于横向比较大小盘 ETF。')
    A('')
    A('## 怎么读这份数据')
    A('')
    A('1. **只看方向不够，要看「和指数涨跌的关系」**。同样是净买入，')
    A('   「跌的时候买」是托底，「涨的时候买」是追涨，含义完全不同。')
    A('2. **单日数据噪音大**。至少看近 5 日、近 20 日累计，以及连续同向天数。')
    A('3. **短期方向不等于长期方向**。近 20 日在买、近 250 日在卖的情况很常见，')
    A('   这通常意味着「回补」而不是「新趋势」，别把两者混为一谈。')
    A('4. **看结构，不只看总量**。沪深300 买、创业板卖，和反过来，透露的信息不同。')
    A('5. **份额变化有 1~2 天披露滞后**（沪深两市公布节奏不同），本页取两市都有数据的最近共同交易日。')
    A('')
    A(f"## 今日读数（{d['asof']} {d['asof_weekday']}）")
    A('')
    net = t['net']
    word = '净买入' if (net or 0) > 0 else ('净卖出' if (net or 0) < 0 else '零变化')
    A('| 项目 | 数值 |')
    A('|---|---|')
    A(f'| **全表{word}** | **{yi(net)} 亿元** |')
    A(f'| 买入 / 卖出家数 | 买 {t["buy_count"]} 家，卖 {t["sell_count"]} 家，持平 {t["flat_count"]} 家 |')
    A(f'| 历史分位 | {ratio(t["pct"])}（近 {d["scope"]["days"]} 个交易日） |')
    A(f'| 历史排名 | {t["rank_text"] or "—"} |')
    if t['streak_days']:
        A(f'| 连续同向 | 已连续 {t["streak_days"]} 个交易日净{"买入" if t["streak_dir"]=="buy" else "卖出"}，累计 {yi(t["streak_sum"])} 亿元 |')
    A(f'| 中证全指当日涨跌 | {pct(t["idx_chg"])} |')
    A(f'| **涨跌与资金的关系** | **{t["tag"]}** |')
    A(f'| 近 5 日累计 | {yi(d["windows"]["d5"])} 亿元 |')
    A(f'| 近 20 日累计 | {yi(d["windows"]["d20"])} 亿元 |')
    A(f'| 近 250 日累计 | {yi(d["windows"]["d250"])} 亿元 |')
    if d['mv']['now'] is not None:
        A(f'| 全表持仓市值 | {d["mv"]["now"]:,.1f} 亿元（历史分位 {ratio(d["mv"]["pct"])}） |')
    A('')

    # 长短期方向相反 —— 最容易误读的地方，单独强调
    if d.get('reversal'):
        A('## ⚠️ 注意：短期与一年维度方向相反')
        A('')
        A(d['reversal']['text'])
        A('')
        if d['mv'].get('hi') is not None:
            A(f"参考：全表持仓市值现在 {d['mv']['now']:,.1f} 亿元，"
              f"近一年最高 {d['mv']['hi']:,.1f} 亿元（{d['mv']['hi_date']}），"
              f"最低 {d['mv']['lo']:,.1f} 亿元（{d['mv']['lo_date']}）。")
            A('')

    # 越跌越买
    A('## 指数涨跌 × 资金流向（最近 10 个交易日）')
    A('')
    A('这张表把「指数当天涨跌」和「ETF 当天净买卖」放在一起，用来判断这笔钱是托底还是追涨。')
    A('')
    A('| 日期 | 中证全指 | 净买卖(亿) | 判断 |')
    A('|---|---|---|---|')
    for r in reversed(d['diverge']['rows']):
        A(f'| {r["d"]} | {pct(r["idx"])} | {yi(r["net"])} | {r["tag"]} |')
    A('')
    if d['diverge']['streak'] >= 2:
        A(f"> 已连续 **{d['diverge']['streak']} 个交易日「越跌越买」**，托底意愿较强。")
        A('')
    A('四种判断的含义：')
    A('')
    A('| 判断 | 含义 |')
    A('|---|---|')
    A('| 越跌越买 | 指数跌、资金却在进。最典型的托底形态 |')
    A('| 顺势加仓 | 指数涨、资金也在进。偏追涨 |')
    A('| 涨了就卖 | 指数涨、资金在出。趁反弹减仓 |')
    A('| 下跌也在卖 | 指数跌、资金也在出。没有托底意愿 |')
    A('')

    # 分类
    A('## 各指数类别流向（亿元）')
    A('')
    A('| 类别 | 今日 | 近 5 日 | 近 20 日 | 倍量 | 标的数 |')
    A('|---|---|---|---|---|---|')
    for c in d['cats']:
        bei = f'{c["bei"]} 倍' if c['bei'] is not None else '—'
        A(f'| {c["name"]} | {yi(c["today"])} | {yi(c["w5"])} | {yi(c["m20"])} | {bei} | {c["n_etf"]} |')
    A('')

    # 明细
    A(f'## {d["scope"]["etfs"]} 只 ETF 明细（按今日净买卖绝对值排序）')
    A('')
    A('| 代码 | 名称 | 类别 | 今日 | 近一周 | 近一月 | 倍量 | 最新份额(亿份) |')
    A('|---|---|---|---|---|---|---|---|')
    for e in d['etfs']:
        bei = f'{e["bei"]} 倍' if e['bei'] is not None else '—'
        A(f'| {e["c"]} | {e["n"]} | {e["cat"]} | {yi(e["today"])} | {yi(e["week"])} | {yi(e["month"])} | '
          f'{bei} | {e["share"]:,.1f} |')
    A('')

    # 名词
    A('## 名词速查')
    A('')
    A('| 名词 | 含义 |')
    A('|---|---|')
    A('| 净申购 / 净赎回 | 基金份额比前一交易日增加 / 减少，通常意味着资金进场 / 离场 |')
    A('| 倍量份额 | 当日净申购份额 ÷ 该 ETF 2025-12-31 份额 × 0.1%，是归一化后的强度。'
      '**基准固定在 2025-12-31**，之后份额变化大的品种（如创业板 ETF）倍量会被放大，'
      '看趋势可以，不宜跨品种横向比较 |')
    A('| 历史分位 | 今天的数值在过去几年中排在什么位置，90% 表示只有 10% 的日子比今天更极端 |')
    A('| 越跌越买 | 指数下跌但 ETF 净买入，典型的逆势托底行为 |')
    A('| 持仓市值 | 全表 ETF 的份额 × 净值加总，可近似看作这批资金的体量 |')
    A('| 数据日 | 份额披露有滞后，本页取沪深两市都有数据的最近共同交易日 |')
    A('')
    A('## 数据来源')
    A('')
    A('| 用途 | 来源 |')
    A('|---|---|')
    A('| 沪市 ETF 份额 | 上交所官方披露，经自建 `a-stock-api` Worker `/api/etf-shares` 逐日代理 |')
    A('| 深市 ETF 份额 | 深交所官方 xlsx，本仓直连解析 |')
    A('| 单位净值 | 天天基金（东财 pingzhongdata），经同一 Worker `/api/fund-nav` |')
    A('| 交易日历 | 同一 Worker `/api/trade-calendar` |')
    A('| 指数 K 线 | 东方财富（主），失败时降级腾讯行情 |')
    A('')
    A('> 2026-11 起**不再依赖 AKShare**：份额 / 净值 / 交易日历统一走自建 Worker 代理，'
      '深市份额直连交易所官方 xlsx；新增标的的历史份额从同源镜像一次性导入。'
      '详见仓库 README 第四章「数据来源」与附录 H。')
    A('')
    A('## 免责声明')
    A('')
    A('本摘要由脚本从公开数据自动生成，仅为数据观察，**不构成任何投资建议**。')
    A('ETF 份额变化不等同于国家队实际操作，仅作为公开信息的代理观察口径。')
    A('')
    return '\n'.join(L)


def main():
    src = None
    for a in sys.argv[1:]:
        if a.startswith('--src='):
            src = a.split('=', 1)[1]
    src = src or os.path.join(HERE, '_etf_data.json')

    if not os.path.exists(src):
        print(f'找不到输入文件：{src}')
        print('请先运行 python refresh.py 生成 _etf_data.json')
        sys.exit(1)

    with open(src, encoding='utf-8') as f:
        p = json.load(f)

    d = build(p)
    md = render(d)
    payload = json.dumps(d, ensure_ascii=False, indent=1)

    if '--check' in sys.argv:
        print(md)
        return

    with open(os.path.join(HERE, 'latest.json'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(payload)
    with open(os.path.join(HERE, 'llms.txt'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(md)
    with open(os.path.join(HERE, 'snapshot.md'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(md)

    print(f'latest.json   {len(payload.encode("utf-8")):>8,} bytes')
    print(f'llms.txt      {len(md.encode("utf-8")):>8,} bytes')
    print(f'snapshot.md   {len(md.encode("utf-8")):>8,} bytes')
    print(f'数据日 {d["asof"]}（{d["asof_weekday"]}），全表{("净买入" if (d["today"]["net"] or 0)>0 else "净卖出")} {d["today"]["net"]} 亿元')


if __name__ == '__main__':
    main()
