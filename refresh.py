# -*- coding: utf-8 -*-
"""
国家队(中央汇金)核心宽基ETF 每日净买入/净卖出 自动刷新脚本
运行方式: python refresh_national_team_etf.py
每次运行: 重新拉取沪深两市 ETF 份额与净值, 计算每日净买卖, 生成自包含HTML报表
净买入/净卖出 = (当日份额 - 前一日份额) × 当日单位净值 (正=净买入, 负=净卖出)
"""
from datetime import datetime, timedelta, timezone
import json, time, os, io, sys, shutil, requests, re
import pandas as pd

# 数据源统一走 a-stock-api Worker（份额/净值/交易日历）；深交所规模直连官方 xlsx（见 _szse_scale）
API_BASE = "https://a-stock-api.wj4020410.workers.dev"

def _api_get(path, params=None, timeout=60):
    """GET a-stock-api 端点，返回 JSON。"""
    r = requests.get(API_BASE + path, params=params, timeout=timeout)
    r.raise_for_status()
    return r.json()

def _szse_scale(start, end):
    """深交所基金规模(日频, ETF) — 直连官方 xlsx，等价于原 akshare.fund_scale_daily_szse。
    start/end 为 YYYYMMDD；返回 [{date:'YYYYMMDD', code:'6位', shares:份}]。日期区间≤6个月。"""
    import random
    params = {
        "SHOWTYPE": "xlsx", "CATALOGID": "scsj_fund_jjgm", "TABKEY": "tab1",
        "txtStart": f"{start[:4]}-{start[4:6]}-{start[6:]}",
        "txtEnd": f"{end[:4]}-{end[4:6]}-{end[6:]}",
        "jjlb": "ETF", "random": str(random.random()),
    }
    r = requests.get("https://www.szse.cn/api/report/ShowReport", params=params, headers={
        "Host": "www.szse.cn",
        "Referer": "https://www.szse.cn/market/fund/volume/etf/index.html",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/132",
    }, timeout=60)
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content), engine="openpyxl").dropna(how="all")
    if df.empty:
        return []
    df = df.rename(columns={"基金规模(份)": "基金份额"})
    out = []
    for _, row in df.iterrows():
        try:
            dd = pd.to_datetime(row.get("日期")).strftime("%Y%m%d")
            code = str(int(float(row.get("基金代码")))).zfill(6)
            v = float(str(row.get("基金份额")).replace(",", ""))
        except Exception:
            continue
        out.append({"date": dd, "code": code, "shares": v})
    return out

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(OUT_DIR, "国家队ETF每日净买入卖出报表.html")
SHARES_CACHE = os.path.join(OUT_DIR, "_etf_shares_cache.json")
# 同源镜像兜底: 与本站同源(均为上交所份额数据)的另一份部署, 其份额缓存用于本地取数失败/取不全时的整行兜底。
# 注意这是"镜像"不是"独立源"——两站取的是同一份上交所数据, 只用于解决"本仓云端连不上上交所"的问题。
MIRROR_CACHE_URL = "https://raw.githubusercontent.com/wwww1998/national-team-etf-flow/main/_etf_shares_cache.json"

def gen_cal_days(start_d, end_d):
    """start_d..end_d 之间的工作日(周一~周五), 返回 YYYYMMDD 序列"""
    out, cur = [], start_d
    while cur <= end_d:
        if cur.weekday() < 5:
            out.append(cur.strftime("%Y%m%d"))
        cur += timedelta(days=1)
    return out

def incremental_dates(cache, cal_days):
    """增量拉取窗口: 有缓存则只拉最近若干天(含补拉最近12天以便捕获晚发布), 无缓存则全量构建;
    缓存起始晚于5年窗口时, 从5年起点补齐缺失历史(仅需一次, 补足后缓存即含完整5年)."""
    hist = (cache or {}).get("shares", {})
    today = datetime.now()
    if not hist:
        return natural_dates(cal_days)
    try:
        lastc = datetime.strptime(max(hist), "%Y%m%d")
        firstc = datetime.strptime(min(hist), "%Y%m%d")
    except Exception:
        lastc = today - timedelta(days=1); firstc = today
    five_years = today - timedelta(days=cal_days)   # 5年窗口起点(约5年前同日期)
    start = lastc + timedelta(days=1)
    if firstc > five_years:
        start = five_years   # 历史不足, 从5年起点补齐
    refresh = today - timedelta(days=12)
    if start > refresh:
        start = refresh
    return gen_cal_days(start, today)

FRESH_RETRY = 2          # 数据滞后时最多重试次数
FRESH_WAIT = 150         # 每次重试间隔秒数

def load_trade_calendar(cache):
    """A股交易日历(YYYYMMDD集合). 优先复用缓存(须覆盖到今日之后), 否则拉取新浪交易日历并写回缓存.
    失败时返回已缓存的部分或 None, 上层退化为仅按工作日过滤(不会中断报表生成)."""
    today = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d")
    td = cache.get("trade_days") or {}
    days = sorted(td.get("days") or [])
    if days and days[-1] >= today:
        return set(days)
    try:
        j = _api_get("/api/trade-calendar", {"start": "2010-01-01"})
        days = sorted({str(x)[:10].replace("-", "") for x in (j.get("data") or [])})
        if days:
            cache["trade_days"] = {"days": days, "end": days[-1]}
            print("交易日历已更新: %s ~ %s (共%d个交易日)" % (days[0], days[-1], len(days)), flush=True)
            return set(days)
    except Exception as e:
        print("交易日历拉取失败(退化为工作日过滤):", str(e)[:80], flush=True)
    return set(days) if days else None

def expected_latest_cut(trade_days=None):
    """最近一个已结束交易日(YYYYMMDD), 按北京时间; 周末与法定节假日自动回落到上一交易日.
    日期超出日历覆盖范围(如跨年日历尚未更新)时退化为仅按工作日判断, 避免误判."""
    d = (datetime.now(timezone(timedelta(hours=8))) - timedelta(days=1)).date()
    cal_last = max(trade_days) if trade_days else ""
    for _ in range(400):
        ds = d.strftime("%Y%m%d")
        if d.weekday() < 5 and (ds > cal_last or ds in trade_days):
            return int(ds)
        d -= timedelta(days=1)
    return int(d.strftime("%Y%m%d"))

def sync_day_of(shares, valid):
    """沪深两市都有数据的最近共同交易日; 仅单边有数据时退化为该单边最新日"""
    mkt_of = {c: m for c, _, m in TARGETS}
    sh, sz = set(), set()
    for ds in valid:
        cds = set(shares.get(ds, {}))
        if any(mkt_of.get(c) == "sh" for c in cds):
            sh.add(ds)
        if any(mkt_of.get(c) == "sz" for c in cds):
            sz.add(ds)
    inter = sorted(sh & sz)
    if inter:
        return inter[-1]
    uni = sorted(sh | sz)
    return uni[-1] if uni else None

def last_published_state():
    """读取上一份报表的 (数据日, 更新时间), 用于数据未推进时沿用旧时间戳, 使输出字节不变、工作流跳过提交"""
    cands = [REPORT]
    if os.path.exists(os.path.join(os.getcwd(), "index.html")):
        cands.append(os.path.join(os.getcwd(), "index.html"))
    for p in cands:
        if not os.path.exists(p):
            continue
        try:
            s = open(p, encoding="utf-8-sig").read()
        except Exception:
            continue
        day = re.search(r"数据日\s*<b>(\d{8})</b>", s)
        upd = re.search(r"更新时间\s*<b>([^<]+)</b>", s)
        if day:
            return {"day": day.group(1), "updated": (upd.group(1) if upd else None)}
    return None

def load_shares_cached(window_dates, cache, trade_days=None):
    """拉取窗口数据并合并进持久化缓存, 返回 (合并后shares, 所有有数据日期, deep_start).
    trade_days 非空时: 丢弃非交易日(法定节假日/周末)份额, 防止节假日被快照兜底误填成"幽灵交易日"污染X轴."""
    hist = cache.setdefault("shares", {})
    fetched, valid_fetch, ds_start = load_shares(window_dates)
    for d, m in fetched.items():
        if m and (not trade_days or d in trade_days):
            hist.setdefault(d, {}).update(m)
    if trade_days:
        cal_last = max(trade_days)
        for d in [x for x in list(hist) if x not in trade_days and x <= cal_last]:
            del hist[d]
    if ds_start:
        cached_ds = cache.get("deep_start")
        if not cached_ds or str(ds_start) < str(cached_ds):
            cache["deep_start"] = ds_start
    cached_days = sorted(hist.keys())
    valid = sorted(d for d in cached_days if hist[d])
    return hist, valid, cache.get("deep_start")

TARGETS = [
    ("510300","沪深300ETF 华泰柏瑞","sh"),("510310","沪深300ETF 易方达","sh"),
    ("510330","沪深300ETF 华夏","sh"),("159919","沪深300ETF 嘉实","sz"),
    ("515330","沪深300ETF 天弘","sh"),
    ("510050","上证50ETF 华夏","sh"),("510100","上证50ETF 易方达","sh"),
    ("510500","中证500ETF 南方","sh"),("512500","中证500ETF 华夏","sh"),
    ("159922","中证500ETF 嘉实","sz"),
    ("512100","中证1000ETF 南方","sh"),("159845","中证1000ETF 华夏","sz"),
    ("560010","中证1000ETF 广发","sh"),("159629","中证1000ETF 富国","sz"),
    ("159633","中证1000ETF 易方达","sz"),
    ("159915","创业板ETF 易方达","sz"),("159952","创业板ETF 广发","sz"),
    ("159977","创业板ETF 天弘","sz"),
    ("588080","科创50ETF 易方达","sh"),("588050","科创50ETF 工银","sh"),
    ("588000","科创50ETF 华夏","sh"),
    ("563360","A500ETF 华泰柏瑞","sh"),("512050","A500ETF 华夏","sh"),
    ("563800","A500ETF 广发","sh"),("563220","A500ETF 富国","sh"),
    ("159338","中证A500ETF 国泰","sz"),("159352","A500ETF 南方","sz"),
    ("159361","A500ETF 易方达","sz"),("159351","A500ETF 嘉实","sz"),
    ("510180","上证180ETF 华安","sh"),("510230","上证180金融ETF 国泰","sh"),
    ("159901","深证100ETF 易方达","sz"),
    ("515800","中证800ETF 汇添富","sh"),("560050","MSCI中国A50ETF 汇添富","sh"),
]

# 2026-11-13 对齐上游 9cb6ddc: 34 只 = 23 只汇金持仓基础口径 + 11 只救市专项口径
# 2026-10 救市口径新增纳入的标的(原有23只之外的全部), 在明细行内打标
NEW_TAG = "2026.10救市"
# 原有23只口径: 2025 年报披露的汇金系持仓宽基ETF, 在明细行内打标
BASE_TAG = "2025末汇金持仓"
BASE23 = {"510300","510310","510330","159919","510050","510100","510500","512500",
          "159922","512100","159845","560010","159629","159915","159952","159977",
          "588080","588050","159901","510180","510230","515800","560050"}
# 动态派生: 以后加减标的无需再手工维护集合(见 README 附录 K)
NEW_TAG_CODES = {c for c, _, _ in TARGETS} - BASE23


# 同类指数归集
CAT_MAP = {
    "510300":"沪深300","510310":"沪深300","510330":"沪深300","159919":"沪深300",
    "515330":"沪深300",
    "510050":"上证50","510100":"上证50",
    "510500":"中证500","512500":"中证500","159922":"中证500",
    "512100":"中证1000","159845":"中证1000","560010":"中证1000","159629":"中证1000",
    "159633":"中证1000",
    "159915":"创业板","159952":"创业板","159977":"创业板",
    "588080":"科创板50","588050":"科创板50","588000":"科创板50",
    "563360":"中证A500","512050":"中证A500","563800":"中证A500","563220":"中证A500",
    "159338":"中证A500","159352":"中证A500","159361":"中证A500","159351":"中证A500",
    "510180":"上证180","510230":"上证180金融",
    "159901":"深证100",
    "515800":"中证800","560050":"MSCI中国A50",
}


# 各类指数 → 行情K线源(东财secid, 腾讯code)。用于"各类指数K线 × 该类ETF净买入联动"图
CAT_INDEX = {
    "沪深300":    ("1.000300", "sh000300"),
    "上证50":     ("1.000016", "sh000016"),
    "中证500":    ("1.000905", "sh000905"),
    "中证1000":   ("1.000852", "sh000852"),
    "创业板":     ("0.399006", "sz399006"),
    "科创板50":   ("1.000688", "sh000688"),
    "上证180":    ("1.000010", "sh000010"),
    "上证180金融":  ("1.000018", "sh000018"),
    "深证100":    ("0.399330", "sz399330"),
    "中证800":    ("1.000906", "sh000906"),
    "中证A500":   ("1.000510", "sh000510"),
    # MSCI 中国A50 互联互通指数在东财/腾讯均无行情, 用跟踪它的 560050 价格走势代理
    "MSCI中国A50": ("1.560050", "sh560050"),
}
CAL_DAYS = 1850  # 约5年(约1300个工作日) 覆盖全报表5年历史

def retry(fn, tries=3, gap=2):
    last = None
    for i in range(tries):
        try:
            return fn()
        except Exception as e:
            last = f"{type(e).__name__}: {str(e)[:90]}"
            time.sleep(gap)
    raise RuntimeError(last)

def natural_dates(n):
    dates, d = [], datetime.now()
    while len(dates) < n:
        if d.weekday() < 5:
            dates.append(d.strftime("%Y%m%d"))
        d -= timedelta(days=1)
    return dates

def fetch_kline(dates, secid="1.000985"):
    """直连东财拉取指数(secid, 默认中证全指000985)日K, 按 dates 对齐返回 [{d,o,h,l,c}|None]. 东财单次限约999条."""
    if not dates:
        return []
    start = dates[0].replace("-", ""); end = dates[-1].replace("-", "")
    url = ("https://push2his.eastmoney.com/api/qt/stock/kline/get"
           "?fields1=f1,f2,f3,f4,f5,f6&fields2=f51,f52,f53,f54,f55,f56"
           "&ut=7eea3edcaed734bea9cbfc24409ed989&klt=101&fqt=0&secid=%s"
           "&beg=%s&end=%s&smplmt=1000&lmt=200000" % (secid, start, end))
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://quote.eastmoney.com/"}

    def _get():
        j = requests.get(url, headers=headers, timeout=45).json()
        if not isinstance(j, dict):
            raise ValueError("bad resp type")
        ks = (j or {}).get("data", {}).get("klines")
        if not ks:
            raise ValueError("empty klines")
        return ks

    data = retry(_get, tries=5, gap=6)
    m = {}
    for line in data:
        f = line.split(",")
        k = f[0].replace("-", "")
        if k.isdigit() and len(k) == 8:
            # 顺序: 日期,开,收,高,低,...
            m[k] = [round(float(f[1]), 3), round(float(f[3]), 3),
                    round(float(f[4]), 3), round(float(f[2]), 3)]
    if not m:
        return [None for _ in dates]
    return [({"d": d, "o": m[k][0], "h": m[k][1], "l": m[k][2], "c": m[k][3]}
             if (k := d.replace("-", "")) in m else None) for d in dates]

def fetch_kline_tx(dates, tx_code="sh000985"):
    """备选源(腾讯) 指数(sh000985)日K, 与 fetch_kline 同返回格式. 行: 日期,开,收,高,低,量"""
    if not dates:
        return []
    def _d8(x):
        s = str(x)
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}" if len(s) == 8 and s.isdigit() else s
    start = _d8(dates[0]); end = _d8(dates[-1])
    url = ("https://web.ifzq.gtimg.cn/appstock/app/fqkline/get"
           "?param=%s,day,%s,%s,1600,qfq" % (tx_code, start, end))
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://gu.qq.com/"}

    def _get():
        j = requests.get(url, headers=headers, timeout=45).json()
        if not isinstance(j, dict):
            raise ValueError("bad resp type")
        d = j.get("data")
        rows = None
        if isinstance(d, dict):
            for sym in (tx_code, "qfqday", "day"):
                v = d.get(sym)
                if isinstance(v, dict):
                    rows = v.get("qfqday") or v.get("day") or v.get("data")
                elif isinstance(v, list) and v:
                    rows = v
                if rows:
                    break
        elif isinstance(d, list) and d:
            rows = d
        if not rows:
            raise ValueError("empty tx klines")
        return rows

    rows = retry(_get, tries=8, gap=8)
    m = {}
    for f in rows:
        try:
            if len(f) >= 5 and f[0].isdigit() is False and "-" in f[0]:
                k = f[0].replace("-", "")
                m[k] = [round(float(f[1]), 3), round(float(f[3]), 3),
                        round(float(f[4]), 3), round(float(f[2]), 3)]
        except Exception:
            pass
    return [({"d": d, "o": m[k][0], "h": m[k][1], "l": m[k][2], "c": m[k][3]}
             if (k := d.replace("-", "")) in m else None) for d in dates]

def load_shares(dates):
    """沪: fund_etf_scale_sse 逐日; 深: fund_scale_daily_szse 区间批量"""
    sh_codes = [c for c, _, mk in TARGETS if mk == "sh"]
    sz_codes = [c for c, _, mk in TARGETS if mk == "sz"]
    daily_shares = {ds: {} for ds in dates}
    valid = []
    for ds in dates:
        try:
            j = retry(lambda ds=ds: _api_get("/api/etf-shares",
                     {"exchange": "SH", "date": f"{ds[:4]}-{ds[4:6]}-{ds[6:]}"}))
            rows = (j or {}).get("data") or []
            if not rows:
                continue
            for row in rows:
                try:
                    # Worker 返回 shares_10k(万份)，×10000 转「份」，与原 akshare 口径一致
                    daily_shares[ds][str(row.get("code")).zfill(6)] = float(row.get("shares_10k")) * 10000.0
                except Exception:
                    pass
            valid.append(ds)
        except Exception as e:
            print(f"  沪 fail {ds}: {str(e)[:70]}", flush=True)
    valid.sort()
    deep_start = None
    if sz_codes:
        sz_days = sorted(dates)
        CH = 45
        for i in range(0, len(sz_days), CH):
            s = sz_days[i]; e = sz_days[min(i + CH - 1, len(sz_days) - 1)]
            try:
                srows = retry(lambda s=s, e=e: _szse_scale(s, e))
            except Exception as ex:
                print(f"  深批 fail {s}~{e}: {str(ex)[:50]}", flush=True); continue
            if not srows:
                continue
            for row in srows:
                dd = row["date"]; code = row["code"]
                if code in sz_codes:
                    daily_shares.setdefault(dd, {})[code] = row["shares"]
        sz_have = sorted(d for d, m in daily_shares.items() if any(c in m for c in sz_codes))
        if sz_have:
            valid = sorted(set(valid) | set(sz_have))
            deep_start = sz_have[0]
    return daily_shares, valid, deep_start

def fetch_mirror_row(ds, codes):
    """同源镜像兜底: 从同源仓库(wwww1998)的份额缓存读取指定交易日的整行数据。
    安全约束(宁缺勿错):
      1. 该日必须存在于镜像缓存中(日期精确匹配, 不接受"最近一天"顶替);
      2. 镜像整行必须齐全——除本表标的齐备外, 该行代码数须达到完整行量级, 防止镜像自身也是半行;
      3. 任一条件不满足即返回 None, 由调用方退回"上一完整交易日"。
    返回值: {code: 份额(份)} 或 None。任何异常都只打印日志, 不阻断主流程。"""
    try:
        r = requests.get(MIRROR_CACHE_URL, timeout=25,
                         headers={"User-Agent": "Mozilla/5.0", "Cache-Control": "no-cache"})
        r.raise_for_status()
        row = ((r.json().get("shares") or {}).get(str(ds))) or {}
        if not row:
            print("  镜像兜底: 镜像缓存无 %s, 跳过" % ds, flush=True)
            return None
        # 门槛随上游镜像规模自适应: 原版 2026-11-05 起改用 keep 过滤只存跟踪标的(2026-11-13 起为 34 只),
        # 写死 100 会让同源镜像兜底永久静默失效。判定语义不变: 整行连本表沪市标的数都不到 => 半行。
        # 真正的完整性仍由下方 len(got) != len(codes) 逐只校验兜住, 一道都没松。
        if len(row) < len(codes):
            print("  镜像兜底: %s 行仅 %d 个代码(< %d 只沪市标的), 疑似半行, 拒绝"
                  % (ds, len(row), len(codes)), flush=True)
            return None
        got = {}
        for c in codes:
            try:
                v = float(row[c])
            except Exception:
                v = 0
            if v > 0:
                got[c] = v
        if len(got) != len(codes):
            print("  镜像兜底: %s 行缺 %s, 拒绝"
                  % (ds, ",".join([c for c in codes if c not in got][:5])), flush=True)
            return None
        return got
    except Exception as e:
        print("  镜像兜底失败:", str(e)[:90], flush=True)
        return None

def import_new_codes_history(cache, trade_days=None):
    """口径 29->34 只切换时, 从同源镜像一次性补入新增标的的历史份额(单位: 份)。
    设计要点(见 README 附录 H / K):
      - 只补空缺, 不覆盖本仓现有 Worker 数据; 重叠值差异 > 0.5% 只报警不改。
      - 镜像与本站同源(同为交易所披露份额), 已实测重叠 44 个交易日逐日差值为 0。
      - 写 hist_done 标记保证幂等: 成功一次后不再下载镜像, 省流量、不产生运行时垃圾。
      - 镜像日期集合是本仓缓存日期集合的子集, 故只补已有日期行, 绝不新建半行、不动 valid。
      - 任何失败都只打印日志并返回 0, 不阻断主流程(当日数据仍由 Worker/深交所提供)。
    """
    hist = cache.setdefault("shares", {})
    done = cache.setdefault("hist_done", {})
    if all(done.get(c) for c in NEW_TAG_CODES):
        return 0
    try:
        r = requests.get(MIRROR_CACHE_URL, timeout=30,
                         headers={"User-Agent": "Mozilla/5.0", "Cache-Control": "no-cache"})
        r.raise_for_status()
        src_shares = r.json().get("shares") or {}
    except Exception as e:
        print("新增标的历史导入: 镜像读取失败(%s), 跳过(不影响当日报表)" % str(e)[:80], flush=True)
        return 0
    filled = 0
    conflict = 0
    for ds, row in src_shares.items():
        ds = str(ds)
        if len(ds) != 8 or not ds.isdigit():
            continue
        if trade_days and ds not in trade_days:   # 不把非交易日写进缓存
            continue
        if ds not in hist:                       # 绝不新建日期行(防半行)
            continue
        if not isinstance(row, dict):
            continue
        for c in NEW_TAG_CODES:
            try:
                v = float(row[c])
            except (KeyError, TypeError, ValueError):
                continue
            if v <= 0:
                continue
            cur = hist[ds].get(c)
            if not cur:
                hist[ds][c] = v
                filled += 1
            elif cur > 0 and abs(cur - v) / cur > 0.005:
                conflict += 1                    # 以现有 Worker 数据为准
    for c in sorted(NEW_TAG_CODES):
        n = sum(1 for m in hist.values() if c in m)
        print("新增标的历史导入 %s: 累计 %d 个交易日" % (c, n), flush=True)
        if n >= 5:
            done[c] = "1"
    if conflict:
        print("新增标的历史导入: %d 处镜像与现有值差异>0.5%%, 已保留现有值(Worker 优先)" % conflict, flush=True)
    print("新增标的历史导入完成: 补入 %d 个数值" % filled, flush=True)
    return filled

def load_nav():
    nav = {}
    for code, _, _ in TARGETS:
        try:
            j = retry(lambda c=code: _api_get("/api/fund-nav", {"code": c}))
            m = {}
            for row in (j.get("data") or []):
                d = str(row.get("date"))[:10].replace("-", "")
                try:
                    m[d] = float(row.get("nav"))
                except Exception:
                    pass
            nav[code] = m
        except Exception as e:
            print(f"  净值 fail {code}: {str(e)[:70]}", flush=True)
    return nav

def main():
    print("拉取数据中 ...", flush=True)
    cache = {}
    if os.path.exists(SHARES_CACHE):
        try:
            with open(SHARES_CACHE, encoding="utf-8-sig") as f:
                cache = json.load(f)
        except Exception as e:
            print("缓存损坏, 将全量重建:", e, flush=True)
            cache = {}
    window = incremental_dates(cache, CAL_DAYS)
    print("本次拉取日期窗口: %s ~ %s (共%d个工作日)" %
          (window[0] if window else "-", window[-1] if window else "-", len(window)), flush=True)
    trade_days = load_trade_calendar(cache)
    shares, valid, deep_start = load_shares_cached(window, cache, trade_days)
    # 口径由 29 只扩到 34 只后, 新增标的缺历史份额: 从同源镜像一次性导入(幂等, 失败不阻断)
    import_new_codes_history(cache, trade_days)
    # 数据新鲜度自检: 若展示日落后于最近已结束交易日(通常因数据源晚发布/单边失败), 等待后重试抓取
    exp = expected_latest_cut(trade_days)
    for _attempt in range(FRESH_RETRY):
        cur = sync_day_of(shares, valid)
        stale = cur is not None and int(cur) < exp
        print("数据新鲜度: 最新同步日=%s 期望=%s %s" %
              (cur, exp, "正常" if not stale else "滞后, 稍后重试"), flush=True)
        if not stale or not window:
            break
        time.sleep(FRESH_WAIT)
        shares, valid, deep_start = load_shares_cached(window, cache, trade_days)
    # 沪市最新交易日兜底(同源镜像): 云端常连不上上交所(fund_etf_scale_sse 失败), 若最近已结束交易日
    # 沪市份额缺失或不齐, 从【同源镜像仓库】读取同一交易日的整行份额补齐。
    # 设计原则: 宁缺勿错——镜像行必须整行齐全才采纳; 否则丢弃该日半行数据, 让同步日停留在上一完整交易日,
    # 绝不写入"用不可靠替代源凑出来的值"(20261102 事故复盘: 东财快照返回陈旧值导致假减持, 已彻底移除该路径)。
    # 仅当 exp 确为交易日时才兜底, 避免节假日(如国庆/中秋)被误填成"幽灵交易日"在K线图上留空。
    _sh_live = [c for c, _, _m in TARGETS if _m == "sh"]
    _exp_s = str(exp)
    _exp_is_trade = (not trade_days) or (_exp_s in trade_days)
    if _sh_live and _exp_is_trade:
        _row = shares.get(_exp_s) or {}
        _missing = [c for c in _sh_live if not _row.get(c)]
        if _missing:
            _mir = fetch_mirror_row(_exp_s, _sh_live)
            if _mir:
                shares.setdefault(_exp_s, {}).update(_mir)
                valid = sorted(set(valid) | {_exp_s})
                print("同源镜像兜底: 补入 %s 沪市 %d 只(镜像整行齐全)" % (_exp_s, len(_mir)), flush=True)
            else:
                _still = [c for c in _sh_live if not (shares.get(_exp_s) or {}).get(c)]
                if _still:
                    # 沪市不齐且镜像不可用: 整段沪市一起丢弃(而不是只丢缺的那几只), 否则残留的几只
                    # 仍会被 sync_day_of 算作"沪市当日有数据"而推进页面, 产生半行报表(20261102 即此形态)。
                    for _c in _sh_live:
                        shares.get(_exp_s, {}).pop(_c, None)
                    print("沪市 %s 缺 %d 只且镜像不可用: 不写半行, 保留上一完整交易日"
                          % (_exp_s, len(_still)), flush=True)
    with open(SHARES_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)
    if not valid:
        print("未取到任何交易日份额数据"); return 1
    nav = load_nav()

    daily = []
    for code, name, mkt in TARGETS:
        prev = None; lnav = None; prev_nav = None
        for ds in valid:
            v = shares.get(ds, {}).get(code)
            if v is None:
                continue
            nn = nav.get(code, {}).get(ds)
            if nn is not None:
                prev_nav = lnav   # 上一有净值日的净值
                lnav = nn
            net = (v - prev) * lnav if (prev is not None and lnav is not None) else None
            # 份额合并(折算)判定: 份额等比骤变 与 净值反向等比联动(乘积≈1) => 非实际申赎, 净买卖置0
            if prev is not None and v != prev and net is not None and prev_nav and lnav:
                rs = v / prev          # 份额变化比(合并折半<1 / 拆分加倍>1)
                rn = lnav / prev_nav   # 净值反向比(合并后>1 / 拆分后<1)
                if (abs(rs - 1.0) > 0.30 and (rs - 1.0) * (rn - 1.0) < 0
                        and 0.6 < rs * rn < 1.6):
                    net = 0.0
            daily.append({"d": ds, "c": code, "n": name, "m": mkt, "s": v,
                          "w": lnav, "f": net})
            prev = v
        # 缺失份额时补全当日0? 不处理, 汇总时按实际值

    # 建立 日期->每条, 日期->合计
    by_date = {}
    items_by = {}
    for it in daily:
        items_by.setdefault(it["d"], {})[it["c"]] = it
    dates_sorted = sorted({it["d"] for it in daily})
    date_total = {}
    cat_total = {}   # 类别 → {日期: 该类别下所有ETF当日净买入合计}
    for ds in dates_sorted:
        s = sum(x["f"] for x in items_by[ds].values() if x["f"] is not None)
        date_total[ds] = s
        per_cat = {}
        for c, it in items_by[ds].items():
            if it["f"] is None:
                continue
            cat = CAT_MAP.get(c, "其他")
            per_cat[cat] = per_cat.get(cat, 0) + it["f"]
        cat_total[ds] = per_cat

    # 沪深ETF申购数据公布时间不一致: 只有当沪、深两市都更新到同一交易日时, 才把该日作为最新展示日;
    # 若两市最新公布日不同步, 则回退到两市都有数据的最近共同交易日(通常为前一天)。
    mkt_of = {c: m for c, _, m in TARGETS}
    sh_have, sz_have = set(), set()
    for ds in dates_sorted:
        codes = set(items_by.get(ds, {}) or {})
        if any(mkt_of.get(c) == "sh" for c in codes):
            sh_have.add(ds)
        if any(mkt_of.get(c) == "sz" for c in codes):
            sz_have.add(ds)
    sync_days = sorted(sh_have & sz_have)
    cut = sync_days[-1] if sync_days else dates_sorted[-1]
    if cut < dates_sorted[-1]:
        print("沪深最新公布日不同步, 展示回退到两市同步日:", cut,
              "(两市最新: 沪=%s 深=%s)" % (sh_have and sorted(sh_have)[-1], sz_have and sorted(sz_have)[-1]), flush=True)
    dates_sorted = [d for d in dates_sorted if d <= cut]
    today = cut
    week_days = dates_sorted[-5:]
    month_days = dates_sorted[-21:]

    def cum(days):
        dc = {x["c"]: 0.0 for x in daily}
        for ds in days:
            for c, it in items_by.get(ds, {}).items():
                if it["f"] is not None:
                    dc[c] += it["f"]
        return dc
    w_sum = cum(week_days); m_sum = cum(month_days)

    lpub = last_published_state()
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    if lpub and lpub.get("day") == today:
        stamp = lpub.get("updated") or stamp
    # 各类指数K线: 逐类尝试东财secid, 失败退腾讯; 全部失败则该类为 None 列表
    cat_kline = {}
    for _cat, (_e_sec, _tx) in CAT_INDEX.items():
        try:
            cat_kline[_cat] = fetch_kline(dates_sorted, _e_sec)
        except Exception as _e:
            try:
                cat_kline[_cat] = fetch_kline_tx(dates_sorted, _tx)
            except Exception as _e2:
                print(f"指数K线失败 {_cat}: {_e2}", flush=True)
                cat_kline[_cat] = [None] * len(dates_sorted)
    payload = {
        "updated": stamp,
        "data_asof": today,
        "deep_start": deep_start,
        "dates": [f"{d[:4]}-{d[4:6]}-{d[6:]}" for d in dates_sorted],
        "date_total": {k: round(v / 1e8, 3) for k, v in date_total.items()},
        "cat_total": {str(k): {cc: round(vv / 1e8, 3) for cc, vv in v.items()} for k, v in cat_total.items()},
        "cat_kline": cat_kline,
        "etfs": [],
    }
    try:
        payload["kline"] = fetch_kline(dates_sorted)
        print("中证全指K线已更新", flush=True)
    except Exception as e:
        print("东财K线失败, 尝试腾讯源:", e)
        try:
            payload["kline"] = fetch_kline_tx(dates_sorted)
            print("腾讯K线已更新", flush=True)
        except Exception as e2:
            print("K线拉取失败(报表其余部分不受影响):", e2)
            payload["kline"] = [None] * len(dates_sorted)
    dsh_all, s5, n5 = {}, {}, {}
    for code, _, _ in TARGETS:
        # 2025-12-31 份额/净值 基准备锚，用于"每日倍量"：单只按份额、同类/全表按金额
        _prev = None; _dd = {}
        for _ds in dates_sorted:
            _it = items_by.get(_ds, {}).get(code)
            if _it:
                _k = f"{_ds[:4]}-{_ds[4:6]}-{_ds[6:]}"
                if _prev is not None:
                    _dd[_k] = round((_it["s"] - _prev) / 1e8, 4)
                _prev = _it["s"]
        dsh_all[code] = _dd
        s5[code] = (shares.get("20251231", {}).get(code) or 0) / 1e8
        _nn = None
        for _ds in sorted(nav.get(code, {})):
            if str(_ds) <= "20251231":
                _nn = nav[code][_ds]
        n5[code] = _nn or 0

    for code, name, mkt in TARGETS:
        payload["etfs"].append({
            "c": code, "n": name, "m": "沪" if mkt == "sh" else "深",
            "cat": CAT_MAP.get(code, "其他"),
            "today": (items_by.get(today, {}).get(code, {}).get("f") or 0) / 1e8,
            "week": round(w_sum.get(code, 0) / 1e8, 3),
            "month": round(m_sum.get(code, 0) / 1e8, 3),
            "last_share": (lambda c=code: (items_by.get(today, {}).get(c, {}).get("s") or 0) / 1e8)(),
            "last_nav": (items_by.get(today, {}).get(code, {}).get("w") or 0),
            "s2025": s5.get(code, 0),
            "n2025": n5.get(code, 0),
            "dsh": dsh_all.get(code, {}),
            "series": {f"{ds[:4]}-{ds[4:6]}-{ds[6:]}": round((items_by.get(ds, {}).get(code, {}).get("f") or 0) / 1e8, 3)
                       for ds in dates_sorted},
        })
    # 每日总持仓市值(全市场口径, 亿元): 每只 份额×净值, 缺日用前一交易日前值填充
    per_day = {}
    for it in daily:
        per_day.setdefault(it["c"], {})[it["d"]] = (it["s"], it["w"])
    last_seen = {}
    mv_total = {}
    for ds in dates_sorted:
        tot = 0.0
        for c, _, _ in TARGETS:
            pd = per_day.get(c, {})
            if ds in pd:
                last_seen[c] = pd[ds]
            s, w = last_seen.get(c, (None, None))
            if s is not None and w is not None:
                tot += s * w
        mv_total[ds] = round(tot / 1e8, 1)
    payload["mv_total"] = mv_total
    return payload

def escape_html(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def render_html(p):
    def fnum(v, sign=False):
        if v is None:
            return "—"
        s = f"{v:+,.2f}" if sign else f"{v:,.2f}"
        return s
    def cls(v):
        return "pos" if v >= 0 else "neg"
    def ftag(code):
        if code in NEW_TAG_CODES:
            return f'<span class="tag">{NEW_TAG}</span>'
        if code in BASE23:
            return f'<span class="tag tagb">{BASE_TAG}</span>'
        return ''
    dates = p["dates"]
    dt = {}
    for k, v in p["date_total"].items():
        kk = str(k)
        if len(kk) == 8:
            kk = f"{kk[:4]}-{kk[4:6]}-{kk[6:]}"
        dt[kk] = v
    # 近21交易日用于柱状图
    chart = dates[-21:]
    chart_total = [dt.get(d, 0) for d in chart]
    maxabs = max([abs(x) for x in chart_total] + [1])
    # 近一年滚动累计 (曲线), 起点对齐到深市有数据的日期
    deep_start = p.get("deep_start")
    if deep_start and len(str(deep_start)) == 8:
        ds = str(deep_start)
        deep_start = f"{ds[:4]}-{ds[4:6]}-{ds[6:]}"
    bidx_sz = dates.index(deep_start) if deep_start in dates else 0
    bidx = bidx_sz  # 用尽深市可用历史, 曲线取最长
    y_dates = dates[bidx:]
    _mv_src = p.get("mv_total", {})
    _mvm = {}
    for _k, _v in _mv_src.items():
        _kk = str(_k)
        if len(_kk) == 8:
            _kk = f"{_kk[:4]}-{_kk[4:6]}-{_kk[6:]}"
        _mvm[_kk] = _v
    y_mv = [_mvm.get(d, 0) for d in y_dates]
    # 近一周 / 近一月 / 近半年 / 近一年 / 今年以来
    week_days = dates[-5:]
    month_days = dates[-21:]
    halfyear_days = dates[-120:]
    oneyear_days = dates[-242:]
    cur_year = str(datetime.now().year)
    ytd_days = [d for d in dates if str(d)[:4] == cur_year]
    week_tot = sum(dt.get(d, 0) for d in week_days)
    month_tot = sum(dt.get(d, 0) for d in month_days)
    halfyear_tot = sum(dt.get(d, 0) for d in halfyear_days)
    oneyear_tot = sum(dt.get(d, 0) for d in oneyear_days)
    ytd_tot = sum(dt.get(d, 0) for d in ytd_days)
    today_tot = dt.get(dates[-1], 0)

    # --- 各类指数K线联动数据 (下拉切换) ---
    _dates_fmt = dates
    _cat_dt = {}   # 转键: YYYYMMDD -> YYYY-MM-DD, 供每类每日净买入取值
    for _ck, _cv in (p.get("cat_total") or {}).items():
        _ckk = str(_ck)
        if len(_ckk) == 8:
            _ckk = f"{_ckk[:4]}-{_ckk[4:6]}-{_ckk[6:]}"
        _cat_dt[_ckk] = _cv
    _tot_dt = {}   # 中证全指: 全市场每日净买入合计(键转YYYY-MM-DD)
    for _tk, _tv in (p.get("date_total") or {}).items():
        _tkk = str(_tk)
        if len(_tkk) == 8:
            _tkk = f"{_tkk[:4]}-{_tkk[4:6]}-{_tkk[6:]}"
        _tot_dt[_tkk] = _tv
    # 每类: 该类当日ETF净买入合计(k键=YYYY-MM-DD, 含缺失天填0); 中证全指在首项
    def _cat_series(cat):
        return [round((_cat_dt.get(d, {}) or {}).get(cat, 0), 3) for d in _dates_fmt]
    i_kl_names = ["中证全指"] + list(CAT_INDEX.keys())
    _zzkl = p.get("kline") or [None] * len(dates)
    i_kl_series = [_zzkl] + [p.get("cat_kline", {}).get(cat) or [None] * len(dates) for cat in i_kl_names[1:]]
    _tot_series = [round(_tot_dt.get(d, 0), 3) for d in _dates_fmt]
    i_qd_series = [_tot_series] + [_cat_series(cat) for cat in i_kl_names[1:]]
    INAMES = json.dumps(i_kl_names, ensure_ascii=False)
    ISERIES = json.dumps(i_kl_series, ensure_ascii=False)
    ISERIES_Q = json.dumps(i_qd_series, ensure_ascii=False)
    D_ARR = json.dumps(dates, ensure_ascii=False)
    RANK_JS = r'''(function(){
 var box=document.getElementById('rankBox');
 if(!box) return;
 function nOf(b){
  var n=parseInt(b.getAttribute('data-n'),10);
  return n===0?0:(isNaN(n)?120:n);
 }
 function initN(){
  var b=document.querySelector('.rng.sel');
  return b?nOf(b):120;
 }
 function fmt(v){
  var neg=v<0, x=Math.abs(v).toFixed(2), p=x.split('.'), a=p[0], r='', c=0, k;
  for(k=a.length-1;k>=0;k--){ r=a[k]+r; c++; if(c%3===0&&k>0) r=','+r; }
  return (neg?'-':'+')+r+'.'+p[1];
 }
 function calc(n){
  var s=(n>0)?Math.max(0,DATA_D.length-n):0, arr=[], i, j, t, q, tot=0, q0=IDX_Q[0];
  for(i=1;i<IDX_NAMES.length;i++){
   q=IDX_Q[i]; t=0;
   for(j=s;j<q.length;j++){ t+=q[j]; }
   arr.push({nm:IDX_NAMES[i], v:t});
  }
  arr.sort(function(a,b){return b.v-a.v;});
  for(j=s;j<q0.length;j++){ tot+=q0[j]; }
  return {arr:arr, tot:tot, days:q0.length-s, n:n};
 }
 function render(n){
  var r=calc(n), mx=0, i, o, pos, w, col, h='';
  for(i=0;i<r.arr.length;i++){ mx=Math.max(mx,Math.abs(r.arr[i].v)); }
  if(mx<=0) mx=1;
  var scope=(r.n>0?('近 '+Math.min(r.n,DATA_D.length)+' 个交易日'):('全部 '+DATA_D.length+' 个交易日'));
  h+='<div style="font-size:12px;color:#7a8194;margin-bottom:9px">';
  h+='区间 <b style="color:#1a1d24">'+scope+'</b> ｜ 全市场合计 <b style="color:'+(r.tot>=0?'#d43d2a':'#1f9d7b')+'">'+fmt(r.tot)+' 亿</b>';
  h+=' <span style="color:#b0b6c4">（共 '+r.arr.length+' 个板块 · '+r.days+' 个交易日逐日加总）</span></div>';
  for(i=0;i<r.arr.length;i++){
   o=r.arr[i]; pos=o.v>=0; w=Math.abs(o.v)/mx*48; col=pos?'#d43d2a':'#1f9d7b';
   h+='<div style="display:flex;align-items:center;gap:8px;margin:3px 0;font-size:12px">';
   h+='<span style="width:96px;flex:none;text-align:right;color:#7a8194;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+o.nm+'</span>';
   h+='<div style="flex:1;position:relative;height:14px;background:#f2f4f8;border-radius:3px;overflow:hidden">';
   h+='<div style="position:absolute;left:50%;top:0;bottom:0;width:1px;background:#c9cfdd"></div>';
   h+='<div style="position:absolute;top:2px;bottom:2px;'+(pos?'left:50%':'right:50%')+';width:'+w.toFixed(2)+'%;background:'+col+';border-radius:2px"></div>';
   h+='</div>';
   h+='<span style="width:100px;flex:none;text-align:right;color:'+col+';font-variant-numeric:tabular-nums;white-space:nowrap">'+fmt(o.v)+' 亿</span>';
   h+='</div>';
  }
  box.innerHTML=h;
 }
 function boot(){
  render(initN());
  var btns=document.querySelectorAll('.rng');
  Array.prototype.forEach.call(btns,function(b){
   b.addEventListener('click',function(){ render(nOf(b)); });
  });
 }
 if(document.readyState==='loading'){document.addEventListener('DOMContentLoaded',boot);}else{boot();}
})();
'''

    KLN_JS = ("<script>\n"
      "var IDX_NAMES=" + INAMES + ";\n"
      "var IDX_K=" + ISERIES + ";\n"
      "var IDX_Q=" + ISERIES_Q + ";\n"
      "var DATA_D=" + D_ARR + ";\n"
      "(function(){\n"
      "function go(){\n"
      " var box=document.getElementById('klnBox');\n"
      " var sel=document.getElementById('klnCat');\n"
      " if(!window.echarts){var t=document.getElementById('klnTip');if(t)t.textContent='图表库未加载(需联网)';return;}\n"
      " var chart=echarts.init(box),cur=120,ci=0;\n"
      " function opt(n,ci2){\n"
      "  var K=IDX_K[ci2],Q=IDX_Q[ci2];\n"
      "  var s=(n>0)?Math.max(0,DATA_D.length-n):0,tail=(s>0)?{startValue:s,endValue:DATA_D.length-1} : null;\n"
      "  var ds=DATA_D.slice(s),ks=K.slice(s),qs=Q.slice(s);\n"
      "  return {axisPointer:{link:[{xAxisIndex:[0,1]}],lineStyle:{color:'#98a2b8'}},\n"
      "   grid:[{left:50,right:56,top:16,height:300,gridIndex:0},{left:50,right:56,bottom:28,height:120,gridIndex:1}],\n"
      "   xAxis:[{type:'category',data:ds,gridIndex:0,boundaryGap:true,axisLine:{lineStyle:{color:'#c9cfdd'}},axisLabel:{color:'#7a8194',fontSize:10},axisTick:{show:false}},\n"
      "          {type:'category',data:ds,gridIndex:1,axisLine:{lineStyle:{color:'#c9cfdd'}},axisLabel:{show:false},axisTick:{show:false}}],\n"
      "   yAxis:[{scale:true,gridIndex:0,position:'right',splitLine:{lineStyle:{color:'#eef1f6'}},axisLabel:{color:'#7a8194',fontSize:10}},\n"
      "          {scale:true,gridIndex:1,position:'right',splitLine:{show:false},axisLabel:{color:'#7a8194',fontSize:10,formatter:function(v){return v.toFixed(0);}}}],\n"
      "   dataZoom:[{type:'inside',xAxisIndex:[0,1],zoomOnMouseWheel:true,moveOnMouseMove:true,moveOnMouseWheel:true,throttle:50},\n"
      "             {type:'slider',xAxisIndex:[0,1],height:18,bottom:0,startValue:0,endValue:ds.length-1,brushSelect:false,textStyle:{fontSize:9,color:'#7a8194'},borderColor:'#dfe4ef',fillerColor:'rgba(99,126,188,0.12)',handleStyle:{color:'#b7c2da'},dataBackground:{lineStyle:{color:'#c9d2e4'},areaStyle:{color:'rgba(99,126,188,0.06)'}}}],\n"
      "   series:[\n"
      "    {name:'指数K线',type:'candlestick',data:ks.map(function(x){return x?[x.o,x.c,x.l,x.h]:[null,null,null,null];}),\n"
      "     itemStyle:{color:'#d43d2a',color0:'#1f9d7b',borderColor:'#d43d2a',borderColor0:'#1f9d7b'}},\n"
      "    {name:'该类ETF净买卖',type:'bar',xAxisIndex:1,yAxisIndex:1,data:qs,\n"
      "     itemStyle:{color:function(p){return p.value>=0?'#d43d2a':'#1f9d7b';}}}\n"
      "   ],\n"
      "   tooltip:{trigger:'axis',axisPointer:{type:'cross'},position:[8,338],confine:true,backgroundColor:'#2a2f3a',borderWidth:0,textStyle:{color:'#fff',fontSize:12},extraCssText:'white-space:nowrap;max-width:880px;padding:6px 10px;box-shadow:0 2px 8px rgba(0,0,0,.15);',\n"
      "    formatter:function(ps){var i=ps[0].dataIndex,g=i+s,d=DATA_D[g],k=K[g]||{},q=Q[g];\n"
      "     var s1='';\n"
      "     if(k.o!=null){s1=' &nbsp;开'+k.o+' 高'+k.h+' 低'+k.l+' 收'+k.c;}\n"
      "     var col=q>=0?'#f0705c':'#46c7a6';\n"
      "     var lab=(ci2===0)?'当日净买卖':'该类ETF净买卖';\n"
      "     return '<b>'+d+'</b>'+s1+' &nbsp;'+lab+' <b style=\"color:'+col+'\">'+(q>=0?'+':'')+q.toFixed(2)+' 亿</b>';}}\n"
      "  };\n"
      " }\n"
      " chart.setOption(opt(cur,ci));\n"
      " function upd(){chart.setOption(opt(cur,ci),true);}\n"
      " sel.addEventListener('change',function(){ci=sel.selectedIndex;chart.setOption(opt(cur,ci));});\n"
      " var btns=document.querySelectorAll('.rng');\n"
      " Array.prototype.forEach.call(btns,function(b){b.addEventListener('click',function(){\n"
      "   var _n2=parseInt(b.getAttribute('data-n'),10);cur=(_n2===0)?0:(isNaN(_n2)?120:_n2);\n"
      "   Array.prototype.forEach.call(btns,function(x){x.classList.toggle('sel',x===b);});\n"
      "   chart.setOption(opt(cur,ci));\n"
      " });});\n"
      "window.addEventListener('resize',function(){chart.resize();});\n"
      "}\n"
      "if(document.readyState==='loading'){document.addEventListener('DOMContentLoaded',go);}else{go();}\n"
      "})();\n" +
      RANK_JS +
      "</script>")

    _cat_opts = "".join('<option%s>%s</option>' % (' selected' if i == 0 else '', _n) for i, _n in enumerate(i_kl_names))

    # 内联 ECharts，避免 CDN 不可达导致图表空白。两级查找:
    #   ① 仓库根目录 echarts.min.js(随仓库分发, 云端 checkout 即得 -> 页面自包含, 零外链)
    #   ② 以上都没有才退回 CDN(最后兜底)
    # 不再硬编码任何本机绝对路径（含用户名），换机器 / CI 时不会误命中，也不泄露本机信息。
    _EC_CANDS = [
        os.path.join(OUT_DIR, "echarts.min.js"),
    ]
    ECHARTS_HTML = None
    for _p in _EC_CANDS:
        try:
            if os.path.exists(_p) and os.path.getsize(_p) > 100000:
                _ec = io.open(_p, encoding="utf-8").read().replace("</script>", "<\\/script>")
                ECHARTS_HTML = "<script>" + _ec + "</script>"
                print("ECharts 内联来源:", _p, flush=True)
                break
        except Exception as _e:
            print("ECharts 本地副本读取失败:", _p, str(_e)[:60], flush=True)
    if not ECHARTS_HTML:
        ECHARTS_HTML = ('<script src="https://registry.npmmirror.com/echarts/5.5.0/files/dist/echarts.min.js"></script>')
        print("ECharts 未找到本地副本, 回退 CDN 外链", flush=True)

    etfs = sorted(p["etfs"], key=lambda e: e["month"])

    # 同类(指数)品种归集
    agg = {}
    for e in etfs:
        agg.setdefault(e.get("cat", "其他"), []).append(e)
    cat_rows_html = ""
    cat_share_tot = cat_today_tot = cat_week_tot = cat_month_tot = cat_mv_tot = 0.0
    cat_s25_tot = cat_mv25_tot = 0.0
    for cat in sorted(agg, key=lambda c: sum(x["month"] for x in agg[c]), reverse=True):
        es = agg[cat]
        td = sum(x["today"] for x in es); wk = sum(x["week"] for x in es)
        mo = sum(x["month"] for x in es); sh = sum(x["last_share"] for x in es)
        mv = sum(x["last_share"] * x["last_nav"] for x in es)
        s25 = sum(x.get("s2025", 0) for x in es)
        mv25 = sum(x.get("s2025", 0) * x.get("n2025", 0) for x in es)
        cat_today_tot += td; cat_week_tot += wk; cat_month_tot += mo; cat_share_tot += sh; cat_mv_tot += mv
        cat_s25_tot += s25; cat_mv25_tot += mv25
        w = min(abs(mo) / 200 * 100, 100)
        cat_rows_html += f"""<tr>
          <td><b>{escape_html(cat)}</b><span class="code">{len(es)}只</span></td>
          <td class="r {cls(td)}">{fnum(td, True)}</td>
          <td class="r {cls(wk)}">{fnum(wk, True)}</td>
          <td class="r {cls(mo)}">{fnum(mo, True)}</td>
          <td class="r">{s25:,.1f}</td>
          <td class="r">{mv25:,.1f}</td>
          <td class="r">{sh:,.1f}</td>
          <td class="r">{mv:,.1f}</td>
          <td class="tbar"><div class="tbf"><i class="{cls(mo)}" style="width:{w:.0f}%"></i></div></td></tr>"""
    cat_tot_row = f"""<tr class="ctot"><td><b>全部（{len(agg)}类{len(etfs)}只）</b></td>
      <td class="r {cls(cat_today_tot)}">{fnum(cat_today_tot, True)}</td>
      <td class="r {cls(cat_week_tot)}">{fnum(cat_week_tot, True)}</td>
      <td class="r {cls(cat_month_tot)}">{fnum(cat_month_tot, True)}</td>
      <td class="r">{cat_s25_tot:,.1f}</td>
      <td class="r">{cat_mv25_tot:,.1f}</td>
      <td class="r">{cat_share_tot:,.1f}</td>
      <td class="r">{cat_mv_tot:,.1f}</td><td></td></tr>"""

    # 每日明细矩阵 (最近 HOT 个交易日)
    HOT = 10
    hot_dates = dates[-HOT:]
    hot_html = ""
    mcol_headers = "".join(f"<th>{d[5:]}</th>" for d in hot_dates)
    # 单元格: 上行=净买卖金额(亿元,单位'亿'), 下行=倍量(单位'倍'); v=净买卖(亿), m=倍量
    def _dm_cell(v, m=None):
        if v is None:
            return '<td class="r" style="color:var(--mut)">—</td>'
        if abs(v) < 0.001:
            top, c = "0", ""
        else:
            top, c = fnum(v, True) + "亿", cls(v)
        if m is None:
            return f'<td class="r {c}">{top}</td>'
        return f'<td class="r {c}">{top}<div class="mult {cls(m)}">{m:+.2f}倍</div></td>'
    for e in etfs:
        ser = e["series"]
        bz = e.get("s2025", 0) * 0.001 if e.get("s2025", 0) > 0 else None
        dsh = e.get("dsh", {})
        cells = ""
        for d in hot_dates:
            v = ser.get(d)
            m = (dsh.get(d) / bz) if (bz and dsh.get(d) is not None) else None
            cells += _dm_cell(v, m)
        hot_html += f"<tr><td><span class='mkt'>{e['m']}</span>{escape_html(e['n'])}{ftag(e['c'])}</td>{cells}</tr>"

    # 同类指数归集 · 每日净买卖明细矩阵 (最近 HOT 个交易日)
    cat_daily = {}
    for e in etfs:
        cat = e.get("cat", "其他")
        cd = cat_daily.setdefault(cat, {})
        for d in hot_dates:
            v = e["series"].get(d)
            if v is not None:
                cd[d] = cd.get(d, 0) + v
    cat_order = sorted(cat_daily, key=lambda c: sum(cat_daily[c].get(d, 0) for d in hot_dates), reverse=True)
    cat_mv = {}
    for e in etfs:
        g = e.get("cat", "其他")
        cat_mv[g] = cat_mv.get(g, 0) + e.get("s2025", 0) * e.get("n2025", 0)
    cat_daily_html = ""
    for cat in cat_order:
        cbz = cat_mv.get(cat, 0) * 0.001 if cat_mv.get(cat, 0) > 0 else None
        cells = ""
        for d in hot_dates:
            v = cat_daily[cat].get(d)
            m = (v / cbz) if (cbz and v is not None) else None
            cells += _dm_cell(v, m)
        cat_daily_html += f'<tr><td><b>{escape_html(cat)}</b><span class="code">{len(agg.get(cat, []))}只</span></td>{cells}</tr>'
    _all_mv = sum(e.get("s2025", 0) * e.get("n2025", 0) for e in etfs)
    _all_bz = _all_mv * 0.001 if _all_mv > 0 else None
    tot_cells = ""
    for d in hot_dates:
        v = dt.get(d, 0)
        m = (v / _all_bz) if _all_bz and v is not None else None
        tot_cells += _dm_cell(v, m)
    cat_daily_tot = f'<tr class="ctot"><td><b>全部（{len(cat_daily)}类{len(etfs)}只）</b></td>{tot_cells}</tr>'

    # 表格行
    rows = ""
    for e in etfs:
        rows += f"""<tr>
          <td><span class="mkt">{e['m']}</span>{escape_html(e['n'])}{ftag(e['c'])}<span class="code">{e['c']}</span></td>
          <td class="r">{e['last_share']:,.1f}</td>
          <td class="r {cls(e['today'])}">{fnum(e['today'], True)}</td>
          <td class="r {cls(e['week'])}">{fnum(e['week'], True)}</td>
          <td class="r {cls(e['month'])}">{fnum(e['month'], True)}</td>
          <td class="tbar"><div class="tbf"><i class="{cls(e['month'])}" style="width:{min(abs(e['month'])/50*100,100):.0f}%"></i></div></td>
        </tr>"""

    # 柱状图
    bars = ""
    for i, d in enumerate(chart):
        v = chart_total[i]
        h = abs(v) / maxabs * 100
        bars += f"""<div class="col" title="{d} {fnum(v,True)}亿">
          <div class="bwrap"><div class="bar {cls(v)}" style="height:{max(h,1.5):.1f}%"></div></div>
          <div class="dt">{d[5:]}</div></div>"""

    # 累计持仓金额折线(SVG)
    W, H = 860, 120; pad = 12
    mn, mx = min(y_mv), max(y_mv)
    if mn == mx: mn -= 1; mx += 1
    npts = len(y_mv)
    pts = []
    for i, v in enumerate(y_mv):
        x = pad + i * (W - 2 * pad) / (npts - 1)
        y = H - pad - (v - mn) / (mx - mn) * (H - 2 * pad)
        pts.append(f"{x:.1f},{y:.1f}")
    poly = " ".join(pts)
    last_pts = pts[-1]; last_v = y_mv[-1]
    ypt_json = json.dumps([{"d": y_dates[i], "v": round(y_mv[i], 1)} for i in range(len(y_dates))], ensure_ascii=False)
    JS_CODE = ("<script>\n"
      "const YPT=" + ypt_json + ";\n"
      "const C0=" + str(pad) + ",C1=" + str(W - pad) + ",CH=" + str(H) +
      ",CMN=" + f"{mn:.0f}" + ",CMX=" + f"{mx:.0f}" + ",WW=" + str(W) + ";\n"
      "const CVD=document.getElementById('cumSvg'),TIP=document.getElementById('cumTip'),G=document.getElementById('cumGuide');\n"
      "const N=YPT.length;\n"
      "function xAt(i){return C0+i*(C1-C0)/(N-1);}\n"
      "function FMT(v){return v.toFixed(1)+'亿';}\n"
      "CVD.addEventListener('mousemove',function(e){"
      "var r=CVD.getBoundingClientRect();"
      "var sx=(e.clientX-r.left)/r.width*WW;"
      "var i=Math.round((sx-C0)/(C1-C0)*(N-1));"
      "if(i<0)i=0;if(i>N-1)i=N-1;"
      "var p=YPT[i],gx=xAt(i);"
      "G.setAttribute('x1',gx);G.setAttribute('x2',gx);"
      "TIP.style.color='var(--accent)';"
      "TIP.textContent='日期 '+p.d+'  持仓金额 '+FMT(p.v);"
      "});"
      "CVD.addEventListener('mouseleave',function(){G.setAttribute('x1',-10);G.setAttribute('x2',-10);});"
      "\n</script>")

    def fmt_dt(d):
        return f"{d[0:4]}-{d[5:7]}-{d[8:10]}"

    html = f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>国家队ETF每日净买入卖出报表</title>
<style>
:root{{--bg:#f6f7f9;--card:#fff;--ink:#1a1d24;--mut:#7a8194;--line:#e6e8ee;
  --pos:#d43d2a;--posbg:#fdecea;--neg:#1f9d7b;--negbg:#e6f6f0;--accent:#2f54eb;}}
*{{box-sizing:border-box}}
body{{margin:0;font-family:-apple-system,"PingFang SC","Microsoft YaHei",system-ui,sans-serif;
  background:var(--bg);color:var(--ink);line-height:1.55;}}
.wrap{{max-width:980px;margin:0 auto;padding:20px 16px 60px;}}
.top{{display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:10px;margin-bottom:18px;}}
h1{{font-size:22px;margin:0;font-weight:700}}
.sub{{color:var(--mut);font-size:12px;margin-top:6px;display:flex;gap:14px;flex-wrap:wrap}}
.sub b{{font-weight:600;color:var(--accent)}}
.kpis{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));grid-auto-rows:1fr;gap:12px;margin-bottom:22px}}
.kpi{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;display:flex;flex-direction:column;justify-content:center}}
.kpi .lb{{font-size:12px;color:var(--mut)}}
.kpi .vl{{font-size:22px;font-weight:700;margin-top:4px;font-variant-numeric:tabular-nums}}
.kpi .tg{{font-size:12px;margin-top:2px}}
.pos{{color:var(--pos)}} .neg{{color:var(--neg)}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin-bottom:22px}}
.card h2{{font-size:15px;margin:0 0 14px;font-weight:600}}
.bars{{display:flex;align-items:flex-end;gap:4px;height:190px;}}
.col{{flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;height:100%;}}
.bwrap{{flex:1;display:flex;align-items:flex-end;width:100%}}
.bar{{width:100%;border-radius:3px;min-height:2px}}
.bar.pos{{background:linear-gradient(180deg,#e8836f,var(--pos))}}
.bar.neg{{background:linear-gradient(180deg,#54c2a6,var(--neg))}}
.dt{{font-size:10px;color:var(--mut);transform:rotate(-40deg);transform-origin:top right}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{font-size:12px;color:var(--mut);font-weight:500;text-align:right;padding:6px 8px;border-bottom:2px solid var(--line);white-space:nowrap}}
th:first-child{{text-align:left}}
td{{padding:7px 8px;border-bottom:1px solid var(--line);font-variant-numeric:tabular-nums}}
td.r{{text-align:right;white-space:nowrap}}
tr.ctot td{{border-top:2px solid var(--ink);font-weight:600;background:#fafbfc}}
td.zero{{color:var(--mut);font-size:11px;opacity:.65}}
.mult{{display:block;font-size:10px;opacity:.72;margin-top:1px;font-family:var(--font-mono,"SFMono-Regular",Consolas,monospace)}}
table.matrix td{{font:11px var(--font-mono,"SFMono-Regular",Consolas,monospace)}}
table.matrix th{{font-size:10px}}
.mkt{{display:inline-block;font-size:10px;color:var(--mut);border:1px solid var(--line);border-radius:4px;padding:0 4px;margin-right:6px}}
.code{{color:var(--mut);font-size:11px;margin-left:6px}}
.tag{{display:inline-block;font-size:10px;line-height:15px;color:#b45309;background:#fff7ed;
  border:1px solid #fed7aa;border-radius:4px;padding:0 5px;margin-left:6px;vertical-align:1px;font-weight:500;white-space:nowrap}}
.tagb{{color:#1d4ed8;background:#eff6ff;border-color:#bfdbfe}}
.tbar{{width:90px}}.tbf{{background:#f0f1f4;border-radius:3px;height:10px;overflow:hidden}}
.tbf i{{display:block;height:100%;border-radius:3px}}
.tbf i.pos{{background:var(--pos)}}.tbf i.neg{{background:var(--neg)}}

.note{{color:var(--mut);font-size:12px;margin-top:18px;line-height:1.7}}
.svgwrap{{margin-top:16px}}
.chartbox{{position:relative}}
.cumtip{{display:block;text-align:center;margin:2px auto 4px;min-height:18px;
  border:none;box-shadow:none;background:transparent;color:var(--mut);
  font:12px "SFMono-Regular",Consolas,monospace;pointer-events:none;white-space:nowrap}}
svg text{{fill:var(--mut);font-size:10px}}
.foot{{color:var(--mut);font-size:11px;text-align:center;margin-top:30px}}
.klnctl{{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px}}
.klnctl .kp{{font-size:12px;color:var(--mut)}}
.rng{{border:1px solid var(--line);background:#fff;border-radius:8px;padding:5px 12px;font-size:12px;cursor:pointer;color:var(--ink)}}
.rng:hover{{border-color:var(--accent)}}
.rng.sel{{background:var(--accent);color:#fff;border-color:var(--accent)}}
.kpt{{font-size:12px;color:var(--mut);margin-left:auto;font-family:SFMono-Regular,Consolas,monospace}}
.ktip{{font-size:12px;color:var(--mut);text-align:center;margin-top:8px;min-height:16px;font-family:SFMono-Regular,Consolas,monospace}}
</style>
  {ECHARTS_HTML}
  </head><body><div class="wrap">
  <div class="top">
    <div>
      <h1>国家队ETF · 每日净买入 / 净卖出</h1>
      <div class="sub"><span>数据日 <b>{escape_html(p['data_asof'])}</b></span>
        <span>更新时间 <b>{escape_html(p['updated'])}</b></span>
        <span>标的口径：汇金系相关宽基ETF篮子 共 <b>{len(etfs)}</b> 只 · 全市场净申赎</span>
        <button onclick="location.reload()" title="点击重新加载，并按沪深两市同步日期展示最新数据"
           style="cursor:pointer;margin-left:6px;padding:4px 12px;border:1px solid var(--line,#ccc);border-radius:6px;background:var(--card,#fff);color:var(--txt,#111);font-size:12px">🔄 刷新数据</button></div>
      <div style="color:var(--mut);font-size:12px;margin-top:6px">汇金 2025 年 12 月 31 日持仓约 1.5 万亿、本表跟踪 {len(etfs)} 只。本数据无法准确检测汇金实际持仓，仅反映这 {len(etfs)} 只 ETF 的每日净申购赎回情况。</div>
    </div>
  </div>
  <div class="kpis">
    <div class="kpi"><div class="lb">最新交易日 {fmt_dt(dates[-1])} 合计</div>
      <div class="vl {cls(today_tot)}">{fnum(today_tot, True)}</div>
      <div class="tg {cls(today_tot)}">{ "净买入" if today_tot>=0 else "净卖出" } 亿元</div></div>
    <div class="kpi"><div class="lb">近一周累计 (5个交易日)</div>
      <div class="vl {cls(week_tot)}">{fnum(week_tot, True)}</div>
      <div class="tg {cls(week_tot)}">{ "净买入" if week_tot>=0 else "净卖出" } 亿元</div></div>
    <div class="kpi"><div class="lb">近一月累计 (21个交易日)</div>
      <div class="vl {cls(month_tot)}">{fnum(month_tot, True)}</div>
      <div class="tg {cls(month_tot)}">{ "净买入" if month_tot>=0 else "净卖出" } 亿元</div></div>
    <div class="kpi"><div class="lb">近半年累计 (约120个交易日)</div>
      <div class="vl {cls(halfyear_tot)}">{fnum(halfyear_tot, True)}</div>
      <div class="tg {cls(halfyear_tot)}">{ "净买入" if halfyear_tot>=0 else "净卖出" } 亿元</div></div>
    <div class="kpi"><div class="lb">近一年累计 (约242个交易日)</div>
      <div class="vl {cls(oneyear_tot)}">{fnum(oneyear_tot, True)}</div>
      <div class="tg {cls(oneyear_tot)}">{ "净买入" if oneyear_tot>=0 else "净卖出" } 亿元</div></div>
    <div class="kpi"><div class="lb">今年以来累计 ({cur_year}年{len(ytd_days)}个交易日)</div>
      <div class="vl {cls(ytd_tot)}">{fnum(ytd_tot, True)}</div>
      <div class="tg {cls(ytd_tot)}">{ "净买入" if ytd_tot>=0 else "净卖出" } 亿元</div></div>
  </div>
  <div class="card"><h2>全部品种 · 每日净买卖合计（近21个交易日，亿元）</h2>
    <div class="bars">{bars}</div>
    <div class="svgwrap"><div class="chartbox">
      <h2 style="font-size:14px">累计持仓金额 · 最长 {len(y_dates)}个交易日（亿元 · 悬停查看当日）</h2>
      <div id="cumTip" class="cumtip"></div>
      <svg id="cumSvg" viewBox="0 0 {W} {H}" width="100%" height="{H}" preserveAspectRatio="none">
        <line id="cumGuide" x1="-10" y1="{pad}" x2="-10" y2="{H-pad}" stroke="#c0c6d4" stroke-width="1" stroke-dasharray="3 3"/>
        <polyline points="{poly}" fill="none" stroke="#2f54eb" stroke-width="1.5"/>
        <circle cx="{last_pts.split(',')[0]}" cy="{last_pts.split(',')[1]}" r="3.5" fill="#2f54eb"/>
        <text x="{pad}" y="{pad-4}">{mn:.0f}</text>
        <text x="{W-pad}" y="{pad-4}" text-anchor="end">{mx:.0f}亿</text>
        <text x="{W-pad}" y="{H-6}" text-anchor="end">{last_v:,.0f}亿</text>
      </svg>
    </div></div>
    {JS_CODE}
  </div>
  <div class="card"><h2>指数K线 × 每日净买入联动（上K线 下当日净买卖 · 悬停查看）</h2>
    <div class="klnctl"><span class="kp">指数</span>
      <select id="klnCat" style="font-size:12px;padding:2px 6px;border:1px solid var(--line);border-radius:4px;background:var(--bg);color:var(--fg)">{_cat_opts}</select>
      <span class="kp" style="margin-left:10px">显示范围</span>
      <button class="rng" data-n="60">60日</button>
      <button class="rng sel" data-n="120">120日</button>
      <button class="rng" data-n="250">250日</button>
      <button class="rng" data-n="0">全部</button>
      <span class="kpt" id="klnDate"></span></div>
    <div id="klnBox" style="height:520px;width:100%"></div>
    <div id="klnTip" class="ktip">默认显示中证全指K线与全市场国家队ETF每日净买入联动；下拉可切换至{len(CAT_INDEX)}类细分指数（沪深300、创业板等）K线 × 该类ETF净买入合计。上方为指数日K线，下方为当日净买入金额（亿元），鼠标悬停联动查看；红=净买入、绿=净卖出。</div>
    <div style="margin-top:18px;border-top:1px dashed var(--line);padding-top:14px">
      <div style="font-size:13px;font-weight:600;margin-bottom:2px">各板块累计净买入排行 <span style="font-size:11px;font-weight:400;color:var(--mut)">（按所选区间累计降序 · 排最上 = 综合入场金额最多）</span></div>
      <div style="font-size:11px;color:var(--mut);margin-bottom:10px">随上方「显示范围」按钮联动：60日 / 120日 / 250日 / 全部</div>
      <div id="rankBox"></div>
      <div style="font-size:11px;color:var(--mut);line-height:1.7;margin-top:10px"><b>口径：</b>把所选显示范围内每个指数类别的「每日净买卖」逐日加总，按累计金额从高到低排列，用来横向比较哪个板块综合加下来入场金额更多。数据直接取自上方联动图同一份逐日金额（<code>cat_total</code> 逐类 / <code>date_total</code> 全市场），不新增取数、不改任何计算口径；红 = 净买入（资金流入）、绿 = 净卖出，条长以本表最大绝对值为基准、零轴居中，仅用于板块间相对比较。<br><b style="color:var(--accent)">说明：本模块为本仓独有增强，原版无此排行，二者不一致属预期；与原版对表校验时排除本项对照</b>（见 README 附录 I 与巡检任务 6.5.8）。</div>
    </div>
  </div>
  <div class="card"><h2>分品种汇总 · 近一周 / 近一月净买入卖出（亿元，按近一月降序）</h2>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>品种</th><th>最新份额(亿)</th><th>今日</th><th>近一周</th><th>近一月</th><th style="width:90px">近一月强度</th></tr></thead>
      <tbody>{rows}</tbody>
    </table></div>
    <div style="color:var(--mut);font-size:11px;margin-top:8px">标注「{BASE_TAG}」的 {len(BASE23)} 只为基础口径（2025 年报披露的汇金系持仓宽基ETF）；标注「{NEW_TAG}」的 {len(NEW_TAG_CODES)} 只为救市专项口径中新增纳入、此前报表未跟踪的品种；本次口径切换点见 README 附录 K。</div>
  </div>
  <div class="card"><h2>同类指数品种归集汇总（亿元）</h2>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>指数类别</th><th>今日</th><th>近一周</th><th>近一月</th><th>2025-12-31份额(亿)</th><th>2025-12-31市值(亿)</th><th>最新份额(亿)</th><th>持仓市值(亿)</th><th style="width:90px">近一月强度</th></tr></thead>
      <tbody>{cat_rows_html}{cat_tot_row}</tbody>
    </table></div>
    <div style="color:var(--mut);font-size:11px;margin-top:8px">将跟踪同一指数的多只 ETF 归集加总，观察该指数方向上的整体净买卖。</div>
  </div>
  <div class="card"><h2>每日净买卖明细 · 金额 / 倍量份额（最近{HOT}个交易日，亿元 · 可横向滚动）</h2>
    <div style="overflow-x:auto"><table class="matrix">
      <thead><tr><th>品种</th>{mcol_headers}</tr></thead>
      <tbody>{hot_html}</tbody>
    </table></div>
    <div style="color:var(--mut);font-size:11px;margin-top:8px">每格上行=净买卖金额(亿元)，下行=倍量份额（＝当日净申购(赎回)份额 ÷ 该 ETF 自身 2025-12-31 份额×0.1%，单位倍）。
      「0」表示该日基金份额较上一交易日无变化（申购与赎回相抵），属正常现象，非数据缺失。正值=净买入(红)，负值=净卖出(绿)。</div>
  </div>
  <div class="card"><h2>同类指数品种 · 每日净买卖明细（金额 / 倍量金额 · 最近{HOT}个交易日，亿元 · 可横向滚动）</h2>
    <div style="overflow-x:auto"><table class="matrix">
      <thead><tr><th>指数类别</th>{mcol_headers}</tr></thead>
      <tbody>{cat_daily_html}{cat_daily_tot}</tbody>
    </table></div>
    <div style="color:var(--mut);font-size:11px;margin-top:8px">将跟踪同一指数的多只 ETF 归集加总，观察该指数方向每日整体净买卖。
      每格上行=净买卖金额(亿元)，下行=倍量金额（＝当日净买入金额 ÷ 该类 2025-12-31 市值×0.1%，单位倍）。
      「0」表示该类当日份额较前日无净变化。正值=净买入(红)，负值=净卖出(绿)。</div>
  </div>
  <div class="note"><b>口径说明：</b>净买入 / 净卖出 =（当日基金份额 − 前一日基金份额）× 当日单位净值，正值=净买入(净申购)，负值=净卖出(净赎回)。
  份额源：上交所 / 深交所交易所官方披露（经自建 a-stock-api Worker 代理）；净值源：天天基金（经同一 Worker）。
  份额为日终口径，当日净买卖反映当日申购赎回结果。数据仅为展示来源，不构成投资建议。</div>
  <div class="foot">Generated by TRAE · 每日运行脚本自动刷新</div>
  {KLN_JS}
</div></body></html>"""
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write(html)
    print("报表已生成:", REPORT)

if __name__ == "__main__":
    p = main()
    if not p:
        sys.exit(1)
    with open(os.path.join(OUT_DIR, "_etf_data.json"), "w", encoding="utf-8") as f:
        json.dump(p, f, ensure_ascii=False)
    render_html(p)
    # ASCII 文件名副本, 供外部脚本(如 Windows 计划任务)在 ANSI 编码下安全引用
    shutil.copyfile(REPORT, os.path.join(OUT_DIR, "report_latest.html"))
    print("数据更新完成, 最新数据日:", p["data_asof"], "共", len(p["dates"]), "个交易日")