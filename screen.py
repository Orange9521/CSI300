"""Auditable CSI300 screening. No missing history is interpreted as zero."""
import argparse
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from io import BytesIO
import json
from pathlib import Path
import re

import pandas as pd
import requests

SOURCES = {
    '中证指数成分股': 'https://oss-ch.csindex.com.cn/static/html/csindex/public/uploads/file/autofile/cons/000300cons.xls',
    '巨潮公告查询': 'https://www.cninfo.com.cn/new/hisAnnouncement/query',
    '东方财富分红记录（辅助）': 'https://datacenter-web.eastmoney.com/api/data/v1/get?reportName=RPT_SHAREBONUS_DET&columns=ALL&pageSize=1&pageNumber=1',
}
KINDS = ('dividend', 'buyback_cancelled', 'equity_financing')
FINANCING = {'IPO', 'SEO', 'rights', 'preferred', 'convertible_conversion', 'other_equity'}
COLUMNS = ['股票代码', '股票名称', '成分股日期', '累计现金分红（元）', '累计回购注销（元）', '累计股权融资（元）', '股东回报（元）', '净股东回报（元）', '股东回报/融资比率', '状态', '数据来源']


def probe_sources():
    """Probe before any computation; retain official roster bytes only on success."""
    rows, roster = [], None
    for name, url in SOURCES.items():
        row = {'来源': name, 'URL': url, '检查时间UTC': datetime.now(timezone.utc).isoformat()}
        try:
            if name == '巨潮公告查询':
                r = requests.post(url, data={'pageNum': 1, 'pageSize': 1, 'column': 'szse', 'tabName': 'fulltext', 'stock': '', 'searchkey': '', 'secid': '', 'plate': '', 'category': '', 'trade': '', 'seDate': '', 'sortName': '', 'sortType': '', 'isHLtitle': 'true'}, timeout=25)
            else:
                r = requests.get(url, timeout=25)
            r.raise_for_status()
            row['状态'] = 'HTTP成功；不代表历史覆盖完整'
            if name == '中证指数成分股':
                pd.read_excel(BytesIO(r.content))  # reject HTML masquerading as XLS
                roster = r.content
            elif name.startswith('东方财富'):
                payload = r.json()
                if not payload.get('success') or not payload.get('result'):
                    raise ValueError('接口未返回有效数据')
        except Exception as exc:
            row['状态'] = '不可用: ' + str(exc)
        rows.append(row)
    return rows, roster


def code(value):
    s = str(value).strip()
    if re.fullmatch(r'\d{1,6}(\.0)?', s):
        return s.split('.')[0].zfill(6)
    raise ValueError(f'非法股票代码: {s}')


def read_roster(data, as_of):
    frame = pd.read_excel(data, dtype=str)
    frame.columns = [str(c).strip() for c in frame.columns]
    code_col = next((c for c in ('成份券代码Constituent Code', '成分券代码', '证券代码', '股票代码') if c in frame), None)
    name_col = next((c for c in ('成份券名称Constituent Name', '成分券名称', '证券简称', '股票名称') if c in frame), None)
    date_col = next((c for c in ('日期Date', '日期', '数据日期') if c in frame), None)
    if not code_col or not name_col or not date_col:
        raise ValueError('官方文件缺少代码、名称或数据日期，拒绝推测列含义')
    date_values = frame[date_col].str.replace(r'^(\d{4})(\d{2})(\d{2})$', r'\1-\2-\3', regex=True)
    dates = pd.to_datetime(date_values, errors='raise').dt.date
    if dates.nunique() != 1 or dates.iloc[0] > as_of:
        raise ValueError('成分股日期不一致或晚于统计截止日')
    result = [{'code': code(c), 'name': str(n), 'date': d.isoformat()} for c, n, d in zip(frame[code_col], frame[name_col], dates)]
    if len(result) != 300 or len({r['code'] for r in result}) != 300:
        raise ValueError('官方名单必须包含300个唯一代码')
    return result


def iso(value):
    return date.fromisoformat(value)


def validate_ledger(ledger, as_of):
    seen = set()
    for e in ledger.get('events', []):
        e['code'] = code(e['code'])
        key = (e['code'], e['kind'], e['event_id'])
        if key in seen:
            raise ValueError(f'重复经济事件: {key}')
        seen.add(key)
        if e['kind'] not in KINDS or not e.get('source') or not e.get('evidence'):
            raise ValueError('事件必须有合法类别、来源及凭证说明')
        iso(e['date'])
        allowed_statuses = {'completed', 'pending', 'cancelled'} if e['kind'] == 'buyback_cancelled' else {'completed', 'pending'}
        if e.get('status') not in allowed_statuses:
            raise ValueError('未知事件状态，不能静默忽略')
        try:
            amount = Decimal(str(e['amount_cny']))
        except InvalidOperation as exc:
            raise ValueError('金额非法') from exc
        if not amount.is_finite() or amount < 0 or amount != amount.quantize(Decimal('.01')):
            raise ValueError('金额必须为非负有限人民币元，精确到分')
        if e['kind'] == 'equity_financing' and e.get('method') not in FINANCING:
            raise ValueError('融资方式不适用，不能计入债务或二级市场交易')
        if e['kind'] == 'buyback_cancelled' and e.get('status') == 'cancelled':
            if not e.get('cancellation_source'):
                raise ValueError('已注销回购必须有注销凭证')
    for c in ledger.get('coverage', []):
        c['code'] = code(c['code'])
        if c['kind'] not in KINDS or not c.get('source') or not c.get('scope'):
            raise ValueError('覆盖声明必须有类别、来源和跨市场统计范围')
        if iso(c['through']) < iso(c['first_listing_date']):
            raise ValueError('覆盖截止日早于首次上市日')
    for issuer in {c['code'] for c in ledger.get('coverage', [])}:
        claims = [c for c in ledger['coverage'] if c['code'] == issuer]
        if len({c['first_listing_date'] for c in claims}) != 1 or len({c['scope'] for c in claims}) != 1:
            raise ValueError('同一发行人各类别首次上市日期与统计范围必须一致')
    return ledger


def calculate(roster, ledger, as_of):
    validate_ledger(ledger, as_of)
    rows = []
    for stock in roster:
        amounts, missing, sources = {}, [], {SOURCES['中证指数成分股']}
        for kind in KINDS:
            claims = [c for c in ledger.get('coverage', []) if c['code'] == stock['code'] and c['kind'] == kind and c.get('complete') is True and iso(c['through']) >= as_of]
            if len(claims) != 1:
                amounts[kind] = None
                missing.append(kind)
                continue
            claim = claims[0]
            sources.add(claim['source'])
            total = Decimal('0')
            for e in ledger.get('events', []):
                initial_ipo = kind == 'equity_financing' and e.get('method') == 'IPO'
                if e['code'] != stock['code'] or e['kind'] != kind or iso(e['date']) > as_of or (iso(e['date']) < iso(claim['first_listing_date']) and not initial_ipo):
                    continue
                accepted = e.get('status') == ('cancelled' if kind == 'buyback_cancelled' else 'completed')
                if accepted:
                    total += Decimal(str(e['amount_cny']))
                    sources.add(e['source'])
                    if kind == 'buyback_cancelled':
                        sources.add(e['cancellation_source'])
            amounts[kind] = total
        d, b, f = (amounts[k] for k in KINDS)
        ret = d + b if d is not None and b is not None else None
        net = ret - f if ret is not None and f is not None else None
        ratio = ret / f if ret is not None and f is not None and f > 0 else None
        rows.append(dict(zip(COLUMNS, [stock['code'], stock['name'], stock['date'], d, b, f, ret, net, ratio, '缺失: ' + ', '.join(missing) if missing else ('完整；融资为0，比率未定义' if f == 0 else '完整'), '\n'.join(sorted(sources))])))
    all_rows = pd.DataFrame(rows, columns=COLUMNS)
    chosen = [r for r in rows if r['净股东回报（元）'] is not None and r['净股东回报（元）'] > 0]
    chosen.sort(key=lambda r: r['净股东回报（元）'], reverse=True)
    return all_rows, pd.DataFrame(chosen, columns=COLUMNS)


def export(path, all_rows, selected, diagnostics, notes, ledger):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        for sheet, frame in [('筛选结果', selected), ('全量与缺失', all_rows), ('数据源检查', pd.DataFrame(diagnostics)), ('统计说明', pd.DataFrame({'说明': notes})), ('逐笔凭证', pd.DataFrame(ledger.get('events', []))), ('覆盖声明', pd.DataFrame(ledger.get('coverage', [])))]:
            # Excel numeric cells are limited to 15 significant digits; exact money stays text.
            frame = frame.map(lambda v: format(v, 'f') if isinstance(v, Decimal) else v)
            frame.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 25
            for row in ws:
                for cell in row:
                    if isinstance(cell.value, str):
                        cell.data_type = 's'  # prevent spreadsheet formula injection


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--as-of', type=iso, default=date.today())
    p.add_argument('--constituents', type=Path, help='官方成分股xls/xlsx；保留日期列')
    p.add_argument('--ledger', type=Path, help='经核实的逐笔事件及完整性声明JSON')
    p.add_argument('--output', type=Path, default=Path('output/csi300_screen.xlsx'))
    args = p.parse_args(argv)
    if args.as_of > date.today():
        p.error('截止日不能晚于今天')
    diagnostics, raw = probe_sources()
    notes = [f'统计截止日: {args.as_of}', '金额统一为人民币元；精确金额以文本保存。', '仅完整历史参与筛选；空白不是0。融资为0时比率未定义。', '本程序不自动将公告搜索结果认定为完整历史。', '首次上市指同一发行人最早公开上市，覆盖A/H及其他适用市场；排除债务和二级市场交易。']
    roster, ledger = [], {'events': [], 'coverage': []}
    try:
        if args.constituents:
            roster = read_roster(args.constituents, args.as_of)
            notes.append(f'使用本地官方文件: {args.constituents}；未验证该文件为最新版本。')
        elif raw:
            roster = read_roster(BytesIO(raw), args.as_of)
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.with_suffix('.constituents.xls').write_bytes(raw)
            notes.append('使用官方当前发布文件；数据日期见成分股日期列，并保留原文件。')
        else:
            notes.append('成分股名单不可用，未进行正式计算；空筛选表不表示没有符合条件的公司。')
        if args.ledger:
            ledger = json.loads(args.ledger.read_text(encoding='utf-8'))
        all_rows, selected = calculate(roster, ledger, args.as_of)
    except (ValueError, KeyError, OSError) as exc:
        notes.append('输入校验失败，未计算: ' + str(exc))
        all_rows = selected = pd.DataFrame(columns=COLUMNS)
        roster = []
    export(args.output, all_rows, selected, diagnostics, notes, ledger)
    print(json.dumps({'output': str(args.output), 'constituents': len(roster), 'selected': len(selected), 'complete': sum(str(x).startswith('完整') for x in all_rows['状态']), 'notes': notes}, ensure_ascii=False, indent=2))
    return 0 if roster and all(str(x).startswith('完整') for x in all_rows['状态']) else 2


if __name__ == '__main__':
    raise SystemExit(main())
