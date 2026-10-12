from datetime import date
from decimal import Decimal
import copy
import pytest
from fetch_public import financing_amount, summarize, fetch_chunk


def test_ipo_unit_validation():
    amount, note = financing_amount('ipo', {'TOTAL_RAISE_FUNDS': '22.44385', 'ISSUE_NUM': '7150', 'ISSUE_PRICE': '31.39'})
    assert amount == Decimal('2244385000')
    assert '亿元' in note
    amount, note = financing_amount('ipo', {'TOTAL_RAISE_FUNDS': '2244385000', 'ISSUE_NUM': '7150', 'ISSUE_PRICE': '31.39'})
    assert amount is None


def test_old_share_transfer_and_missing_not_assumed():
    assert financing_amount('ipo', {'TOTAL_RAISE_FUNDS': 1, 'ORIG_TRANSFER_NUM': 5})[0] is None
    assert financing_amount('seo', {'TOTAL_RAISE_FUNDS': None, 'ISSUE_NUM': 100, 'ISSUE_PRICE': 20})[0] is None


def datasets():
    return {'ipo': [], 'seo': [], 'rights': [], 'dividends': [], 'buybacks': []}


def test_no_records_does_not_mean_zero():
    rows, *_ = summarize([{'code': '000001', 'name': 'test', 'date': '2026-10-09'}], datasets(), date(2026, 10, 12))
    assert rows[0]['已取得三类融资小计（元；非完整股权融资）'] is None
    assert rows[0]['累计股权融资（元）'] is None


def test_dividends_and_buybacks_never_inferred():
    data = datasets()
    common = {'_source': 'https://example.org', '_raw_sha256': 'test'}
    data['dividends'] = [dict(common, SECURITY_CODE='000001', ASSIGN_PROGRESS='实施分配', EX_DIVIDEND_DATE='2026-06-01', TOTAL_SHARES=1000000, PRETAX_BONUS_RMB=10)]
    data['buybacks'] = [dict(common, DIM_SCODE='000001', REPURAMOUNT=999999, REPURPROGRESS='007')]
    summary, _, dividends, buybacks = summarize([{'code': '000001', 'name': 'test', 'date': '2026-10-09'}], data, date(2026, 10, 12))
    assert summary[0]['已实施分红记录数'] == 1
    assert dividends[0]['实际现金分红总额（元）'] is None
    assert buybacks[0]['已核实注销金额（元）'] is None
    assert summary[0]['净股东回报（元）'] is None


def test_financing_duplicates_and_future_records():
    data = datasets()
    r = {'SECURITY_CODE': '000001', 'FINANCE_CODE': 'abc', 'ISSUE_LISTING_DATE': '2026-06-01', 'TOTAL_RAISE_FUNDS': Decimal('123.45'), '_source': 'https://example.org', '_raw_sha256': 'test'}
    data['seo'] = [r, dict(r, FINANCE_CODE='future', ISSUE_LISTING_DATE='2027-01-01')]
    roster = [{'code': '000001', 'name': 'test', 'date': '2026-10-09'}]
    summary, rows, *_ = summarize(roster, data, date(2026, 10, 12))
    assert summary[0]['已取得增发募集金额小计（元）'] == Decimal('123.45')
    assert len(rows) == 1
    data['seo'].append(copy.deepcopy(r))
    with pytest.raises(ValueError, match='重复融资'):
        summarize(roster, data, date(2026, 10, 12))


def test_fetch_rejects_incomplete_pagination(monkeypatch, tmp_path):
    class Response:
        text = '{"success":true,"result":{"pages":1,"count":2,"data":[{"SECURITY_CODE":"000001"}]}}'
        content = text.encode()
        url = 'https://example.org'
        def raise_for_status(self):
            pass
    monkeypatch.setattr('fetch_public.requests.get', lambda *a, **k: Response())
    with pytest.raises(ValueError, match='分页行数'):
        fetch_chunk('ipo', ['000001'], tmp_path)
