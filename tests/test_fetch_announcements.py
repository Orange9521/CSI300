from decimal import Decimal
from fetch_announcements import extract_amount, deduplicate


def test_exact_cash_dividend_not_per_share():
    text = '每股派发现金红利28.02423元（含税），共计派\n发现金红利35,032,574,305.19元。'
    amount, quote = extract_amount('dividend', text)
    assert amount == Decimal('35032574305.19')
    assert '35032' not in quote  # original comma-formatted quote retained
    assert extract_amount('dividend', '每股现金红利28.02423元')[0] is None


def test_partial_cancellation_corresponding_cost_only():
    text = '回购支付总金额400,080,671.30元。本次注销3,000,000股，对应回购金额为147,268,958.76元。'
    assert extract_amount('buyback_cancelled', text)[0] == Decimal('147268958.76')
    assert extract_amount('buyback_cancelled', '计划回购10亿元；已回购3亿元，拟注销。')[0] is None


def test_conflicting_cash_totals_not_selected():
    text = '现金红利总额为100万元。共计派发现金红利200万元。'
    assert extract_amount('dividend', text)[0] is None


def test_duplicate_economic_event_excluded():
    rows = [{'经济事件ID': 'a', '是否纳入已取得小计': True}, {'经济事件ID': 'a', '是否纳入已取得小计': True}]
    deduplicate(rows)
    assert all(not r['是否纳入已取得小计'] for r in rows)


def test_repeated_cninfo_page_rejected(monkeypatch, tmp_path):
    import pytest
    from fetch_announcements import query_batch
    class Response:
        content = b'{}'
        def raise_for_status(self):
            pass
        def json(self):
            return {'totalAnnouncement': 2, 'announcements': [{'secCode': '000001', 'announcementId': 'same'}]}
    monkeypatch.setattr('fetch_announcements.requests.post', lambda *a, **k: Response())
    with pytest.raises(ValueError, match='分页重复'):
        query_batch([{'code': '000001', 'orgId': 'test'}], 'test', '', tmp_path)


def test_all_bought_shares_cancelled_cost_allowed_but_partial_rejected():
    text = '累计回购股份数量为100,000股，支付的总金额为2,000,000.00元。公司办理完毕上述100,000股回购股份注销手续。'
    assert extract_amount('buyback_cancelled', text)[0] == Decimal('2000000.00')
    assert extract_amount('buyback_cancelled', text.replace('上述100,000股', '上述70,000股'))[0] is None


def test_eastmoney_cancellation_wording_and_units():
    text = '本次注销已回购股份71,452,577股。累计回购公司股份数量为71,452,577股，已使用资金总额为99,988.54万元（不含交易费用）。'
    assert extract_amount('buyback_cancelled', text)[0] == Decimal('999885400')
    assert extract_amount('buyback_cancelled', text.replace('不含交易费用', '含交易费用'))[0] is None


def test_commissions_and_stamp_duty_are_not_shareholder_return():
    text = '累计回购公司股份44,804,006股，支付的总金额为430,266,583.09元（含印花税、交易佣金等交易费用）。本次注销的股份为44,804,006股。'
    assert extract_amount('buyback_cancelled', text)[0] is None
    assert extract_amount('buyback_cancelled', text.replace('（含印花税', '（不含印花税'))[0] == Decimal('430266583.09')
