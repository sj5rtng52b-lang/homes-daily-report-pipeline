#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Homes店长大盘 → MTD营收对比表 自动填表脚本
用法: python fill_template.py <raw.json> <template.xlsx> <out.xlsx>

raw.json 结构: {"date":"YYYYMMDD","cat":[...],"loss":[...],"overall":[[...]],"channel":[[...]]}
  cat     = 品类看板-部门   (widget 497734, mtd, 门店19005)
  loss    = 商品管理-部门   (widget 497737)
  overall = 门店看板-整体表现(widget 497731)
  channel = 门店看板-分渠道成交(widget 497780)

口径说明（全部经 20260831 八月模板逐格对账验证，误差=0）:
  部门行 E..AM 直接取 API；H/J/K/AN/AP/AQ/AR..BE 为派生列:
    H=E/E18  J=J18*I  K=E/J  AN=AF/AF18  AP=AP18*AO  AQ=AF/AP
    AR=N-AF(线上=整体-POS)  AS=AR/天数  AT 由 cur/(1+环比) 倒推上期计算
    AV=Q-AJ  AW/AY 同理倒推  AX=AS/AV  AZ=AR/AR18
    BA=AI BB=AU BC=AA BD=BB-BC(端外) BE=上期端外占比
  总计行: E18=ΣE  F18/G18=ΣE/Σ(E/(1+环比/同比))-1  L18/M18=营收加权平均
    N18..V18 取整体表现  X18..AM18 取分渠道成交  AR18=N18-AF18
  预算列(I/J/K整体 + AO/AP/AQ的POS) 从《预算分解.xlsx》读取部门未税预算：
    单日=当日预算，MTD=本月1日至统计日累计；占比=部门/总计；达成率=未税营收/预算。
"""
import json, sys, io, os
from openpyxl import load_workbook
from budget_lib import BudgetJsonReader

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

STORE = '青岛盒马鲜生卓悦里店'
BUDGET_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '大盘取数看板_公共资产', 'store', 'budget_cat.json')

def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

def prev_val(cur, mom):
    """上期值 = 本期/(1+环比)"""
    if cur is None or mom is None or mom <= -1:
        return None
    return cur / (1 + mom)

def chain_mom(cur, prev):
    if cur is None or prev is None or prev == 0:
        return None
    return cur / prev - 1

def main(raw_path, tpl_path, out_path, period='mtd', budget_path=None):
    with open(raw_path, encoding='utf-8') as f:
        raw = json.load(f)
    date = str(raw['date'])
    is_1d = (period == '1d')
    # 单日：日均=当日值，天数按 1；月度累计：天数=当月已过几天
    days = 1 if is_1d else int(date[6:8])
    cat_rows, loss_rows = raw['cat'], raw['loss']
    ov, ch = raw['overall'][0], raw['channel'][0]
    cat = {r[4]: r for r in cat_rows}
    loss = {r[4]: r for r in loss_rows}

    checks = {'date': date, 'days': days, 'period': period}
    # 数据日期回显校验
    echo = {r[0] for r in cat_rows}
    checks['date_echo_ok'] = echo == {date}

    wb = load_workbook(tpl_path)
    ws = wb['Sheet1']

    # ---- 部门行序（从模板读取） ----
    tpl_depts = []
    for r in range(3, 18):
        d = ws.cell(row=r, column=4).value
        d = '-' if d in (None, '') else str(d).strip()
        tpl_depts.append(d)

    # ---- 从预算文件读取当期部门未税预算 ----
    br = BudgetJsonReader(budget_path or BUDGET_PATH)
    budget = br.daily_pretax(date) if is_1d else br.mtd_pretax(date)
    # 当期税比：单日用当日税比，MTD用本月混合税比（供POS含税营收折未税）
    tax_ratio = br.daily_tax(date) if is_1d else br.mtd_tax(date)
    J18 = sum(v['total'] for v in budget.values())   # 整体未税预算合计
    AP18 = sum(v['pos'] for v in budget.values())    # POS未税预算合计
    checks['budget_date_ok'] = bool(budget)
    checks['budget_range'] = f"{br.min_date()}~{br.max_date()}"
    checks['budget_covered'] = (br.min_date() or '') <= date <= (br.max_date() or '')
    checks['tax_ratio'] = round(tax_ratio, 4) if tax_ratio else None
    checks['J18_budget'] = round(J18, 0)
    checks['AP18_budget'] = round(AP18, 0)

    missing = [d for d in tpl_depts if d not in cat]
    extra = [d for d in cat if d not in tpl_depts]
    checks['missing_in_api'] = missing
    checks['new_in_api'] = extra

    # ---- 部门行数据 ----
    D = {}
    for d in tpl_depts:
        r = cat.get(d)
        lr = loss.get(d)
        if r is None:
            D[d] = None
            continue
        g = lambda i: num(r[i]) if i < len(r) else None
        e, f, gy = g(5), g(6), g(7)
        n, o = g(10), g(11)
        x, z = g(20), g(22)
        af, ah = g(28), g(30)
        D[d] = {
            'E': e, 'F': f, 'G': gy,
            'L': num(lr[5]) if lr else None, 'M': num(lr[6]) if lr else None,
            'N': n, 'O': o, 'P': g(12), 'Q': g(13), 'R': g(14), 'S': g(15), 'T': g(16),
            'U': g(17), 'V': g(18), 'W': g(19),
            'X': x, 'Y': g(21), 'Z': z, 'AA': g(23), 'AB': g(24), 'AC': g(25),
            'AD': g(26), 'AE': g(27),
            'AF': af, 'AG': g(29), 'AH': ah, 'AI': g(31), 'AJ': g(32), 'AK': g(33),
            'AL': g(34), 'AM': g(35),
        }

    # ---- 总计行汇总 ----
    act = [D[d] for d in tpl_depts if D[d]]
    E18 = sum(x['E'] for x in act if x['E'] is not None)
    def wsum(key):
        s = 0.0
        for x in act:
            v = x[key]
            if v is not None:
                s += v
        return s
    Ep = sum(x['E'] / (1 + x['F']) for x in act if x['E'] is not None and x['F'] is not None and x['F'] > -1)
    Ey = sum(x['E'] / (1 + x['G']) for x in act if x['E'] is not None and x['G'] is not None and x['G'] > -1)
    F18 = chain_mom(E18, Ep)
    G18 = chain_mom(E18, Ey)
    Lw = sum(x['L'] * x['E'] for x in act if x['L'] is not None and x['E'] is not None)
    Mw = sum(x['M'] * x['E'] for x in act if x['M'] is not None and x['E'] is not None)
    Ew = sum(x['E'] for x in act if x['L'] is not None and x['E'] is not None)
    L18 = Lw / Ew if Ew else None
    M18 = Mw / Ew if Ew else None

    N18 = num(ov[4])  # 门店级总成交金额取整体表现（模板验证：与部门ΣN相差约8845，模板采用门店级）
    checks['sum_dept_N'] = round(wsum('N'), 2)
    checks['diff_N18_vs_sumDept'] = round((N18 or 0) - wsum('N'), 2)
    O18, P18, Q18, R18 = num(ov[6]), num(ov[5]), num(ov[7]), num(ov[8])
    U18, V18 = num(ov[9]), num(ov[10])
    X18, Z18, Y18 = num(ch[4]), num(ch[5]), num(ch[7])
    AA18, AB18, AC18 = num(ch[8]), num(ch[10]), num(ch[11])
    AD18, AE18 = num(ch[15]), num(ch[16])
    AF18, AH18, AG18 = num(ch[23]), num(ch[24]), num(ch[25])
    AI18, AJ18, AK18 = num(ch[26]), num(ch[28]), num(ch[29])
    AL18, AM18 = num(ch[31]), num(ch[32])
    AR18 = N18 - AF18 if (N18 is not None and AF18 is not None) else None
    AS18 = AR18 / days if AR18 is not None else None
    AT18 = chain_mom(AR18, (prev_val(N18, O18) or 0) - (prev_val(AF18, AH18) or 0)
                     if prev_val(N18, O18) is not None and prev_val(AF18, AH18) is not None else None)
    AU18 = AR18 / N18 if (AR18 is not None and N18) else None
    AV18 = Q18 - AJ18 if (Q18 is not None and AJ18 is not None) else None
    pq = prev_val(Q18, R18); paj = prev_val(AJ18, AK18)
    AW18 = chain_mom(AV18, pq - paj if pq is not None and paj is not None else None)
    AX18 = AS18 / AV18 if (AS18 is not None and AV18) else None
    AY18 = chain_mom(1 + AT18, 1 + AW18) if AT18 is not None and AW18 is not None else None
    n18p = prev_val(N18, O18); af18p = prev_val(AF18, AH18); x18p = prev_val(X18, Z18)
    BE18 = ((n18p - af18p) / n18p - x18p / n18p) if (n18p and af18p is not None and x18p is not None) else None
    AQ18 = (AF18 / tax_ratio) / AP18 if (AF18 is not None and tax_ratio and AP18) else None
    K18 = E18 / J18 if J18 else None
    checks['online_plus_pos_eq_1'] = abs((AU18 or 0) + (AI18 or 0) - 1) < 0.001
    checks['E18'] = round(E18, 2)
    checks['K18'] = round(K18, 6) if K18 is not None else None
    checks['AQ18'] = round(AQ18, 6) if AQ18 is not None else None

    # ---- 写入部门行 ----
    for i, d in enumerate(tpl_depts):
        r = 3 + i
        x = D[d]
        ws.cell(row=r, column=1).value = date
        ws.cell(row=r, column=2).value = period
        ws.cell(row=r, column=3).value = STORE
        ws.cell(row=r, column=4).value = '' if d == '-' else d
        # ---- 预算列（整体 I/J/K + POS AO/AP/AQ），MTD 与单日通用 ----
        b = budget.get(d)
        if b:
            J = b['total']                       # 整体未税预算
            AP = b['pos']                        # POS未税预算
            I = J / J18 if J18 else None         # 整体预算占比
            AO = AP / AP18 if AP18 else None     # POS预算占比
        else:
            I = AO = J = AP = None
        if x is None:  # API 无此部门 → 营收清零占位，预算照常、达成率记 0
            for c in range(5, 58):
                if c not in (9, 10, 11, 41, 42, 43):
                    ws.cell(row=r, column=c).value = 0 if c in (5, 14) else None
            ws.cell(row=r, column=9).value = I
            ws.cell(row=r, column=10).value = J
            ws.cell(row=r, column=11).value = 0.0 if J else None
            ws.cell(row=r, column=41).value = AO
            ws.cell(row=r, column=42).value = AP
            ws.cell(row=r, column=43).value = 0.0 if AP else None
            continue
        e, n, af = x['E'], x['N'], x['AF']
        K = e / J if (e is not None and J) else None
        # POS达成率：POS成交af为含税，按当期税比折未税，再比未税POS预算
        AQ = (af / tax_ratio) / AP if (af is not None and tax_ratio and AP) else None
        H = e / E18 if (e is not None and E18) else None
        AN = af / AF18 if (af is not None and AF18) else None
        AR = n - af if (n is not None and af is not None) else None
        # 部门日均按部门有效天数(N/P)，非常规日历天数（社区服务26天/其他部门5天等场景）
        days_eff = (n / x['P']) if (n is not None and x['P']) else days
        AS = AR / days_eff if AR is not None else None
        np_, afp_ = prev_val(n, x['O']), prev_val(af, x['AH'])
        AT = chain_mom(AR, np_ - afp_ if np_ is not None and afp_ is not None else None)
        AU = AR / n if (AR is not None and n) else None
        AV = x['Q'] - x['AJ'] if (x['Q'] is not None and x['AJ'] is not None) else None
        qp, ajp = prev_val(x['Q'], x['R']), prev_val(x['AJ'], x['AK'])
        AW = chain_mom(AV, qp - ajp if qp is not None and ajp is not None else None)
        AX = AS / AV if (AS is not None and AV) else None
        AY = chain_mom(1 + AT, 1 + AW) if AT is not None and AW is not None else None
        AZ = AR / AR18 if (AR is not None and AR18) else None
        xp = prev_val(x['X'], x['Z'])
        BE = ((np_ - afp_) / np_ - xp / np_) if (np_ and afp_ is not None and xp is not None) else None
        vals = {
            5: e, 6: x['F'], 7: x['G'], 8: H, 9: I, 10: J, 11: K,
            12: x['L'], 13: x['M'],
            14: n, 15: x['O'], 16: x['P'], 17: x['Q'], 18: x['R'], 19: x['S'], 20: x['T'],
            21: x['U'], 22: x['V'], 23: x['W'],
            24: x['X'], 25: x['Y'], 26: x['Z'], 27: x['AA'], 28: x['AB'], 29: x['AC'],
            30: x['AD'], 31: x['AE'],
            32: af, 33: x['AG'], 34: x['AH'], 35: x['AI'], 36: x['AJ'], 37: x['AK'],
            38: x['AL'], 39: x['AM'],
            40: AN, 41: AO, 42: AP, 43: AQ,
            44: AR, 45: AS, 46: AT, 47: AU, 48: AV, 49: AW, 50: AX, 51: AY, 52: AZ,
            53: None, 54: None, 55: None, 56: None, 57: BE,
        }
        # BA=pos占比 BB=线上占比 BC=app占比 BD=端外占比(线上-app)
        vals[53] = x['AI']
        vals[54] = AU
        vals[55] = x['AA']
        vals[56] = (AU - x['AA']) if (AU is not None and x['AA'] is not None) else None
        for c, v in vals.items():
            ws.cell(row=r, column=c).value = v

    # ---- 写入总计行 ----
    tot = {
        1: date, 2: period, 3: STORE, 4: '总计',
        5: E18, 6: F18, 7: G18, 8: 1,
        9: 1.0, 10: J18, 11: K18,
        12: L18, 13: M18,
        14: N18, 15: O18, 16: P18, 17: Q18, 18: R18, 19: Q18, 20: None,
        21: U18, 22: V18, 23: None,
        24: X18, 25: Y18, 26: Z18, 27: AA18, 28: AB18, 29: AC18, 30: AD18, 31: AE18,
        32: AF18, 33: AG18, 34: AH18, 35: AI18, 36: AJ18, 37: AK18, 38: AL18, 39: AM18,
        40: AI18,
        41: 1.0, 42: AP18, 43: AQ18,
        44: AR18, 45: AS18, 46: AT18, 47: AU18, 48: AV18, 49: AW18, 50: AX18, 51: AY18, 52: AU18,
        53: AI18, 54: AU18, 55: AA18,
        56: (AU18 - AA18) if (AU18 is not None and AA18 is not None) else None, 57: BE18,
    }
    for c, v in tot.items():
        ws.cell(row=18, column=c).value = v

    wb.save(out_path)
    checks['status'] = 'ok'
    checks['out'] = out_path
    print(json.dumps(checks, ensure_ascii=False))

if __name__ == '__main__':
    period = sys.argv[4] if len(sys.argv) > 4 else 'mtd'
    budget_path = sys.argv[5] if len(sys.argv) > 5 else None
    main(sys.argv[1], sys.argv[2], sys.argv[3], period, budget_path)
