# -*- coding: utf-8 -*-
"""
ArmChio — финансовая модель адаптации формата «Чио Чио» (фикс-прайс экспресс-парикмахерская) под Армению.
Запуск:  python -I model.py
Выходы (в папке скрипта):
  ArmChio_model.xlsx        — модель с формулами (лист PnL_36m и CAPEX считаются формулами Excel)
  pnl_36m_<scenario>.csv    — помесячный P&L на 36 мес. по сценариям
  model_output.json         — ключевые числа
  financial_model.md        — полный отчёт (копируется также в scratchpad, см. OUT_MD)
Все входные параметры — в блоке ASSUMPTIONS с источником и степенью достоверности.
"""
import csv, json, math, os, sys
from copy import deepcopy

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_MD = r"C:\Users\Tigran\AppData\Local\Temp\claude\C-------------\8c8c1948-a2aa-44fe-b054-2a9727d74e60\scratchpad\financial_model.md"

F1 = "chio_chio_model.md"
F2 = "armenia_prices_competitors.md"
F3 = "armenia_costs_regulation_macro.md"

# ----------------------------------------------------------------------------------------------
# 1. ДОПУЩЕНИЯ  (key, value, unit, source, confidence, comment)
# ----------------------------------------------------------------------------------------------
A = {}
ASSUMPTIONS = []
def add(key, value, unit, source, conf, comment=""):
    A[key] = value
    ASSUMPTIONS.append((key, value, unit, source, conf, comment))

# --- курсы, макро
add("fx_usd", 361.36, "AMD/USD", f"{F2}, {F3}: ЦБА 19.10.2026", "факт")
add("fx_rub", 4.3731, "AMD/RUB", f"{F2}: ЦБА 19.10.2026", "факт")
add("cpi", 0.043, "доля/год", f"{F3}: ИПЦ сен-2026 г/г 4,3%", "факт", "индексация цен и фикс. расходов с 13-го и 25-го мес.")
add("refi_rate", 0.0675, "доля/год", f"{F3}: ставка ЦБ РА, 09.2026", "факт")
add("discount_rate", 0.18, "доля/год", "оценка: ставка ЦБ 6,75% + премия за риск малого бизнеса/неликвидность ≈ 11 п.п.", "оценка",
    "диапазон 15–20%; 18% как середина, проверка при 15% и 20% в разделе F")
add("avg_wage_am", 314196, "AMD/мес", f"{F3}: Armstat авг-2026", "факт")
add("avg_wage_ru", 103612, "RUB/мес", "Росстат, январь 2026 (цит. по glavkniga.ru/situations/k518592 и t-j.ru/news/salary-2026-10/)", "факт (вторичный источник)",
    "общероссийская средняя начисленная; 2025 г. — 100 360 ₽ (brobank.ru по Росстату)")
add("pop_yerevan", 1147600, "чел", f"{F3}: Armstat 01.01.2026", "факт")

# --- прайс (вариант A: спальный район; вариант B: центр/ТЦ)
SERVICES = [  # key, name, price A, price B, mix share, minutes (с уборкой места)
    ("mens",   "Мужская стрижка (ножницы/модельная)", 2500, 3000, 0.40, 20),
    ("clip",   "Стрижка машинкой (одна длина)",        2000, 2500, 0.20, 15),
    ("kids",   "Детская стрижка (до 12 лет)",          2000, 2500, 0.12, 15),
    ("wshort", "Женская стрижка короткая",             2500, 3000, 0.08, 20),
    ("wlong",  "Женская стрижка длинная (без укладки)",3500, 4000, 0.05, 30),
    ("beard",  "Борода / контур",                      1500, 2000, 0.08, 10),
    ("style",  "Укладка феном (без мытья)",            1500, 2000, 0.04, 15),
    ("roots",  "Окрашивание корней (опция, 1 тон)",    5000, 6000, 0.03, 45),
]
assert abs(sum(s[4] for s in SERVICES) - 1.0) < 1e-9
add("price_mens_A", 2500, "AMD", f"{F2} §4: медиана эконом 3 000, дно 1 000–1 500, True Cost 3 000; позиционирование −17% к медиане эконом", "оценка")
add("price_mens_B", 3000, "AMD", f"{F2}: центр 3 000–5 000; = True Cost (но за 15 мин вместо 30)", "оценка")
add("mix", "40/20/12/8/5/8/4/3", "% услуг", f"{F1}: ЦА — мужчины, дети, пенсионеры; {F2} §3: прокси ЮMoney (мужчины — эконом)", "допущение",
    "муж/машинка/дет/жен.кор/жен.дл/борода/укладка/корни; проверять по CRM после 3 мес.")
add("minutes_avg", round(sum(s[4]*s[5] for s in SERVICES), 2), "мин/услуга", f"{F1}: норматив 15 мин, 4–5 стрижек/час; взвешено по миксу", "оценка")

# --- операционная модель
add("hours_per_day", 11, "ч", f"{F1}: Казань 09–21; {F2}: Ереван 10–20/10–22 → 10:00–21:00", "оценка")
add("days_per_month", 365/12, "дн", "календарь, без выходных (7/7)", "факт")
add("chairs", 3, "шт", f"{F1}: 2–3 кресла на точку", "факт")
add("masters", 4, "чел", "3 кресла × 7 дней × 11 ч = 231 кресло-ч/нед ≈ 4 мастера по 5–6 смен", "оценка", "ТК РА 40 ч/нед — проверить юридически схему смен")
add("load_mature_base", 0.45, "доля ёмкости", f"{F1}: зрелая точка 40–70 стрижек/день при ёмкости ~100–115; консервативно для нового бренда", "оценка")
add("load_mature_pess", 0.32, "доля", f"{F1}: «треть франчайзи убыточны»; Подмосковье 80–100 тыс.₽", "оценка")
add("load_mature_opt", 0.55, "доля", f"{F1}: Хабаровск 50–80/день, Казань ~1,1 млн ₽", "оценка")
LOAD_SHAPE = [0.28,0.37,0.47,0.56,0.63,0.70,0.74,0.77,0.79,0.81,0.84,0.86] \
           + [0.88 + (0.95-0.88)*i/11 for i in range(12)] \
           + [0.97 + (1.00-0.97)*i/11 for i in range(12)]
add("load_shape", "м1 28% → м6 70% → м12 86% → м24 95% → м36 100% от зрелой", "кривая", f"{F1}: кейсы Екб (15 тыс. ₽ в 1-й мес.), spark (10–15 → 90–100/день за 9 мес.), «первые 2 недели клиентов почти не было»", "оценка")

# --- расходы
add("area_sqm", 40, "м²", f"{F3}: модель 40 м² (Чио Чио 20–25 м² — вариант CAPEX)", "факт")
add("rent_sqm_sleep", 6500, "AMD/м²/мес", f"{F3} §1.2: медиана спальных районов", "факт (цены предложения)")
add("rent_sqm_center", 8500, "AMD/м²/мес", f"{F3} §1.2: медиана Кентрон", "факт (цены предложения)")
add("rent_sqm_center_1line", 13000, "AMD/м²/мес", f"{F3} §1.2: Кентрон первая линия/витрина", "факт (цены предложения)")
add("deposit_months", 2, "мес", f"{F3}: депозит 1–2 мес.", "факт")
add("fot_share_market", 0.50, "доля выручки", f"{F3} §2.1: доминирует 50/50", "факт")
add("fot_share_chio", 0.40, "доля выручки", f"{F1} §2.3: ФОТ ≈35–45%", "оценка")
add("fot_floor_per_master", 150000, "AMD/мес gross", f"{F3}: минималка 75 000, оклады 200–400 тыс. → гарантия на ramp-up", "допущение")
add("manager_cost", 120000, "AMD/мес", "0,5 ставки управляющего-наставника (Чио Чио: администратора нет, но контроль нужен)", "допущение")
add("cleaning_cost", 40000, "AMD/мес", "уборка 1 ч/день", "допущение")
add("consumables_per_client", 95, "AMD/клиент", f"{F3}: дезсредства/расходники 20–40 тыс. на 3 кресла; {F1}: ~30 тыс.₽ при ~500 тыс.₽ (6%) — без мытья/косметики ниже; ≈4% чека A", "оценка", "одноразовые воротнички/салфетки, дезраствор, лезвия, масло для машинок")
add("consumables_floor", 30000, "AMD/мес", f"{F3} §5.1", "оценка")
add("utilities", 75000, "AMD/мес", f"{F3} §7: 60–90 тыс. на 40 м²", "оценка")
add("it_crm", 17000, "AMD/мес", "alteg.io цены не публикует (проверено alteg.io/ru и KB 20.10.2026); ориентир $10–12/сотрудник × 4 ≈ 15 тыс. + моб. связь 2 тыс.; альтернатива Fresha (0 ₽ + комиссия)", "оценка")
add("hdm_monthly", 3000, "AMD/мес", f"{F3} §5.2: e-HDM 3 000/мес", "факт")
add("hdm_setup", 30000, "AMD", f"{F3} §5.2: e-HDM регистрация 30 000", "факт")
add("card_share", 0.50, "доля выручки", f"{F2} §3: 33% только наличные, POS 19% оборота карт, QR Idram — допущение 40–60%", "допущение")
add("acq_rate", 0.012, "доля безнала", f"{F3} §5.3: 0,5–2,5%; взвешено ArCa/Visa/Idram", "оценка")
add("acq_terminal", 4600, "AMD/мес", f"{F3} §5.3: IDBank", "факт")
add("marketing_pct", 0.03, "доля выручки", f"{F1}: модель УК 10 тыс.₽ при 150–300 тыс. (3–7%)", "оценка")
add("marketing_floor", 60000, "AMD/мес", "карты/соцсети/флаеры минимум", "оценка")
add("accounting", 35000, "AMD/мес", "аутсорс бухгалтерии ИП на оборотном налоге (ежемесячная отчётность, ФОТ)", "оценка")
add("other_pct", 0.02, "доля выручки", "ремонт/замена инструмента, форс-мажор", "допущение")
add("other_floor", 20000, "AMD/мес", "", "допущение")
add("other_documented_share", 0.5, "доля", "часть «прочих» расходов без документов", "допущение")
add("ip_fixed", 30800, "AMD/мес", f"{F3} §4.2: подоходный 5 000 + соц. 5 000 + гербовый 120 000/год + ОМС 129 600/год", "факт")
add("tt_rate", 0.10, "доля", f"{F3} §4.2: налог с оборота, услуги 10%", "факт")
add("tt_deduct", 0.06, "доля расходов", f"{F3} §4.2: вычет 6% документированных расходов", "факт")
add("tt_floor", 0.045, "доля", f"{F3} §4.2: минимум 4,5% оборота", "факт")
add("fot_documented", True, "да/нет", f"{F3} §4.3: в примере ФОТ включён в документированные расходы; официальное трудоустройство", "оценка", "альтернатива (ФОТ не зачитывается) — в разделе C")
add("royalty_chio", 0.04, "доля выручки", f"{F1}: 4%, мин. 5 000 ₽, со 2-го мес.", "факт")
add("royalty_chio_min_rub", 5000, "RUB/мес", F1, "факт")
add("fee_chio_lite_rub", 290000, "RUB", f"{F1}: пакет «Лайт»/2.0", "факт")
add("fee_chio_std_rub", 590000, "RUB", f"{F1}: основной пакет 2025–26", "факт")
add("depr_months", 60, "мес", "5 лет на ремонт/оборудование; паушальный — срок договора 5 лет", "допущение")

# --- CAPEX
add("fitout_sqm", 100000, "AMD/м²", f"{F3} §6: 80–120 тыс./м² (салон «средний+»)", "оценка")
add("chair_price", 90000, "AMD", f"{F3} §6: РФ 56–110 тыс. (ЕАЭС без пошлины, +НДС 20% при ввозе)", "факт/оценка")
add("post_price", 120000, "AMD", f"{F3} §6: рабочее место (зеркало, консоль)", "оценка")
add("wash_unit", 220000, "AMD", f"{F3} §6: кресло-мойка 130–304 тыс. (1 шт. для окрашивания/платного мытья)", "факт/оценка")
add("vacuum_unit", 125000, "AMD", "Sibel Eye Vac £215–307 (sallybeauty.co.uk), Head Jog Hair Vacuum £148,5 (astonandfincher.co.uk); ≈$190–390 + доставка/НДС 20% → 100–170 тыс.; 2 шт. на 3 кресла", "оценка",
    "цены аппарата Чио Чио не публикуются; «воздушная мойка» = пылесос с насадкой (файл 1 §7)")
add("tools_per_post", 150000, "AMD", f"{F3} §6: машинки, ножницы, фен, тележка на пост", "оценка")
add("sterilizer", 80000, "AMD", f"{F3} §5.1: УФ/сухожар 30–120 тыс.", "факт/оценка")
add("waiting_furniture", 350000, "AMD", f"{F3} §6: мебель зоны ожидания/стойка 300–600 тыс.", "оценка")
add("signage", 300000, "AMD", f"{F3} §6: вывеска 150–400 тыс.", "оценка")
add("branding_interior", 200000, "AMD", "брендинг интерьера, навигация, прайс-борды, форма", "оценка")
add("selfservice_it", 250000, "AMD", "планшет/киоск самообслуживания + стойка + ПК + камеры (замена терминала Чио Чио)", "оценка")
add("crm_setup", 30000, "AMD", "настройка CRM/онлайн-записи, сайт-визитка", "оценка")
add("registration", 3000, "AMD", f"{F3} §4.4: госпошлина ИП", "факт")
add("legal_setup", 60000, "AMD", "договоры, ЭЦП, консультация по налоговому режиму", "оценка")
add("medbooks", 60000, "AMD", f"{F3} §5.1: 10–20 тыс./чел × 4", "оценка")
add("launch_marketing", 400000, "AMD", f"{F1}: бесплатные стрижки на открытии (возврат 70–80%), флаеры, карты; ≈ 12% годового маркетинга", "оценка")
add("contingency", 0.10, "доля hard-CAPEX", "непредвиденное (ремонт по факту +20–50% в кейсах Чио Чио)", "допущение")

FX = A["fx_usd"]; RUB = A["fx_rub"]
DAYS = A["days_per_month"]

# ----------------------------------------------------------------------------------------------
# 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ----------------------------------------------------------------------------------------------
def avg_check(price_col):  # 2 -> A, 3 -> B
    return sum(s[price_col]*s[4] for s in SERVICES)

def capacity_per_day(chairs, hours=None):
    minutes = A["minutes_avg"]
    hours = hours or A["hours_per_day"]
    return chairs * hours * 60.0 / minutes

def default_params(**kw):
    p = dict(
        name="base", chairs=3, masters=4, price_mult=1.0, price_col=2,
        load_mature=A["load_mature_base"], rent_sqm=A["rent_sqm_sleep"], area=A["area_sqm"],
        fot_share=A["fot_share_market"], consumables_per_client=A["consumables_per_client"],
        royalty=0.0, royalty_min=0.0, royalty_from_month=2, franchise_fee=0.0,
        fitout_area=40, manager_cost=A["manager_cost"], fot_documented=A["fot_documented"],
        marketing_pct=A["marketing_pct"], indexation=True, load_mult=1.0,
    )
    p.update(kw)
    return p

def capex(p):
    chairs = p["chairs"]
    hard = {
        "Ремонт/отделка (fit-out) %d м² × %s" % (p["fitout_area"], f"{A['fitout_sqm']:,}"): p["fitout_area"]*A["fitout_sqm"],
        "Кресла парикмахерские × %d" % chairs: chairs*A["chair_price"],
        "Рабочие посты (зеркало, консоль) × %d" % chairs: chairs*A["post_price"],
        "Мойка (1 шт., для окрашивания/платного мытья)": A["wash_unit"],
        "«Воздушная мойка» — вакуумные аппараты × 2": 2*A["vacuum_unit"],
        "Инструмент и техника × %d постов" % chairs: chairs*A["tools_per_post"],
        "Стерилизатор УФ/сухожар": A["sterilizer"],
        "Мебель зоны ожидания, стойка, хранение": A["waiting_furniture"],
        "Вывеска световая": A["signage"],
        "Брендинг интерьера, прайс-борды, форма": A["branding_interior"],
        "Киоск самообслуживания/ПК/камеры": A["selfservice_it"],
        "Касса e-HDM (регистрация)": A["hdm_setup"],
    }
    hard_total = sum(hard.values())
    soft = {
        "Непредвиденное %d%% от hard-CAPEX" % int(A["contingency"]*100): round(hard_total*A["contingency"]),
        "Настройка CRM/онлайн-записи/сайт": A["crm_setup"],
        "Регистрация ИП": A["registration"],
        "Юридическое/налоговое оформление": A["legal_setup"],
        "Медкнижки персонала": A["medbooks"],
        "Депозит аренды × %d мес." % A["deposit_months"]: A["deposit_months"]*p["rent_sqm"]*p["area"],
        "Маркетинг открытия": A["launch_marketing"],
        "Паушальный взнос": p["franchise_fee"],
    }
    items = dict(hard); items.update(soft)
    return items, hard_total

def month_pnl(p, m, load, royalty_on=True):
    """P&L месяца m (1..36) при загрузке load (доля ёмкости)."""
    idx = (1+A["cpi"])**((m-1)//12) if p["indexation"] else 1.0
    check = avg_check(p["price_col"])*p["price_mult"]*idx
    cap_day = capacity_per_day(p["chairs"])
    clients_day = cap_day*load
    rev = clients_day*check*DAYS
    fot = max(p["fot_share"]*rev, p["masters"]*A["fot_floor_per_master"]*idx)
    rent = p["rent_sqm"]*p["area"]*idx
    mgr = (p["manager_cost"]+A["cleaning_cost"])*idx
    cons = max(p["consumables_per_client"]*idx*clients_day*DAYS, A["consumables_floor"]*idx)
    util = A["utilities"]*idx
    it = A["it_crm"]*idx
    hdm = A["hdm_monthly"]
    acq = rev*A["card_share"]*A["acq_rate"] + A["acq_terminal"]
    mkt = max(p["marketing_pct"]*rev, A["marketing_floor"]*idx)
    acc = A["accounting"]*idx
    oth = max(A["other_pct"]*rev, A["other_floor"]*idx)
    roy = 0.0
    if royalty_on and p["royalty"] > 0 and m >= p["royalty_from_month"]:
        roy = max(p["royalty"]*rev, p["royalty_min"])
    documented = rent + mgr + cons + util + it + hdm + acq + mkt + acc + roy + oth*A["other_documented_share"]
    if p["fot_documented"]:
        documented += fot
    tt = max(A["tt_rate"]*rev - A["tt_deduct"]*documented, A["tt_floor"]*rev)
    ip = A["ip_fixed"]
    opex = fot+rent+mgr+cons+util+it+hdm+acq+mkt+acc+oth+roy+tt+ip
    ebitda = rev - opex
    return dict(month=m, load=load, clients_day=clients_day, check=check, revenue=rev, fot=fot, rent=rent, manager=mgr,
                consumables=cons, utilities=util, it=it, hdm=hdm, acquiring=acq, marketing=mkt, accounting=acc,
                other=oth, royalty=roy, documented=documented, turnover_tax=tt, ip_fixed=ip, opex=opex, ebitda=ebitda)

def npv(rate_annual, cfs):  # cfs[0] at t=0
    r = (1+rate_annual)**(1/12)-1
    return sum(cf/(1+r)**t for t, cf in enumerate(cfs))

def irr_annual(cfs):
    lo, hi = -0.99, 10.0
    f = lambda r: sum(cf/(1+r)**t for t, cf in enumerate(cfs))
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo+hi)/2
        if f(lo)*f(mid) <= 0: hi = mid
        else: lo = mid
    return (1+(lo+hi)/2)**12-1

def run(p, months=36):
    shape = LOAD_SHAPE
    rows = []
    for m in range(1, months+1):
        load = min(0.95, shape[m-1]*p["load_mature"]*p["load_mult"])
        rows.append(month_pnl(p, m, load))
    # резерв = максимальная просадка накопленного операционного потока (до CAPEX)
    cum, mn = 0.0, 0.0
    for r in rows:
        cum += r["ebitda"]; mn = min(mn, cum)
    reserve = -mn
    items, hard_total = capex(p)
    items["Оборотный резерв (покрытие убытков ramp-up, расчёт)"] = round(reserve)
    capex_total = sum(items.values())
    depr_base = hard_total + items["Непредвиденное %d%% от hard-CAPEX" % int(A["contingency"]*100)] + A["crm_setup"] + A["branding_interior"]*0
    depr = depr_base/A["depr_months"] + p["franchise_fee"]/A["depr_months"]
    cfs = [-capex_total]
    cum = -capex_total; payback = None; cum_disc = -capex_total; dpayback = None
    r_m = (1+A["discount_rate"])**(1/12)-1
    for r in rows:
        r["depreciation"] = depr
        r["net_profit"] = r["ebitda"] - depr
        r["cf"] = r["ebitda"]           # налог с оборота уже в opex; CAPEX в t=0
        cum += r["cf"]; r["cum_cf"] = cum
        cum_disc += r["cf"]/(1+r_m)**r["month"]; r["cum_dcf"] = cum_disc
        if payback is None and cum >= 0: payback = r["month"]
        if dpayback is None and cum_disc >= 0: dpayback = r["month"]
        cfs.append(r["cf"])
    # возврат депозита в конце горизонта (не включаем в NPV консервативно)
    tv = 3*sum(r["ebitda"] for r in rows[24:36])  # терминальная стоимость = 3× EBITDA года 3 (справочно)
    res = dict(params=p, rows=rows, capex_items=items, capex_total=capex_total, hard_capex=hard_total, reserve=reserve,
               depr=depr, payback=payback, dpayback=dpayback, npv=npv(A["discount_rate"], cfs), npv_tv=npv(A["discount_rate"], cfs[:-1]+[cfs[-1]+tv]), tv=tv,
               npv15=npv(0.15, cfs), npv20=npv(0.20, cfs), irr=irr_annual(cfs), cfs=cfs)
    for y in (1,2,3):
        yr = rows[(y-1)*12:y*12]
        res[f"y{y}"] = {k: sum(r[k] for r in yr) for k in ("revenue","fot","rent","opex","ebitda","net_profit","turnover_tax","royalty","depreciation")}
        res[f"y{y}"]["clients_day"] = sum(r["clients_day"] for r in yr)/12
    res["m36"] = rows[35]
    return res

def breakeven(p, month=36):
    """Загрузка, при которой чистая прибыль (EBITDA − амортизация) = 0 при структуре расходов месяца `month`."""
    base = run(p)
    depr = base["depr"]
    def f(load):
        r = month_pnl(p, month, load)
        return r["ebitda"] - depr
    lo, hi = 0.001, 0.95
    if f(hi) < 0: return None
    for _ in range(100):
        mid = (lo+hi)/2
        if f(mid) < 0: lo = mid
        else: hi = mid
    r = month_pnl(p, month, hi)
    # EBITDA-безубыточность (без амортизации)
    lo2, hi2 = 0.001, 0.95
    for _ in range(100):
        mid = (lo2+hi2)/2
        if month_pnl(p, month, mid)["ebitda"] < 0: lo2 = mid
        else: hi2 = mid
    r2 = month_pnl(p, month, hi2)
    return dict(load=hi, clients_day=r["clients_day"], revenue=r["revenue"], load_ebitda=hi2, clients_day_ebitda=r2["clients_day"], revenue_ebitda=r2["revenue"])

# ----------------------------------------------------------------------------------------------
# 3. РАСЧЁТЫ
# ----------------------------------------------------------------------------------------------
fmt = lambda x, d=0: f"{x:,.{d}f}".replace(",", " ")
usd = lambda x: x/FX
rub = lambda x: x/RUB

check_A = avg_check(2); check_B = avg_check(3)
cap3 = capacity_per_day(3); cap2 = capacity_per_day(2)

# сценарии (спальный район, своя сеть)
P_BASE = default_params(name="base")
P_PESS = default_params(name="pessimistic", price_mult=0.90, load_mature=A["load_mature_pess"], rent_sqm=7500, consumables_per_client=120)
P_OPT  = default_params(name="optimistic", load_mature=A["load_mature_opt"], fot_share=0.45)
R = {p["name"]: run(p) for p in (P_BASE, P_PESS, P_OPT)}

# варианты локации / схем (зрелый месяц 36)
P_CENTER = default_params(name="center", price_col=3, rent_sqm=A["rent_sqm_center"], load_mature=0.50)
P_CENTER1 = default_params(name="center_1line", price_col=3, rent_sqm=A["rent_sqm_center_1line"], load_mature=0.55)
P_2CH = default_params(name="2chairs", chairs=2, masters=3, fitout_area=25, area=25)
P_FOT40 = default_params(name="fot40", fot_share=A["fot_share_chio"])
P_NOFOTDOC = default_params(name="fot_not_documented", fot_documented=False)
P_OWNER = default_params(name="owner_operator", manager_cost=0)
# стратегические варианты
P_FR_LITE = default_params(name="franchise_lite", royalty=A["royalty_chio"], royalty_min=A["royalty_chio_min_rub"]*RUB, franchise_fee=A["fee_chio_lite_rub"]*RUB)
P_FR_STD  = default_params(name="franchise_std",  royalty=A["royalty_chio"], royalty_min=A["royalty_chio_min_rub"]*RUB, franchise_fee=A["fee_chio_std_rub"]*RUB)
VAR = {p["name"]: run(p) for p in (P_CENTER, P_CENTER1, P_2CH, P_FOT40, P_NOFOTDOC, P_OWNER, P_FR_LITE, P_FR_STD)}

BE = breakeven(P_BASE)
BE_PESS = breakeven(P_PESS); BE_OPT = breakeven(P_OPT); BE_2 = breakeven(P_2CH); BE_C = breakeven(P_CENTER)

# --- чувствительность
def sens_table():
    out = []
    for factor, label in (("price_mult","Цена"),("load_mult","Загрузка"),("rent_sqm","Аренда"),("fot_share","Доля ФОТ")):
        row = {"factor": label}
        for d in (-0.2,-0.1,0.0,0.1,0.2):
            p = deepcopy(P_BASE)
            if factor in ("price_mult","load_mult"): p[factor] = 1+d
            elif factor == "rent_sqm": p[factor] = A["rent_sqm_sleep"]*(1+d)
            else: p[factor] = A["fot_share_market"]*(1+d)
            r = run(p)
            row[d] = dict(net3=r["y3"]["net_profit"], net36=sum(x["net_profit"] for x in r["rows"]), payback=r["payback"], npv=r["npv"])
        out.append(row)
    return out
SENS = sens_table()
TORNADO = sorted([(s["factor"], s[-0.2]["net36"], s[0.2]["net36"]) for s in SENS], key=lambda t: -abs(t[2]-t[1]))

# --- рынок сверху вниз (спальный район)
add("men_15_64_share", 0.31, "доля населения", "допущение: мужчины ≈47,5% × возраст 15–64 ≈66% (структура населения РА — вне файлов, проверить по Armstat)", "допущение")
add("kids_share", 0.20, "доля населения", "допущение: дети 0–14 (вне файлов, проверить по Armstat)", "допущение")
add("haircut_freq_men", 10, "раз/год", f"{F2} §3: прокси ЮMoney (21% раз в месяц), оценка 3–5 недель → 8–12/год", "допущение")
add("haircut_freq_kids", 6, "раз/год", "допущение", "допущение")
add("catchment_radius_km", 1.2, "км", f"{F1}: «у дома»; радиус 1–1,5 км", "допущение")
DISTRICTS = [  # name, pop 2025, density
    ("Малатия-Себастия",150100,5966),("Шенгавит",140400,3458),("Нор-Норк",137100,9717),("Кентрон",132800,9948),
    ("Эребуни",129500,2670),("Ачапняк",122500,4744),("Арабкир",119000,8981),("Канакер-Зейтун",75400,9754),
    ("Аван",58700,7229),("Давташен",51900,7960),("Норк-Мараш",13900,2920),("Нубарашен",11100,644)]
area_catch = math.pi*A["catchment_radius_km"]**2
def catchment(density, pop): return min(density*area_catch, pop)
TOPDOWN = []
for name, pop, dens in DISTRICTS:
    c = catchment(dens, pop)
    men_cuts_day = c*A["men_15_64_share"]*A["haircut_freq_men"]/365
    kids_cuts_day = c*A["kids_share"]*A["haircut_freq_kids"]/365
    TOPDOWN.append((name, pop, dens, c, men_cuts_day, kids_cuts_day))
# требуемая доля: базовая точка — услуги муж.+машинка+борода = 68% визитов, дети 12%
m36 = R["base"]["m36"]
need_men = m36["clients_day"]*(0.40+0.20+0.08)
need_kids = m36["clients_day"]*0.12
salons_per_capita = 1529/A["pop_yerevan"]

# --- ёмкость сети
add("pop_per_point_base", 50000, "жит./точка", f"{F1}: Казань 41 точка / ~1,3 млн ≈ 1 на 32 тыс. (насыщение с каннибализацией); {F2}: 1 529 парикмахерских → консервативно 50 тыс.", "оценка")
add("pop_per_point_aggr", 35000, "жит./точка", "агрессивный ориентир ~ Казань", "оценка")
add("regions_pop", "Гюмри ~112 тыс., Ванадзор ~78 тыс., Абовян ~44 тыс., Вагаршапат ~47 тыс.", "чел", "ВНЕ ФАЙЛОВ: перепись 2022 (ориентир, проверить по Armstat); в файлах только число парикмахерских 2ГИС: 119/3/61/95", "допущение")
REGIONS = [("Гюмри",112000,119,2),("Ванадзор",78000,3,1),("Абовян",44000,61,1),("Вагаршапат",47000,95,1)]
NETCAP = []
for name, pop, dens in DISTRICTS:
    base_n = int(round(pop/A["pop_per_point_base"])); aggr_n = int(round(pop/A["pop_per_point_aggr"]))
    if name == "Кентрон": base_n += 1; aggr_n += 1   # дневной трафик
    if dens < 3000: base_n = max(0, base_n-1)       # низкая плотность — меньше точек «у дома»
    NETCAP.append((name, pop, dens, base_n, aggr_n))
yerevan_base = sum(x[3] for x in NETCAP); yerevan_aggr = sum(x[4] for x in NETCAP)
regions_base = sum(r[3] for r in REGIONS)

# --- график открытий (36 мес.) и консолидация сети (вариант 3)
OPENINGS = [(1,"own","Нор-Норк"),(7,"own","Малатия-Себастия"),(10,"own","Ачапняк"),
            (13,"fr","Арабкир"),(15,"fr","Шенгавит"),(17,"fr","Эребуни"),(19,"fr","Канакер-Зейтун"),(21,"fr","Кентрон"),(23,"fr","Давташен"),
            (25,"fr","Малатия-Себастия 2"),(26,"fr","Нор-Норк 2"),(27,"fr","Гюмри"),(28,"fr","Аван"),(29,"fr","Кентрон 2"),
            (31,"fr","Ачапняк 2"),(32,"fr","Абовян"),(33,"fr","Шенгавит 2"),(34,"fr","Вагаршапат"),(35,"fr","Эребуни 2"),(36,"fr","Арабкир 2")]
# условия своей франшизы
add("own_fee", 1500000, "AMD", "предложение: ≈$4,2 тыс. / 343 тыс.₽ — чуть выше «Лайт» Чио Чио (1,27 млн AMD); включает обучение, дизайн-проект, запуск", "предложение", "подобрано из юнит-экономики: маржа субфранчайзи ≥10%, окупаемость ≤36 мес.")
add("own_royalty", 0.03, "доля выручки", "предложение: 3% (Чио Чио 4%, Прядки в порядке 3%, отрасль 1–7%) — обоснование в E: при 4–5% субфранчайзи теряет 10% маржи", "предложение")
add("own_royalty_min", 50000, "AMD/мес", "≈ 2,3× минимума Чио Чио (5 000 ₽ = 21 866 AMD); со 2-го месяца", "предложение")
add("own_mkt_fund", 0.01, "доля выручки", "общий маркетинговый фонд внутри стандартного бюджета маркетинга 3% (2% локально + 1% в фонд сети); тратится УК на сеть", "предложение")
add("supply_margin", 0.20, "доля закупок", "маржа УК на расходниках/косметике (Чио Чио — наценка 120–130% по отзыву; берём умеренно)", "оценка")
add("uk_payroll", 800000, "AMD/мес gross", "2 чел.: директор сети/запуск 450 тыс., тренер-технолог 350 тыс. (ср. з/п Ереван ≈348 тыс.); SMM — на аутсорсе в «прочем»", "оценка")
add("uk_other", 500000, "AMD/мес", "офис/коворкинг 100, IT/CRM/сайт/приложение 100, маркетинг бренда и SMM-аутсорс 150, юр./бух. 100, прочее 50 тыс. (сверх маркетингового фонда сети)", "оценка")
add("uk_setup", 3000000, "AMD", "бренд-бук, ТЗ, стандарты, бизнес-бук, учебная программа, IT-настройка (разово, перед первой франшизой)", "оценка")
P_OWN_FR = default_params(name="own_franchisee", royalty=A["own_royalty"], royalty_min=A["own_royalty_min"], franchise_fee=A["own_fee"])
R_OWN_FR = run(P_OWN_FR)
P_OWN_FR4 = default_params(name="own_franchisee_4pct", royalty=0.04, royalty_min=A["own_royalty_min"], franchise_fee=A["own_fee"])
R_OWN_FR4 = run(P_OWN_FR4)
P_OWN_FR6 = default_params(name="own_franchisee_5pct", royalty=0.05, royalty_min=A["own_royalty_min"], franchise_fee=A["own_fee"])
R_OWN_FR6 = run(P_OWN_FR6)

def franchisor_steady(n_points, own_points=3, openings_per_year=0):
    """УК при зрелой сети n_points (из них own_points собственных), всё на уровне месяца 36 базовой точки."""
    rev_pt = m36["revenue"]
    fr_pts = max(0, n_points-own_points)
    royalty = fr_pts*max(A["own_royalty"]*rev_pt, A["own_royalty_min"])
    mkt_fund = fr_pts*A["own_mkt_fund"]*rev_pt
    supply = fr_pts*m36["consumables"]*A["supply_margin"]  # собственные точки — внутригрупповой оборот, не доход
    fees = openings_per_year*A["own_fee"]/12
    own_profit = own_points*m36["ebitda"]
    uk_cost = A["uk_payroll"]+A["uk_other"]+mkt_fund  # фонд тратится полностью
    uk_tax = max(A["tt_rate"]*(royalty+fees+supply) - A["tt_deduct"]*(A["uk_payroll"]+A["uk_other"]), A["tt_floor"]*(royalty+fees+supply))
    return dict(n=n_points, fr=fr_pts, royalty=royalty, mkt_fund=mkt_fund, supply=supply, fees=fees, own_profit=own_profit,
                uk_cost=uk_cost, uk_tax=uk_tax, uk_net_ex_own=royalty+mkt_fund+supply+fees-uk_cost-uk_tax,
                total=royalty+mkt_fund+supply+fees+own_profit-uk_cost-uk_tax)
FRANCHISOR = [franchisor_steady(n, openings_per_year=o) for n, o in ((5,2),(10,4),(20,6))]
FR_PT_INCOME = max(A["own_royalty"]*m36["revenue"], A["own_royalty_min"]) + m36["consumables"]*A["supply_margin"]
N_BE_UK = (A["uk_payroll"]+A["uk_other"]) / FR_PT_INCOME

def master_franchise_benchmark():
    """Вариант 2: мастер-франшиза Чио Чио — отраслевые ориентиры (franchising.com, global-franchise.com)."""
    rev_pt = m36["revenue"]
    out = []
    for n in (5,10,20):
        gross_roy = n*A["royalty_chio"]*rev_pt
        master_share = 0.5*gross_roy  # 50/50 — нижняя граница диапазона 50/50…75/25 в пользу мастера
        sub_fee = A["fee_chio_lite_rub"]*RUB  # паушальный субфранчайзи ≈ «Лайт»
        fee_share = 0.5*sub_fee*(n/3)/12  # ~n/3 открытий в год, 50% мастеру
        uk_cost = A["uk_payroll"]+A["uk_other"]
        out.append(dict(n=n, master_royalty=master_share, master_fees=fee_share, uk_cost=uk_cost, net=master_share+fee_share-uk_cost))
    return out
MASTER = master_franchise_benchmark()

def network_consolidated():
    """Помесячная консолидация сети (вариант 3): собственные точки — полный P&L; франчайзи — роялти+фонд+поставки+паушальный; УК — расходы."""
    rows = []
    own_runs = {}; fr_runs = {}
    for (m0, kind, nm) in OPENINGS:
        p = default_params(name=nm, indexation=False)
        (own_runs if kind=="own" else fr_runs)[(m0,nm)] = run(p)
    cum = 0.0; cum_inv = 0.0
    for m in range(1, 37):
        rev_own = ebitda_own = capex_m = 0.0
        rev_fr = royalty = fund = supply = fees = 0.0
        n_own = n_fr = 0
        for (m0,nm), r in own_runs.items():
            if m == m0: capex_m += r["capex_total"]
            if m >= m0:
                n_own += 1; row = r["rows"][m-m0]; rev_own += row["revenue"]; ebitda_own += row["ebitda"]
        for (m0,nm), r in fr_runs.items():
            if m == m0: fees += A["own_fee"]
            if m >= m0:
                n_fr += 1; row = r["rows"][m-m0]; rev_fr += row["revenue"]
                if m-m0 >= 1: royalty += max(A["own_royalty"]*row["revenue"], A["own_royalty_min"])
                fund += A["own_mkt_fund"]*row["revenue"]; supply += row["consumables"]*A["supply_margin"]
        uk = 0.0
        if m >= 7: uk = (A["uk_payroll"]+A["uk_other"])*(0.5 if m < 13 else 1.0)
        if m == 10: capex_m += A["uk_setup"]
        uk_income = royalty+fund+supply+fees
        uk_tax = max(A["tt_rate"]*uk_income - A["tt_deduct"]*uk, A["tt_floor"]*uk_income) if uk_income>0 else 0.0
        net = ebitda_own + uk_income - fund - uk - uk_tax   # фонд тратится на маркетинг сети
        cf = net - capex_m
        cum += cf; cum_inv += capex_m
        rows.append(dict(month=m, n_own=n_own, n_fr=n_fr, rev_own=rev_own, rev_fr=rev_fr, ebitda_own=ebitda_own, royalty=royalty, fund=fund,
                         supply=supply, fees=fees, uk=uk, uk_tax=uk_tax, net=net, capex=capex_m, cf=cf, cum=cum))
    return rows, cum_inv
NET_ROWS, NET_INV = network_consolidated()
fr_capex_total = sum(1 for o in OPENINGS if o[1]=="fr")*R_OWN_FR["capex_total"]

# --- сравнение с Чио Чио РФ (оценка структуры)
def chio_rf():
    rev = 50*600*DAYS          # 50 клиентов × 600 ₽ (прайс Казани 500–800) — ≈ 913 тыс. ₽ (между 564 тыс. и 1,1 млн)
    fot = 0.40*rev; rent = 50000; roy = 0.04*rev; cons = 0.04*rev; mkt = 10000+0.02*rev; util = 15000; tax = 0.03*rev  # ПСН+взносы ИП ≈3% (оценка)
    other = 0.02*rev
    opex = fot+rent+roy+cons+mkt+util+tax+other
    return dict(revenue=rev, fot=fot, rent=rent, royalty=roy, consumables=cons, marketing=mkt, utilities=util, tax=tax, other=other, opex=opex, ebitda=rev-opex)
CHIO = chio_rf()

# ----------------------------------------------------------------------------------------------
# 4. ПРОВЕРКИ
# ----------------------------------------------------------------------------------------------
def verify(res):
    errs = []
    cum = -res["capex_total"]
    for r in res["rows"]:
        lines = r["fot"]+r["rent"]+r["manager"]+r["consumables"]+r["utilities"]+r["it"]+r["hdm"]+r["acquiring"]+r["marketing"]+r["accounting"]+r["other"]+r["royalty"]+r["turnover_tax"]+r["ip_fixed"]
        if abs(lines - r["opex"]) > 0.01: errs.append(f"m{r['month']}: opex mismatch")
        if abs(r["revenue"]-r["opex"]-r["ebitda"]) > 0.01: errs.append(f"m{r['month']}: EBITDA mismatch")
        if abs(r["ebitda"]-r["depreciation"]-r["net_profit"]) > 0.01: errs.append(f"m{r['month']}: net mismatch")
        tt_rule = max(0.10*r["revenue"]-0.06*r["documented"], 0.045*r["revenue"])
        if abs(tt_rule - r["turnover_tax"]) > 0.01: errs.append(f"m{r['month']}: turnover tax rule violated")
        if r["turnover_tax"] < 0.045*r["revenue"]-0.01: errs.append(f"m{r['month']}: tax below 4.5% floor")
        cum += r["cf"]
        if abs(cum - r["cum_cf"]) > 0.01: errs.append(f"m{r['month']}: cum CF mismatch")
    if abs(sum(res["capex_items"].values()) - res["capex_total"]) > 0.01: errs.append("capex sum mismatch")
    return errs
ALL_ERRS = {}
for k, v in list(R.items())+list(VAR.items())+[("own_fr",R_OWN_FR)]:
    e = verify(v)
    if e: ALL_ERRS[k] = e
# сеть: накопленный поток
c = 0.0
for r in NET_ROWS:
    c += r["cf"]
    if abs(c - r["cum"]) > 0.01: ALL_ERRS["network"] = ["cum mismatch"]
# пример из файла 3 (§4.3): 2,5 млн выручки, 1,65 млн документированных → 151 000
ex = max(0.10*2_500_000 - 0.06*1_650_000, 0.045*2_500_000)
if abs(ex-151000) > 0.01: ALL_ERRS["file3_example"] = [f"got {ex}"]
print("ПРОВЕРКИ:", "OK — все тождества сходятся" if not ALL_ERRS else ALL_ERRS)

# ----------------------------------------------------------------------------------------------
# 5. CSV
# ----------------------------------------------------------------------------------------------
CSV_COLS = ["month","load","clients_day","check","revenue","fot","rent","manager","consumables","utilities","it","hdm","acquiring","marketing","accounting","other","royalty","documented","turnover_tax","ip_fixed","opex","ebitda","depreciation","net_profit","cf","cum_cf","cum_dcf"]
for name, res in R.items():
    path = os.path.join(HERE, f"pnl_36m_{name}.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(CSV_COLS)
        for r in res["rows"]:
            w.writerow([round(r[c], 4) if c in ("load",) else round(r[c], 1) if c in ("clients_day",) else round(r[c]) for c in CSV_COLS])
    print("CSV:", path)

# ----------------------------------------------------------------------------------------------
# 6. EXCEL (openpyxl, с формулами)
# ----------------------------------------------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
wb = Workbook()
ws = wb.active; ws.title = "Assumptions"
bold = Font(bold=True); hdr_fill = PatternFill("solid", fgColor="DDEBF7"); inp_fill = PatternFill("solid", fgColor="FFF2CC")
ws.append(["Параметр","Значение","Ед.","Источник","Достоверность","Комментарий"])
for c in ws[1]: c.font = bold; c.fill = hdr_fill
CELL = {}
for i, (k, v, u, s, cf, cm) in enumerate(ASSUMPTIONS, start=2):
    ws.append([k, v, u, s, cf, cm]); CELL[k] = f"Assumptions!$B${i}"
    ws.cell(row=i, column=2).fill = inp_fill
# параметры сценария (ссылки для PnL)
r0 = ws.max_row + 2
ws.cell(row=r0, column=1, value="ПАРАМЕТРЫ БАЗОВОГО СЦЕНАРИЯ (меняйте жёлтые ячейки)").font = bold
scen = [("sc_chairs", P_BASE["chairs"], "кресел"),("sc_masters", P_BASE["masters"], "мастеров"),("sc_check", round(check_A,2), "AMD средний чек (из листа Prices)"),
        ("sc_price_mult", 1.0, "множитель цены"),("sc_load_mature", P_BASE["load_mature"], "зрелая загрузка"),("sc_rent_sqm", P_BASE["rent_sqm"], "AMD/м²"),
        ("sc_area", P_BASE["area"], "м²"),("sc_fot_share", P_BASE["fot_share"], "доля ФОТ"),("sc_royalty", 0.0, "роялти, доля"),("sc_royalty_min", 0.0, "роялти минимум AMD"),
        ("sc_franchise_fee", 0.0, "паушальный AMD"),("sc_consumables_per_client", P_BASE["consumables_per_client"], "расходники AMD/клиент"),("sc_manager", P_BASE["manager_cost"], "управляющий AMD/мес"),
        ("sc_fot_documented", 1, "1 = ФОТ документирован для вычета 6%"),("sc_capacity_day", round(cap3,3), "ёмкость услуг/день при 3 креслах (формула ниже)")]
for j, (k, v, u) in enumerate(scen):
    rr = r0+1+j
    ws.cell(row=rr, column=1, value=k); ws.cell(row=rr, column=2, value=v).fill = inp_fill; ws.cell(row=rr, column=3, value=u)
    CELL[k] = f"Assumptions!$B${rr}"
# ёмкость — формулой
ws[CELL["sc_capacity_day"].split("!")[1].replace("$","")] = f"={CELL['sc_chairs']}*{CELL['hours_per_day']}*60/{CELL['minutes_avg']}"
ws[CELL["sc_check"].split("!")[1].replace("$","")] = "=Prices!C12"
for col, w_ in zip("ABCDEF", (28,16,14,80,22,60)): ws.column_dimensions[col].width = w_

# Prices
wp = wb.create_sheet("Prices")
wp.append(["Услуга","Цена A (спальный)","Цена B (центр/ТЦ)","Доля в миксе","Мин/услуга","Вклад A","Вклад B","Чио Чио РФ Казань 2026, ₽","В AMD по 4,3731"])
for c in wp[1]: c.font = bold; c.fill = hdr_fill
chio_px = {"mens":700,"clip":500,"kids":500,"wshort":800,"wlong":800,"beard":300,"style":None,"roots":2000}
for i, s in enumerate(SERVICES, start=2):
    wp.append([s[1], s[2], s[3], s[4], s[5], f"=B{i}*D{i}", f"=C{i}*D{i}", chio_px[s[0]], f"=H{i}*{CELL['fx_rub']}" if chio_px[s[0]] else None])
    for col in (2,3,4,5): wp.cell(row=i, column=col).fill = inp_fill
n = len(SERVICES)+1
wp.append([]); wp.append(["Сумма долей", None, None, f"=SUM(D2:D{n})"])
wp.append(["СРЕДНИЙ ЧЕК, AMD", None, f"=SUM(F2:F{n})", None, None, None, f"=SUM(G2:G{n})"])  # row 12: C12 = A, G12 = B
wp.cell(row=12, column=1).font = bold
wp.append(["Средний чек в USD", None, f"=C12/{CELL['fx_usd']}", None, None, None, f"=G12/{CELL['fx_usd']}"])
wp.append(["Средний чек в RUB", None, f"=C12/{CELL['fx_rub']}", None, None, None, f"=G12/{CELL['fx_rub']}"])
wp.append(["Средневзвеш. мин/услуга", None, f"=SUMPRODUCT(D2:D{n},E2:E{n})"])
for col, w_ in zip("ABCDEFGHI", (40,16,16,12,10,12,12,22,16)): wp.column_dimensions[col].width = w_

# CAPEX (формулы)
wc = wb.create_sheet("CAPEX")
wc.append(["Статья","Своя сеть, 40 м² (AMD)","Чио Чио «Лайт» (AMD)","Чио Чио стандарт (AMD)","Своя сеть 25 м², 2 кресла (AMD)","Своя сеть 40 м², USD"])
for c in wc[1]: c.font = bold; c.fill = hdr_fill
capex_rows = [
 ("Ремонт/отделка", f"={CELL['sc_area']}*{CELL['fitout_sqm']}", f"={CELL['sc_area']}*{CELL['fitout_sqm']}", f"={CELL['sc_area']}*{CELL['fitout_sqm']}", f"=25*{CELL['fitout_sqm']}"),
 ("Кресла", f"={CELL['sc_chairs']}*{CELL['chair_price']}", f"={CELL['sc_chairs']}*{CELL['chair_price']}", f"={CELL['sc_chairs']}*{CELL['chair_price']}", f"=2*{CELL['chair_price']}"),
 ("Рабочие посты", f"={CELL['sc_chairs']}*{CELL['post_price']}", f"={CELL['sc_chairs']}*{CELL['post_price']}", f"={CELL['sc_chairs']}*{CELL['post_price']}", f"=2*{CELL['post_price']}"),
 ("Мойка (1 шт.)", f"={CELL['wash_unit']}", f"={CELL['wash_unit']}", f"={CELL['wash_unit']}", f"={CELL['wash_unit']}"),
 ("Вакуумные аппараты × 2", f"=2*{CELL['vacuum_unit']}", f"=2*{CELL['vacuum_unit']}", f"=2*{CELL['vacuum_unit']}", f"=2*{CELL['vacuum_unit']}"),
 ("Инструмент/техника", f"={CELL['sc_chairs']}*{CELL['tools_per_post']}", f"={CELL['sc_chairs']}*{CELL['tools_per_post']}", f"={CELL['sc_chairs']}*{CELL['tools_per_post']}", f"=2*{CELL['tools_per_post']}"),
 ("Стерилизатор", f"={CELL['sterilizer']}", f"={CELL['sterilizer']}", f"={CELL['sterilizer']}", f"={CELL['sterilizer']}"),
 ("Мебель ожидания/стойка", f"={CELL['waiting_furniture']}", f"={CELL['waiting_furniture']}", f"={CELL['waiting_furniture']}", f"={CELL['waiting_furniture']}"),
 ("Вывеска", f"={CELL['signage']}", f"={CELL['signage']}", f"={CELL['signage']}", f"={CELL['signage']}"),
 ("Брендинг интерьера", f"={CELL['branding_interior']}", f"={CELL['branding_interior']}", f"={CELL['branding_interior']}", f"={CELL['branding_interior']}"),
 ("Киоск/ПК/камеры", f"={CELL['selfservice_it']}", f"={CELL['selfservice_it']}", f"={CELL['selfservice_it']}", f"={CELL['selfservice_it']}"),
 ("Касса e-HDM", f"={CELL['hdm_setup']}", f"={CELL['hdm_setup']}", f"={CELL['hdm_setup']}", f"={CELL['hdm_setup']}"),
]
for i, row in enumerate(capex_rows, start=2):
    wc.append([row[0], row[1], row[2], row[3], row[4], f"=B{i}/{CELL['fx_usd']}"])
hard_end = len(capex_rows)+1
wc.append(["Итого hard-CAPEX", f"=SUM(B2:B{hard_end})", f"=SUM(C2:C{hard_end})", f"=SUM(D2:D{hard_end})", f"=SUM(E2:E{hard_end})", f"=B{hard_end+1}/{CELL['fx_usd']}"])
hr = hard_end+1; wc.cell(row=hr, column=1).font = bold
soft_rows = [
 ("Непредвиденное", f"=B{hr}*{CELL['contingency']}", f"=C{hr}*{CELL['contingency']}", f"=D{hr}*{CELL['contingency']}", f"=E{hr}*{CELL['contingency']}"),
 ("CRM/сайт", f"={CELL['crm_setup']}", f"={CELL['crm_setup']}", f"={CELL['crm_setup']}", f"={CELL['crm_setup']}"),
 ("Регистрация ИП", f"={CELL['registration']}", f"={CELL['registration']}", f"={CELL['registration']}", f"={CELL['registration']}"),
 ("Юр./налоговое оформление", f"={CELL['legal_setup']}", f"={CELL['legal_setup']}", f"={CELL['legal_setup']}", f"={CELL['legal_setup']}"),
 ("Медкнижки", f"={CELL['medbooks']}", f"={CELL['medbooks']}", f"={CELL['medbooks']}", f"={CELL['medbooks']}*3/4"),
 ("Депозит аренды", f"={CELL['deposit_months']}*{CELL['sc_rent_sqm']}*{CELL['sc_area']}", f"={CELL['deposit_months']}*{CELL['sc_rent_sqm']}*{CELL['sc_area']}", f"={CELL['deposit_months']}*{CELL['sc_rent_sqm']}*{CELL['sc_area']}", f"={CELL['deposit_months']}*{CELL['sc_rent_sqm']}*25"),
 ("Маркетинг открытия", f"={CELL['launch_marketing']}", f"={CELL['launch_marketing']}", f"={CELL['launch_marketing']}", f"={CELL['launch_marketing']}"),
 ("Паушальный взнос", "0", f"={CELL['fee_chio_lite_rub']}*{CELL['fx_rub']}", f"={CELL['fee_chio_std_rub']}*{CELL['fx_rub']}", "0"),
 ("Оборотный резерв (расчёт в Python, см. PnL)", round(R["base"]["reserve"]), round(VAR["franchise_lite"]["reserve"]), round(VAR["franchise_std"]["reserve"]), round(VAR["2chairs"]["reserve"])),
]
for j, row in enumerate(soft_rows):
    i = hr+1+j
    wc.append([row[0], row[1], row[2], row[3], row[4], f"=B{i}/{CELL['fx_usd']}"])
se = hr+len(soft_rows)
wc.append(["ИТОГО CAPEX", f"=B{hr}+SUM(B{hr+1}:B{se})", f"=C{hr}+SUM(C{hr+1}:C{se})", f"=D{hr}+SUM(D{hr+1}:D{se})", f"=E{hr}+SUM(E{hr+1}:E{se})", f"=B{se+1}/{CELL['fx_usd']}"])
tot_r = se+1; wc.cell(row=tot_r, column=1).font = bold
wc.append(["Итого CAPEX, USD", f"=B{tot_r}/{CELL['fx_usd']}", f"=C{tot_r}/{CELL['fx_usd']}", f"=D{tot_r}/{CELL['fx_usd']}", f"=E{tot_r}/{CELL['fx_usd']}"])
wc.append(["Итого CAPEX, RUB", f"=B{tot_r}/{CELL['fx_rub']}", f"=C{tot_r}/{CELL['fx_rub']}", f"=D{tot_r}/{CELL['fx_rub']}", f"=E{tot_r}/{CELL['fx_rub']}"])
wc.append(["Амортизируемая база (hard + непредвиденное + CRM)", f"=B{hr}+B{hr+1}+B{hr+2}", f"=C{hr}+C{hr+1}+C{hr+2}+C{hr+8}", f"=D{hr}+D{hr+1}+D{hr+2}+D{hr+8}", f"=E{hr}+E{hr+1}+E{hr+2}"])
CAPEX_TOTAL_CELL = f"CAPEX!$B${tot_r}"; DEPR_BASE_CELL = f"CAPEX!$B${tot_r+3}"
for col, w_ in zip("ABCDEF", (44,22,22,22,26,18)): wc.column_dimensions[col].width = w_

# PnL_36m (формулы)
wl = wb.create_sheet("PnL_36m")
cols = ["Мес.","Индекс цен","Загрузка","Клиентов/день","Средний чек","ВЫРУЧКА","ФОТ мастеров","Аренда","Управл.+уборка","Расходники","Коммуналка","IT/связь","Касса","Эквайринг","Маркетинг","Бухгалтерия","Прочее","Роялти","Документ. расходы","Налог с оборота","Фиксы ИП","ИТОГО РАСХОДЫ","EBITDA","Амортизация","ЧИСТАЯ ПРИБЫЛЬ","CF","Накопл. CF","Дисконт-фактор","Накопл. DCF"]
wl.append(cols)
for c in wl[1]: c.font = bold; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True)
# строка 2: t=0 — CAPEX
wl.append([0]+[None]*24+[f"=-{CAPEX_TOTAL_CELL}", f"=Z2", 1, f"=Z2"])
r_m = f"((1+{CELL['discount_rate']})^(1/12)-1)"
for m in range(1, 37):
    i = m+2
    load = round(R["base"]["rows"][m-1]["load"]/P_BASE["load_mature"], 6)  # относительная форма кривой
    wl.append([
        m,
        f"=(1+{CELL['cpi']})^INT((A{i}-1)/12)",
        f"=MIN(0.95,{load}*{CELL['sc_load_mature']})",
        f"=C{i}*{CELL['sc_capacity_day']}",
        f"={CELL['sc_check']}*{CELL['sc_price_mult']}*B{i}",
        f"=D{i}*E{i}*{CELL['days_per_month']}",
        f"=MAX({CELL['sc_fot_share']}*F{i},{CELL['sc_masters']}*{CELL['fot_floor_per_master']}*B{i})",
        f"={CELL['sc_rent_sqm']}*{CELL['sc_area']}*B{i}",
        f"=({CELL['sc_manager']}+{CELL['cleaning_cost']})*B{i}",
        f"=MAX({CELL['sc_consumables_per_client']}*B{i}*D{i}*{CELL['days_per_month']},{CELL['consumables_floor']}*B{i})",
        f"={CELL['utilities']}*B{i}",
        f"={CELL['it_crm']}*B{i}",
        f"={CELL['hdm_monthly']}",
        f"=F{i}*{CELL['card_share']}*{CELL['acq_rate']}+{CELL['acq_terminal']}",
        f"=MAX({CELL['marketing_pct']}*F{i},{CELL['marketing_floor']}*B{i})",
        f"={CELL['accounting']}*B{i}",
        f"=MAX({CELL['other_pct']}*F{i},{CELL['other_floor']}*B{i})",
        f"=IF(AND({CELL['sc_royalty']}>0,A{i}>=2),MAX({CELL['sc_royalty']}*F{i},{CELL['sc_royalty_min']}),0)",
        f"=H{i}+I{i}+J{i}+K{i}+L{i}+M{i}+N{i}+O{i}+P{i}+R{i}+Q{i}*{CELL['other_documented_share']}+IF({CELL['sc_fot_documented']}=1,G{i},0)",
        f"=MAX({CELL['tt_rate']}*F{i}-{CELL['tt_deduct']}*S{i},{CELL['tt_floor']}*F{i})",
        f"={CELL['ip_fixed']}",
        f"=SUM(G{i}:R{i})+T{i}+U{i}",
        f"=F{i}-V{i}",
        f"=({DEPR_BASE_CELL}+{CELL['sc_franchise_fee']})/{CELL['depr_months']}",
        f"=W{i}-X{i}",
        f"=W{i}",
        f"=AA{i-1}+Z{i}",
        f"=1/(1+{r_m})^A{i}",
        f"=AC{i-1}+Z{i}*AB{i}",
    ])
wl.append([])
tr = wl.max_row+1
wl.cell(row=tr, column=1, value="Год 1").font = bold; wl.cell(row=tr+1, column=1, value="Год 2").font = bold; wl.cell(row=tr+2, column=1, value="Год 3").font = bold
for y in range(3):
    a, b = 3+12*y, 14+12*y
    for col in range(6, 26):
        L = get_column_letter(col)
        wl.cell(row=tr+y, column=col, value=f"=SUM({L}{a}:{L}{b})")
wl.cell(row=tr+4, column=1, value="NPV (ставка Assumptions.discount_rate)").font = bold
wl.cell(row=tr+4, column=6, value=f"=Z2+SUMPRODUCT(Z3:Z38,AB3:AB38)")
wl.cell(row=tr+5, column=1, value="IRR годовая").font = bold
wl.cell(row=tr+5, column=6, value="=(1+IRR(Z2:Z38))^12-1")
wl.cell(row=tr+6, column=1, value="Окупаемость простая, мес.").font = bold
wl.cell(row=tr+6, column=6, value='=IFERROR(INDEX(A3:A38,MATCH(TRUE,INDEX(AA3:AA38>=0,0),0)),"нет за 36 мес.")')
wl.cell(row=tr+7, column=1, value="Окупаемость дисконтированная, мес.").font = bold
wl.cell(row=tr+7, column=6, value='=IFERROR(INDEX(A3:A38,MATCH(TRUE,INDEX(AC3:AC38>=0,0),0)),"нет за 36 мес.")')
wl.cell(row=tr+8, column=1, value="Проверка: Σ(выручка − расходы − EBITDA) по 36 мес. (должно быть 0)")
wl.cell(row=tr+8, column=6, value="=SUM(F3:F38)-SUM(V3:V38)-SUM(W3:W38)")
wl.cell(row=tr+9, column=1, value="Проверка: мин(налог − 4,5% выручки) ≥ 0")
wl.cell(row=tr+9, column=6, value="=MIN(INDEX(T3:T38-0.045*F3:F38,0))")
wl.column_dimensions["A"].width = 8
for col in range(2, 30): wl.column_dimensions[get_column_letter(col)].width = 14
wl.freeze_panes = "B2"

# Scenarios (значения из Python)
wsx = wb.create_sheet("Scenarios")
wsx.append(["Показатель","Пессимистичный","Базовый","Оптимистичный"]);
for c in wsx[1]: c.font = bold; c.fill = hdr_fill
def scen_rows():
    out = [("Параметры: цена (×), зрелая загрузка, аренда AMD/м², ФОТ", "0,9 / 32% / 7 500 / 50%", "1,0 / 45% / 6 500 / 50%", "1,0 / 55% / 6 500 / 45%")]
    for lab, key in (("Выручка год 1","revenue"),("Выручка год 2","revenue"),("Выручка год 3","revenue")):
        pass
    for y in (1,2,3):
        for lab, key in (("Выручка","revenue"),("ФОТ","fot"),("Налог с оборота","turnover_tax"),("EBITDA","ebitda"),("Чистая прибыль","net_profit")):
            out.append((f"{lab}, год {y}", *(round(R[s][f"y{y}"][key]) for s in ("pessimistic","base","optimistic"))))
        out.append((f"Клиентов/день, среднее за год {y}", *(round(R[s][f"y{y}"]["clients_day"],1) for s in ("pessimistic","base","optimistic"))))
    out.append(("CAPEX всего (с резервом)", *(round(R[s]["capex_total"]) for s in ("pessimistic","base","optimistic"))))
    out.append(("Окупаемость простая, мес.", *(R[s]["payback"] or "нет за 36" for s in ("pessimistic","base","optimistic"))))
    out.append(("Окупаемость дисконтированная, мес.", *(R[s]["dpayback"] or "нет за 36" for s in ("pessimistic","base","optimistic"))))
    out.append(("NPV 36 мес. @18%", *(round(R[s]["npv"]) for s in ("pessimistic","base","optimistic"))))
    out.append(("IRR годовая", *((f"{R[s]['irr']*100:.1f}%" if R[s]['irr'] is not None else "n/a") for s in ("pessimistic","base","optimistic"))))
    return out
for row in scen_rows(): wsx.append(list(row))
wsx.column_dimensions["A"].width = 48
for col in "BCD": wsx.column_dimensions[col].width = 26

# Sensitivity
wsn = wb.create_sheet("Sensitivity")
wsn.append(["Фактор","−20%","−10%","база","+10%","+20%","Показатель"])
for c in wsn[1]: c.font = bold; c.fill = hdr_fill
for s in SENS:
    wsn.append([s["factor"]]+[round(s[d]["net36"]) for d in (-0.2,-0.1,0,0.1,0.2)]+["Чистая прибыль за 36 мес., AMD"])
    wsn.append([s["factor"]]+[s[d]["payback"] or ">36" for d in (-0.2,-0.1,0,0.1,0.2)]+["Окупаемость, мес."])
wsn.column_dimensions["A"].width = 16; wsn.column_dimensions["G"].width = 34

# Network
wn = wb.create_sheet("Network")
wn.append(["Мес.","Собств. точек","Франч. точек","Выручка собств.","Выручка франч.","EBITDA собств.","Роялти","Маркет. фонд","Маржа поставок","Паушальные","Расходы УК","Налог УК","Чистый результат","CAPEX","CF","Накопл. CF"])
for c in wn[1]: c.font = bold; c.fill = hdr_fill
for r in NET_ROWS:
    wn.append([r["month"], r["n_own"], r["n_fr"]]+[round(r[k]) for k in ("rev_own","rev_fr","ebitda_own","royalty","fund","supply","fees","uk","uk_tax","net","capex","cf","cum")])
for col in range(1, 17): wn.column_dimensions[get_column_letter(col)].width = 15

wb.calculation.fullCalcOnLoad = True
XLSX = os.path.join(HERE, "ArmChio_model.xlsx")
wb.save(XLSX); print("XLSX:", XLSX)

# ----------------------------------------------------------------------------------------------
# 7. JSON с ключевыми числами
# ----------------------------------------------------------------------------------------------
KEY = dict(check_A=check_A, check_B=check_B, capacity_3=cap3, capacity_2=cap2,
           base=dict(m36=R["base"]["m36"], y1=R["base"]["y1"], y2=R["base"]["y2"], y3=R["base"]["y3"], capex=R["base"]["capex_total"], payback=R["base"]["payback"], dpayback=R["base"]["dpayback"], npv=R["base"]["npv"], irr=R["base"]["irr"]),
           pess=dict(capex=R["pessimistic"]["capex_total"], payback=R["pessimistic"]["payback"], npv=R["pessimistic"]["npv"], irr=R["pessimistic"]["irr"], y3=R["pessimistic"]["y3"]),
           opt=dict(capex=R["optimistic"]["capex_total"], payback=R["optimistic"]["payback"], npv=R["optimistic"]["npv"], irr=R["optimistic"]["irr"], y3=R["optimistic"]["y3"]),
           breakeven=BE, franchise_lite=dict(capex=VAR["franchise_lite"]["capex_total"], payback=VAR["franchise_lite"]["payback"], npv=VAR["franchise_lite"]["npv"]),
           franchise_std=dict(capex=VAR["franchise_std"]["capex_total"], payback=VAR["franchise_std"]["payback"], npv=VAR["franchise_std"]["npv"]),
           network=dict(yerevan_base=yerevan_base, yerevan_aggr=yerevan_aggr, regions=regions_base, total_inv=NET_INV, cum36=NET_ROWS[-1]["cum"]),
           errors=ALL_ERRS)
with open(os.path.join(HERE, "model_output.json"), "w", encoding="utf-8") as f: json.dump(KEY, f, ensure_ascii=False, indent=1, default=str)

# ----------------------------------------------------------------------------------------------
# 8. ОТЧЁТ (Markdown)
# ----------------------------------------------------------------------------------------------
def T(headers, rows):
    s = "| " + " | ".join(headers) + " |\n|" + "---|"*len(headers) + "\n"
    for r in rows: s += "| " + " | ".join(str(x) for x in r) + " |\n"
    return s
def money(x): return fmt(x)
def pct(x, d=1): return f"{x*100:.{d}f}%"
b = R["base"]; pe = R["pessimistic"]; op = R["optimistic"]
m36b = b["m36"]
share_men = need_men; share_kids = need_kids

md = []
md.append(f"""# ArmChio — финансовая модель адаптации формата «Чио Чио» под Армению

Дата: 20.10.2026. Роль: финансовый аналитик / CFO франчайзинговой сети. Валюта модели — AMD; курс ЦБА 19.10.2026: 1 USD = {A['fx_usd']} AMD, 1 RUB = {A['fx_rub']} AMD.
Источники входных данных — только четыре файла исследования ({F1}, {F2}, {F3}, repo_inventory.md) плюс явно помеченные допущения и 4 веб-проверки (Росстат-2026, alteg.io, цены вакуумных аппаратов, условия мастер-франшиз). Репозиторий ArmChio пуст (repo_inventory.md) — модель строится с нуля.
Все расчёты выполнены скриптом `model/model.py` (python -I), результаты проверены тождествами (выручка − расходы = EBITDA; EBITDA − амортизация = чистая прибыль; накопленный поток; правило налога с оборота «10% − 6% документированных расходов, но не менее 4,5%» — в т.ч. на контрольном примере файла 3 §4.3: 2,5 млн / 1,65 млн → 151 000 AMD). Результат проверок: **{'OK — все тождества сходятся' if not ALL_ERRS else ALL_ERRS}**.

Файлы: `C:\\Новая папка\\ArmChio-analysis\\repo\\model\\model.py`, `ArmChio_model.xlsx` (листы Assumptions / Prices / CAPEX / PnL_36m — на формулах; Scenarios / Sensitivity / Network — значения), `pnl_36m_base.csv`, `pnl_36m_pessimistic.csv`, `pnl_36m_optimistic.csv`, `model_output.json`.

## Резюме (базовый сценарий: 3 кресла, 40 м², спальный район Еревана, своя сеть, ФОТ 50%)

| Показатель | Значение |
|---|---|
| Средний чек (спальный / центр) | **{money(check_A)} / {money(check_B)} AMD** (${usd(check_A):.2f} / ${usd(check_B):.2f}; {rub(check_A):.0f} / {rub(check_B):.0f} ₽) |
| Ёмкость точки (3 кресла × 11 ч, {A['minutes_avg']} мин/услуга) | {cap3:.0f} услуг/день; 2 кресла — {cap2:.0f} |
| Зрелая точка (мес. 36): загрузка / клиентов в день | {pct(m36b['load'],0)} / **{m36b['clients_day']:.0f}** |
| Выручка зрелой точки | **{money(m36b['revenue'])} AMD/мес** (${usd(m36b['revenue']):,.0f}; {rub(m36b['revenue'])/1000:,.0f} тыс. ₽) |
| EBITDA зрелой точки | **{money(m36b['ebitda'])} AMD/мес ({pct(m36b['ebitda']/m36b['revenue'])})** |
| Чистая прибыль зрелой точки (после амортизации и всех налогов) | **{money(m36b['net_profit'])} AMD/мес ({pct(m36b['net_profit']/m36b['revenue'])})** ≈ ${usd(m36b['net_profit']):,.0f} |
| Год 1 / 2 / 3: выручка | {money(b['y1']['revenue'])} / {money(b['y2']['revenue'])} / {money(b['y3']['revenue'])} AMD |
| Год 1 / 2 / 3: чистая прибыль | {money(b['y1']['net_profit'])} / {money(b['y2']['net_profit'])} / {money(b['y3']['net_profit'])} AMD |
| CAPEX (своя сеть 40 м², с депозитом, маркетингом открытия и резервом) | **{money(b['capex_total'])} AMD ≈ ${usd(b['capex_total']):,.0f} ≈ {rub(b['capex_total'])/1e6:.2f} млн ₽** |
| CAPEX франшиза Чио Чио «Лайт» / стандарт | {money(VAR['franchise_lite']['capex_total'])} / {money(VAR['franchise_std']['capex_total'])} AMD (${usd(VAR['franchise_lite']['capex_total']):,.0f} / ${usd(VAR['franchise_std']['capex_total']):,.0f}) |
| Безубыточность (чистая прибыль = 0, структура мес. 36) | **{BE['clients_day']:.0f} клиентов/день**, выручка {money(BE['revenue'])} AMD/мес, загрузка {pct(BE['load'],0)}; EBITDA-безубыточность {BE['clients_day_ebitda']:.0f} кл./день |
| Окупаемость простая / дисконтированная @18% (база) | **{b['payback']} / {b['dpayback']} мес.** |
| Окупаемость: пессимистичный / оптимистичный | {pe['payback'] or 'нет за 36 мес.'} / {op['payback']} мес. |
| NPV 36 мес. @18% (без терминальной стоимости): пессим. / база / оптим. | {money(pe['npv'])} / **{money(b['npv'])}** / {money(op['npv'])} AMD |
| IRR (36 мес.): пессим. / база / оптим. | {(f"{pe['irr']*100:.0f}%" if pe['irr'] is not None else 'n/a')} / **{b['irr']*100:.0f}%** / {op['irr']*100:.0f}% |
| Рекомендуемый вариант | **(3) собственный локальный бренд** (3 собственные точки + лёгкая франшиза 1,5 млн AMD / 3%); (1) франшиза Чио Чио — окупаемость {VAR['franchise_lite']['payback']}–{VAR['franchise_std']['payback']} мес. и NPV {money(VAR['franchise_lite']['npv'])} / {money(VAR['franchise_std']['npv'])} при тех же операциях; (2) мастер-франшиза — не покрывает УК при ≤20 точках |
| Ёмкость сети | Ереван **{yerevan_base}** точек (база) / {yerevan_aggr} (агрессивно); регионы +{regions_base}; Армения ≈ **{yerevan_base+regions_base}–{yerevan_aggr+regions_base+2}** |

""")

# --- Таблица допущений
md.append("## 0. Таблица допущений\n\nКолонки: значение | источник (файл/URL) | достоверность (факт / оценка / допущение / предложение).\n\n")
md.append(T(["Параметр","Значение","Ед.","Источник","Достоверность","Комментарий"],
            [(k, (fmt(v) if isinstance(v,int) or (isinstance(v,float) and v==int(v)) else (f"{v:.4f}".rstrip("0").rstrip(".") if isinstance(v,float) else v)), u, s, c, cm) for k,v,u,s,c,cm in ASSUMPTIONS]))

# --- A
md.append(f"""
## A. Прайс-лист ArmChio (AMD)

**Позиционирование.** По файлу 2: реальное дно рынка Еревана — 1 000–1 500 AMD («павильонный» мастер, Нор-Норк), медиана эконом-сегмента (n=27) — 3 000, медиана всей выборки — 4 000, среднее ≈ 4 900; ближайший аналог формата True Cost — 3 000 за 30 мин с мытьём; барбершопы 6 000–7 000. Файл 3: аренда 40 м² в спальном районе 260 000 AMD/мес, т.е. при цене 1 500 одна только аренда требует 173 стрижек/мес, при 2 500 — 104. Поэтому базовая мужская стрижка ставится на **2 500 AMD в спальном районе (−17% к медиане эконом, +67% к дну)** и **3 000 в центре/ТЦ (= True Cost, но 15 минут и без записи)**. Это $6,9 / $8,3 — в 2–2,5 раза дороже российской маркетинговой цены Чио Чио 300 ₽ (1 312 AMD), но ниже реального казанского прайса 700 ₽ (3 061 AMD).

Правило фикс-прайса «любая стрижка — одна цена» в чистом виде в Армении **не рекомендуется**: армянский рынок дифференцирует машинку (2 000–4 000) и модельную (3 000–5 000), а Чио Чио в 2026 г. сама ушла от единой цены (Казань: 700 муж. / 800 жен. / 500 дет.). Используем «короткий фикс-прайс»: 4 цены стрижек + 3 допуслуги.

""")
md.append(T(["Услуга","A: спальный район","B: центр/ТЦ","Доля в миксе (допущение)","Мин/услуга","Рынок Ереван эконом (файл 2)","Чио Чио Казань 2026, ₽ → AMD"],
            [(s[1], fmt(s[2]), fmt(s[3]), pct(s[4],0), s[5],
              {"mens":"1 500–3 000 (медиана 3 000)","clip":"2 000–4 000","kids":"2 000–3 000","wshort":"1 500–3 000","wlong":"2 000–4 000","beard":"1 500–3 000","style":"1 000–2 000","roots":"2 000–2 900 (эконом) / 5 000–10 000 (центр)"}[s[0]],
              (f"{chio_px[s[0]]} → {fmt(chio_px[s[0]]*RUB)}" if chio_px[s[0]] else "—")) for s in SERVICES]
            + [("**Средний чек (взвешенный по миксу)**", f"**{money(check_A)}**", f"**{money(check_B)}**", "100%", f"{A['minutes_avg']:.1f}", "медиана эконом 3 000", "оценка чека 500–700 ₽ → 2 187–3 061")]))
chio_check_rub = 600
md.append(f"""
Обоснование микса (допущение, проверять по CRM через 3 месяца): ЦА Чио Чио — мужчины, дети, пенсионеры (файл 1 §5); по прокси ЮMoney (файл 2 §3) 39% мужчин стригутся в эконом-парикмахерских; женщины в Ереване массово пользуются салонами с мытьём и укладкой (файл 2 §3, Flacon), поэтому их доля в формате без мытья ограничена 13%; окрашивание корней — опция (3%), требует одной мойки.

**Сравнение с Чио Чио РФ.**

| Метрика | Чио Чио РФ (Казань 2026) | ArmChio A (спальный) | ArmChio B (центр) |
|---|---|---|---|
| Базовая муж. стрижка | 700 ₽ = {fmt(700*RUB)} AMD | 2 500 AMD = {2500/RUB:.0f} ₽ | 3 000 AMD = {3000/RUB:.0f} ₽ |
| Маркетинговая цена «от» | 300 ₽ = {fmt(300*RUB)} AMD | 2 000 AMD (машинка) = {2000/RUB:.0f} ₽ | 2 500 AMD = {2500/RUB:.0f} ₽ |
| Средний чек | ~{chio_check_rub} ₽ (оценка файла 1: 500–700) = {fmt(chio_check_rub*RUB)} AMD | {money(check_A)} AMD = {rub(check_A):.0f} ₽ | {money(check_B)} AMD = {rub(check_B):.0f} ₽ |
| Средняя з/п | {fmt(A['avg_wage_ru'])} ₽ (Росстат, янв-2026) | {fmt(A['avg_wage_am'])} AMD (Armstat, авг-2026) | — |
| Паритет: базовая стрижка / средняя з/п | 700/{fmt(A['avg_wage_ru'])} = **{700/A['avg_wage_ru']*100:.2f}%** (300 ₽ → {300/A['avg_wage_ru']*100:.2f}%) | 2 500/{fmt(A['avg_wage_am'])} = **{2500/A['avg_wage_am']*100:.2f}%** | 3 000 → {3000/A['avg_wage_am']*100:.2f}% |
| Паритет: средний чек / средняя з/п | {chio_check_rub/A['avg_wage_ru']*100:.2f}% | {check_A/A['avg_wage_am']*100:.2f}% | {check_B/A['avg_wage_am']*100:.2f}% |
| Стрижек на среднюю з/п | {A['avg_wage_ru']/700:.0f} | {A['avg_wage_am']/2500:.0f} | {A['avg_wage_am']/3000:.0f} |

Вывод: по паритету ArmChio в спальном районе на {(2500/A['avg_wage_am'])/(700/A['avg_wage_ru'])*100-100:.0f}% «дороже» для местного кошелька, чем реальный Чио Чио в Казани, но в 2,8 раза дешевле армянской медианы всей выборки (4 000) относительно дохода. Запас на снижение до 2 000 есть (см. чувствительность), но тогда аренда съедает {A['rent_sqm_sleep']*40/2000:.0f} стрижек/мес.

## B. Операционная модель точки

| Параметр | Значение | Обоснование |
|---|---|---|
| Часы работы | 10:00–21:00, 7/7 (11 ч) | файл 1: Казань 09–21; файл 2: Ереван 10–20/22 |
| Норматив времени | {A['minutes_avg']:.1f} мин/услуга (взвешено: стрижка 15–20, борода 10, окрашивание 45) | файл 1: 15 мин, 4–5 стрижек/час |
| Ёмкость | 3 кресла: **{cap3:.0f} услуг/день**; 2 кресла: {cap2:.0f} | 11 ч × 60 / {A['minutes_avg']:.1f} × кресла |
| Персонал | 4 мастера на 3 кресла (231 кресло-часов/нед.), 0,5 управляющего, уборщица | без администратора — киоск самообслуживания + CRM |
| Зрелая загрузка (база) | 45% → {cap3*0.45:.0f} кл./день | файл 1: норма зрелой точки 40–70 стрижек/день на 2–3 кресла (= 35–60% ёмкости); для неизвестного бренда берём нижнюю треть |

**Кривая выхода на мощность** (доля от зрелой загрузки; из кейсов файла 1: Екатеринбург — 15 тыс. ₽ прибыли в 1-й месяц, spark — рост 10–15 → 90–100 стрижек/день за 9 месяцев, партнёр 2021 — «первые 2 недели клиентов почти не было», СПб — окупаемость первой точки 2,5 года; поправка на незнакомый бренд: Чио Чио выходит на норму за 3–6 мес., ArmChio — за 12–24):

""")
md.append(T(["Период","Мес. 1","Мес. 2","Мес. 3","Мес. 4–6","Мес. 7–12","Год 2","Год 3"],
            [("Доля от зрелой загрузки", pct(LOAD_SHAPE[0],0), pct(LOAD_SHAPE[1],0), pct(LOAD_SHAPE[2],0), f"{pct(LOAD_SHAPE[3],0)}→{pct(LOAD_SHAPE[5],0)}", f"{pct(LOAD_SHAPE[6],0)}→{pct(LOAD_SHAPE[11],0)}", f"{pct(LOAD_SHAPE[12],0)}→{pct(LOAD_SHAPE[23],0)}", f"{pct(LOAD_SHAPE[24],0)}→100%"),
             ("Загрузка кресла (база 45%)", *(pct(LOAD_SHAPE[i]*0.45,0) for i in (0,1,2)), f"{pct(LOAD_SHAPE[3]*0.45,0)}→{pct(LOAD_SHAPE[5]*0.45,0)}", f"{pct(LOAD_SHAPE[6]*0.45,0)}→{pct(LOAD_SHAPE[11]*0.45,0)}", f"{pct(LOAD_SHAPE[12]*0.45,0)}→{pct(LOAD_SHAPE[23]*0.45,0)}", f"{pct(LOAD_SHAPE[24]*0.45,0)}→45%"),
             ("Клиентов/день (3 кресла)", *(f"{b['rows'][i]['clients_day']:.0f}" for i in (0,1,2)), f"{b['rows'][3]['clients_day']:.0f}→{b['rows'][5]['clients_day']:.0f}", f"{b['rows'][6]['clients_day']:.0f}→{b['rows'][11]['clients_day']:.0f}", f"{b['rows'][12]['clients_day']:.0f}→{b['rows'][23]['clients_day']:.0f}", f"{b['rows'][24]['clients_day']:.0f}→{b['rows'][35]['clients_day']:.0f}"),
             ("Выручка, тыс. AMD/мес", *(f"{b['rows'][i]['revenue']/1000:.0f}" for i in (0,1,2)), f"{b['rows'][3]['revenue']/1000:.0f}→{b['rows'][5]['revenue']/1000:.0f}", f"{b['rows'][6]['revenue']/1000:.0f}→{b['rows'][11]['revenue']/1000:.0f}", f"{b['rows'][12]['revenue']/1000:.0f}→{b['rows'][23]['revenue']/1000:.0f}", f"{b['rows'][24]['revenue']/1000:.0f}→{b['rows'][35]['revenue']/1000:.0f}")]))
md.append(f"""
**Рынок сверху вниз (зона обслуживания радиусом {A['catchment_radius_km']} км = {area_catch:.2f} км²).** Плотность районов — файл 3 §3.1. Доля мужчин 15–64 ≈ 31% и детей 0–14 ≈ 20% — допущения (возрастная структура вне файлов). Частота: мужчины 10 стрижек/год (3–5 недель; диапазон 8–12), дети 6/год.

""")
md.append(T(["Район","Население 2025","Плотн., чел/км²","Жителей в радиусе 1,2 км","Муж. стрижек/день в зоне","Детских/день в зоне","Нужная доля рынка муж. для {0:.0f} кл./день*".format(m36b['clients_day'])],
            [(n, fmt(p), fmt(d), fmt(c), f"{mc:.0f}", f"{kc:.0f}", pct(need_men/mc) if mc>0 else "—") for n,p,d,c,mc,kc in TOPDOWN]))
md.append(f"""
\\* Зрелая точка делает {m36b['clients_day']:.0f} услуг/день, из них мужских (стрижка+машинка+борода) 68% = {need_men:.0f}, детских 12% = {need_kids:.0f}. В Нор-Норке/Канакер-Зейтуне/Арабкире это требует ≈{pct(need_men/TOPDOWN[2][4])} рынка мужских стрижек зоны при ≈{salons_per_capita*TOPDOWN[2][3]:.0f} конкурирующих парикмахерских в радиусе (2ГИС: 1 на 720 жителей) — то есть одна точка ArmChio должна собирать поток примерно {need_men/TOPDOWN[2][4]*salons_per_capita*TOPDOWN[2][3]:.1f} средних районных парикмахерских, что соответствует заявке Чио Чио «60 клиентов/день против 30 у обычного салона» (файл 1). В Малатии-Себастии (плотность 5 966) требуемая доля {pct(need_men/TOPDOWN[0][4])}, в Шенгавите/Эребуни ({TOPDOWN[1][2]}/{TOPDOWN[4][2]} чел/км²) — {pct(need_men/TOPDOWN[1][4])}/{pct(need_men/TOPDOWN[4][4])} — там нужна точка у метро/рынка, а не «у дома». Снизу вверх: 4 мастера × 12 клиентов/день = 48 — совпадает с оценкой файла 3 §2.1 (12–18 стрижек/день на мастера). Пессимистичный сценарий (32% → {pe['m36']['clients_day']:.0f} кл./день) соответствует «средней» точке Чио Чио (564 тыс. ₽ ≈ {564000*RUB/1e6:.2f} млн AMD).

## C. Расходы точки в месяц

Зрелый месяц (36) базового сценария и альтернативы. ФОТ: в Армении доминирует 50/50 (файл 3 §2.1) — это **база**; 40% «как у Чио Чио» реально только при выручке на мастера ≥ 400 тыс. AMD gross (иначе мастер уйдёт на 50% в соседний салон) — учтено как оптимистичный вариант. Налоги с ФОТ (20% подоходный + 5% пенсионный + 1 000 гербовый + ОМС) удерживаются **из** gross, работодатель сверх gross не платит (файл 3 §2.4) — поэтому стоимость мастера для точки = gross. На ramp-up действует гарантия 150 000 gross/мастер (файл 3: минималка 75 000, оклады 200–400 тыс.).

""")
def cost_rows(res_list, labels):
    keys = [("Выручка","revenue"),("ФОТ мастеров","fot"),("Аренда","rent"),("Управляющий 0,5 + уборка","manager"),("Расходники/дезинфекция (95 AMD/клиент)","consumables"),("Коммуналка","utilities"),("IT/CRM/связь","it"),("Касса e-HDM","hdm"),("Эквайринг (50% безнал × 1,2% + терминал)","acquiring"),("Маркетинг","marketing"),("Бухгалтерия","accounting"),("Прочее/резерв","other"),("Роялти","royalty"),("Налог с оборота","turnover_tax"),("Фиксы ИП (подоходный, соц., гербовый, ОМС)","ip_fixed"),("ИТОГО расходы","opex"),("EBITDA","ebitda"),("Амортизация","depreciation"),("Чистая прибыль","net_profit")]
    rows = []
    for lab, k in keys:
        rows.append([lab] + [f"{money(r['m36'][k])} ({pct(r['m36'][k]/r['m36']['revenue'])})" if k not in ("revenue",) else money(r['m36'][k]) for r in res_list])
    return T(["Статья, AMD/мес (мес. 36)"]+labels, rows)
md.append(cost_rows([R["base"], VAR["center"], VAR["center_1line"], VAR["fot40"], VAR["2chairs"]],
                    ["База: спальный, 6 500/м², ФОТ 50%","Центр 8 500/м², прайс B, загр. 50%","Центр 1-я линия 13 000/м², прайс B, загр. 55%","Спальный, ФОТ 40% (как Чио Чио)","Спальный, 2 кресла, 25 м²"]))
m36n = VAR["fot_not_documented"]["m36"]
FLOOR_MONTHS = ", ".join(str(r["month"]) for r in b["rows"] if abs(r["turnover_tax"]-0.045*r["revenue"]) < 0.5) or "ни в одном"
md.append(f"""
**Налог с оборота — проверка правила.** Мес. 36 базы: 10% × {money(m36b['revenue'])} = {money(0.1*m36b['revenue'])}; документированные расходы {money(m36b['documented'])} × 6% = {money(0.06*m36b['documented'])}; разница {money(0.1*m36b['revenue']-0.06*m36b['documented'])} против пола 4,5% = {money(0.045*m36b['revenue'])} → налог **{money(m36b['turnover_tax'])} ({pct(m36b['turnover_tax']/m36b['revenue'],2)} выручки)**. Если ФОТ мастеров *не* принимается как документированный расход (риск, помеченный в файле 3 §4.2 «уточнить у бухгалтера»; при оплате мастеров «в конверте» — точно не принимается): налог {money(m36n['turnover_tax'])} ({pct(m36n['turnover_tax']/m36n['revenue'],2)}), чистая прибыль мес. 36 падает с {money(m36b['net_profit'])} до {money(m36n['net_profit'])} AMD. Пол 4,5% в базовом сценарии срабатывает в месяцах {FLOOR_MONTHS} (малая выручка на ramp-up при больших документированных расходах: гарантия ФОТ 600 000 + аренда + фикс), далее налог считается по формуле 10% − 6% и растёт до {pct(m36b['turnover_tax']/m36b['revenue'],2)} выручки (проверено скриптом и формулой Excel «мин(налог − 4,5% выручки) ≥ 0»).

**Роялти Чио Чио (вариант 1)** при выручке {money(m36b['revenue'])}: 4% = {money(0.04*m36b['revenue'])} AMD/мес (мин. 5 000 ₽ = {money(5000*RUB)}); кросс-граничные роялти в РФ дополнительно могут облагаться налогом у источника (не моделируется — требует проверки по НК РА/СИДН РФ–РА).

**Владелец-оператор.** Если собственник сам управляет точкой (как 93% франчайзи Чио Чио — впервые в бизнесе), статья «управляющий» 120 000 исчезает: чистая прибыль мес. 36 {money(VAR['owner_operator']['m36']['net_profit'])} AMD, окупаемость {VAR['owner_operator']['payback']} мес. (vs {b['payback']}).

## D. CAPEX открытия

""")
ci = b["capex_items"]; cl = VAR["franchise_lite"]["capex_items"]; cs = VAR["franchise_std"]["capex_items"]; c2 = VAR["2chairs"]["capex_items"]
keys_order = list(ci.keys())
rows = []
for k in keys_order:
    k2 = k
    if k.startswith("Ремонт"): k2 = [x for x in c2 if x.startswith("Ремонт")][0]
    elif "× 3" in k: k2 = k.replace("× 3", "× 2")
    rows.append((k, money(ci[k]), f"{usd(ci[k]):,.0f}", money(cl[k]), money(cs[k]), money(c2.get(k2, 0))))
rows.append(("**ИТОГО**", f"**{money(b['capex_total'])}**", f"**{usd(b['capex_total']):,.0f}**", f"**{money(VAR['franchise_lite']['capex_total'])}**", f"**{money(VAR['franchise_std']['capex_total'])}**", f"**{money(VAR['2chairs']['capex_total'])}**"))
rows.append(("Итого, USD", f"{usd(b['capex_total']):,.0f}", "", f"{usd(VAR['franchise_lite']['capex_total']):,.0f}", f"{usd(VAR['franchise_std']['capex_total']):,.0f}", f"{usd(VAR['2chairs']['capex_total']):,.0f}"))
rows.append(("Итого, тыс. RUB", f"{rub(b['capex_total'])/1000:,.0f}", "", f"{rub(VAR['franchise_lite']['capex_total'])/1000:,.0f}", f"{rub(VAR['franchise_std']['capex_total'])/1000:,.0f}", f"{rub(VAR['2chairs']['capex_total'])/1000:,.0f}"))
md.append(T(["Статья","Сценарий 1: своя сеть, 40 м², 3 кресла (AMD)","USD","Сценарий 2: франшиза Чио Чио «Лайт» 290 тыс. ₽","Сценарий 3: франшиза Чио Чио стандарт 590 тыс. ₽","Доп.: своя сеть 25 м², 2 кресла"], rows))
md.append(f"""
Комментарии: (1) hard-CAPEX {money(b['hard_capex'])} AMD (${usd(b['hard_capex']):,.0f}) лежит в коридоре файла 3 «5–7,5 млн без депозита» с учётом непредвиденного; (2) по курсу ArmChio стоит {rub(b['capex_total'])/1e6:.2f} млн ₽ — выше заявленных Чио Чио «от 0,4–1 млн», но на уровне фактических 1,15–1,7 млн ₽ (файл 1 §6 п. 2) с поправкой на ремонт в Ереване (100 тыс. AMD/м² ≈ 22,9 тыс. ₽/м²); (3) оборотный резерв — расчётная максимальная просадка накопленного операционного потока на ramp-up (мес. 1–{max((i for i,r in enumerate(b['rows'],1) if r['cum_cf']==min(x['cum_cf'] for x in b['rows'])), default=0)}); (4) паушальный взнос Чио Чио пересчитан по 4,3731: 290 000 ₽ = {money(290000*RUB)} AMD, 590 000 ₽ = {money(590000*RUB)} AMD — это {pct(290000*RUB/b['capex_total'])}–{pct(590000*RUB/b['capex_total'])} от CAPEX своей точки; (5) вакуумные аппараты: аналоги Sibel Eye Vac (£215–307) и Head Jog Hair Vacuum (£148,5) — профессиональные стационарные пылесосы для волос; аппарат Чио Чио не продаётся отдельно, его цена не публикуется.

## E. Три стратегических варианта

### Сравнение для инвестора (одна и та же точка, спальный район, база)

""")
def var_row(lbl, r):
    return (lbl, money(r["capex_total"]), f"{usd(r['capex_total']):,.0f}", money(r["m36"]["royalty"]), money(r["m36"]["net_profit"]), pct(r["m36"]["net_profit"]/r["m36"]["revenue"]), money(sum(x["net_profit"] for x in r["rows"])), r["payback"] or "нет за 36", r["dpayback"] or "нет за 36", money(r["npv"]), money(r["npv_tv"]), (f"{r['irr']*100:.0f}%" if r["irr"] is not None else "n/a"))
md.append(T(["Вариант","CAPEX, AMD","USD","Роялти мес. 36","Чистая прибыль мес. 36","Маржа","Чистая прибыль 36 мес.","Окуп., мес.","Диск. окуп.","NPV @18% (36 мес., без TV)","NPV @18% с TV = 3× EBITDA Y3","IRR 36 мес."],
            [var_row("(1a) Франчайзи Чио Чио «Лайт» (290 тыс. ₽ + 4%)", VAR["franchise_lite"]),
             var_row("(1b) Франчайзи Чио Чио стандарт (590 тыс. ₽ + 4%)", VAR["franchise_std"]),
             var_row("(3) Своя сеть — собственная точка (0 + 0%)", R["base"]),
             var_row("(3) Субфранчайзи ArmChio (1,5 млн AMD + 3%, фонд 1% внутри маркетинга 3%)", R_OWN_FR)]))
md.append(f"""
**(1) Единичный франчайзи Чио Чио в Армении.** Армении нет в списке стран сети (файл 2: Россия, Беларусь, Казахстан, Грузия-Батуми без подтверждённой точки, Молдова); зарубежные салоны «фактически автономны», цены не унифицированы, колл-центр/HR/приложение/эквайринг/ОФД — российские (файл 1 §7). Франчайзи получит бизнес-бук, дизайн-проект и мебель (с НДС 20% при ввозе в ЕАЭС), но не получит работающего IT-контура и маркетинга на армянском. Экономика: роялти 4% ≈ {money(0.04*m36b['revenue'])} AMD/мес + паушальный 1,27–2,58 млн AMD снижают NPV с {money(b['npv'])} до {money(VAR['franchise_lite']['npv'])} / {money(VAR['franchise_std']['npv'])} и удлиняют окупаемость с {b['payback']} до {VAR['franchise_lite']['payback']} / {VAR['franchise_std']['payback']} мес. Плюс: гарантия окупаемости «в договоре» (условия не раскрыты), рассрочка, бренд для 60–80 тыс. релокантов (5–7% населения Еревана). Минус: роялти с выручки при убытке, обязательная косметика с наценкой 120–130%, подсудность в РФ, 1/3 убыточных партнёров.

**(2) Мастер-франшиза Чио Чио на Армению.** У Чио Чио нет публичных условий мастер-франшизы (файл 1: «данных не найдено»; зарубежные точки — единичные, через обычные договоры). Отраслевой ориентир (franchising.com, global-franchise.com, sprintlaw.co.uk, 20.10.2026): мастер платит территориальный взнос (кратный паушальному; за субточку $5–60 тыс. по отраслям), роялти делятся 50/50…75/25 в пользу мастера, график открытий с санкциями, срок 10–20 лет. Расчёт при 50/50 от 4% и паушальном субфранчайзи ≈ «Лайт»:

""")
md.append(T(["Сеть, точек","Доля мастера в роялти, AMD/мес","Доля мастера в паушальных, AMD/мес (n/3 открытий в год)","Расходы УК мастера, AMD/мес","Результат мастера, AMD/мес"],
            [(m["n"], money(m["master_royalty"]), money(m["master_fees"]), money(m["uk_cost"]), money(m["net"])) for m in MASTER]))
md.append(f"""
Мастер-франшиза покрывает УК только при ~20 точках (2% сетевой выручки ≈ {money(0.02*m36b['revenue'])} AMD с точки), при этом мастер несёт расходы на локализацию IT и обучение сам, а бренд и стандарты остаются у Казани. Вариант имеет смысл только как переговорная опция «территориальный эксклюзив за символический взнос + помощь с мебелью/обучением»; как инвестиция — хуже варианта 3.

**(3) Собственный локальный бренд ArmChio.** Технология Чио Чио не защищена: «товарный знак, концепция и идея почти ничего не стоят, ценность — в процессах» (Набиуллин, файл 1 §7); ~45 клонов в РФ. Воспроизводимо: фикс-прайс, 15 минут, без мытья (пылесос), без администратора (киоск + CRM alteg.io/Fresha + Idram QR), мастера на %, микро-формат, бесплатные стрижки на открытии, «6-я бесплатно». Что нужно купить/создать: бренд-бук, стандарты и обучение (≈3 млн AMD разово), УК 3 человека.

**Условия своей франшизы (предложение).** Критерий: паушальный + роялти должны оставлять субфранчайзи **чистую маржу ≥ 10% и простую окупаемость ≤ 36 мес.** в базовом сценарии (собственная точка окупается за {b['payback']} мес. при марже {pct(m36b['net_profit']/m36b['revenue'])}; в РФ Чио Чио при 4% оставляет 20–35% EBITDA-маржи на зрелой точке, но там нет 50% ФОТ и ~5,5% оборотного налога). Маркетинговый фонд 1% не добавляется сверху, а выделяется из стандартного бюджета маркетинга 3% (2% локально + 1% централизованно). Перебор ставок при паушальном 1,5 млн AMD:

""")
md.append(T(["Роялти","Чистая прибыль субфранчайзи мес. 36","Маржа","Окупаемость, мес.","NPV @18%","Роялти УК с точки, AMD/мес"],
            [(f"**3% (предложение)**", money(R_OWN_FR["m36"]["net_profit"]), pct(R_OWN_FR["m36"]["net_profit"]/R_OWN_FR["m36"]["revenue"]), R_OWN_FR["payback"] or ">36", money(R_OWN_FR["npv"]), money(R_OWN_FR["m36"]["royalty"])),
             (f"4%", money(R_OWN_FR4["m36"]["net_profit"]), pct(R_OWN_FR4["m36"]["net_profit"]/R_OWN_FR4["m36"]["revenue"]), R_OWN_FR4["payback"] or ">36", money(R_OWN_FR4["npv"]), money(R_OWN_FR4["m36"]["royalty"])),
             (f"5%", money(R_OWN_FR6["m36"]["net_profit"]), pct(R_OWN_FR6["m36"]["net_profit"]/R_OWN_FR6["m36"]["revenue"]), R_OWN_FR6["payback"] or ">36", money(R_OWN_FR6["npv"]), money(R_OWN_FR6["m36"]["royalty"]))]))
md.append(f"""
Предложение: паушальный **{money(A['own_fee'])} AMD** (≈${usd(A['own_fee']):,.0f}, {rub(A['own_fee'])/1000:.0f} тыс. ₽; включает обучение, дизайн-проект, запуск), роялти **3% от выручки, мин. {money(A['own_royalty_min'])} AMD со 2-го месяца**, маркетинговый фонд 1% внутри бюджета 3%, централизованная закупка расходников с маржой УК 20%. Субфранчайзи сохраняет {pct(R_OWN_FR['m36']['net_profit']/R_OWN_FR['m36']['revenue'])} чистой маржи и окупается за {R_OWN_FR['payback'] or '>36'} мес. При 4% окупаемость {R_OWN_FR4['payback'] or '>36'} мес. и маржа {pct(R_OWN_FR4['m36']['net_profit']/R_OWN_FR4['m36']['revenue'])} — предел; при 5% ({R_OWN_FR6['payback'] or '>36'} мес., {pct(R_OWN_FR6['m36']['net_profit']/R_OWN_FR6['m36']['revenue'])}) франшиза не продаётся. Важно: NPV за строгие 36 мес. без терминальной стоимости отрицателен для **любого** субфранчайзи (и для франчайзи Чио Чио) — горизонт короче срока жизни точки; с терминальной стоимостью 3× EBITDA года 3 NPV субфранчайзи при 3% = {money(R_OWN_FR['npv_tv'])} AMD. Роялти 3% ниже Чио Чио (4%), потому что армянская точка структурно беднее на ~15 п.п. EBITDA-маржи (раздел G).

**Модель франчайзера (вариант 3), стационарное состояние, все точки зрелые (мес. 36 базы), 3 собственные + остальные франчайзинговые:**

""")
md.append(T(["Точек в сети","Франч.","Роялти, AMD/мес","Маркет. фонд (транзит)","Маржа поставок","Паушальные (откр./год → /мес)","EBITDA 3 собств. точек","Расходы УК (2–3 чел., офис, IT, маркетинг) + фонд","Налог УК","Результат УК без собств. точек","Результат холдинга"],
            [(f["n"], f["fr"], money(f["royalty"]), money(f["mkt_fund"]), money(f["supply"]), money(f["fees"]), money(f["own_profit"]), money(f["uk_cost"]), money(f["uk_tax"]), money(f["uk_net_ex_own"]), money(f["total"])) for f in FRANCHISOR]))
md.append(f"""
Вывод: при роялти 3% одна франчайзинговая точка приносит УК {money(FR_PT_INCOME)} AMD/мес (роялти + маржа поставок); фиксированные расходы УК {money(A['uk_payroll']+A['uk_other'])} AMD/мес (≈${usd(A['uk_payroll']+A['uk_other']):,.0f}) окупаются текущими доходами (без паушальных) только при **≈{N_BE_UK:.0f} франчайзинговых точках**; при 10 точках (7 франч., 4 открытия/год) результат УК без собственных точек {money(FRANCHISOR[1]['uk_net_ex_own'])} AMD/мес, при 20 (17 франч., 6 открытий/год) — {money(FRANCHISOR[2]['uk_net_ex_own'])} AMD/мес, из которых паушальные {money(FRANCHISOR[2]['fees'])} — т.е. франчайзинговая надстройка в Армении выходит в плюс только на сети 10+ точек и зависит от темпа продаж франшиз (ёмкость страны ~28–40 точек, разд. H, исчерпывается за 3–4 года). Экономический смысл варианта 3 на горизонте 3 лет — в **3 собственных точках** (EBITDA {money(FRANCHISOR[0]['own_profit'])} AMD/мес) плюс паушальные при росте; франшиза — инструмент быстрого захвата ёмкости до прихода копий формата, а центром прибыли становится лишь на зрелой сети. УК должна быть минимальной (2 человека), первые 3 точки — собственные.

## F. P&L 36 месяцев — базовый сценарий (3 кресла, спальный район Еревана, своя сеть)

Индексация цен и фиксированных расходов +4,3% с 13-го и 25-го месяца (ИПЦ). Все суммы — AMD. Полная таблица по статьям — `pnl_36m_base.csv` и лист PnL_36m в Excel.

""")
md.append(T(["Мес.","Загр.","Кл./день","Выручка","ФОТ","Аренда","Налог с об.","Прочие опер.","EBITDA","Аморт.","Чистая прибыль","CF накопл."],
            [(r["month"], pct(r["load"],0), f"{r['clients_day']:.0f}", money(r["revenue"]), money(r["fot"]), money(r["rent"]), money(r["turnover_tax"]),
              money(r["opex"]-r["fot"]-r["rent"]-r["turnover_tax"]), money(r["ebitda"]), money(r["depreciation"]), money(r["net_profit"]), money(r["cum_cf"])) for r in b["rows"]]))
md.append("\n**Сводные годовые P&L по сценариям**\n\n")
def yr_block():
    hdr = ["Показатель","Пессим. Y1","Пессим. Y2","Пессим. Y3","База Y1","База Y2","База Y3","Оптим. Y1","Оптим. Y2","Оптим. Y3"]
    rows = []
    for lab, k in (("Выручка","revenue"),("ФОТ","fot"),("Аренда","rent"),("Налог с оборота","turnover_tax"),("Итого расходы","opex"),("EBITDA","ebitda"),("Амортизация","depreciation"),("Чистая прибыль","net_profit")):
        rows.append([lab]+[money(R[s][f"y{y}"][k]) for s in ("pessimistic","base","optimistic") for y in (1,2,3)])
    rows.append(["Клиентов/день (среднее)"]+[f"{R[s][f'y{y}']['clients_day']:.0f}" for s in ("pessimistic","base","optimistic") for y in (1,2,3)])
    rows.append(["EBITDA-маржа"]+[pct(R[s][f"y{y}"]["ebitda"]/R[s][f"y{y}"]["revenue"]) for s in ("pessimistic","base","optimistic") for y in (1,2,3)])
    return T(hdr, rows)
md.append(yr_block())
md.append(f"""
Параметры сценариев: **пессимистичный** — цена −10% (чек {money(check_A*0.9)}), зрелая загрузка 32% ({pe['m36']['clients_day']:.0f} кл./день — уровень «средней» точки Чио Чио), аренда 7 500/м², ФОТ 50%, расходники 120 AMD/клиент; **базовый** — чек {money(check_A)}, 45%, 6 500/м², ФОТ 50%; **оптимистичный** — 55% ({op['m36']['clients_day']:.0f} кл./день), ФОТ 45%, остальное как база.

""")
md.append(T(["Метрика","Пессимистичный","Базовый","Оптимистичный"],
            [("CAPEX (с резервом на убытки ramp-up)", money(pe["capex_total"]), money(b["capex_total"]), money(op["capex_total"])),
             ("Резерв в составе CAPEX", money(pe["reserve"]), money(b["reserve"]), money(op["reserve"])),
             ("Безубыточность (чистая прибыль = 0): клиентов/день", f"{BE_PESS['clients_day']:.0f}" if BE_PESS else "—", f"{BE['clients_day']:.0f}", f"{BE_OPT['clients_day']:.0f}"),
             ("Безубыточность: выручка/мес", money(BE_PESS["revenue"]) if BE_PESS else "—", money(BE["revenue"]), money(BE_OPT["revenue"])),
             ("EBITDA-безубыточность: клиентов/день", f"{BE_PESS['clients_day_ebitda']:.0f}" if BE_PESS else "—", f"{BE['clients_day_ebitda']:.0f}", f"{BE_OPT['clients_day_ebitda']:.0f}"),
             ("Окупаемость простая, мес.", pe["payback"] or "нет за 36 мес.", b["payback"], op["payback"]),
             ("Окупаемость дисконтированная @18%, мес.", pe["dpayback"] or "нет за 36 мес.", b["dpayback"], op["dpayback"]),
             ("NPV 36 мес. @15% / @18% / @20%", f"{money(pe['npv15'])} / {money(pe['npv'])} / {money(pe['npv20'])}", f"{money(b['npv15'])} / {money(b['npv'])} / {money(b['npv20'])}", f"{money(op['npv15'])} / {money(op['npv'])} / {money(op['npv20'])}"),
             ("NPV @18% с терминальной стоимостью 3× EBITDA года 3 (справочно)", money(pe["npv_tv"]), money(b["npv_tv"]), money(op["npv_tv"])),
             ("IRR годовая (36 мес., без терминальной стоимости)", f"{pe['irr']*100:.1f}%" if pe["irr"] is not None else "отриц.", f"{b['irr']*100:.1f}%", f"{op['irr']*100:.1f}%"),
             ("Накопленный CF на мес. 36", money(pe["rows"][-1]["cum_cf"]), money(b["rows"][-1]["cum_cf"]), money(op["rows"][-1]["cum_cf"]))]))
md.append(f"""
Ставка дисконтирования 18%: ставка рефинансирования ЦБ РА 6,75% (файл 3) — безрисковый якорь в AMD; доходность гособлигаций РА ориентировочно 9–10% (вне файлов, допущение); премия за риск малого несетевого бизнеса с единственной точкой, неликвидность и валютный риск (аренда де-факто в USD) 8–10 п.п. → 15–20%, берём 18%; NPV показан также при 15% и 20%. Терминальная стоимость не учитывается (консервативно); депозит аренды (2 мес.) возвращается после горизонта и в NPV не включён.

Безубыточность базы: при структуре мес. 36 точка выходит в ноль по чистой прибыли на **{BE['clients_day']:.0f} клиентах/день ({money(BE['revenue'])} AMD/мес, загрузка {pct(BE['load'],0)})**; по EBITDA — {BE['clients_day_ebitda']:.0f} кл./день. Для сравнения, витрина Чио Чио «16 стрижек × 300 ₽» и расчётные 25–37 для крупных городов РФ (файл 1): в Армении порог выше из-за 50% ФОТ, аренды и оборотного налога; чек {money(check_A)} AMD ({rub(check_A):.0f} ₽) сопоставим с реальным казанским 2026 (500–700 ₽), но вдвое выше маркетинговых 300 ₽ Чио Чио. Центр (8 500/м², прайс B): безубыточность {BE_C['clients_day']:.0f} кл./день; 2 кресла/25 м²: {BE_2['clients_day']:.0f} кл./день.

## G. Чувствительность

Чистая прибыль за 36 месяцев (AMD) и окупаемость (мес.) при изменении одного фактора относительно базы (база: {money(sum(x['net_profit'] for x in b['rows']))} AMD, {b['payback']} мес.).

""")
md.append(T(["Фактор","−20%","−10%","База","+10%","+20%"],
            sum([[ (f"{s['factor']}: чистая прибыль 36 мес.", *(money(s[d]['net36']) for d in (-0.2,-0.1,0.0,0.1,0.2))),
                   (f"{s['factor']}: окупаемость, мес.", *((s[d]['payback'] or '>36') for d in (-0.2,-0.1,0.0,0.1,0.2))),
                   (f"{s['factor']}: NPV @18%", *(money(s[d]['npv']) for d in (-0.2,-0.1,0.0,0.1,0.2)))] for s in SENS], [])))
md.append("\n**«Торнадо» (размах чистой прибыли 36 мес. при ±20%):**\n\n")
md.append(T(["Ранг","Фактор","−20%","+20%","Размах, AMD"], [(i+1, t[0], money(t[1]), money(t[2]), money(abs(t[2]-t[1]))) for i, t in enumerate(TORNADO)]))
md.append(f"""
Три критических фактора: **{TORNADO[0][0]}, {TORNADO[1][0]}, {TORNADO[2][0]}**. Цена и загрузка действуют почти симметрично (обе через выручку, ФОТ-процент снимает половину прироста); цена чуть выгоднее — расходники и время мастера на клиента не растут; аренда — слабейший фактор (±20% аренды = ±{money(0.2*A['rent_sqm_sleep']*40)} AMD/мес), что подтверждает: в Ереване не нужно экономить на локации, нужно экономить на ФОТ-схеме и держать цену. Доля ФОТ ±20% (40%↔60% выручки) — это разница между прибыльной и околонулевой точкой.

**Почему маржа в Армении ниже, чем у Чио Чио в РФ** (структура зрелой точки, % выручки; РФ — оценка на базе файла 1: 50 клиентов × 600 ₽, ФОТ 40%, аренда 50 тыс. ₽, роялти 4%, ПСН+взносы ≈3%):

""")
rf = CHIO; am = m36b
md.append(T(["Статья","Чио Чио РФ (оценка)","ArmChio база (мес. 36)","Разница, п.п."],
            [("Выручка/мес", f"{fmt(rf['revenue'])} ₽ = {money(rf['revenue']*RUB)} AMD", f"{money(am['revenue'])} AMD = {rub(am['revenue'])/1000:,.0f} тыс. ₽", ""),
             ("ФОТ мастеров", pct(rf['fot']/rf['revenue']), pct(am['fot']/am['revenue']), f"{(am['fot']/am['revenue']-rf['fot']/rf['revenue'])*100:+.1f}"),
             ("Аренда", pct(rf['rent']/rf['revenue']), pct(am['rent']/am['revenue']), f"{(am['rent']/am['revenue']-rf['rent']/rf['revenue'])*100:+.1f}"),
             ("Налоги бизнеса (ПСН+взносы vs налог с оборота+фиксы ИП)", pct(rf['tax']/rf['revenue']), pct((am['turnover_tax']+am['ip_fixed'])/am['revenue']), f"{((am['turnover_tax']+am['ip_fixed'])/am['revenue']-rf['tax']/rf['revenue'])*100:+.1f}"),
             ("Роялти", pct(rf['royalty']/rf['revenue']), "0% (своя сеть)", f"{(0-rf['royalty']/rf['revenue'])*100:+.1f}"),
             ("Управляющий/уборка/бухгалтерия", "0% (владелец сам, бухгалтерия в «прочем»)", pct((am['manager']+am['accounting'])/am['revenue']), f"{((am['manager']+am['accounting'])/am['revenue'])*100:+.1f}"),
             ("Расходники, маркетинг, коммуналка, IT, эквайринг, прочее", pct((rf['consumables']+rf['marketing']+rf['utilities']+rf['other'])/rf['revenue']), pct((am['consumables']+am['marketing']+am['utilities']+am['it']+am['hdm']+am['acquiring']+am['other'])/am['revenue']), f"{((am['consumables']+am['marketing']+am['utilities']+am['it']+am['hdm']+am['acquiring']+am['other'])/am['revenue']-(rf['consumables']+rf['marketing']+rf['utilities']+rf['other'])/rf['revenue'])*100:+.1f}"),
             ("**EBITDA-маржа**", f"**{pct(rf['ebitda']/rf['revenue'])}**", f"**{pct(am['ebitda']/am['revenue'])}**", f"**{(am['ebitda']/am['revenue']-rf['ebitda']/rf['revenue'])*100:+.1f}**")]))
md.append(f"""
Разбор: (а) **ФОТ 50% vs 40%** — рынок труда Еревана диктует 50/50 (файл 3), это −10 п.п.; (б) **аренда**: 260 000 AMD = ${usd(260000):,.0f} за 40 м² против 50 тыс. ₽ (= {money(50000*RUB)} AMD, ${usd(50000*RUB):,.0f}) — в долларах ереванская аренда на {(260000/(50000*RUB)-1)*100:.0f}% дороже региональной российской при чеке на {(1-check_A/(600*RUB))*100:.0f}% ниже казанского (2 345 AMD против 600 ₽ = {money(600*RUB)} AMD), и номинирована де-факто в USD (файл 3: объявления в $) — валютный риск, которого нет у российского франчайзи; (в) **налог с оборота 10%−6%** даёт эффективные {pct(am['turnover_tax']/am['revenue'],1)} + фиксы ИП против патента (ПСН) в РФ ≈1–3% — Чио Чио сама строит витрину на ПСН («экономия 1,29 млн ₽/год против УСН 6%», файл 1 §1); микро-режим 0% парикмахерским в Армении закрыт с 2025 (файл 3 §4.1); (г) мы считаем управляющего и бухгалтера, а модель УК Чио Чио — нет. Частично компенсирует: отсутствие роялти (+4 п.п.) и более высокий валютный чек.

## H. Ёмкость сети и график открытий

**Ереван.** Норматив «жителей на точку»: Казань (родина сети) — 41 точка на ~1,3 млн = 1 на ~32 тыс. при признанной каннибализации (файл 1 §6 п. 6); Ереван плотнее по конкуренции (1 парикмахерская на 720 жителей vs 1 на 1 340 в РФ, файл 2 §2), поэтому база **1 точка на 50 тыс.**, агрессивно 1 на 35 тыс.; Кентрон +1 за дневной трафик; районы с плотностью <3 000 чел/км² (Эребуни, Шенгавит, Норк-Мараш, Нубарашен) −1 (формат «у дома» требует пешего трафика).

""")
md.append(T(["Район","Население 01.01.2025","Плотность, чел/км²","Точек (база, 1/50 тыс.)","Точек (агрессивно, 1/35 тыс.)","Приоритет"],
            [(n, fmt(p), fmt(d), bn, an, {"Нор-Норк":"1 — пилот (плотность 9 717, цены дна 1 000–1 500 → ниша фикс-прайса)","Малатия-Себастия":"2 — крупнейший район","Ачапняк":"3","Арабкир":"4 — релоканты, 7 300/м²","Кентрон":"5 — прайс B, трафик","Канакер-Зейтун":"6","Шенгавит":"у метро/рынка","Эребуни":"у метро/рынка","Давташен":"7 — релоканты","Аван":"8","Норк-Мараш":"—","Нубарашен":"—"}[n]) for n,p,d,bn,an in NETCAP]
            + [("**Итого Ереван**", fmt(sum(x[1] for x in NETCAP)), "", f"**{yerevan_base}**", f"**{yerevan_aggr}**", "")]))
md.append(f"""
**Регионы** (население — ориентир переписи 2022, вне файлов; число парикмахерских — 2ГИС, файл 2 §1.4; цены дна 1 000–1 500 AMD означают, что прайс A придётся снизить до ~2 000 — отдельный региональный прайс):

""")
md.append(T(["Город","Население (ориентир)","Парикмахерских в 2ГИС","Точек ArmChio (база)","Комментарий"],
            [(n, fmt(p), s, k, {"Гюмри":"2-й город; 2 точки (центр + Муш/Ани); цены 1 500–2 000","Ванадзор":"2ГИС почти не покрывает (3 точки) — полевой обход обязателен","Абовян":"пригород Еревана, 1 точка","Вагаршапат":"дно рынка 1 000 — только при чеке ≤2 000"}[n]) for n,p,s,k in REGIONS]
            + [("**Итого регионы**", "", "", f"**{regions_base}**", f"Армения всего: {yerevan_base+regions_base} (база) – {yerevan_aggr+regions_base+2} (агрессивно)")]))
md.append(f"""
**График открытий на 3 года (вариант 3):** собственные точки — мес. 1 (Нор-Норк, пилот), 7 (Малатия-Себастия), 10 (Ачапняк); с 13-го месяца — субфранчайзи по 1 точке каждые ~2 месяца в год 2 (6 точек) и ~1 в месяц в год 3 (11 точек, включая Гюмри, Абовян, Вагаршапат). Итого к мес. 36: **{NET_ROWS[-1]['n_own']} собственных + {NET_ROWS[-1]['n_fr']} франчайзинговых = {NET_ROWS[-1]['n_own']+NET_ROWS[-1]['n_fr']} точек** — это ~80% базовой ёмкости Армении; дальнейший рост — только за счёт агрессивного сценария и каннибализации. УК включается с 7-го месяца (50% штата), с 13-го — полностью; разовые расходы на систему франшизы (3 млн AMD) — в 10-м месяце.

**Консолидированный P&L сети (холдинг: собственные точки полностью + доходы УК от франчайзи − расходы УК), AMD, без индексации:**

""")
md.append(T(["Мес.","Собств.","Франч.","Выручка собств.","Выручка франч. (справочно)","EBITDA собств.","Роялти","Фонд","Маржа поставок","Паушальные","Расходы УК","Налог УК","Чистый результат холдинга","CAPEX холдинга","Накопл. CF"],
            [(r["month"], r["n_own"], r["n_fr"], money(r["rev_own"]), money(r["rev_fr"]), money(r["ebitda_own"]), money(r["royalty"]), money(r["fund"]), money(r["supply"]), money(r["fees"]), money(r["uk"]), money(r["uk_tax"]), money(r["net"]), money(r["capex"]), money(r["cum"])) for r in NET_ROWS]))
yr_net = [sum(r["net"] for r in NET_ROWS[(y-1)*12:y*12]) for y in (1,2,3)]
yr_rev = [sum(r["rev_own"]+r["rev_fr"] for r in NET_ROWS[(y-1)*12:y*12]) for y in (1,2,3)]
min_cum = min(r["cum"] for r in NET_ROWS); min_cum_m = [r["month"] for r in NET_ROWS if r["cum"]==min_cum][0]
md.append(f"""
Итоги сети: инвестиции холдинга (3 собственные точки по {money(b['capex_total'])} + система франшизы 3 млн) = **{money(NET_INV)} AMD ≈ ${usd(NET_INV):,.0f}**; максимальная потребность в финансировании (минимум накопленного потока) **{money(-min_cum)} AMD (${usd(-min_cum):,.0f}) на мес. {min_cum_m}**; накопленный CF холдинга на мес. 36: **{money(NET_ROWS[-1]['cum'])} AMD** ({'окупился' if NET_ROWS[-1]['cum']>0 else 'не окупился'} за 3 года; собственные точки ещё не все зрелые). Суммарная выручка сети (собств. + франч.) по годам: {money(yr_rev[0])} / {money(yr_rev[1])} / {money(yr_rev[2])} AMD; чистый результат холдинга по годам: {money(yr_net[0])} / {money(yr_net[1])} / {money(yr_net[2])} AMD. Инвестиции субфранчайзи (не холдинга): {sum(1 for o in OPENINGS if o[1]=='fr')} × {money(R_OWN_FR['capex_total'])} = {money(fr_capex_total)} AMD. Суммарные инвестиции в сеть из всех источников ≈ **{money(NET_INV+fr_capex_total)} AMD ≈ ${usd(NET_INV+fr_capex_total):,.0f}**.

## Ограничения модели и что проверить до инвестиций

1. **Спрос** — главный недоказанный параметр: частота стрижек, доля walk-in и готовность платить 2 500 за 15 минут без мытья в Армении не измерены (файл 2 §6: 200 уличных интервью, подсчёт очередей в 5 районных точках, Google Trends). Пессимистичный сценарий (32%) показывает, что единичная точка окупается {'за ' + str(pe['payback']) + ' мес.' if pe['payback'] else 'не окупается за 36 мес.'}.
2. **ФОТ по проценту и ТК РА** — минималка 75 000 и 40-часовая неделя: схему «4 мастера на 3 кресла 7/7 по 11 ч» нужно оформить сменным графиком/суммированным учётом; при найме на оклад 200 000+ ФОТ фиксируется и ломает ramp-up.
3. **Налог с оборота** — принимается ли ФОТ в документированные расходы для вычета 6% (разница {money(m36n['turnover_tax']-m36b['turnover_tax'])} AMD/мес); исключение парикмахерских из микро-режима подтверждено двумя источниками, но не первичным текстом ст. 267 НК.
4. **Аренда в USD** — при ослаблении драма на 10% аренда в AMD +10% (эффект −{money(0.1*260000)} AMD/мес, см. чувствительность).
5. **Данные вне файлов** (помечены): возрастная структура населения, население регионов, доходность гособлигаций, цены alteg.io (не публикуются) и вакуумных аппаратов (британские аналоги). Средняя з/п РФ — Росстат за январь 2026 по вторичным источникам.
6. **Сезонность** не моделируется (в файле 1 — летний спад до −30% в мае 2025 у одного франчайзи); резерв в CAPEX частично это покрывает.
""")

MD = "".join(md)
with open(os.path.join(HERE, "financial_model.md"), "w", encoding="utf-8") as f: f.write(MD)
os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
with open(OUT_MD, "w", encoding="utf-8") as f: f.write(MD)
print("MD:", OUT_MD)

# ----------------------------------------------------------------------------------------------
# 9. КРАТКИЙ ВЫВОД В КОНСОЛЬ
# ----------------------------------------------------------------------------------------------
print(f"Средний чек A/B: {check_A:.0f} / {check_B:.0f} AMD; ёмкость 3 кресла {cap3:.1f}/день")
print(f"База м36: загрузка {m36b['load']:.2f}, {m36b['clients_day']:.1f} кл/день, выручка {m36b['revenue']:,.0f}, EBITDA {m36b['ebitda']:,.0f} ({m36b['ebitda']/m36b['revenue']*100:.1f}%), net {m36b['net_profit']:,.0f}, налог {m36b['turnover_tax']:,.0f} ({m36b['turnover_tax']/m36b['revenue']*100:.2f}%)")
for k in ("pessimistic","base","optimistic"):
    r = R[k]; print(f"{k}: CAPEX {r['capex_total']:,.0f} (reserve {r['reserve']:,.0f}), payback {r['payback']}, dpayback {r['dpayback']}, NPV {r['npv']:,.0f}, IRR {r['irr']}, Y3 net {r['y3']['net_profit']:,.0f}")
print(f"BE: {BE}")
print(f"Franchise lite/std: capex {VAR['franchise_lite']['capex_total']:,.0f}/{VAR['franchise_std']['capex_total']:,.0f}, payback {VAR['franchise_lite']['payback']}/{VAR['franchise_std']['payback']}, NPV {VAR['franchise_lite']['npv']:,.0f}/{VAR['franchise_std']['npv']:,.0f}")
print(f"Own franchisee 3%: payback {R_OWN_FR['payback']}, NPV {R_OWN_FR['npv']:,.0f}, margin {R_OWN_FR['m36']['net_profit']/R_OWN_FR['m36']['revenue']*100:.1f}%")
print(f"Network: Yerevan {yerevan_base}/{yerevan_aggr}, regions {regions_base}, inv {NET_INV:,.0f}, cum36 {NET_ROWS[-1]['cum']:,.0f}, min cum {min_cum:,.0f} @m{min_cum_m}")
print("TORNADO:", TORNADO)
