# -*- coding: utf-8 -*-
"""
ArmChio — финансовая модель адаптации формата «Чио Чио» под Армению. ВЕРСИЯ 2 (после независимой верификации, см. verification_report.md).
Запуск:  python -X utf8 -I model.py
Выходы (в папке скрипта): ArmChio_model.xlsx (Assumptions/Prices/CAPEX/PnL_36m — формулы; Scenarios/Sensitivity/Schemes/Franchise/Network — значения),
pnl_36m_<scenario>.csv, model_output.json, financial_model.md (+ копия в scratchpad, OUT_MD). v1 сохранена как model_v1.py / financial_model_v1.md.
"""
import csv, json, math, os
from copy import deepcopy

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_MD = r"C:\Users\Tigran\AppData\Local\Temp\claude\C-------------\8c8c1948-a2aa-44fe-b054-2a9727d74e60\scratchpad\financial_model.md"
F1, F2, F3, FV = "chio_chio_model.md", "armenia_prices_competitors.md", "armenia_costs_regulation_macro.md", "verification_report.md"

# v1 — для таблицы «что изменилось»
V1 = dict(check=2345, load=0.45, clients=47.8, revenue=3706995, ebitda=658187, ebitda_m=0.178, net=533754, net_m=0.144, tax_pct=0.0551,
          capex=9086362, be=27.5, payback=25, dpayback=30, npv=2724649, irr=0.361, pess_payback=None, opt_payback=16,
          fr_lite_pb=33, fr_std_pb=36, own_fr="1,5 млн + 3% (окуп. 32, NPV −1,21 млн)", yerevan="23 / 33", armenia="28–40",
          holding_cum=9883438, holding_inv=30259086, uk_be=9, schedule="3 собств. + 17 франч.", capex_ex=8509000, reserve=577362, area=40)

# ----------------------------------------------------------------------------------------------
# 1. ДОПУЩЕНИЯ
# ----------------------------------------------------------------------------------------------
A = {}; ASSUMPTIONS = []
def add(key, value, unit, source, conf, comment=""):
    A[key] = value; ASSUMPTIONS.append((key, value, unit, source, conf, comment))

add("fx_usd", 361.36, "AMD/USD", f"{F2}, {F3}: ЦБА 19.10.2026 (подтв. {FV} A.9)", "факт")
add("fx_rub", 4.3731, "AMD/RUB", f"{F2}: ЦБА 19.10.2026", "факт")
add("cpi", 0.043, "доля/год", f"{F3}: ИПЦ сен-2026 г/г 4,3%", "факт", "индексация цен и фикс. расходов с 13-го и 25-го мес.")
add("refi_rate", 0.0675, "доля/год", f"{F3}: ставка ЦБ РА, 09.2026", "факт")
add("discount_rate", 0.18, "доля/год", "оценка: ставка ЦБ 6,75% + премия за риск малого бизнеса/неликвидность/валюта ≈ 11 п.п.", "оценка", "диапазон 15–20%; NPV также при 15% и 20%")
add("avg_wage_am", 314196, "AMD/мес", f"{F3}: Armstat авг-2026 (подтв. {FV} A.10)", "факт")
add("avg_wage_ru", 103612, "RUB/мес", "Росстат, январь 2026 (glavkniga.ru/situations/k518592; 1prime.ru 01.04.2026)", "факт (вторичный источник)")
add("pop_yerevan", 1147600, "чел", f"{F3}: Armstat 01.01.2026", "факт")

SERVICES = [  # key, name, price A (спальный), price B (центр), mix, minutes, add-on?
    ("mens",   "Мужская стрижка (ножницы/модельная)", 2500, 3000, 0.40, 20, False),
    ("clip",   "Стрижка машинкой (одна длина)",        2000, 2500, 0.20, 15, False),
    ("kids",   "Детская стрижка (до 12 лет)",          2000, 2500, 0.12, 15, False),
    ("wshort", "Женская стрижка короткая",             2500, 3000, 0.08, 20, False),
    ("wlong",  "Женская стрижка длинная (без укладки)",3500, 4000, 0.05, 30, False),
    ("beard",  "Борода / контур (доп. к визиту)",      1500, 2000, 0.08, 10, True),
    ("style",  "Укладка феном без мытья (доп.)",       1500, 2000, 0.04, 15, True),
    ("roots",  "Окрашивание корней (опция, 1 тон)",    5000, 6000, 0.03, 45, False),
]
assert abs(sum(s[4] for s in SERVICES) - 1.0) < 1e-9
ADDON_SHARE = sum(s[4] for s in SERVICES if s[6])
add("price_mens_A", 2500, "AMD", f"{F2} §4: медиана эконом 3 000, дно 1 000–1 500, True Cost 3 000; −17% к медиане эконом", "решение", "не повышаем для компенсации v2: потолок — медиана эконом 3 000; чувствительность ±20% в G")
add("price_mens_B", 3000, "AMD", f"{F2}: центр 3 000–5 000; = True Cost (15 мин вместо 30)", "решение")
add("mix", "40/20/12/8/5/8/4/3", "% услуг", f"{F1}: ЦА мужчины/дети/пенсионеры; {F2} §3: прокси ЮMoney", "допущение", "муж/машинка/дет/жен.кор/жен.дл/борода/укладка/корни; вариант «эконом-микс» в A; пересмотр по CRM через 3 мес.")
add("minutes_avg", round(sum(s[4]*s[5] for s in SERVICES), 2), "мин/услуга", f"{F1}: 15 мин, 4–5 стрижек/час; взвешено по миксу", "оценка")
add("addon_share", ADDON_SHARE, "доля услуг", "борода и укладка — доп. услуги к визиту (визитов = услуг × (1 − доля))", "допущение")

add("hours_per_day", 11, "ч", f"{F1}: Казань 09–21; {F2}: Ереван 10–20/22 → 10:00–21:00", "решение")
add("days_per_month", 365/12, "дн", "7/7 без санитарных дней", "допущение")
add("chairs", 3, "шт", f"{F1}: 2–3 кресла", "решение", "вариант 2 кресла в C/D")
add("masters", 4, "чел", "3 кресла × 7 дн × 11 ч = 231 кресло-ч/нед; 4 мастера по 5–6 смен — законно только для ИП-мастеров (ТК РА 40/48 ч)", "решение")
add("load_mature_base", 0.40, "доля ёмкости", f"{FV} C.5/D.2: 45% = средняя зрелая точка Чио Чио с брендом и колл-центром → для нового бренда 40%", "оценка")
add("load_mature_pess", 0.30, "доля", f"{FV} D.2; {F1}: «треть франчайзи убыточны»", "оценка")
add("load_mature_opt", 0.45, "доля", f"{F1}: средняя зрелая точка Чио Чио", "оценка")
LOAD_SHAPE = [0.28,0.37,0.47,0.56,0.63,0.70,0.74,0.77,0.79,0.81,0.84,0.86] + [0.88 + (0.95-0.88)*i/11 for i in range(12)] + [0.97 + (1.00-0.97)*i/11 for i in range(12)] + [1.0]*24
add("load_shape", "м1 28% → м6 70% → м12 86% → м24 95% → м36 100%", "кривая", f"{F1}: кейсы Екб, spark, партнёр-2021; {FV} C.5 — согласуется", "оценка")
add("early_stress", 0.5, "множитель м1–3", f"{FV} C.5: стресс «первые 3 месяца на 50% ниже кривой» — для расчёта оборотного резерва", "допущение")

add("area_sqm", 30, "м²", f"решение (v2, {FV} D.15): 3 поста ≈10,5 м² + ожидание 6–7 + зона персонала/хранение 4 + санузел с раковиной 3 + мойка 2,5 + проходы ≈4 → 30 м² удовлетворяет санправилам (файл 3 §5.1); Чио Чио 20–25 м²; v1 — 40 м²", "решение", "вариант 40 м² в C/G")
add("rent_sqm_sleep", 6500, "AMD/м²/мес", f"{F3} §1.2; {FV} A.24: выборка спальных районов n=3 (5 556 / 5 808 / 9 583), медиана 5 808, единственный малый формат 9 583", "оценка (n=3)", "диапазон 5 800–9 600; стресс 9 500 в C и G; закрыть обзвоном 30 объектов")
add("rent_sqm_center", 8500, "AMD/м²/мес", f"{F3} §1.2: Кентрон (пересчёт {FV}: 8 451–8 667 для 20–120 м²)", "оценка (цены предложения)")
add("rent_sqm_center_1line", 13000, "AMD/м²/мес", f"{F3} §1.2: первая линия/витрина", "оценка (цены предложения)")
add("deposit_months", 2, "мес", f"{F3}", "факт")
add("fitout_months", 1.25, "мес", f"{FV} C.6: аренда за период ремонта 1–1,5 мес. (каникулы не норма)", "оценка")
add("rent_documented", False, "да/нет", f"{FV} C.2: арендодатель-физлицо не выписывает расчётный документ ст. 55 ч. 2 НК → вычета 6% нет", "допущение (база = нет)", "если арендодатель — юрлицо/ИП с документами → True (вариант в C)")
add("fot_share_market", 0.50, "доля выручки", f"{F3} §2.1: вакансии «50/50»; {FV} C.3: в практике это доля на руки при неформальной оплате", "оценка (3–4 вакансии)", "в модели 50% — брутто-выплата (gross сотруднику или сумма ИП); схемы a/b/c в C")
add("fot_guarantee", 200000, "AMD/мес брутто", f"{FV} D.8: 150 000 gross → 111 500 на руки < рынка 150–220 тыс.; 200 000 gross/ИП → ≈170 тыс. на руки", "оценка")
add("masters_ip", True, "да/нет", f"{FV} C.3: базовая схема — мастер-ИП на налоге с оборота, договор оказания услуг; выплата документирована → вычет 6%", "решение", "риск переквалификации в трудовые отношения — см. C")
add("ip_master_doc_exp", 0.05, "доля дохода", "документированные расходы самого ИП-мастера (инструмент, транспорт) — мало", "допущение")
add("manager_cost", 120000, "AMD/мес", "0,5 ставки управляющего-наставника (сотрудник; не документируемый расход)", "допущение")
add("cleaning_cost", 40000, "AMD/мес", "уборка (сотрудник/физлицо)", "допущение")
add("host_cost", 150000, "AMD/мес", f"{FV} D.7: администратор/хост на первые 6 мес. — опция, база 0 (киоск + мастера)", "допущение")
add("consumables_per_client", 95, "AMD/клиент", f"{F3}: 20–40 тыс. на 3 кресла; {F1}: ~6% в РФ с косметикой; без мытья ниже", "оценка", "документированы (чеки поставщиков)")
add("consumables_floor", 30000, "AMD/мес", F3, "оценка")
add("utilities", 75000, "AMD/мес", f"{F3} §7: 60–90 тыс. на 40 м²", "оценка", "документированы (ENA, Veolia)")
add("it_crm", 17000, "AMD/мес", "alteg.io цены не публикует (проверено 20.10.2026); $10–12/сотр. × 4 + связь; альтернатива Fresha", "оценка")
add("hdm_monthly", 3000, "AMD/мес", f"{F3} §5.2: e-HDM", "факт")
add("hdm_setup", 30000, "AMD", f"{F3} §5.2", "факт")
add("card_share", 0.50, "доля выручки", f"{F2} §3; {FV} D.5: спальные 40%, центр 60%", "допущение")
add("acq_rate", 0.012, "доля безнала", f"{F3} §5.3", "оценка")
add("acq_terminal", 4600, "AMD/мес", f"{F3} §5.3: IDBank", "факт")
add("marketing_y1", 0.05, "доля выручки", f"{FV} D.6: новый бренд без колл-центра/приложения УК → 5% в год 1", "оценка")
add("marketing_y1_floor", 120000, "AMD/мес", FV, "оценка")
add("marketing_pct", 0.03, "доля выручки", f"{F1}: модель УК 10 тыс.₽ при 150–300 тыс. (3–7%) → 3% с года 2", "оценка")
add("marketing_floor", 60000, "AMD/мес", "", "оценка")
add("accounting", 35000, "AMD/мес", "аутсорс бухгалтерии ИП на налоге с оборота", "оценка")
add("accounting_ip_extra", 15000, "AMD/мес", "ведение отчётности 4 ИП-мастеров (схема c)", "допущение")
add("other_pct", 0.02, "доля выручки", "ремонт инструмента, форс-мажор", "допущение")
add("other_floor", 20000, "AMD/мес", "", "допущение")
add("other_documented_share", 0.5, "доля", "", "допущение")
add("ip_fixed", 20000, "AMD/мес", f"{F3} §4.2: подоходный 5 000 + соц. 5 000 + гербовый 120 000/год (доход >12 млн) = 20 000; {FV} A.18 подтв.", "факт", f"окончательность подоходного ИП 5 000 при налоге с оборота — подтвердить у бухгалтера ({FV} B.1); иначе +20% с прибыли")
add("ip_oms_year", 129600, "AMD/год", f"{F3}; {FV} C.4/14: ОМС ИП обязателен при доходе прошлого года >2,4 млн → с 13-го месяца", "факт")
add("tt_rate", 0.10, "доля", f"{F3} §4.2; {FV} A.13 подтв. (fincore, alphaaccounting)", "факт")
add("tt_deduct", 0.06, "доля документированных расходов", f"{F3}; закон ՀՕ-72-Ն: расходы, подтверждённые расчётными документами ст. 55 ч. 2 п. 1–5 НК (счёт-фактура, накладная, чек ККМ)", "факт")
add("tt_floor", 0.045, "доля", F3, "факт")
add("tt_threshold", 115e6, "AMD/год", f"{F3}; {FV} A.14: PwC/fincore/profin 115 млн vs armenian-lawyer 120 млн — берём 115", "факт (расхождение источников)")
add("uk_tax_royalty", 0.10, "доля", f"{FV} C.9: роялти/паушальные — доход без вычета документированных расходов", "факт (вторичные источники)")
add("tt_trade_rate", 0.10, "доля", f"{F3} §4.2: торговля 10% − 9,5% закупок, мин. 1% (поставки расходников УК)", "факт")
add("royalty_chio", 0.04, "доля выручки", f"{F1}: 4%, мин. 5 000 ₽, со 2-го мес. ({FV} A.3 подтв.)", "факт")
add("royalty_chio_min_rub", 5000, "RUB/мес", F1, "факт")
add("fee_chio_lite_rub", 290000, "RUB", f"{F1} ({FV} A.2 подтв.)", "факт")
add("fee_chio_std_rub", 590000, "RUB", f"{F1} ({FV} A.1 подтв.)", "факт")
add("depr_months", 60, "мес", "5 лет", "допущение")

# CAPEX
add("fitout_sqm", 100000, "AMD/м²", f"{F3} §6: 80–120 тыс./м²", "оценка")
add("chair_price", 90000, "AMD", f"{F3} §6: РФ 56–110 тыс. (без НДС и доставки)", "оценка")
add("post_price", 120000, "AMD", f"{F3} §6", "оценка")
add("wash_unit", 220000, "AMD", f"{F3} §6: кресло-мойка 130–304 тыс.", "оценка")
add("vacuum_unit", 125000, "AMD", "Sibel Eye Vac £215–307 (sallybeauty.co.uk), Head Jog Hair Vacuum £148,5 (astonandfincher.co.uk) + доставка; 2 шт.", "оценка")
add("tools_per_post", 150000, "AMD", f"{F3} §6", "оценка")
add("import_vat", 0.20, "доля", f"{F3} §6; {FV} C.6: НДС 20% при ввозе из РФ (ЕАЭС, пошлина 0) на кресла/мойку/вакуум/инструмент", "факт")
add("sterilizer", 80000, "AMD", f"{F3} §5.1", "оценка")
add("hvac", 400000, "AMD", f"{FV} C.6: 2 сплит-системы + вытяжка (санправила: ≤28 °C, вентиляция)", "оценка")
add("boiler", 65000, "AMD", f"{FV} C.6: горячая вода обязательна", "оценка")
add("waiting_furniture", 350000, "AMD", f"{F3} §6", "оценка")
add("signage", 300000, "AMD", f"{F3} §6: 150–400 тыс.", "оценка")
add("signage_permit", 100000, "AMD", f"{FV} A.27: тариф согласования вывески мэрии Еревана НЕ НАЙДЕН — плейсхолдер 50–150 тыс.", "оценка (не найдено)")
add("branding_interior", 200000, "AMD", "брендинг интерьера, прайс-борды, форма", "оценка")
add("selfservice_it", 250000, "AMD", "киоск самообслуживания + ПК + камеры", "оценка")
add("crm_setup", 30000, "AMD", "", "оценка")
add("training", 250000, "AMD", f"{FV} C.6: стажировка 4 мастеров 2 недели до открытия (гарантия) — 15-минутная технология", "оценка")
add("registration", 3000, "AMD", f"{F3} §4.4", "факт")
add("legal_setup", 60000, "AMD", "", "оценка")
add("medbooks", 60000, "AMD", f"{F3} §5.1: 4 × 15 тыс.", "оценка")
add("launch_marketing", 400000, "AMD", f"{F1}: бесплатные стрижки на открытии, флаеры", "оценка")
add("contingency", 0.15, "доля hard-CAPEX", f"{FV} C.6: факт CAPEX у франчайзи в 1,5–2 раза выше витрины → 15%", "допущение")

FX = A["fx_usd"]; RUB = A["fx_rub"]; DAYS = A["days_per_month"]

# ----------------------------------------------------------------------------------------------
# 2. НАЛОГИ С ВЫПЛАТ МАСТЕРАМ (три схемы)
# ----------------------------------------------------------------------------------------------
def stamp_duty(g):  # военный гербовый сбор, шкала (verification A.16)
    for lim, v in ((100000,1500),(200000,3000),(500000,5500),(1000000,8500)):
        if g <= lim: return v
    return 15000
def oms_employee(g):
    if g <= 200000: return 0
    if g <= 500000: return 300
    if g <= 1000000: return 3300
    return 10800
def pension(g):
    return 0.05*g if g <= 500000 else min(0.10*g - 25000, 87500)
def net_employee(g):
    """На руки сотруднику при gross g (трудовой договор/ГПХ): 20% подоходный + пенсионный + гербовый + ОМС; работодатель 0% сверх gross."""
    return g - 0.20*g - pension(g) - stamp_duty(g) - oms_employee(g)
assert abs(net_employee(250000) - 181700) < 1  # контрольный пример файла 3 §2.4
def gross_up(net_target):
    lo, hi = net_target, net_target*2
    for _ in range(100):
        mid = (lo+hi)/2
        if net_employee(mid) < net_target: lo = mid
        else: hi = mid
    return hi
def net_ip_master(g, month):
    """На руки мастеру-ИП (налог с оборота 10% − 6% своих документированных расходов, пол 4,5%; фиксы 5 000+5 000; гербовый ИП 12 000/год; ОМС ИП 129 600/год со 2-го года)."""
    tax = max(A["tt_rate"]*g - A["tt_deduct"]*A["ip_master_doc_exp"]*g, A["tt_floor"]*g)
    fixed = 10000 + 1000 + (A["ip_oms_year"]/12 if month >= 13 else 0)
    return g - tax - fixed

# ----------------------------------------------------------------------------------------------
# 3. ЮНИТ-МОДЕЛЬ
# ----------------------------------------------------------------------------------------------
def avg_check(price_col): return sum(s[price_col]*s[4] for s in SERVICES)
def capacity_per_day(chairs, hours=None): return chairs*(hours or A["hours_per_day"])*60.0/A["minutes_avg"]

def default_params(**kw):
    p = dict(name="base", chairs=3, masters=4, price_mult=1.0, price_col=2, load_mature=A["load_mature_base"], load_mult=1.0,
             rent_sqm=A["rent_sqm_sleep"], area=A["area_sqm"], fitout_area=A["area_sqm"], rent_documented=A["rent_documented"],
             fot_share=A["fot_share_market"], fot_mode="share", masters_ip=A["masters_ip"], fot_guarantee=A["fot_guarantee"],
             consumables_per_client=A["consumables_per_client"], manager_cost=A["manager_cost"], host_months=0,
             royalty=0.0, royalty_min=0.0, royalty_from_month=2, franchise_fee=0.0, marketing_mult=1.0, marketing_flat=None,
             indexation=True, early_stress=False, fx_rent_mult=1.0)
    p.update(kw); return p

def capex(p):
    ch = p["chairs"]; rent = p["rent_sqm"]*p["area"]
    imported = ch*A["chair_price"] + A["wash_unit"] + 2*A["vacuum_unit"] + ch*A["tools_per_post"]
    hard = {
        f"Ремонт/отделка (fit-out) {p['fitout_area']} м² × {A['fitout_sqm']:,}".replace(",", " "): p["fitout_area"]*A["fitout_sqm"],
        f"Кресла парикмахерские × {ch}": ch*A["chair_price"],
        f"Рабочие посты (зеркало, консоль) × {ch}": ch*A["post_price"],
        "Мойка (1 шт.)": A["wash_unit"],
        "«Воздушная мойка» — вакуумные аппараты × 2": 2*A["vacuum_unit"],
        f"Инструмент и техника × {ch} постов": ch*A["tools_per_post"],
        "НДС 20% при ввозе оборудования из РФ (кресла, мойка, вакуум, инструмент) [v2]": round(imported*A["import_vat"]),
        "Кондиционеры (2) + вытяжная вентиляция [v2]": A["hvac"],
        "Водонагреватель [v2]": A["boiler"],
        "Стерилизатор УФ/сухожар": A["sterilizer"],
        "Мебель зоны ожидания, стойка, хранение": A["waiting_furniture"],
        "Вывеска световая": A["signage"],
        "Согласование вывески, мэрия Еревана (тариф не найден — оценка) [v2]": A["signage_permit"],
        "Брендинг интерьера, прайс-борды, форма": A["branding_interior"],
        "Киоск самообслуживания/ПК/камеры": A["selfservice_it"],
        "Касса e-HDM (регистрация)": A["hdm_setup"],
    }
    hard_total = sum(hard.values())
    soft = {
        f"Непредвиденное {int(A['contingency']*100)}% от hard-CAPEX [v2: было 10%]": round(hard_total*A["contingency"]),
        "Настройка CRM/онлайн-записи/сайт": A["crm_setup"],
        "Стажировка мастеров до открытия (2 нед. × 4) [v2]": A["training"],
        "Регистрация ИП": A["registration"],
        "Юридическое/налоговое оформление": A["legal_setup"],
        "Медкнижки персонала": A["medbooks"],
        f"Аренда за период ремонта {A['fitout_months']} мес. [v2]": round(A["fitout_months"]*rent),
        f"Депозит аренды × {A['deposit_months']} мес.": A["deposit_months"]*rent,
        "Маркетинг открытия": A["launch_marketing"],
        "Паушальный взнос": p["franchise_fee"],
    }
    items = dict(hard); items.update(soft)
    return items, hard_total

def month_pnl(p, m, load):
    idx = (1+A["cpi"])**((m-1)//12) if p["indexation"] else 1.0
    check = avg_check(p["price_col"])*p["price_mult"]*idx
    clients_day = capacity_per_day(p["chairs"])*load
    rev = clients_day*check*DAYS
    # --- мастера
    share_amt = p["fot_share"]*rev
    floor_amt = p["masters"]*p["fot_guarantee"]*idx
    if p["fot_mode"] == "grossup":   # схема (b): 50% выручки НА РУКИ сотрудникам → gross-up
        per_master_net = max(share_amt, floor_amt)/p["masters"]
        fot = gross_up(per_master_net)*p["masters"]
    else:
        fot = max(share_amt, floor_amt)
    fot_documented = (p["fot_mode"] == "share" and p["masters_ip"])
    rent = p["rent_sqm"]*p["area"]*idx*p["fx_rent_mult"]
    mgr = (p["manager_cost"]+A["cleaning_cost"])*idx + (A["host_cost"]*idx if m <= p["host_months"] else 0)
    cons = max(p["consumables_per_client"]*idx*clients_day*DAYS, A["consumables_floor"]*idx)
    util = A["utilities"]*idx; it = A["it_crm"]*idx; hdm = A["hdm_monthly"]
    acq = rev*A["card_share"]*A["acq_rate"] + A["acq_terminal"]
    if p["marketing_flat"] is not None:
        mkt = max(p["marketing_flat"]*rev, A["marketing_y1_floor"]*idx)
    elif m <= 12:
        mkt = max(A["marketing_y1"]*rev, A["marketing_y1_floor"]*idx)
    else:
        mkt = max(A["marketing_pct"]*rev, A["marketing_floor"]*idx)
    mkt *= p["marketing_mult"]
    acc = (A["accounting"] + (A["accounting_ip_extra"] if fot_documented else 0))*idx
    oth = max(A["other_pct"]*rev, A["other_floor"]*idx)
    roy = max(p["royalty"]*rev, p["royalty_min"]) if (p["royalty"] > 0 and m >= p["royalty_from_month"]) else 0.0
    documented = cons + util + it + hdm + acq + mkt + acc + roy + oth*A["other_documented_share"]
    if fot_documented: documented += fot
    if p["rent_documented"]: documented += rent
    tt = max(A["tt_rate"]*rev - A["tt_deduct"]*documented, A["tt_floor"]*rev)
    ip = A["ip_fixed"] + (A["ip_oms_year"]/12 if m >= 13 else 0)
    opex = fot+rent+mgr+cons+util+it+hdm+acq+mkt+acc+oth+roy+tt+ip
    return dict(month=m, load=load, clients_day=clients_day, visits_day=clients_day*(1-ADDON_SHARE), check=check, revenue=rev, fot=fot, rent=rent,
                manager=mgr, consumables=cons, utilities=util, it=it, hdm=hdm, acquiring=acq, marketing=mkt, accounting=acc, other=oth,
                royalty=roy, documented=documented, turnover_tax=tt, ip_fixed=ip, opex=opex, ebitda=rev-opex, fot_documented=fot_documented)

def npv(rate_annual, cfs):
    r = (1+rate_annual)**(1/12)-1
    return sum(cf/(1+r)**t for t, cf in enumerate(cfs))
def irr_annual(cfs):
    f = lambda r: sum(cf/(1+r)**t for t, cf in enumerate(cfs))
    lo, hi = -0.99, 10.0
    if f(lo)*f(hi) > 0: return None
    for _ in range(200):
        mid = (lo+hi)/2
        if f(lo)*f(mid) <= 0: hi = mid
        else: lo = mid
    return (1+(lo+hi)/2)**12-1

def loads(p, months):
    out = []
    for m in range(1, months+1):
        l = min(0.95, LOAD_SHAPE[m-1]*p["load_mature"]*p["load_mult"])
        if p["early_stress"] and m <= 3: l *= A["early_stress"]
        out.append(l)
    return out

def run(p, months=36):
    rows = [month_pnl(p, m, l) for m, l in enumerate(loads(p, months), start=1)]
    # оборотный резерв — по стресс-кривой (м1–3 × 0,5), максимальная просадка накопленного EBITDA
    ps = dict(p, early_stress=True)
    cum = mn = 0.0
    for m, l in enumerate(loads(ps, 36), start=1):
        cum += month_pnl(ps, m, l)["ebitda"]; mn = min(mn, cum)
    reserve = -mn
    cum = mn2 = 0.0
    for r in rows[:36]:
        cum += r["ebitda"]; mn2 = min(mn2, cum)
    reserve_base = -mn2
    items, hard_total = capex(p)
    capex_total = sum(items.values())          # инвестиции (без оборотного резерва)
    financing = capex_total + reserve           # потребность в финансировании = CAPEX + резерв на стресс ramp-up
    depr = (hard_total + items[[k for k in items if k.startswith("Непредвиденное")][0]] + A["crm_setup"] + p["franchise_fee"])/A["depr_months"]
    cfs = [-capex_total]; cum = -capex_total; cum_d = -capex_total; payback = dpayback = None
    r_m = (1+A["discount_rate"])**(1/12)-1
    for r in rows:
        r["depreciation"] = depr; r["net_profit"] = r["ebitda"]-depr; r["cf"] = r["ebitda"]
        cum += r["cf"]; r["cum_cf"] = cum; cum_d += r["cf"]/(1+r_m)**r["month"]; r["cum_dcf"] = cum_d
        if payback is None and cum >= 0: payback = r["month"]
        if dpayback is None and cum_d >= 0: dpayback = r["month"]
        cfs.append(r["cf"])
    cfs36 = cfs[:37]
    tv = 3*sum(r["ebitda"] for r in rows[24:36])
    res = dict(params=p, rows=rows, capex_items=items, capex_total=capex_total, financing=financing, hard_capex=hard_total, reserve=reserve, reserve_base=reserve_base,
               depr=depr, payback=payback, dpayback=dpayback, npv=npv(A["discount_rate"], cfs36), npv15=npv(0.15, cfs36), npv20=npv(0.20, cfs36),
               npv_tv=npv(A["discount_rate"], cfs36[:-1]+[cfs36[-1]+tv]), irr=irr_annual(cfs36), cfs=cfs, months=months)
    if months >= 60:
        res["npv60"] = npv(A["discount_rate"], cfs[:61]); res["irr60"] = irr_annual(cfs[:61])
    for y in range(1, months//12+1):
        yr = rows[(y-1)*12:y*12]
        res[f"y{y}"] = {k: sum(r[k] for r in yr) for k in ("revenue","fot","rent","opex","ebitda","net_profit","turnover_tax","royalty","depreciation","marketing")}
        res[f"y{y}"]["clients_day"] = sum(r["clients_day"] for r in yr)/12
    res["m36"] = rows[35]
    return res

def breakeven(p, month=36):
    base = run(p); depr = base["depr"]
    def solve(fn):
        lo, hi = 0.001, 0.95
        if fn(hi) < 0: return None
        for _ in range(100):
            mid = (lo+hi)/2
            if fn(mid) < 0: lo = mid
            else: hi = mid
        return hi
    l1 = solve(lambda l: month_pnl(p, month, l)["ebitda"]-depr); l2 = solve(lambda l: month_pnl(p, month, l)["ebitda"])
    if l1 is None: return None
    r1 = month_pnl(p, month, l1); r2 = month_pnl(p, month, l2)
    return dict(load=l1, clients_day=r1["clients_day"], visits_day=r1["visits_day"], revenue=r1["revenue"], load_ebitda=l2, clients_day_ebitda=r2["clients_day"], revenue_ebitda=r2["revenue"])

# ----------------------------------------------------------------------------------------------
# 4. РАСЧЁТЫ
# ----------------------------------------------------------------------------------------------
fmt = lambda x, d=0: f"{x:,.{d}f}".replace(",", " ")
usd = lambda x: x/FX; rub = lambda x: x/RUB
pct = lambda x, d=1: f"{x*100:.{d}f}%"
money = lambda x: fmt(x)
check_A = avg_check(2); check_B = avg_check(3); cap3 = capacity_per_day(3); cap2 = capacity_per_day(2)
check_visit_A = check_A/(1-ADDON_SHARE)
# эконом-микс (машинка 35%, борода 5%)
ECON_MIX = {"mens":0.30,"clip":0.35,"kids":0.12,"wshort":0.08,"wlong":0.03,"beard":0.05,"style":0.04,"roots":0.03}
check_econ = sum(s[2]*ECON_MIX[s[0]] for s in SERVICES)

P_BASE = default_params(name="base")
P_PESS = default_params(name="pessimistic", price_mult=0.90, load_mature=A["load_mature_pess"], rent_sqm=8000, masters_ip=False, consumables_per_client=120)
P_OPT  = default_params(name="optimistic", load_mature=A["load_mature_opt"], rent_sqm=5800)
R = {p["name"]: run(p) for p in (P_BASE, P_PESS, P_OPT)}
b, pe, op = R["base"], R["pessimistic"], R["optimistic"]; m36b = b["m36"]

VAR_P = dict(
    center=default_params(name="center", price_col=3, rent_sqm=A["rent_sqm_center"], load_mature=0.45),
    center_1line=default_params(name="center_1line", price_col=3, rent_sqm=A["rent_sqm_center_1line"], load_mature=0.50),
    chairs2=default_params(name="chairs2", chairs=2, masters=3, area=25, fitout_area=25),
    area40=default_params(name="area40", area=40, fitout_area=40),
    price28=default_params(name="price_2800", price_mult=2568/2345),
    emp=default_params(name="scheme_a_employees", masters_ip=False),
    grossup=default_params(name="scheme_b_grossup", fot_mode="grossup", masters_ip=False),
    rent_doc=default_params(name="rent_documented", rent_documented=True),
    rent9500=default_params(name="rent_9500", rent_sqm=9500),
    rent5800=default_params(name="rent_5800", rent_sqm=5800),
    owner=default_params(name="owner_operator", manager_cost=0),
    host=default_params(name="host_6m", host_months=6),
    fot55=default_params(name="fot_55", fot_share=0.55),
    fx10=default_params(name="fx_minus10", fx_rent_mult=1.10),
    mkt5=default_params(name="marketing_5pct_all", marketing_flat=0.05),
    fr_lite=default_params(name="franchise_lite", royalty=A["royalty_chio"], royalty_min=A["royalty_chio_min_rub"]*RUB, franchise_fee=A["fee_chio_lite_rub"]*RUB),
    fr_std=default_params(name="franchise_std", royalty=A["royalty_chio"], royalty_min=A["royalty_chio_min_rub"]*RUB, franchise_fee=A["fee_chio_std_rub"]*RUB),
)
VAR = {k: run(p) for k, p in VAR_P.items()}
BE = breakeven(P_BASE); BE_PESS = breakeven(P_PESS); BE_OPT = breakeven(P_OPT); BE_EMP = breakeven(VAR_P["emp"]); BE_C = breakeven(VAR_P["center"]); BE_2 = breakeven(VAR_P["chairs2"]); BE_40 = breakeven(VAR_P["area40"])

# схемы оплаты мастеров (мес. 36 базы): выплата на мастера
g36 = m36b["fot"]/P_BASE["masters"]
SCHEMES = [
    ("(a) Трудовой договор, 50% выручки gross", g36, net_employee(g36), m36b["fot"], False),
    ("(b) Трудовой договор, 50% выручки НА РУКИ (gross-up)", gross_up(g36), g36, gross_up(g36)*4, False),
    ("(c) Мастер-ИП (налог с оборота), 50% выручки — базовая", g36, net_ip_master(g36, 36), m36b["fot"], True),
]
g_guar = A["fot_guarantee"]

# чувствительность
def sens_table():
    out = []
    for factor, label in (("price_mult","Цена"),("load_mult","Загрузка"),("rent_sqm","Аренда"),("fot_share","Доля ФОТ")):
        row = {"factor": label}
        for d in (-0.2,-0.1,0.0,0.1,0.2):
            p = deepcopy(P_BASE)
            if factor in ("price_mult","load_mult"): p[factor] = 1+d
            elif factor == "rent_sqm": p[factor] = A["rent_sqm_sleep"]*(1+d)
            else: p[factor] = A["fot_share_market"]*(1+d)
            r = run(p); row[d] = dict(net36=sum(x["net_profit"] for x in r["rows"]), payback=r["payback"], npv=r["npv"])
        out.append(row)
    return out
SENS = sens_table()
TORNADO = sorted([(s["factor"], s[-0.2]["net36"], s[0.2]["net36"]) for s in SENS], key=lambda t: -abs(t[2]-t[1]))
SWITCHES = [("База (ИП-мастера, аренда у физлица, 40 м², маркетинг 5%→3%)", b),
            ("Мастера — сотрудники (ФОТ не вычитается), схема (a)", VAR["emp"]),
            ("Мастера — сотрудники с gross-up до 50% на руки, схема (b)", VAR["grossup"]),
            ("Арендодатель выдаёт документы (аренда вычитается)", VAR["rent_doc"]),
            ("Аренда 9 500/м² (единственное малое объявление в спальном районе)", VAR["rent9500"]),
            ("Аренда 5 800/м² (медиана n=3)", VAR["rent5800"]),
            ("Драм −10% → аренда +10%", VAR["fx10"]),
            ("Маркетинг 5% все 3 года", VAR["mkt5"]),
            ("Доля ФОТ 55%", VAR["fot55"]),
            ("Администратор-хост 150 000 в мес. 1–6", VAR["host"]),
            ("Владелец-оператор (без управляющего)", VAR["owner"]),
            ("Площадь 40 м² как в v1 (аренда 260 000, fit-out 4,0 млн)", VAR["area40"]),
            ("Прайс A+: муж. 2 800, машинка/детская 2 200, жен. 2 800/3 800 (чек +9,5%, −7% к медиане эконом)", VAR["price28"])]

# рынок сверху вниз
add("men_15_64_share", 0.31, "доля населения", "допущение (возрастная структура вне файлов)", "допущение")
add("kids_share", 0.20, "доля населения", "допущение", "допущение")
add("haircut_freq_men", 10, "раз/год", f"{F2} §3: прокси ЮMoney; 3–5 недель → 8–12", "допущение")
add("haircut_freq_kids", 6, "раз/год", "допущение", "допущение")
add("catchment_radius_km", 1.2, "км", f"{F1}: «у дома»", "допущение")
DISTRICTS = [("Малатия-Себастия",150100,5966),("Шенгавит",140400,3458),("Нор-Норк",137100,9717),("Кентрон",132800,9948),("Эребуни",129500,2670),("Ачапняк",122500,4744),
             ("Арабкир",119000,8981),("Канакер-Зейтун",75400,9754),("Аван",58700,7229),("Давташен",51900,7960),("Норк-Мараш",13900,2920),("Нубарашен",11100,644)]
area_catch = math.pi*A["catchment_radius_km"]**2
TOPDOWN = [(n, p_, d, min(d*area_catch, p_), min(d*area_catch, p_)*A["men_15_64_share"]*A["haircut_freq_men"]/365, min(d*area_catch, p_)*A["kids_share"]*A["haircut_freq_kids"]/365) for n, p_, d in DISTRICTS]
need_men = m36b["clients_day"]*(0.40+0.20+0.08); need_kids = m36b["clients_day"]*0.12
salons_per_capita = 1529/A["pop_yerevan"]

# ёмкость сети (v2)
add("pop_per_point_base", 60000, "жит./точка", f"{FV} C.9: Казань откатилась с 41–42 (2022) до 26 адресов (2026) = 1 на ~50 тыс.; Ереван плотнее по конкуренции (1 салон/720 жит.) → 1/60 тыс., диапазон 50–70 тыс.", "оценка")
add("pop_per_point_lo", 70000, "жит./точка", "нижняя граница", "оценка")
add("pop_per_point_hi", 50000, "жит./точка", "верхняя граница (= Казань-2026)", "оценка")
add("regions_pop", "Гюмри ~112 тыс., Ванадзор ~78 тыс., Абовян ~44 тыс., Вагаршапат ~47 тыс.", "чел", "ВНЕ ФАЙЛОВ: перепись 2022 (ориентир); в файлах — число парикмахерских 2ГИС 119/3/61/95", "допущение")
REGIONS = [("Гюмри",112000,119,2,1),("Ванадзор",78000,3,1,1),("Абовян",44000,61,1,1),("Вагаршапат",47000,95,1,0)]
def cap_points(per):
    out = []
    for name, pop, dens in DISTRICTS:
        n = int(round(pop/per))
        if name == "Кентрон": n += 1
        if dens < 3000: n = max(0, n-1)
        out.append(n)
    return out
NETCAP = list(zip([d[0] for d in DISTRICTS], [d[1] for d in DISTRICTS], [d[2] for d in DISTRICTS], cap_points(A["pop_per_point_lo"]), cap_points(A["pop_per_point_base"]), cap_points(A["pop_per_point_hi"])))
yerevan_lo = sum(x[3] for x in NETCAP); yerevan_base = sum(x[4] for x in NETCAP); yerevan_hi = sum(x[5] for x in NETCAP)
regions_base = sum(r[3] for r in REGIONS); regions_lo = sum(r[4] for r in REGIONS)

# своя франшиза: сетка пакетов (субфранчайзи = базовая точка + роялти + паушальный), 36 и 60 мес.
add("own_royalty_min", 50000, "AMD/мес", "≈2,3× минимума Чио Чио (21 866 AMD); со 2-го месяца", "предложение")
add("own_mkt_fund", 0.01, "доля выручки", "фонд внутри бюджета маркетинга франчайзи (не сверху)", "предложение")
add("supply_margin", 0.20, "доля закупок", "маржа УК на расходниках", "оценка")
add("uk_payroll", 800000, "AMD/мес gross", "2 чел.: директор сети/запуск 450 тыс., тренер-технолог 350 тыс.", "оценка")
add("uk_other", 500000, "AMD/мес", "офис 100, IT 100, маркетинг бренда/SMM 150, юр./бух. 100, прочее 50 тыс.", "оценка")
add("uk_setup", 3000000, "AMD", "бренд-бук, стандарты, бизнес-бук, обучение, IT (разово)", "оценка")
GRID = []
for fee in (800000, 1000000, 1200000, 1500000):
    for roy in (0.025, 0.03, 0.04):
        p = default_params(name=f"fr_{fee}_{roy}", royalty=roy, royalty_min=A["own_royalty_min"], franchise_fee=fee)
        r = run(p, 60)
        GRID.append(dict(fee=fee, roy=roy, payback=r["payback"], margin=r["m36"]["net_profit"]/r["m36"]["revenue"], npv36=r["npv"], npv60=r["npv60"], irr60=r["irr60"], capex=r["capex_total"]))
# выбор пакета: минимальный паушальный + роялти, при которых NPV60 > 0; окупаемость ≤ 30 — проверяем отдельно
OWN_FEE, OWN_ROY = 1000000, 0.03
PRICE_PLUS = 2568/2345
CONDITIONS = []
for load in (0.40, 0.45, 0.50):
    for pm, pl in ((1.0, "A (2 500)"), (PRICE_PLUS, "A+ (2 800)")):
        p = default_params(name="cond", royalty=OWN_ROY, royalty_min=A["own_royalty_min"], franchise_fee=OWN_FEE, load_mature=load, price_mult=pm)
        r = run(p, 60)
        CONDITIONS.append(dict(load=load, price=pl, payback=r["payback"], npv36=r["npv"], npv60=r["npv60"], irr36=r["irr"], margin=r["m36"]["net_profit"]/r["m36"]["revenue"]))
COND_OK = [c for c in CONDITIONS if c["npv36"] > 0 and c["payback"] and c["payback"] <= 30]
add("own_fee", OWN_FEE, "AMD", f"предложение (пакет «founders» для первых 5–7 договоров): NPV субфранчайзи за 60 мес. > 0 @18%; {FV} D.11", "предложение")
add("own_royalty", OWN_ROY, "доля выручки", "предложение: 3%, мин. 50 000 (Чио Чио 4%, Прядки 3%)", "предложение")
P_OWN_FR = default_params(name="own_franchisee", royalty=OWN_ROY, royalty_min=A["own_royalty_min"], franchise_fee=OWN_FEE)
R_OWN_FR = run(P_OWN_FR, 60)
P_OWN_FR45 = default_params(name="own_franchisee_load45", royalty=OWN_ROY, royalty_min=A["own_royalty_min"], franchise_fee=OWN_FEE, load_mature=0.45)
R_OWN_FR45 = run(P_OWN_FR45, 60)

def uk_taxes(royalty_income, fees, supply_cost):
    """УК: роялти/паушальные — 10% без вычета; поставки расходников — торговля 10% − 9,5% закупок, мин. 1% оборота."""
    t_roy = A["uk_tax_royalty"]*(royalty_income+fees)
    sales = supply_cost*(1+A["supply_margin"])
    t_trade = max(A["tt_trade_rate"]*sales - 0.095*supply_cost, 0.01*sales)
    return t_roy + t_trade

def franchisor_steady(n_points, own_points=3, openings_per_year=0):
    rev_pt = m36b["revenue"]; fr = max(0, n_points-own_points)
    royalty = fr*max(OWN_ROY*rev_pt, A["own_royalty_min"]); fund = fr*A["own_mkt_fund"]*rev_pt
    supply_cost = fr*m36b["consumables"]; supply = supply_cost*A["supply_margin"]
    fees = openings_per_year*OWN_FEE/12; own_profit = own_points*m36b["ebitda"]
    uk_cost = A["uk_payroll"]+A["uk_other"]+fund; tax = uk_taxes(royalty, fees, supply_cost)
    return dict(n=n_points, fr=fr, royalty=royalty, mkt_fund=fund, supply=supply, fees=fees, own_profit=own_profit, uk_cost=uk_cost, uk_tax=tax,
                uk_net_ex_own_nofee=royalty+fund+supply-uk_cost-uk_taxes(royalty, 0, supply_cost), uk_net_ex_own=royalty+fund+supply+fees-uk_cost-tax,
                total=royalty+fund+supply+fees+own_profit-uk_cost-tax)
FRANCHISOR = [franchisor_steady(n, openings_per_year=o) for n, o in ((8,2),(12,3),(20,4))]
FR_PT_INCOME = max(OWN_ROY*m36b["revenue"], A["own_royalty_min"])*(1-A["uk_tax_royalty"]) + m36b["consumables"]*A["supply_margin"] - (max(0.10*m36b["consumables"]*1.2-0.095*m36b["consumables"], 0.012*m36b["consumables"]))
N_BE_UK = (A["uk_payroll"]+A["uk_other"])/FR_PT_INCOME

def master_franchise_benchmark():
    rev_pt = m36b["revenue"]; out = []
    for n in (8,12,20):
        gross_roy = n*A["royalty_chio"]*rev_pt; master_share = 0.5*gross_roy
        fee_share = 0.5*A["fee_chio_lite_rub"]*RUB*(n/4)/12
        uk_cost = A["uk_payroll"]+A["uk_other"]; tax = A["uk_tax_royalty"]*(master_share+fee_share)
        out.append(dict(n=n, master_royalty=master_share, master_fees=fee_share, uk_cost=uk_cost, tax=tax, net=master_share+fee_share-uk_cost-tax))
    return out
MASTER = master_franchise_benchmark()

OPENINGS = [(1,"own","Нор-Норк (пилот)"),(8,"own","Малатия-Себастия"),(14,"own","Ачапняк"),
            (18,"fr","Арабкир"),(21,"fr","Шенгавит (у метро)"),(24,"fr","Кентрон"),(27,"fr","Канакер-Зейтун"),(30,"fr","Эребуни (у метро)"),(33,"fr","Гюмри"),(36,"fr","Давташен")]
def network_consolidated():
    runs = {}
    for (m0, kind, nm) in OPENINGS:
        p = default_params(name=nm, indexation=False) if kind == "own" else default_params(name=nm, indexation=False, royalty=OWN_ROY, royalty_min=A["own_royalty_min"], franchise_fee=OWN_FEE)
        runs[(m0, kind, nm)] = run(p)
    rows = []; cum = cum_nofee = 0.0; inv = 0.0
    for m in range(1, 37):
        rev_own = ebitda_own = capex_m = rev_fr = royalty = fund = supply_cost = fees = 0.0; n_own = n_fr = 0
        for (m0, kind, nm), r in runs.items():
            if m < m0: continue
            row = r["rows"][m-m0]
            if kind == "own":
                if m == m0: capex_m += r["capex_total"]
                n_own += 1; rev_own += row["revenue"]; ebitda_own += row["ebitda"]
            else:
                if m == m0: fees += OWN_FEE
                n_fr += 1; rev_fr += row["revenue"]; royalty += row["royalty"]; fund += A["own_mkt_fund"]*row["revenue"]; supply_cost += row["consumables"]
        uk = (A["uk_payroll"]+A["uk_other"])*(0.5 if 7 <= m < 13 else (1.0 if m >= 13 else 0.0))
        if m == 12: capex_m += A["uk_setup"]
        supply = supply_cost*A["supply_margin"]; tax = uk_taxes(royalty, fees, supply_cost); tax_nofee = uk_taxes(royalty, 0, supply_cost)
        net = ebitda_own + royalty + supply + fees - uk - tax      # фонд тратится на маркетинг сети (транзит)
        net_nofee = ebitda_own + royalty + supply - uk - tax_nofee
        cf = net - capex_m; cum += cf; cum_nofee += net_nofee - capex_m; inv += capex_m
        rows.append(dict(month=m, n_own=n_own, n_fr=n_fr, rev_own=rev_own, rev_fr=rev_fr, ebitda_own=ebitda_own, royalty=royalty, fund=fund, supply=supply, fees=fees,
                         uk=uk, uk_tax=tax, net=net, net_nofee=net_nofee, capex=capex_m, cf=cf, cum=cum, cum_nofee=cum_nofee))
    return rows, inv
NET_ROWS, NET_INV = network_consolidated()
N_FR = sum(1 for o in OPENINGS if o[1] == "fr"); N_OWN = sum(1 for o in OPENINGS if o[1] == "own")

def chio_rf():
    rev = 50*700*DAYS  # 50 клиентов × 700 ₽ (реальный прайс Казани 700/800/500)
    fot = 0.40*rev; rent = 50000; roy = 0.04*rev; cons = 0.04*rev; mkt = 10000+0.02*rev; util = 15000; tax = 0.03*rev; other = 0.02*rev
    opex = fot+rent+roy+cons+mkt+util+tax+other
    return dict(revenue=rev, fot=fot, rent=rent, royalty=roy, consumables=cons, marketing=mkt, utilities=util, tax=tax, other=other, opex=opex, ebitda=rev-opex)
CHIO = chio_rf()

# ----------------------------------------------------------------------------------------------
# 5. ПРОВЕРКИ
# ----------------------------------------------------------------------------------------------
def verify(res):
    errs = []; cum = -res["capex_total"]
    for r in res["rows"]:
        lines = r["fot"]+r["rent"]+r["manager"]+r["consumables"]+r["utilities"]+r["it"]+r["hdm"]+r["acquiring"]+r["marketing"]+r["accounting"]+r["other"]+r["royalty"]+r["turnover_tax"]+r["ip_fixed"]
        if abs(lines-r["opex"]) > 0.01: errs.append(f"m{r['month']} opex")
        if abs(r["revenue"]-r["opex"]-r["ebitda"]) > 0.01: errs.append(f"m{r['month']} ebitda")
        if abs(r["ebitda"]-r["depreciation"]-r["net_profit"]) > 0.01: errs.append(f"m{r['month']} net")
        if abs(max(0.10*r["revenue"]-0.06*r["documented"], 0.045*r["revenue"]) - r["turnover_tax"]) > 0.01: errs.append(f"m{r['month']} tax rule")
        doc_expected = r["consumables"]+r["utilities"]+r["it"]+r["hdm"]+r["acquiring"]+r["marketing"]+r["accounting"]+r["royalty"]+0.5*r["other"] + (r["fot"] if r["fot_documented"] else 0) + (r["rent"] if res["params"]["rent_documented"] else 0)
        if abs(doc_expected-r["documented"]) > 0.01: errs.append(f"m{r['month']} documented")
        cum += r["cf"]
        if abs(cum-r["cum_cf"]) > 0.01: errs.append(f"m{r['month']} cum")
    if abs(sum(res["capex_items"].values())-res["capex_total"]) > 0.01: errs.append("capex sum")
    return errs
ALL_ERRS = {}
for k, v in list(R.items())+list(VAR.items())+[("own_fr", R_OWN_FR)]:
    e = verify(v)
    if e: ALL_ERRS[k] = e
c1 = c2 = 0.0
for r in NET_ROWS:
    c1 += r["cf"]; c2 += r["net_nofee"]-r["capex"]
    if abs(c1-r["cum"]) > 0.01 or abs(c2-r["cum_nofee"]) > 0.01: ALL_ERRS["network"] = ["cum mismatch"]
if abs(max(0.10*2_500_000-0.06*1_650_000, 0.045*2_500_000)-151000) > 0.01: ALL_ERRS["file3_example"] = ["mismatch"]
# gross-up самосогласованность
if abs(net_employee(gross_up(300000))-300000) > 1: ALL_ERRS["grossup"] = ["mismatch"]
print("ПРОВЕРКИ:", "OK — все тождества сходятся" if not ALL_ERRS else ALL_ERRS)

# ----------------------------------------------------------------------------------------------
# 6. CSV
# ----------------------------------------------------------------------------------------------
CSV_COLS = ["month","load","clients_day","visits_day","check","revenue","fot","rent","manager","consumables","utilities","it","hdm","acquiring","marketing","accounting","other","royalty","documented","turnover_tax","ip_fixed","opex","ebitda","depreciation","net_profit","cf","cum_cf","cum_dcf"]
for name, res in R.items():
    path = os.path.join(HERE, f"pnl_36m_{name}.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";"); w.writerow(CSV_COLS)
        for r in res["rows"]:
            w.writerow([round(r[c], 4) if c == "load" else round(r[c], 1) if c in ("clients_day","visits_day") else round(r[c]) for c in CSV_COLS])
    print("CSV:", path)

# ----------------------------------------------------------------------------------------------
# 7. EXCEL
# ----------------------------------------------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
wb = Workbook(); ws = wb.active; ws.title = "Assumptions"
bold = Font(bold=True); hdr_fill = PatternFill("solid", fgColor="DDEBF7"); inp_fill = PatternFill("solid", fgColor="FFF2CC")
ws.append(["Параметр","Значение","Ед.","Источник","Достоверность","Комментарий"])
for c in ws[1]: c.font = bold; c.fill = hdr_fill
CELL = {}
for i, (k, v, u, s, cf, cm) in enumerate(ASSUMPTIONS, start=2):
    ws.append([k, (int(v) if isinstance(v, bool) else v), u, s, cf, cm]); CELL[k] = f"Assumptions!$B${i}"; ws.cell(row=i, column=2).fill = inp_fill
r0 = ws.max_row+2
ws.cell(row=r0, column=1, value="ПАРАМЕТРЫ БАЗОВОГО СЦЕНАРИЯ (жёлтые — вход)").font = bold
scen = [("sc_chairs",3,"кресел"),("sc_masters",4,"мастеров"),("sc_check",round(check_A,2),"средний чек (=Prices!C12)"),("sc_price_mult",1.0,"множитель цены"),
        ("sc_load_mature",P_BASE["load_mature"],"зрелая загрузка"),("sc_rent_sqm",P_BASE["rent_sqm"],"AMD/м²"),("sc_area",P_BASE["area"],"м²"),("sc_fot_share",0.5,"доля выплат мастерам (схемы a/c)"),
        ("sc_fot_documented",1,"1 = мастера-ИП, выплаты документированы (схема c); 0 = сотрудники"),("sc_rent_documented",0,"1 = арендодатель выдаёт документы"),
        ("sc_royalty",0.0,"роялти доля"),("sc_royalty_min",0.0,"роялти минимум"),("sc_franchise_fee",0.0,"паушальный AMD"),("sc_consumables_per_client",95,"AMD/клиент"),
        ("sc_manager",P_BASE["manager_cost"],"управляющий AMD/мес"),("sc_capacity_day",round(cap3,3),"ёмкость услуг/день (формула)")]
for j, (k, v, u) in enumerate(scen):
    rr = r0+1+j; ws.cell(row=rr, column=1, value=k); ws.cell(row=rr, column=2, value=v).fill = inp_fill; ws.cell(row=rr, column=3, value=u); CELL[k] = f"Assumptions!$B${rr}"
addr = lambda k: CELL[k].split("!")[1].replace("$","")
ws[addr("sc_capacity_day")] = f"={CELL['sc_chairs']}*{CELL['hours_per_day']}*60/{CELL['minutes_avg']}"
ws[addr("sc_check")] = "=Prices!C12"
for col, w_ in zip("ABCDEF", (28,16,14,80,24,70)): ws.column_dimensions[col].width = w_

wp = wb.create_sheet("Prices")
wp.append(["Услуга","Цена A (спальный)","Цена B (центр)","Доля в миксе","Мин/услуга","Вклад A","Вклад B","Чио Чио Казань 2026, ₽","В AMD по 4,3731","Доп. услуга"])
for c in wp[1]: c.font = bold; c.fill = hdr_fill
chio_px = {"mens":700,"clip":500,"kids":500,"wshort":800,"wlong":800,"beard":300,"style":None,"roots":2000}
for i, s in enumerate(SERVICES, start=2):
    wp.append([s[1], s[2], s[3], s[4], s[5], f"=B{i}*D{i}", f"=C{i}*D{i}", chio_px[s[0]], (f"=H{i}*{CELL['fx_rub']}" if chio_px[s[0]] else None), int(s[6])])
    for col in (2,3,4,5): wp.cell(row=i, column=col).fill = inp_fill
n = len(SERVICES)+1
wp.append([]); wp.append(["Сумма долей", None, None, f"=SUM(D2:D{n})"])
wp.append(["СРЕДНИЙ ЧЕК за услугу, AMD", None, f"=SUM(F2:F{n})", None, None, None, f"=SUM(G2:G{n})"])  # row 12
wp.append(["Средний чек в USD", None, f"=C12/{CELL['fx_usd']}", None, None, None, f"=G12/{CELL['fx_usd']}"])
wp.append(["Средний чек в RUB", None, f"=C12/{CELL['fx_rub']}", None, None, None, f"=G12/{CELL['fx_rub']}"])
wp.append(["Средневзвеш. мин/услуга", None, f"=SUMPRODUCT(D2:D{n},E2:E{n})"])
wp.append(["Доля доп. услуг", None, f"=SUMPRODUCT(D2:D{n},J2:J{n})"])
wp.append(["Средний чек за ВИЗИТ, AMD", None, "=C12/(1-C16)", None, None, None, "=G12/(1-C16)"])
for col, w_ in zip("ABCDEFGHIJ", (40,16,16,12,10,12,12,22,16,10)): wp.column_dimensions[col].width = w_

wc = wb.create_sheet("CAPEX")
wc.append(["Статья","Своя сеть 40 м² (AMD)","Чио Чио «Лайт»","Чио Чио стандарт","Своя сеть 25 м², 2 кресла","Своя сеть 40 м², USD"])
for c in wc[1]: c.font = bold; c.fill = hdr_fill
def C4(expr3, expr2): return (expr3, expr3, expr3, expr2)
imp3 = f"({CELL['sc_chairs']}*{CELL['chair_price']}+{CELL['wash_unit']}+2*{CELL['vacuum_unit']}+{CELL['sc_chairs']}*{CELL['tools_per_post']})"
imp2 = f"(2*{CELL['chair_price']}+{CELL['wash_unit']}+2*{CELL['vacuum_unit']}+2*{CELL['tools_per_post']})"
capex_rows = [
 ("Ремонт/отделка", *C4(f"={CELL['sc_area']}*{CELL['fitout_sqm']}", f"=25*{CELL['fitout_sqm']}")),
 ("Кресла", *C4(f"={CELL['sc_chairs']}*{CELL['chair_price']}", f"=2*{CELL['chair_price']}")),
 ("Рабочие посты", *C4(f"={CELL['sc_chairs']}*{CELL['post_price']}", f"=2*{CELL['post_price']}")),
 ("Мойка", *C4(f"={CELL['wash_unit']}", f"={CELL['wash_unit']}")),
 ("Вакуумные аппараты × 2", *C4(f"=2*{CELL['vacuum_unit']}", f"=2*{CELL['vacuum_unit']}")),
 ("Инструмент/техника", *C4(f"={CELL['sc_chairs']}*{CELL['tools_per_post']}", f"=2*{CELL['tools_per_post']}")),
 ("НДС 20% на ввоз оборудования [v2]", *C4(f"={imp3}*{CELL['import_vat']}", f"={imp2}*{CELL['import_vat']}")),
 ("Кондиционеры + вентиляция [v2]", *C4(f"={CELL['hvac']}", f"={CELL['hvac']}")),
 ("Водонагреватель [v2]", *C4(f"={CELL['boiler']}", f"={CELL['boiler']}")),
 ("Стерилизатор", *C4(f"={CELL['sterilizer']}", f"={CELL['sterilizer']}")),
 ("Мебель ожидания/стойка", *C4(f"={CELL['waiting_furniture']}", f"={CELL['waiting_furniture']}")),
 ("Вывеска", *C4(f"={CELL['signage']}", f"={CELL['signage']}")),
 ("Согласование вывески (оценка) [v2]", *C4(f"={CELL['signage_permit']}", f"={CELL['signage_permit']}")),
 ("Брендинг интерьера", *C4(f"={CELL['branding_interior']}", f"={CELL['branding_interior']}")),
 ("Киоск/ПК/камеры", *C4(f"={CELL['selfservice_it']}", f"={CELL['selfservice_it']}")),
 ("Касса e-HDM", *C4(f"={CELL['hdm_setup']}", f"={CELL['hdm_setup']}")),
]
for i, row in enumerate(capex_rows, start=2): wc.append([row[0], row[1], row[2], row[3], row[4], f"=B{i}/{CELL['fx_usd']}"])
hard_end = len(capex_rows)+1; hr = hard_end+1
wc.append(["Итого hard-CAPEX", f"=SUM(B2:B{hard_end})", f"=SUM(C2:C{hard_end})", f"=SUM(D2:D{hard_end})", f"=SUM(E2:E{hard_end})", f"=B{hr}/{CELL['fx_usd']}"]); wc.cell(row=hr, column=1).font = bold
soft_rows = [
 ("Непредвиденное 15% [v2]", f"=B{hr}*{CELL['contingency']}", f"=C{hr}*{CELL['contingency']}", f"=D{hr}*{CELL['contingency']}", f"=E{hr}*{CELL['contingency']}"),
 ("CRM/сайт", *C4(f"={CELL['crm_setup']}", f"={CELL['crm_setup']}")),
 ("Стажировка мастеров [v2]", *C4(f"={CELL['training']}", f"={CELL['training']}")),
 ("Регистрация ИП", *C4(f"={CELL['registration']}", f"={CELL['registration']}")),
 ("Юр./налоговое оформление", *C4(f"={CELL['legal_setup']}", f"={CELL['legal_setup']}")),
 ("Медкнижки", *C4(f"={CELL['medbooks']}", f"={CELL['medbooks']}*3/4")),
 ("Аренда за период ремонта [v2]", *C4(f"={CELL['fitout_months']}*{CELL['sc_rent_sqm']}*{CELL['sc_area']}", f"={CELL['fitout_months']}*{CELL['sc_rent_sqm']}*25")),
 ("Депозит аренды", *C4(f"={CELL['deposit_months']}*{CELL['sc_rent_sqm']}*{CELL['sc_area']}", f"={CELL['deposit_months']}*{CELL['sc_rent_sqm']}*25")),
 ("Маркетинг открытия", *C4(f"={CELL['launch_marketing']}", f"={CELL['launch_marketing']}")),
 ("Паушальный взнос", "0", f"={CELL['fee_chio_lite_rub']}*{CELL['fx_rub']}", f"={CELL['fee_chio_std_rub']}*{CELL['fx_rub']}", "0"),
]
for j, row in enumerate(soft_rows):
    i = hr+1+j; wc.append([row[0], row[1], row[2], row[3], row[4], f"=B{i}/{CELL['fx_usd']}"])
se = hr+len(soft_rows); tot_r = se+1
wc.append(["ИТОГО CAPEX", f"=B{hr}+SUM(B{hr+1}:B{se})", f"=C{hr}+SUM(C{hr+1}:C{se})", f"=D{hr}+SUM(D{hr+1}:D{se})", f"=E{hr}+SUM(E{hr+1}:E{se})", f"=B{tot_r}/{CELL['fx_usd']}"]); wc.cell(row=tot_r, column=1).font = bold
wc.append(["Итого CAPEX, USD", f"=B{tot_r}/{CELL['fx_usd']}", f"=C{tot_r}/{CELL['fx_usd']}", f"=D{tot_r}/{CELL['fx_usd']}", f"=E{tot_r}/{CELL['fx_usd']}"])
wc.append(["Итого CAPEX, RUB", f"=B{tot_r}/{CELL['fx_rub']}", f"=C{tot_r}/{CELL['fx_rub']}", f"=D{tot_r}/{CELL['fx_rub']}", f"=E{tot_r}/{CELL['fx_rub']}"])
fee_row = hr+10
wc.append(["Амортизируемая база (hard + непредвиденное + CRM + паушальный)", f"=B{hr}+B{hr+1}+B{hr+2}", f"=C{hr}+C{hr+1}+C{hr+2}+C{fee_row}", f"=D{hr}+D{hr+1}+D{hr+2}+D{fee_row}", f"=E{hr}+E{hr+1}+E{hr+2}"])
wc.append(["Оборотный резерв на стресс ramp-up (значение из Python; не входит в CAPEX)", round(b["reserve"]), round(VAR["fr_lite"]["reserve"]), round(VAR["fr_std"]["reserve"]), round(VAR["chairs2"]["reserve"])])
wc.append(["ПОТРЕБНОСТЬ В ФИНАНСИРОВАНИИ = CAPEX + резерв", f"=B{tot_r}+B{tot_r+4}", f"=C{tot_r}+C{tot_r+4}", f"=D{tot_r}+D{tot_r+4}", f"=E{tot_r}+E{tot_r+4}"])
CAPEX_TOTAL_CELL = f"CAPEX!$B${tot_r}"; DEPR_BASE_CELL = f"CAPEX!$B${tot_r+3}"
for col, w_ in zip("ABCDEF", (52,22,20,20,26,18)): wc.column_dimensions[col].width = w_

wl = wb.create_sheet("PnL_36m")
cols = ["Мес.","Индекс цен","Загрузка","Услуг/день","Средний чек","ВЫРУЧКА","Выплаты мастерам","Аренда","Управл.+уборка","Расходники","Коммуналка","IT/связь","Касса","Эквайринг","Маркетинг","Бухгалтерия","Прочее","Роялти","Документ. расходы","Налог с оборота","Фиксы ИП","ИТОГО РАСХОДЫ","EBITDA","Амортизация","ЧИСТАЯ ПРИБЫЛЬ","CF","Накопл. CF","Дисконт-фактор","Накопл. DCF"]
wl.append(cols)
for c in wl[1]: c.font = bold; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True)
wl.append([0]+[None]*24+[f"=-{CAPEX_TOTAL_CELL}", "=Z2", 1, "=Z2"])
r_m = f"((1+{CELL['discount_rate']})^(1/12)-1)"
for m in range(1, 37):
    i = m+2; shape = round(LOAD_SHAPE[m-1], 6)
    wl.append([m, f"=(1+{CELL['cpi']})^INT((A{i}-1)/12)", f"=MIN(0.95,{shape}*{CELL['sc_load_mature']})", f"=C{i}*{CELL['sc_capacity_day']}",
        f"={CELL['sc_check']}*{CELL['sc_price_mult']}*B{i}", f"=D{i}*E{i}*{CELL['days_per_month']}",
        f"=MAX({CELL['sc_fot_share']}*F{i},{CELL['sc_masters']}*{CELL['fot_guarantee']}*B{i})",
        f"={CELL['sc_rent_sqm']}*{CELL['sc_area']}*B{i}", f"=({CELL['sc_manager']}+{CELL['cleaning_cost']})*B{i}",
        f"=MAX({CELL['sc_consumables_per_client']}*B{i}*D{i}*{CELL['days_per_month']},{CELL['consumables_floor']}*B{i})",
        f"={CELL['utilities']}*B{i}", f"={CELL['it_crm']}*B{i}", f"={CELL['hdm_monthly']}", f"=F{i}*{CELL['card_share']}*{CELL['acq_rate']}+{CELL['acq_terminal']}",
        f"=IF(A{i}<=12,MAX({CELL['marketing_y1']}*F{i},{CELL['marketing_y1_floor']}*B{i}),MAX({CELL['marketing_pct']}*F{i},{CELL['marketing_floor']}*B{i}))",
        f"=({CELL['accounting']}+IF({CELL['sc_fot_documented']}=1,{CELL['accounting_ip_extra']},0))*B{i}",
        f"=MAX({CELL['other_pct']}*F{i},{CELL['other_floor']}*B{i})",
        f"=IF(AND({CELL['sc_royalty']}>0,A{i}>=2),MAX({CELL['sc_royalty']}*F{i},{CELL['sc_royalty_min']}),0)",
        f"=J{i}+K{i}+L{i}+M{i}+N{i}+O{i}+P{i}+R{i}+Q{i}*{CELL['other_documented_share']}+IF({CELL['sc_fot_documented']}=1,G{i},0)+IF({CELL['sc_rent_documented']}=1,H{i},0)",
        f"=MAX({CELL['tt_rate']}*F{i}-{CELL['tt_deduct']}*S{i},{CELL['tt_floor']}*F{i})",
        f"={CELL['ip_fixed']}+IF(A{i}>=13,{CELL['ip_oms_year']}/12,0)",
        f"=SUM(G{i}:R{i})+T{i}+U{i}", f"=F{i}-V{i}", f"=({DEPR_BASE_CELL}+{CELL['sc_franchise_fee']})/{CELL['depr_months']}", f"=W{i}-X{i}", f"=W{i}",
        f"=AA{i-1}+Z{i}", f"=1/(1+{r_m})^A{i}", f"=AC{i-1}+Z{i}*AB{i}"])
wl.append([]); tr = wl.max_row+1
for y in range(3):
    wl.cell(row=tr+y, column=1, value=f"Год {y+1}").font = bold
    a_, b_ = 3+12*y, 14+12*y
    for col in range(6, 26):
        L = get_column_letter(col); wl.cell(row=tr+y, column=col, value=f"=SUM({L}{a_}:{L}{b_})")
labels = [("NPV @discount_rate, 36 мес.", "=Z2+SUMPRODUCT(Z3:Z38,AB3:AB38)"), ("IRR годовая", "=(1+IRR(Z2:Z38))^12-1"),
          ("Окупаемость простая, мес.", '=IFERROR(INDEX(A3:A38,MATCH(TRUE,INDEX(AA3:AA38>=0,0),0)),"нет за 36 мес.")'),
          ("Окупаемость дисконтированная, мес.", '=IFERROR(INDEX(A3:A38,MATCH(TRUE,INDEX(AC3:AC38>=0,0),0)),"нет за 36 мес.")'),
          ("Проверка: Σ(выручка − расходы − EBITDA) = 0", "=SUM(F3:F38)-SUM(V3:V38)-SUM(W3:W38)"),
          ("Проверка: мин(налог − 4,5% выручки) ≥ 0", "=MIN(INDEX(T3:T38-0.045*F3:F38,0))")]
for j, (lab, f_) in enumerate(labels):
    wl.cell(row=tr+4+j, column=1, value=lab).font = bold; wl.cell(row=tr+4+j, column=6, value=f_)
wl.column_dimensions["A"].width = 8
for col in range(2, 30): wl.column_dimensions[get_column_letter(col)].width = 14
wl.freeze_panes = "B2"

def sheet_values(title, header, rows, widths=None):
    w_ = wb.create_sheet(title); w_.append(header)
    for c in w_[1]: c.font = bold; c.fill = hdr_fill
    for r in rows: w_.append(list(r))
    for k, wd in enumerate(widths or [], start=1): w_.column_dimensions[get_column_letter(k)].width = wd
    return w_
S3 = ("pessimistic","base","optimistic")
sc_rows = [("Параметры: цена ×, зрелая загрузка, аренда AMD/м², схема мастеров", "0,9 / 30% / 8 000 / сотрудники (a)", "1,0 / 40% / 6 500 / ИП (c)", "1,0 / 45% / 5 800 / ИП (c)")]
for y in (1,2,3):
    for lab, key in (("Выручка","revenue"),("Выплаты мастерам","fot"),("Налог с оборота","turnover_tax"),("EBITDA","ebitda"),("Чистая прибыль","net_profit")):
        sc_rows.append((f"{lab}, год {y}", *(round(R[s][f"y{y}"][key]) for s in S3)))
    sc_rows.append((f"Услуг/день, среднее за год {y}", *(round(R[s][f"y{y}"]["clients_day"],1) for s in S3)))
sc_rows += [("CAPEX (без резерва)", *(round(R[s]["capex_total"]) for s in S3)), ("Резерв на стресс ramp-up", *(round(R[s]["reserve"]) for s in S3)), ("Потребность в финансировании", *(round(R[s]["financing"]) for s in S3)), ("Окупаемость простая, мес.", *(R[s]["payback"] or "нет за 36" for s in S3)),
            ("Окупаемость дисконтированная, мес.", *(R[s]["dpayback"] or "нет за 36" for s in S3)), ("NPV 36 мес. @18%", *(round(R[s]["npv"]) for s in S3)),
            ("NPV 36 мес. @18% с TV 3×EBITDA Y3", *(round(R[s]["npv_tv"]) for s in S3)), ("IRR годовая", *((f"{R[s]['irr']*100:.1f}%" if R[s]['irr'] is not None else "n/a") for s in S3))]
sheet_values("Scenarios", ["Показатель","Пессимистичный","Базовый","Оптимистичный"], sc_rows, [52,30,30,30])
sn_rows = []
for s in SENS:
    sn_rows.append([s["factor"]]+[round(s[d]["net36"]) for d in (-0.2,-0.1,0,0.1,0.2)]+["Чистая прибыль 36 мес."])
    sn_rows.append([s["factor"]]+[s[d]["payback"] or ">36" for d in (-0.2,-0.1,0,0.1,0.2)]+["Окупаемость, мес."])
    sn_rows.append([s["factor"]]+[round(s[d]["npv"]) for d in (-0.2,-0.1,0,0.1,0.2)]+["NPV @18%"])
sn_rows.append([]); sn_rows.append(["Переключатели", "Налог м36", "Чистая прибыль м36", "Чистая прибыль 36 мес.", "Окупаемость", "NPV @18%", ""])
for lab, r in SWITCHES: sn_rows.append([lab, round(r["m36"]["turnover_tax"]), round(r["m36"]["net_profit"]), round(sum(x["net_profit"] for x in r["rows"])), r["payback"] or ">36", round(r["npv"]), ""])
sheet_values("Sensitivity", ["Фактор","−20%","−10%","база","+10%","+20%","Показатель"], sn_rows, [60,16,16,16,16,16,30])
sheet_values("Schemes", ["Схема оплаты мастеров (мес. 36 базы)","Брутто на мастера, AMD","На руки мастеру, AMD","На руки, % выручки точки","Стоимость для салона (4 мастера)","Вычет 6% для салона"],
             [(nm, round(g), round(net_), round(net_*4/m36b["revenue"]*1000)/10, round(cost), "да" if doc else "нет") for nm, g, net_, cost, doc in SCHEMES], [60,22,22,22,26,18])
sheet_values("Franchise", ["Паушальный","Роялти","Окупаемость, мес.","Маржа м36","NPV 36 мес. @18%","NPV 60 мес. @18%","IRR 60 мес.","CAPEX субфранчайзи"],
             [(g["fee"], g["roy"], g["payback"] or ">36", round(g["margin"],4), round(g["npv36"]), round(g["npv60"]), (round(g["irr60"],4) if g["irr60"] is not None else "n/a"), round(g["capex"])) for g in GRID], [14,10,18,12,18,18,12,20])
sheet_values("Network", ["Мес.","Собств.","Франч.","Выручка собств.","Выручка франч.","EBITDA собств.","Роялти","Фонд (транзит)","Маржа поставок","Паушальные","Расходы УК","Налог УК","Результат холдинга","Результат без паушальных","CAPEX","CF","Накопл. CF","Накопл. CF без паушальных"],
             [(r["month"], r["n_own"], r["n_fr"])+tuple(round(r[k]) for k in ("rev_own","rev_fr","ebitda_own","royalty","fund","supply","fees","uk","uk_tax","net","net_nofee","capex","cf","cum","cum_nofee")) for r in NET_ROWS], [8]+[15]*17)
wb.calculation.fullCalcOnLoad = True
XLSX = os.path.join(HERE, "ArmChio_model.xlsx"); wb.save(XLSX); print("XLSX:", XLSX)

# ----------------------------------------------------------------------------------------------
# 8. JSON
# ----------------------------------------------------------------------------------------------
KEY = dict(version=2, check_A=check_A, check_B=check_B, check_visit_A=check_visit_A, capacity_3=cap3,
           base=dict(m36=m36b, capex=b["capex_total"], reserve=b["reserve"], payback=b["payback"], dpayback=b["dpayback"], npv=b["npv"], npv_tv=b["npv_tv"], irr=b["irr"], y1=b["y1"], y2=b["y2"], y3=b["y3"]),
           pess=dict(capex=pe["capex_total"], payback=pe["payback"], npv=pe["npv"], irr=pe["irr"], y3=pe["y3"]), opt=dict(capex=op["capex_total"], payback=op["payback"], npv=op["npv"], irr=op["irr"], y3=op["y3"]),
           breakeven=BE, schemes=[(s[0], s[1], s[2], s[3]) for s in SCHEMES], franchise_grid=GRID, own_package=dict(fee=OWN_FEE, royalty=OWN_ROY, payback=R_OWN_FR["payback"], npv36=R_OWN_FR["npv"], npv60=R_OWN_FR["npv60"]),
           fr_lite=dict(capex=VAR["fr_lite"]["capex_total"], payback=VAR["fr_lite"]["payback"], npv=VAR["fr_lite"]["npv"]), fr_std=dict(capex=VAR["fr_std"]["capex_total"], payback=VAR["fr_std"]["payback"], npv=VAR["fr_std"]["npv"]),
           network=dict(yerevan=[yerevan_lo, yerevan_base, yerevan_hi], regions=[regions_lo, regions_base], inv=NET_INV, cum36=NET_ROWS[-1]["cum"], cum36_nofee=NET_ROWS[-1]["cum_nofee"], n_own=N_OWN, n_fr=N_FR, uk_be=N_BE_UK), errors=ALL_ERRS)
with open(os.path.join(HERE, "model_output.json"), "w", encoding="utf-8") as f: json.dump(KEY, f, ensure_ascii=False, indent=1, default=str)

# ----------------------------------------------------------------------------------------------
# 9. ОТЧЁТ
# ----------------------------------------------------------------------------------------------
def T(headers, rows):
    s = "| " + " | ".join(headers) + " |\n|" + "---|"*len(headers) + "\n"
    for r in rows: s += "| " + " | ".join(str(x) for x in r) + " |\n"
    return s
def pb(r): return r["payback"] if r["payback"] else "нет за 36"
def dpb(r): return r["dpayback"] if r["dpayback"] else "нет за 36"
def irr_s(x): return f"{x*100:.0f}%" if x is not None else "n/a"
net36 = lambda r: sum(x["net_profit"] for x in r["rows"][:36])
sum_men_dom = TOPDOWN[2]
FLOOR_MONTHS = ", ".join(str(r["month"]) for r in b["rows"] if abs(r["turnover_tax"]-0.045*r["revenue"]) < 0.5) or "ни в одном"
m36e = VAR["emp"]["m36"]; m36g = VAR["grossup"]["m36"]
md = []
md.append(f"""# ArmChio — финансовая модель адаптации формата «Чио Чио» под Армению (v2)

Дата: 20.10.2026 (v2 после независимой верификации — `verification_report.md`). Роль: финансовый аналитик / CFO франчайзинговой сети. Валюта — AMD; курс ЦБА 19.10.2026: 1 USD = {A['fx_usd']} AMD, 1 RUB = {A['fx_rub']} AMD.
Источники входных данных — четыре файла исследования ({F1}, {F2}, {F3}, repo_inventory.md), отчёт верификации и явно помеченные допущения. Все расчёты — `model/model.py` (`python -X utf8 -I model.py`); тождества (выручка − расходы = EBITDA; EBITDA − амортизация = чистая прибыль; накопленный поток; состав документированных расходов; налог с оборота «10% − 6% документированных, но не менее 4,5%», контрольный пример файла 3 §4.3 → 151 000; gross-up ↔ net; пример gross 250 000 → 181 700 на руки) проверены скриптом: **{'OK — все тождества сходятся' if not ALL_ERRS else ALL_ERRS}**. Формулы Excel проверены только через openpyxl и python-пересчёт (без COM).

Файлы: `C:\\Новая папка\\ArmChio-analysis\\repo\\model\\model.py`, `ArmChio_model.xlsx` (Assumptions/Prices/CAPEX/PnL_36m — формулы; Scenarios/Sensitivity/Schemes/Franchise/Network — значения), `pnl_36m_base|pessimistic|optimistic.csv`, `model_output.json`; v1 — `financial_model_v1.md`, `model_v1.py`.

## Что изменилось в v2 и почему

Верификация (FAIL условный) выявила два допущения, каждое из которых обнуляло NPV базы: (1) ФОТ и аренда у физлица засчитывались в документированные расходы для вычета 6% из налога с оборота — закон ՀՕ-72-Ն требует расчётных документов ст. 55 ч. 2 п. 1–5 НК (счёт-фактура, накладная, чек ККМ), зарплата и аренда у физлица ими не оформляются; (2) «50% мастеру» трактовалось как gross при рыночной практике «на руки». Плюс: ёмкость сети и график открытий завышены (Казань Чио Чио откатилась с 41–42 до 26 адресов), CAPEX недоучтён на 1,0–1,6 млн, загрузка 45% = средняя зрелая точка Чио Чио с брендом и колл-центром.

""")
md.append(T(["Показатель","v1","v2","Почему"],
 [("Налог с оборота: вычет 6%", "по ФОТ, аренде, управляющему и всем расходам", "только по расходам с расчётными документами: расходники, коммуналка, IT, касса, эквайринг, маркетинг, бухгалтерия, роялти, 50% прочих; **выплаты мастерам — только если мастера ИП (база)**; аренда — параметр (база: нет)", f"{FV} C.2, A.15"),
  ("Схема оплаты мастеров", "50% gross сотрудникам (на руки ≈37% выручки)", "три схемы: (a) сотрудник 50% gross, (b) gross-up до 50% на руки, **(c) мастер-ИП 50% — база**", f"{FV} C.3"),
  ("Зрелая загрузка: пессим./база/оптим.", "32% / 45% / 55%", f"{pct(A['load_mature_pess'],0)} / {pct(A['load_mature_base'],0)} / {pct(A['load_mature_opt'],0)}", f"{FV} C.5, D.2"),
  ("Маркетинг", "3% (пол 60 тыс.)", "год 1 — 5% (пол 120 тыс.), далее 3%", f"{FV} D.6"),
  ("Гарантия мастеру на ramp-up", "150 000 gross", "200 000 брутто (≈170 тыс. на руки ИП)", f"{FV} D.8"),
  ("Непредвиденное в CAPEX", "10%", "15%", f"{FV} C.6"),
  ("Новые строки CAPEX", "—", "аренда на ремонт 1,25 мес., НДС 20% на ввоз, кондиционеры/вентиляция, бойлер, согласование вывески (оценка), стажировка мастеров", f"{FV} C.6"),
  ("Оборотный резерв", "просадка по базовой кривой (577 тыс.), включён в CAPEX и одновременно учтён в CF (двойной счёт)", "просадка при стрессе (м1–3 × 0,5); показан отдельно как потребность в финансировании, в базу NPV/окупаемости не входит (операционные убытки ramp-up уже в CF)", f"{FV} C.5; методическая правка v2"),
  ("Площадь точки", "40 м² («факт»)", "30 м² («решение»): 3 поста + ожидание + зона персонала + санузел; 40 м² — вариант", f"{FV} D.15, B.3"),
  ("Прайс", "муж. 2 500", "без изменений (2 500); прайс A+ 2 800 (−7% к медиане) — рычаг, не база", f"спрос на 2 000–2 500 не доказан ({FV} E.1); потолок — медиана эконом 3 000"),
  ("Фиксы ИП", "30 800 с 1-го мес. (гербовый 1 000 в тексте)", "20 000 + ОМС 10 800 со 2-го года; гербовый с з/п — шкала 1 500–15 000", f"{FV} C.4, A.16"),
  ("Налог УК на роялти/паушальные", "10% − 6%, пол 4,5%", "10% без вычета; поставки — торговля 10% − 9,5%", f"{FV} C.9"),
  ("Аренда спальных районов", "6 500 «факт»", "6 500 «оценка, n=3» (5 556/5 808/9 583); стресс 9 500 и 5 800", f"{FV} A.24"),
  ("Ёмкость сети", f"Ереван {V1['yerevan']}, Армения {V1['armenia']}", f"Ереван {yerevan_lo}–{yerevan_hi} (база {yerevan_base}), Армения {yerevan_lo+regions_lo}–{yerevan_hi+regions_base}", f"{FV} C.9"),
  ("График открытий к мес. 36", V1["schedule"], f"{N_OWN} собств. + {N_FR} франч.; холдинг показан с паушальными и без", f"{FV} C.9"),
  ("Пакет своей франшизы", V1["own_fr"], f"{money(OWN_FEE)} + {pct(OWN_ROY)} мин. 50 000; NPV субфранчайзи показан за 36 и 60 мес.", f"{FV} D.11"),
  ("Сравнение с Чио Чио РФ", "чек 600 ₽", "реальный прайс Казани 700/800/500 ₽ → 700 ₽", f"{FV} B.1"),
  ("Метрика клиентов", "услуг/день", "услуг/день и визитов/день (борода/укладка — доп.), чек за услугу и за визит", f"{FV} C.5"),
  ("Пометки «факт» у chairs/area/rent/fot/days", "факт", "решение / оценка (n=3) / допущение", f"{FV} B.3")]))
md.append("\n**Ключевые показатели: старое → новое**\n\n")
md.append(T(["Показатель","v1","v2 (база)"],
 [("Средний чек за услугу / за визит", f"{money(V1['check'])} / —", f"{money(check_A)} / {money(check_visit_A)}"),
  ("Зрелая точка (м36): загрузка, услуг/день", f"{pct(V1['load'],0)}, {V1['clients']:.0f}", f"{pct(m36b['load'],0)}, {m36b['clients_day']:.0f} (визитов {m36b['visits_day']:.0f})"),
  ("Выручка м36", money(V1["revenue"]), money(m36b["revenue"])),
  ("EBITDA м36", f"{money(V1['ebitda'])} ({pct(V1['ebitda_m'])})", f"{money(m36b['ebitda'])} ({pct(m36b['ebitda']/m36b['revenue'])})"),
  ("Чистая прибыль м36", f"{money(V1['net'])} ({pct(V1['net_m'])})", f"{money(m36b['net_profit'])} ({pct(m36b['net_profit']/m36b['revenue'])})"),
  ("Налог с оборота м36, % выручки", pct(V1["tax_pct"],2), pct(m36b["turnover_tax"]/m36b["revenue"],2)),
  ("CAPEX без резерва / резерв / потребность в финансировании", f"{money(V1['capex_ex'])} / {money(V1['reserve'])} / {money(V1['capex'])}", f"{money(b['capex_total'])} / {money(b['reserve'])} / {money(b['financing'])}"),
  ("Безубыточность (чистая прибыль = 0), услуг/день", f"{V1['be']:.0f}", f"{BE['clients_day']:.0f}"),
  ("Окупаемость простая / дисконтированная", f"{V1['payback']} / {V1['dpayback']}", f"{pb(b)} / {dpb(b)}"),
  ("NPV 36 мес. @18%", money(V1["npv"]), money(b["npv"])),
  ("IRR 36 мес.", pct(V1["irr"]), irr_s(b["irr"])),
  ("Окупаемость пессим. / оптим.", f"нет за 36 / {V1['opt_payback']}", f"{pb(pe)} / {pb(op)}"),
  ("Франшиза Чио Чио «Лайт»/стандарт: окупаемость", f"{V1['fr_lite_pb']} / {V1['fr_std_pb']}", f"{pb(VAR['fr_lite'])} / {pb(VAR['fr_std'])}"),
  ("Холдинг (сеть): инвестиции / накопл. CF м36", f"{money(V1['holding_inv'])} / {money(V1['holding_cum'])}", f"{money(NET_INV)} / {money(NET_ROWS[-1]['cum'])} (без паушальных {money(NET_ROWS[-1]['cum_nofee'])})"),
  ("Самоокупаемость УК, франч. точек", V1["uk_be"], f"{N_BE_UK:.0f}")]))

md.append(f"""
## Резюме (база v2: 3 кресла, 30 м², спальный район Еревана, своя сеть, мастера-ИП 50%, аренда у физлица без документов, загрузка 40%, прайс 2 500)

| Показатель | Значение |
|---|---|
| Средний чек за услугу (спальный / центр) | **{money(check_A)} / {money(check_B)} AMD** (${usd(check_A):.2f} / ${usd(check_B):.2f}; {rub(check_A):.0f} / {rub(check_B):.0f} ₽); за визит {money(check_visit_A)} AMD |
| Ёмкость (3 кресла × 11 ч, {A['minutes_avg']} мин/услуга) | {cap3:.0f} услуг/день (2 кресла — {cap2:.0f}) |
| Зрелая точка (м36) | {pct(m36b['load'],0)} загрузки → **{m36b['clients_day']:.0f} услуг/день ≈ {m36b['visits_day']:.0f} визитов**; выручка **{money(m36b['revenue'])} AMD/мес** (${usd(m36b['revenue']):,.0f}; {rub(m36b['revenue'])/1000:,.0f} тыс. ₽) |
| EBITDA / чистая прибыль м36 | **{money(m36b['ebitda'])} ({pct(m36b['ebitda']/m36b['revenue'])}) / {money(m36b['net_profit'])} ({pct(m36b['net_profit']/m36b['revenue'])})** ≈ ${usd(m36b['net_profit']):,.0f}/мес |
| Чистая прибыль Y1 / Y2 / Y3 | {money(b['y1']['net_profit'])} / {money(b['y2']['net_profit'])} / {money(b['y3']['net_profit'])} AMD |
| CAPEX (своя сеть, 30 м², с депозитом, арендой на ремонт, маркетингом открытия, 15% непредвиденного) | **{money(b['capex_total'])} AMD ≈ ${usd(b['capex_total']):,.0f} ≈ {rub(b['capex_total'])/1e6:.2f} млн ₽** (hard {money(b['hard_capex'])}); 40 м² — {money(VAR['area40']['capex_total'])} |
| Оборотный резерв (стресс ramp-up) / потребность в финансировании | {money(b['reserve'])} / **{money(b['financing'])} AMD ≈ ${usd(b['financing']):,.0f}** |
| CAPEX франшиза Чио Чио «Лайт» / стандарт | {money(VAR['fr_lite']['capex_total'])} / {money(VAR['fr_std']['capex_total'])} AMD |
| Безубыточность (структура м36) | **{BE['clients_day']:.0f} услуг/день ≈ {BE['visits_day']:.0f} визитов** ({money(BE['revenue'])} AMD/мес, загрузка {pct(BE['load'],0)}); по EBITDA {BE['clients_day_ebitda']:.0f} |
| Окупаемость простая / дисконтированная @18% | **{pb(b)} / {dpb(b)} мес.**; пессим. {pb(pe)}; оптим. {pb(op)} |
| NPV 36 мес. @18% (без TV): пессим. / база / оптим. | {money(pe['npv'])} / **{money(b['npv'])}** / {money(op['npv'])}; с TV 3× EBITDA Y3: {money(pe['npv_tv'])} / {money(b['npv_tv'])} / {money(op['npv_tv'])} |
| IRR 36 мес.: пессим. / база / оптим. | {irr_s(pe['irr'])} / **{irr_s(b['irr'])}** / {irr_s(op['irr'])} |
| Схема мастеров (база) | (c) мастер-ИП, 50% выручки: на руки {money(SCHEMES[2][2])} при брутто {money(g36)} ({pct(SCHEMES[2][2]/g36,0)}); сотрудник (a) получил бы {money(SCHEMES[0][2])}; gross-up (b) стоил бы салону {money(SCHEMES[1][3])} (ФОТ {pct(SCHEMES[1][3]/m36b['revenue'])}) → EBITDA м36 {money(m36g['ebitda'])} |
| Если мастера — сотрудники (ФОТ не вычитается) | налог {pct(m36e['turnover_tax']/m36e['revenue'],1)}, чистая прибыль м36 {money(m36e['net_profit'])}, окупаемость {pb(VAR['emp'])}, NPV {money(VAR['emp']['npv'])} |
| Площадь | 30 м² (3 поста, ожидание, зона персонала, санузел); при 40 м² окупаемость {pb(VAR['area40'])}, NPV {money(VAR['area40']['npv'])} |
| Прайс как рычаг | A+ (муж. 2 800, −7% к медиане эконом): окупаемость {pb(VAR['price28'])}, NPV {money(VAR['price28']['npv'])} — не включён в базу, спрос не доказан |
| Рекомендуемый вариант | **(3) собственный бренд: пилот + 2 собственные точки, франшиза после доказанного спроса**; франшиза Чио Чио (1) — окупаемость {pb(VAR['fr_lite'])}/{pb(VAR['fr_std'])}, NPV {money(VAR['fr_lite']['npv'])} / {money(VAR['fr_std']['npv'])}; мастер-франшиза (2) — не покрывает УК при ≤20 точках |
| Пакет своей франшизы | паушальный {money(OWN_FEE)} + роялти {pct(OWN_ROY)} (мин. 50 000): при базе субфранчайзи окупаемость {pb(R_OWN_FR)} мес., NPV 36 {money(R_OWN_FR['npv'])}, NPV 60 {money(R_OWN_FR['npv60'])}; условие «NPV36 > 0 и окупаемость ≤ 30» выполняется при {('; '.join(f"загрузке {pct(c['load'],0)} и прайсе {c['price']}" for c in COND_OK) if COND_OK else 'ни одном из проверенных сочетаний')} → продавать только после пилота с такими показателями |
| Ёмкость сети | Ереван **{yerevan_lo}–{yerevan_hi}** (база {yerevan_base}), регионы {regions_lo}–{regions_base}; Армения **{yerevan_lo+regions_lo}–{yerevan_hi+regions_base}** |
| График и холдинг | {N_OWN} собств. ({'/'.join('м'+str(m) for m,k,_ in OPENINGS if k=='own')}) + {N_FR} франч. ({'–'.join(['м'+str(min(m for m,k,_ in OPENINGS if k=='fr')), 'м'+str(max(m for m,k,_ in OPENINGS if k=='fr'))])}); инвестиции холдинга {money(NET_INV)}; накопл. CF м36 **{money(NET_ROWS[-1]['cum'])}** с паушальными / **{money(NET_ROWS[-1]['cum_nofee'])}** без |

## 0. Таблица допущений

""")
def fmtv(v):
    if isinstance(v, bool): return "да" if v else "нет"
    if isinstance(v, int) or (isinstance(v, float) and v == int(v)): return fmt(v)
    if isinstance(v, float): return f"{v:.4f}".rstrip("0").rstrip(".")
    return v
md.append(T(["Параметр","Значение","Ед.","Источник","Достоверность","Комментарий"], [(k, fmtv(v), u, s, c, cm) for k, v, u, s, c, cm in ASSUMPTIONS]))

md.append(f"""
## A. Прайс-лист ArmChio (AMD)

**Позиционирование.** Дно рынка Еревана 1 000–1 500 AMD (павильонный мастер), медиана эконом-сегмента (n=27) 3 000, медиана всей выборки 4 000, True Cost 3 000 за 30 мин с мытьём, барбершопы 6 000–7 000 (файл 2). Аренда 40 м² в спальном районе 260 000 AMD/мес = {260000/2500:.0f} мужских стрижек по 2 500 (173 — по 1 500). Базовая мужская стрижка **2 500 в спальном районе (−17% к медиане эконом) и 3 000 в центре/ТЦ**. В v2 цены **не повышены** для компенсации выросших расходов: потолок позиционирования — медиана эконом 3 000, и верификация подчёркивает, что спрос на 2 000–2 500 за 15 минут без мытья не доказан (E.1); эффект ±20% цены — в разделе G. Единая цена «любая стрижка» не используется: рынок различает машинку/модельную, и сама Чио Чио в 2026 ушла от единой цены (700/800/500 ₽).

""")
md.append(T(["Услуга","A: спальный","B: центр/ТЦ","Доля (допущение)","Мин","Рынок Ереван эконом (файл 2)","Чио Чио Казань 2026, ₽ → AMD"],
 [(s[1], fmt(s[2]), fmt(s[3]), pct(s[4],0), s[5], {"mens":"1 500–3 000 (медиана 3 000)","clip":"2 000–4 000","kids":"2 000–3 000","wshort":"1 500–3 000","wlong":"2 000–4 000","beard":"1 500–3 000","style":"1 000–2 000","roots":"2 000–2 900 / центр 5 000–10 000"}[s[0]],
   (f"{chio_px[s[0]]} → {fmt(chio_px[s[0]]*RUB)}" if chio_px[s[0]] else "—")) for s in SERVICES]
 + [("**Средний чек за услугу**", f"**{money(check_A)}**", f"**{money(check_B)}**", "100%", f"{A['minutes_avg']:.1f}", "медиана эконом 3 000", "700–800 ₽ → 3 061–3 498"),
    ("Средний чек за визит (борода/укладка — доп. к визиту, 12% услуг)", money(check_visit_A), money(check_B/(1-ADDON_SHARE)), "", "", "", ""),
    ("«Эконом-микс» (машинка 35%, модельная 30%, борода 5%)", money(check_econ), "", "", "", f"−{(1-check_econ/check_A)*100:.0f}% к чеку", "")]))
md.append(f"""
**Сравнение с Чио Чио РФ (реальный прайс Казани 2026, не маркетинговые 300 ₽).**

| Метрика | Чио Чио РФ (Казань) | ArmChio A | ArmChio B |
|---|---|---|---|
| Базовая муж. стрижка | 700 ₽ = {fmt(700*RUB)} AMD | 2 500 AMD = {2500/RUB:.0f} ₽ (**−{(1-2500/(700*RUB))*100:.0f}%**) | 3 000 AMD = {3000/RUB:.0f} ₽ (−{(1-3000/(700*RUB))*100:.0f}%) |
| Средний чек | ~700 ₽ (прайс 500–800) = {fmt(700*RUB)} AMD | {money(check_A)} AMD = {rub(check_A):.0f} ₽ | {money(check_B)} AMD = {rub(check_B):.0f} ₽ |
| Средняя з/п | {fmt(A['avg_wage_ru'])} ₽ (Росстат, янв-2026) | {fmt(A['avg_wage_am'])} AMD (Armstat, авг-2026) | — |
| Паритет: стрижка / з/п | 700/{fmt(A['avg_wage_ru'])} = **{700/A['avg_wage_ru']*100:.2f}%** | **{2500/A['avg_wage_am']*100:.2f}%** | {3000/A['avg_wage_am']*100:.2f}% |
| Стрижек на среднюю з/п | {A['avg_wage_ru']/700:.0f} | {A['avg_wage_am']/2500:.0f} | {A['avg_wage_am']/3000:.0f} |

ArmChio в спальном районе дешевле реального казанского прайса на {(1-2500/(700*RUB))*100:.0f}% в валюте, но по паритету к зарплате на {(2500/A['avg_wage_am'])/(700/A['avg_wage_ru'])*100-100:.0f}% «дороже» для местного кошелька — армянский клиент беднее, ценовая власть ограничена (файл 3 §3.2).

## B. Операционная модель

| Параметр | Значение | Обоснование |
|---|---|---|
| Часы | 10:00–21:00, 7/7 | файл 1 (Казань 09–21), файл 2 (Ереван 10–20/22); 11 ч = +10% к типичным районным точкам |
| Норматив | {A['minutes_avg']:.1f} мин/услуга | файл 1: 15 мин, 4–5/час; взвешено по миксу |
| Ёмкость | 3 кресла — **{cap3:.0f} услуг/день**; 2 кресла — {cap2:.0f} | 11 × 60 / {A['minutes_avg']:.1f} × кресла |
| Персонал | 4 мастера-ИП (сменный график под 231 кресло-час/нед), 0,5 управляющего, уборка | без администратора: киоск + CRM; ТК РА 40/48 ч делает 11-часовые смены законными только для ИП (верификация C.5) |
| Зрелая загрузка (база) | **40% → {cap3*0.40:.0f} услуг/день ≈ {cap3*0.40*(1-ADDON_SHARE):.0f} визитов** | 45% = средняя зрелая точка Чио Чио (564 тыс. ₽ / ~600–700 ₽ ≈ 30–47/день) с 8-летним брендом; новый бренд — ниже |
| Gate-KPI пилота | ≥{BE['visits_day']:.0f} визитов/день к мес. 6, иначе стоп сети | = безубыточность по чистой прибыли |

**Кривая выхода** (доля от зрелой; кейсы файла 1 — Екатеринбург 15 тыс. ₽ в 1-й мес., spark 10–15 → 90–100/день за 9 мес., «первые 2 недели клиентов почти не было»): м1 28% → м3 47% → м6 70% → м12 86% → м24 95% → м36 100%. Для резерва — стресс ×0,5 в м1–3.

""")
md.append(T(["Период","М1","М2","М3","М4–6","М7–12","Год 2","Год 3"],
 [("Загрузка (база 40%)", *(pct(b['rows'][i]['load'],0) for i in (0,1,2)), f"{pct(b['rows'][3]['load'],0)}→{pct(b['rows'][5]['load'],0)}", f"{pct(b['rows'][6]['load'],0)}→{pct(b['rows'][11]['load'],0)}", f"{pct(b['rows'][12]['load'],0)}→{pct(b['rows'][23]['load'],0)}", f"{pct(b['rows'][24]['load'],0)}→{pct(b['rows'][35]['load'],0)}"),
  ("Услуг/день", *(f"{b['rows'][i]['clients_day']:.0f}" for i in (0,1,2)), f"{b['rows'][3]['clients_day']:.0f}→{b['rows'][5]['clients_day']:.0f}", f"{b['rows'][6]['clients_day']:.0f}→{b['rows'][11]['clients_day']:.0f}", f"{b['rows'][12]['clients_day']:.0f}→{b['rows'][23]['clients_day']:.0f}", f"{b['rows'][24]['clients_day']:.0f}→{b['rows'][35]['clients_day']:.0f}"),
  ("Выручка, тыс. AMD", *(f"{b['rows'][i]['revenue']/1000:.0f}" for i in (0,1,2)), f"{b['rows'][3]['revenue']/1000:.0f}→{b['rows'][5]['revenue']/1000:.0f}", f"{b['rows'][6]['revenue']/1000:.0f}→{b['rows'][11]['revenue']/1000:.0f}", f"{b['rows'][12]['revenue']/1000:.0f}→{b['rows'][23]['revenue']/1000:.0f}", f"{b['rows'][24]['revenue']/1000:.0f}→{b['rows'][35]['revenue']/1000:.0f}")]))
md.append(f"""
**Рынок сверху вниз** (радиус {A['catchment_radius_km']} км = {area_catch:.2f} км²; мужчины 15–64 ≈ 31%, дети 0–14 ≈ 20% — допущения; частота 10 и 6 раз/год):

""")
md.append(T(["Район","Население 2025","Плотн., чел/км²","Жителей в радиусе","Муж. стрижек/день в зоне","Детских/день","Нужная доля рынка муж. для {0:.0f} услуг/день".format(m36b['clients_day'])],
 [(n, fmt(p_), fmt(d), fmt(c), f"{mc:.0f}", f"{kc:.0f}", pct(need_men/mc) if mc > 0 else "—") for n, p_, d, c, mc, kc in TOPDOWN]))
md.append(f"""
Зрелая точка — {m36b['clients_day']:.0f} услуг/день, из них мужских 68% = {need_men:.0f}: в Нор-Норке это {pct(need_men/sum_men_dom[4])} мужских стрижек зоны при ≈{salons_per_capita*sum_men_dom[3]:.0f} парикмахерских в радиусе — поток ≈{need_men/sum_men_dom[4]*salons_per_capita*sum_men_dom[3]:.1f} средних районных точек (заявка Чио Чио: «60 клиентов против 30 у обычного салона»). Снизу вверх: 4 мастера × {m36b['clients_day']/4:.0f} услуг/день — в оценке файла 3 (12–18 стрижек/день на мастера). Пессимистичный сценарий (30% → {pe['m36']['clients_day']:.0f}/день) — уровень слабой точки Чио Чио.

## C. Расходы точки, налог с оборота и схема оплаты мастеров

### C.1. Три схемы оформления мастеров (мес. 36 базы, выплата 50% выручки = {money(m36b['fot'])} AMD на 4 мастеров)

Удержания с сотрудника (файл 3 §2.4, верификация A.16/17): подоходный 20%, пенсионный 5% (до 500 тыс.; выше 10% − 25 000, макс. 87 500), военный гербовый сбор по шкале 1 500 / 3 000 / 5 500 / 8 500 / 15 000 (до 100 / 200 / 500 тыс. / 1 млн / выше), ОМС 300 (200–500 тыс.) / 3 300 / 10 800; работодатель сверх gross не платит. Контроль: gross 250 000 → {money(net_employee(250000))} на руки (пример файла 3 ✓). Мастер-ИП на налоге с оборота (микро-режим парикмахерам закрыт с 2025 — файл 3 §4.1, верификация A.12): 10% − 6% собственных документированных расходов (≈5% дохода) = {pct(0.10-0.06*0.05,1)}, фиксы 5 000 + 5 000, гербовый ИП 12 000/год, ОМС ИП 129 600/год со 2-го года.

""")
md.append(T(["Схема","Брутто на мастера","На руки мастеру","На руки, % от 50%","Стоимость для салона (4)","ФОТ, % выручки","Вычет 6% у салона","EBITDA м36","Чистая прибыль м36"],
 [(SCHEMES[0][0], money(g36), money(SCHEMES[0][2]), pct(SCHEMES[0][2]/g36,0), money(m36e['fot']), pct(m36e['fot']/m36e['revenue']), "нет", money(m36e['ebitda']), money(m36e['net_profit'])),
  (SCHEMES[1][0], money(SCHEMES[1][1]), money(g36), "100%", money(m36g['fot']), pct(m36g['fot']/m36g['revenue']), "нет", money(m36g['ebitda']), money(m36g['net_profit'])),
  (f"**{SCHEMES[2][0]}**", money(g36), f"**{money(SCHEMES[2][2])}**", pct(SCHEMES[2][2]/g36,0), money(m36b['fot']), pct(m36b['fot']/m36b['revenue']), "**да**", money(m36b['ebitda']), money(m36b['net_profit']))]))
md.append(f"""
Рынок Еревана («50/50» в вакансиях — на руки при неформальной оплате) требует ≈{money(g36)} на руки с мастера при такой выручке. Схема (a) даёт на {(1-SCHEMES[0][2]/g36)*100:.0f}% меньше — мастера уйдут в соседний салон на «50 на руки» (риск текучки в мес. 1–6, когда вложено обучение). Схема (b) — честная легализация при найме сотрудников: ФОТ {pct(m36g['fot']/m36g['revenue'])} выручки, EBITDA м36 {money(m36g['ebitda'])} ({pct(m36g['ebitda']/m36g['revenue'])}) — точка {'едва прибыльна' if m36g['net_profit']>0 else 'убыточна'} по чистой прибыли ({money(m36g['net_profit'])}), окупаемость {pb(VAR['grossup'])}. Схема (c) — **база**: мастер получает {pct(SCHEMES[2][2]/g36,0)} от брутто (на {money(SCHEMES[2][2]-SCHEMES[0][2])} больше, чем сотрудник, при той же стоимости для салона), выплата ИП оформляется расчётным документом и возвращает салону вычет 6% (налог {pct(m36b['turnover_tax']/m36b['revenue'],2)} против {pct(m36e['turnover_tax']/m36e['revenue'],2)}). Гарантия на ramp-up 200 000 брутто → ≈{money(net_ip_master(200000,1))} на руки ИП (рынок 150–220 тыс.). Издержки схемы: +15 000/мес бухгалтерии на 4 ИП, регистрация ИП мастером (3 000 AMD, 1 день), медкнижки — независимо от схемы.

**Правовые риски схемы (c), честно:** (1) переквалификация договора оказания услуг в трудовой (ТК РА) — признаки: фиксированный график, подчинение правилам салона, единственный заказчик, гарантированный минимум; смягчение — ИП сам выбирает смены, может обслуживать клиентов вне салона/в других салонах, оплата за результат, минимум оформлен как аванс/ставка аренды места (практика «аренда места или процент — по выбору мастера», файл 3 §2.1); (2) ИП-мастера попадают под те же санправила (медкнижки 2 раза в год); (3) при переквалификации салон доначислит 20% подоходного и пенсионные как агент — финансовый эффект равен переходу к схеме (a): чистая прибыль м36 {money(m36e['net_profit'])} вместо {money(m36b['net_profit'])}, NPV {money(VAR['emp']['npv'])}; (4) заключение армянского юриста/бухгалтера — обязательный шаг до найма (верификация E.2–E.3).

### C.2. Налог с оборота: что вычитается

База (ст. 258 НК, закон ՀՕ-72-Ն): вычет 6% — только по расходам с расчётными документами ст. 55 ч. 2 п. 1–5 (счёт-фактура/налоговый счёт, накладная, чек ККМ). В модели документированы: расходники (чеки), коммуналка (ENA/Veolia), IT, касса, эквайринг (банк), маркетинг (счета), бухгалтерия, роялти (счёт УК), 50% прочих, **выплаты ИП-мастерам (схема c)**. Не документированы: зарплата сотрудников (управляющий, уборка; мастера при схемах a/b), **аренда у физлица** (арендатор лишь удерживает 10% подоходного как агент; параметр `rent_documented`, база = нет; при аренде у юрлица/ИП — да).

Мес. 36 базы: 10% × {money(m36b['revenue'])} = {money(0.1*m36b['revenue'])}; документированные {money(m36b['documented'])} × 6% = {money(0.06*m36b['documented'])}; налог **{money(m36b['turnover_tax'])} ({pct(m36b['turnover_tax']/m36b['revenue'],2)})** против пола 4,5% = {money(0.045*m36b['revenue'])}. Пол 4,5% срабатывает в месяцах {FLOOR_MONTHS}. Если арендодатель выдаёт документы: налог {money(VAR['rent_doc']['m36']['turnover_tax'])} ({pct(VAR['rent_doc']['m36']['turnover_tax']/VAR['rent_doc']['m36']['revenue'],2)}), NPV {money(VAR['rent_doc']['npv'])}. Если мастера — сотрудники: {money(m36e['turnover_tax'])} ({pct(m36e['turnover_tax']/m36e['revenue'],2)}). Порог режима 115 млн AMD/год (PwC, fincore; armenian-lawyer — 120 млн) — точка с выручкой ≈{m36b['revenue']*12/1e6:.0f} млн/год далеко от порога. Открытый вопрос (верификация B.1): является ли фикс 5 000 окончательным подоходным ИП при налоге с оборота; если ИП платит 20% с прибыли — чистая прибыль м36 ≈{money(m36b['net_profit']*0.8)}; альтернатива — ООО (налог с оборота заменяет налог на прибыль, дивиденды 5%). Фиксы ИП в модели: 20 000/мес + ОМС 10 800/мес с 13-го месяца.

### C.3. Структура расходов (мес. 36) — база и варианты

""")
def cost_rows(res_list, labels):
    keys = [("Выручка","revenue"),("Выплаты мастерам","fot"),("Аренда","rent"),("Управляющий 0,5 + уборка","manager"),("Расходники (95 AMD/клиент)","consumables"),("Коммуналка","utilities"),("IT/CRM/связь","it"),("Касса e-HDM","hdm"),("Эквайринг","acquiring"),("Маркетинг (5% Y1 → 3%)","marketing"),("Бухгалтерия","accounting"),("Прочее/резерв","other"),("Роялти","royalty"),("Налог с оборота","turnover_tax"),("Фиксы ИП + ОМС","ip_fixed"),("ИТОГО расходы","opex"),("EBITDA","ebitda"),("Амортизация","depreciation"),("Чистая прибыль","net_profit")]
    rows = []
    for lab, k in keys:
        rows.append([lab]+[(money(r['m36'][k]) if k == "revenue" else f"{money(r['m36'][k])} ({pct(r['m36'][k]/r['m36']['revenue'])})") for r in res_list])
    return T(["AMD/мес (м36)"]+labels, rows)
md.append(cost_rows([b, VAR["emp"], VAR["grossup"], VAR["center"], VAR["center_1line"], VAR["chairs2"], VAR["area40"], VAR["rent9500"], VAR["price28"]],
                    ["База: 30 м², ИП, 6 500/м², 40%","Сотрудники (a)","Gross-up (b)","Центр 8 500, прайс B, 45%","Центр 1-я линия 13 000, прайс B, 50%","2 кресла, 25 м²","40 м² (аренда 260 тыс.)","Аренда 9 500/м² (стресс)","Прайс A+ (муж. 2 800)"]))
md.append(f"""
Аренда спальных районов: выборка файла 3 — n=3 (5 556 / 5 808 / 9 583 AMD/м²), медиана 5 808, единственное малое помещение (24 м², пр. Гая) — 9 583; 6 500 — экспертная оценка внутри диапазона 5 800–9 600. При 9 500/м² окупаемость {pb(VAR['rent9500'])} мес. и NPV {money(VAR['rent9500']['npv'])}; при 5 800 — {pb(VAR['rent5800'])} мес., {money(VAR['rent5800']['npv'])}. Закрыть полевым обзвоном 30 объектов 25–45 м² первой линии. Владелец-оператор без управляющего: чистая прибыль м36 {money(VAR['owner']['m36']['net_profit'])}, окупаемость {pb(VAR['owner'])}. Администратор-хост 150 000 в мес. 1–6 (если киоск не принимают клиенты, 33% — только наличные): NPV {money(VAR['host']['npv'])}, окупаемость {pb(VAR['host'])}.

## D. CAPEX открытия (v2)

""")
ci = b["capex_items"]; cl = VAR["fr_lite"]["capex_items"]; cs = VAR["fr_std"]["capex_items"]; c2 = VAR["chairs2"]["capex_items"]
rows = []
for k in ci:
    k2 = k
    if k.startswith("Ремонт"): k2 = [x for x in c2 if x.startswith("Ремонт")][0]
    elif "× 3" in k: k2 = k.replace("× 3", "× 2")
    rows.append((k, money(ci[k]), f"{usd(ci[k]):,.0f}", money(cl[k]), money(cs[k]), money(c2.get(k2, 0))))
rows += [("**ИТОГО**", f"**{money(b['capex_total'])}**", f"**{usd(b['capex_total']):,.0f}**", f"**{money(VAR['fr_lite']['capex_total'])}**", f"**{money(VAR['fr_std']['capex_total'])}**", f"**{money(VAR['chairs2']['capex_total'])}**"),
         ("Итого, USD", f"{usd(b['capex_total']):,.0f}", "", f"{usd(VAR['fr_lite']['capex_total']):,.0f}", f"{usd(VAR['fr_std']['capex_total']):,.0f}", f"{usd(VAR['chairs2']['capex_total']):,.0f}"),
         ("Итого, тыс. RUB", f"{rub(b['capex_total'])/1000:,.0f}", "", f"{rub(VAR['fr_lite']['capex_total'])/1000:,.0f}", f"{rub(VAR['fr_std']['capex_total'])/1000:,.0f}", f"{rub(VAR['chairs2']['capex_total'])/1000:,.0f}")]
rows += [("Оборотный резерв на стресс ramp-up (не в CAPEX)", money(b["reserve"]), f"{usd(b['reserve']):,.0f}", money(VAR["fr_lite"]["reserve"]), money(VAR["fr_std"]["reserve"]), money(VAR["chairs2"]["reserve"])),
         ("**Потребность в финансировании**", f"**{money(b['financing'])}**", f"**{usd(b['financing']):,.0f}**", f"**{money(VAR['fr_lite']['financing'])}**", f"**{money(VAR['fr_std']['financing'])}**", f"**{money(VAR['chairs2']['financing'])}**")]
md.append(T(["Статья","Своя сеть 30 м², 3 кресла (AMD)","USD","Франшиза Чио Чио «Лайт» 290 тыс. ₽","Франшиза Чио Чио стандарт 590 тыс. ₽","Своя сеть 25 м², 2 кресла"], rows))
v2_add = sum(v for k, v in ci.items() if "[v2]" in k) - b["reserve"]
md.append(f"""
Комментарии: (1) новые строки v2 (аренда на ремонт, НДС на ввоз, кондиционеры/вентиляция, бойлер, согласование вывески, стажировка) и непредвиденное 15% дают +{money(VAR['area40']['capex_total']-V1['capex_ex'])} AMD к CAPEX 40 м² v1 без резерва ({money(V1['capex_ex'])} → {money(VAR['area40']['capex_total'])}); переход базы на 30 м² возвращает {money(VAR['area40']['capex_total']-b['capex_total'])} → CAPEX базы {money(b['capex_total'])} ≈ ${usd(b['capex_total']):,.0f} ≈ {rub(b['capex_total'])/1e6:.2f} млн ₽ (факт франчайзи Чио Чио 1,15–1,7 млн ₽); оборотный резерв по стрессу {money(b['reserve'])} (по базовой кривой {money(b['reserve_base'])}; v1 — 577 362) показан отдельно: потребность в финансировании {money(b['financing'])}; (2) тариф согласования вывески в Ереване не найден (решение Совета старейшин о местных пошлинах) — плейсхолдер 100 000; (3) паушальный Чио Чио 290 000 ₽ = {money(290000*RUB)} AMD, 590 000 ₽ = {money(590000*RUB)} AMD; (4) вакуумные аппараты — британские аналоги (Sibel Eye Vac, Head Jog), цена аппарата Чио Чио не публикуется.

## E. Три стратегических варианта

### Сравнение для инвестора (одна и та же точка, база v2)

""")
def var_row(lbl, r, r60=None):
    return (lbl, money(r["capex_total"]), f"{usd(r['capex_total']):,.0f}", money(r["m36"]["royalty"]), money(r["m36"]["net_profit"]), pct(r["m36"]["net_profit"]/r["m36"]["revenue"]), money(net36(r)), pb(r), dpb(r), money(r["npv"]), money(r["npv_tv"]), (money(r60["npv60"]) if r60 else "—"), irr_s(r["irr"]))
R_BASE60 = run(P_BASE, 60); R_LITE60 = run(VAR_P["fr_lite"], 60); R_STD60 = run(VAR_P["fr_std"], 60)
md.append(T(["Вариант","CAPEX","USD","Роялти м36","Чистая прибыль м36","Маржа","Чистая прибыль 36 мес.","Окуп.","Диск. окуп.","NPV 36 @18%","NPV 36 с TV","NPV 60 мес. @18%","IRR 36"],
 [var_row("(1a) Франчайзи Чио Чио «Лайт» (290 тыс. ₽ + 4%)", VAR["fr_lite"], R_LITE60), var_row("(1b) Франчайзи Чио Чио стандарт (590 тыс. ₽ + 4%)", VAR["fr_std"], R_STD60),
  var_row("(3) Своя собственная точка (0 + 0%)", b, R_BASE60), var_row(f"(3) Субфранчайзи ArmChio ({money(OWN_FEE)} + {pct(OWN_ROY)})", R_OWN_FR, R_OWN_FR)]))
md.append(f"""
**(1) Единичный франчайзи Чио Чио.** Армении нет в списке стран сети; зарубежные точки автономны, IT/колл-центр/эквайринг — российские (файл 1 §7). Роялти 4% ({money(0.04*m36b['revenue'])} AMD/мес) + паушальный 1,27–2,58 млн AMD дают окупаемость {pb(VAR['fr_lite'])}/{pb(VAR['fr_std'])} мес. и NPV {money(VAR['fr_lite']['npv'])} / {money(VAR['fr_std']['npv'])} при тех же операциях, что у собственной точки ({pb(b)} мес., {money(b['npv'])}). Дополнительно: налог у источника на роялти в РФ (СИДН РА–РФ, не моделируется), обязательная косметика с наценкой, подсудность в РФ, «гарантия окупаемости» с нераскрытыми условиями, права на бренд в Армении не проверены.

**(2) Мастер-франшиза Чио Чио.** Публичных условий нет (файл 1). Отраслевой ориентир (franchising.com, global-franchise.com, sprintlaw.co.uk): территориальный взнос кратный паушальному, деление роялти 50/50…75/25 в пользу мастера, график открытий, срок 10–20 лет. При 50/50 от 4%, субфранчайзи ≈ «Лайт», налог 10% на доходы мастера:

""")
md.append(T(["Сеть, точек","Доля мастера в роялти","Доля мастера в паушальных (n/4 откр./год)","Налог 10%","Расходы УК мастера","Результат мастера, AMD/мес"],
 [(m["n"], money(m["master_royalty"]), money(m["master_fees"]), money(m["tax"]), money(m["uk_cost"]), money(m["net"])) for m in MASTER]))
md.append(f"""
При ёмкости Армении {yerevan_lo+regions_lo}–{yerevan_hi+regions_base} точек мастер-франшиза выходит в плюс лишь на верхней границе ёмкости — экономически несостоятельна; имеет смысл только как переговорная опция (эксклюзив за символический взнос + мебель/обучение).

**(3) Собственный бренд ArmChio.** Формат воспроизводим (Набиуллин: «ценность — в процессах», ~45 клонов в РФ): фикс-прайс, 15 минут, без мытья (вакуумный аппарат), без администратора (киоск + alteg.io/Fresha + Idram QR), мастера на проценте, микро-формат. Экономика варианта на горизонте 3 лет — **3–4 собственные точки**; франшиза — инструмент захвата ёмкости после доказанного пилота.

**Пакет своей франшизы — сетка (субфранчайзи = базовая точка + паушальный + роялти; 36 и 60 мес.):**

""")
md.append(T(["Паушальный","Роялти","Окупаемость, мес.","Маржа м36","NPV 36 @18%","NPV 60 @18%","IRR 60","CAPEX субфранчайзи"],
 [(money(g["fee"]), pct(g["roy"]), g["payback"] or ">36", pct(g["margin"]), money(g["npv36"]), money(g["npv60"]), irr_s(g["irr60"]), money(g["capex"])) for g in GRID]))
md.append(f"""
Вывод: при базе (40%, прайс 2 500) NPV за 36 мес. отрицателен для любого пакета (как и для собственной точки: {money(b['npv'])}); за 60 мес. положителен при паушальном ≤1,2 млн и роялти ≤3%. **При каких условиях субфранчайзи получает NPV36 > 0 @18% и окупаемость ≤ 30 мес.** (пакет {money(OWN_FEE)} + {pct(OWN_ROY)}):

""")
md.append(T(["Зрелая загрузка","Прайс","Окупаемость, мес.","Маржа м36","NPV 36 @18%","NPV 60 @18%","IRR 36","Условие выполнено"],
 [(pct(c["load"],0), c["price"], c["payback"] or ">36", pct(c["margin"]), money(c["npv36"]), money(c["npv60"]), irr_s(c["irr36"]), "**да**" if (c["npv36"] > 0 and c["payback"] and c["payback"] <= 30) else "нет") for c in CONDITIONS]))
md.append(f"""
**Предложение:** пакет «founders» для первых 5–7 договоров — паушальный **{money(OWN_FEE)} AMD** (≈${usd(OWN_FEE):,.0f}), роялти **{pct(OWN_ROY)}, мин. 50 000 со 2-го месяца**, фонд 1% внутри бюджета маркетинга 3%; продажа франшизы допустима только когда пилот фактически показал загрузку ≥45% (≈{cap3*0.45:.0f} услуг/день) или подтвердил спрос по прайсу A+ — иначе продавать нечего: при 40% субфранчайзи окупается за {pb(R_OWN_FR)} мес. Это и есть gate для франчайзинга.

**Модель франчайзера (вариант 3), стационарно, все точки зрелые, {N_OWN} собственные:** налог УК — 10% с роялти и паушальных без вычета; поставки расходников — торговля (10% − 9,5% закупок).

""")
md.append(T(["Точек","Франч.","Роялти","Фонд (транзит)","Маржа поставок","Паушальные (откр./год→/мес)","EBITDA собств. точек","Расходы УК + фонд","Налог УК","УК без собств. точек, без паушальных","УК без собств. точек, с паушальными","Холдинг"],
 [(f["n"], f["fr"], money(f["royalty"]), money(f["mkt_fund"]), money(f["supply"]), money(f["fees"]), money(f["own_profit"]), money(f["uk_cost"]), money(f["uk_tax"]), money(f["uk_net_ex_own_nofee"]), money(f["uk_net_ex_own"]), money(f["total"])) for f in FRANCHISOR]))
md.append(f"""
Одна франчайзинговая точка приносит УК после налогов {money(FR_PT_INCOME)} AMD/мес; фиксированные расходы УК {money(A['uk_payroll']+A['uk_other'])} AMD/мес окупаются текущими доходами при **≈{N_BE_UK:.0f} франчайзинговых точках** — при ёмкости страны {yerevan_lo+regions_lo}–{yerevan_hi+regions_base} УК на горизонте 3 лет — **центр затрат**, живущий на паушальные и прибыль собственных точек.

## F. P&L 36 месяцев — базовый сценарий v2

Индексация цен и фиксированных расходов +4,3% с 13-го и 25-го месяца. AMD. Полная таблица — `pnl_36m_base.csv`, лист PnL_36m.

""")
md.append(T(["Мес.","Загр.","Услуг/день","Выручка","Мастера (ИП)","Аренда","Маркетинг","Налог с об.","Прочие опер.","EBITDA","Аморт.","Чистая прибыль","CF накопл."],
 [(r["month"], pct(r["load"],0), f"{r['clients_day']:.0f}", money(r["revenue"]), money(r["fot"]), money(r["rent"]), money(r["marketing"]), money(r["turnover_tax"]),
   money(r["opex"]-r["fot"]-r["rent"]-r["turnover_tax"]-r["marketing"]), money(r["ebitda"]), money(r["depreciation"]), money(r["net_profit"]), money(r["cum_cf"])) for r in b["rows"]]))
md.append("\n**Сводные годовые P&L по сценариям**\n\n")
hdr = ["Показатель"]+[f"{s} Y{y}" for s in ("Пессим.","База","Оптим.") for y in (1,2,3)]
rows = []
for lab, k in (("Выручка","revenue"),("Выплаты мастерам","fot"),("Аренда","rent"),("Маркетинг","marketing"),("Налог с оборота","turnover_tax"),("Итого расходы","opex"),("EBITDA","ebitda"),("Амортизация","depreciation"),("Чистая прибыль","net_profit")):
    rows.append([lab]+[money(R[s][f"y{y}"][k]) for s in S3 for y in (1,2,3)])
rows.append(["Услуг/день (среднее)"]+[f"{R[s][f'y{y}']['clients_day']:.0f}" for s in S3 for y in (1,2,3)])
rows.append(["EBITDA-маржа"]+[pct(R[s][f"y{y}"]["ebitda"]/R[s][f"y{y}"]["revenue"]) for s in S3 for y in (1,2,3)])
md.append(T(hdr, rows))
md.append(f"""
Сценарии: **пессимистичный** — цена −10% (чек {money(check_A*0.9)}), зрелая загрузка 30% ({pe['m36']['clients_day']:.0f}/день), аренда 8 000/м², мастера-сотрудники (схема a, ФОТ не вычитается), расходники 120/клиент; **базовый** — {money(check_A)}, 40%, 6 500, ИП; **оптимистичный** — 45% ({op['m36']['clients_day']:.0f}/день — средняя зрелая точка Чио Чио), аренда 5 800, ИП.

""")
md.append(T(["Метрика","Пессимистичный","Базовый","Оптимистичный"],
 [("CAPEX (с резервом по стрессу)", money(pe["capex_total"]), money(b["capex_total"]), money(op["capex_total"])),
  ("Резерв в составе CAPEX", money(pe["reserve"]), money(b["reserve"]), money(op["reserve"])),
  ("Безубыточность (чистая прибыль = 0): услуг/день", f"{BE_PESS['clients_day']:.0f}" if BE_PESS else "—", f"{BE['clients_day']:.0f}", f"{BE_OPT['clients_day']:.0f}"),
  ("Безубыточность: выручка/мес", money(BE_PESS["revenue"]) if BE_PESS else "—", money(BE["revenue"]), money(BE_OPT["revenue"])),
  ("EBITDA-безубыточность: услуг/день", f"{BE_PESS['clients_day_ebitda']:.0f}" if BE_PESS else "—", f"{BE['clients_day_ebitda']:.0f}", f"{BE_OPT['clients_day_ebitda']:.0f}"),
  ("Окупаемость простая, мес.", pb(pe), pb(b), pb(op)), ("Окупаемость дисконтированная @18%", dpb(pe), dpb(b), dpb(op)),
  ("NPV 36 мес. @15% / @18% / @20%", f"{money(pe['npv15'])} / {money(pe['npv'])} / {money(pe['npv20'])}", f"{money(b['npv15'])} / {money(b['npv'])} / {money(b['npv20'])}", f"{money(op['npv15'])} / {money(op['npv'])} / {money(op['npv20'])}"),
  ("NPV 36 @18% с TV 3× EBITDA Y3 (справочно)", money(pe["npv_tv"]), money(b["npv_tv"]), money(op["npv_tv"])),
  ("NPV 60 мес. @18% (справочно)", money(run(P_PESS,60)["npv60"]), money(R_BASE60["npv60"]), money(run(P_OPT,60)["npv60"])),
  ("IRR 36 мес.", irr_s(pe["irr"]), irr_s(b["irr"]), irr_s(op["irr"])), ("Накопленный CF м36", money(pe["rows"][-1]["cum_cf"]), money(b["rows"][-1]["cum_cf"]), money(op["rows"][-1]["cum_cf"]))]))
md.append(f"""
Ставка 18%: ставка ЦБ 6,75% + ориентир гособлигаций 9–10% (вне файлов) + премия за риск единичной точки/неликвидность/валюту 8–10 п.п. → 15–20%. Терминальная стоимость и возврат депозита — вне базового NPV.

Безубыточность базы при структуре м36: **{BE['clients_day']:.0f} услуг/день ≈ {BE['visits_day']:.0f} визитов ({money(BE['revenue'])} AMD/мес, {pct(BE['load'],0)} ёмкости)**; по EBITDA — {BE['clients_day_ebitda']:.0f}. При сотрудниках (схема a) — {BE_EMP['clients_day']:.0f}; центр (8 500, прайс B) — {BE_C['clients_day']:.0f}; 2 кресла/25 м² — {BE_2['clients_day']:.0f}. Чио Чио: витрина 16, расчётно 25–37 для крупных городов РФ (файл 1).

## G. Чувствительность

Чистая прибыль за 36 мес. (база {money(net36(b))}), окупаемость и NPV при изменении одного фактора:

""")
md.append(T(["Фактор","−20%","−10%","База","+10%","+20%"],
 sum([[(f"{s['factor']}: чистая прибыль 36 мес.", *(money(s[d]['net36']) for d in (-0.2,-0.1,0.0,0.1,0.2))),
       (f"{s['factor']}: окупаемость, мес.", *((s[d]['payback'] or '>36') for d in (-0.2,-0.1,0.0,0.1,0.2))),
       (f"{s['factor']}: NPV @18%", *(money(s[d]['npv']) for d in (-0.2,-0.1,0.0,0.1,0.2)))] for s in SENS], [])))
md.append("\n**Торнадо (размах чистой прибыли 36 мес. при ±20%):**\n\n")
md.append(T(["Ранг","Фактор","−20%","+20%","Размах"], [(i+1, t[0], money(t[1]), money(t[2]), money(abs(t[2]-t[1]))) for i, t in enumerate(TORNADO)]))
md.append("\n**Переключатели (структурные допущения):**\n\n")
md.append(T(["Вариант","Налог м36 (% выручки)","Чистая прибыль м36","Чистая прибыль 36 мес.","Окупаемость","NPV @18%"],
 [(lab, f"{money(r['m36']['turnover_tax'])} ({pct(r['m36']['turnover_tax']/r['m36']['revenue'],1)})", money(r["m36"]["net_profit"]), money(net36(r)), pb(r), money(r["npv"])) for lab, r in SWITCHES]))
rf = CHIO; am = m36b
md.append(f"""
Критичны: **{TORNADO[0][0]}, {TORNADO[1][0]}, {TORNADO[2][0]}**; из структурных — схема оплаты мастеров (сотрудники: NPV {money(VAR['emp']['npv'])}; gross-up: {money(VAR['grossup']['npv'])}) и аренда (9 500/м²: {money(VAR['rent9500']['npv'])}). Доля ФОТ 55% (если рынок потребует больше 50% брутто) — NPV {money(VAR['fot55']['npv'])}, окупаемость {pb(VAR['fot55'])}.

**Почему маржа ниже, чем у Чио Чио в РФ** (РФ — оценка: 50 клиентов × 700 ₽ (реальный прайс Казани), ФОТ 40%, аренда 50 тыс. ₽, роялти 4%, ПСН+взносы ≈3%):

""")
md.append(T(["Статья","Чио Чио РФ (оценка)","ArmChio база м36","Разница, п.п."],
 [("Выручка/мес", f"{fmt(rf['revenue'])} ₽ = {money(rf['revenue']*RUB)} AMD", f"{money(am['revenue'])} AMD = {rub(am['revenue'])/1000:,.0f} тыс. ₽", ""),
  ("Выплаты мастерам", pct(rf['fot']/rf['revenue']), pct(am['fot']/am['revenue']), f"{(am['fot']/am['revenue']-rf['fot']/rf['revenue'])*100:+.1f}"),
  ("Аренда", pct(rf['rent']/rf['revenue']), pct(am['rent']/am['revenue']), f"{(am['rent']/am['revenue']-rf['rent']/rf['revenue'])*100:+.1f}"),
  ("Налоги бизнеса", pct(rf['tax']/rf['revenue']), pct((am['turnover_tax']+am['ip_fixed'])/am['revenue']), f"{((am['turnover_tax']+am['ip_fixed'])/am['revenue']-rf['tax']/rf['revenue'])*100:+.1f}"),
  ("Роялти", pct(rf['royalty']/rf['revenue']), "0%", f"{-rf['royalty']/rf['revenue']*100:+.1f}"),
  ("Управляющий/уборка/бухгалтерия", "0% (владелец сам)", pct((am['manager']+am['accounting'])/am['revenue']), f"{(am['manager']+am['accounting'])/am['revenue']*100:+.1f}"),
  ("Расходники, маркетинг, коммуналка, IT, эквайринг, прочее", pct((rf['consumables']+rf['marketing']+rf['utilities']+rf['other'])/rf['revenue']), pct((am['consumables']+am['marketing']+am['utilities']+am['it']+am['hdm']+am['acquiring']+am['other'])/am['revenue']), f"{((am['consumables']+am['marketing']+am['utilities']+am['it']+am['hdm']+am['acquiring']+am['other'])/am['revenue']-(rf['consumables']+rf['marketing']+rf['utilities']+rf['other'])/rf['revenue'])*100:+.1f}"),
  ("**EBITDA-маржа**", f"**{pct(rf['ebitda']/rf['revenue'])}**", f"**{pct(am['ebitda']/am['revenue'])}**", f"**{(am['ebitda']/am['revenue']-rf['ebitda']/rf['revenue'])*100:+.1f}**")]))
md.append(f"""
Разбор: (а) выплаты мастерам 50% vs 40% (−10 п.п.); (б) аренда 260 000 AMD = ${usd(260000):,.0f} против 50 тыс. ₽ = ${usd(50000*RUB):,.0f} — на {(260000/(50000*RUB)-1)*100:.0f}% дороже при чеке на {(1-check_A/(700*RUB))*100:.0f}% ниже казанского ({money(check_A)} vs 700 ₽ = {money(700*RUB)} AMD), и де-факто в USD; (в) налог с оборота {pct(am['turnover_tax']/am['revenue'],1)} + фиксы ИП против ПСН ≈1–3% в РФ (Чио Чио строит витрину на патенте); микро-режим 0% парикмахерским в Армении закрыт; (г) управляющий и бухгалтер учтены, в модели УК Чио Чио их нет. Компенсирует: нет роялти (+4 п.п.).

## H. Ёмкость сети и график открытий (v2)

**Ереван.** Казань (родина сети): 41–42 точки в 2022 → **26 адресов в 2026** (kazan.chio-chio.ru) = 1 на ~50 тыс. жителей после отката; Ереван плотнее по конкуренции (1 парикмахерская на 720 жителей vs 1 на 1 340 в РФ). Норматив **1 точка на 60 тыс. (диапазон 50–70 тыс.)**; Кентрон +1 (дневной трафик); районы с плотностью <3 000 чел/км² −1.

""")
md.append(T(["Район","Население 01.01.2025","Плотность","1/70 тыс.","1/60 тыс. (база)","1/50 тыс. (= Казань-2026)","Приоритет"],
 [(n, fmt(p_), fmt(d), lo, ba, hi, {"Нор-Норк":"1 — пилот (9 717 чел/км², дно цен 1 000–1 500)","Малатия-Себастия":"2 — крупнейший район","Ачапняк":"3","Арабкир":"4 — релоканты","Шенгавит":"франч., у метро","Канакер-Зейтун":"франч.","Кентрон":"франч., прайс B","Эребуни":"франч., у метро","Давташен":"франч.","Аван":"франч.","Норк-Мараш":"—","Нубарашен":"—"}[n]) for n, p_, d, lo, ba, hi in NETCAP]
 + [("**Итого Ереван**", fmt(sum(x[1] for x in NETCAP)), "", f"**{yerevan_lo}**", f"**{yerevan_base}**", f"**{yerevan_hi}**", "")]))
md.append(T(["Город","Население (ориентир, вне файлов)","Парикмахерских в 2ГИС","Точек (база)","Точек (консервативно)","Комментарий"],
 [(n, fmt(p_), s, k, kl, {"Гюмри":"2-й город; цены 1 500–2 000 → региональный прайс ≈2 000","Ванадзор":"2ГИС почти не покрывает (3) — полевой обход","Абовян":"пригород Еревана","Вагаршапат":"дно 1 000 — под вопросом"}[n]) for n, p_, s, k, kl in REGIONS]
 + [("**Итого регионы**", "", "", f"**{regions_base}**", f"**{regions_lo}**", f"Армения: **{yerevan_lo+regions_lo}–{yerevan_hi+regions_base}** (база {yerevan_base+regions_base})")]))
md.append(f"""
**График открытий (36 мес.):** собственные — м1 Нор-Норк (пилот), м8 Малатия-Себастия (после gate-KPI пилота ≥{BE['visits_day']:.0f} визитов/день), м12 Ачапняк, м18 Арабкир; франчайзинговые — только с м18, после 12+ месяцев истории пилота, по 1 каждые 3 месяца: {', '.join(f'м{m} {nm}' for m, k, nm in OPENINGS if k=='fr')}. Итого к м36: **{N_OWN} собственных + {N_FR} франчайзинговых = {N_OWN+N_FR}** из ёмкости {yerevan_lo+regions_lo}–{yerevan_hi+regions_base}. УК: с м7 50% штата, с м13 полный; система франшизы (3 млн) — м12.

**Консолидированный P&L холдинга** (собственные точки полностью + роялти/поставки/паушальные от франчайзи − УК − налог УК; фонд 1% — транзит; без индексации):

""")
md.append(T(["Мес.","Собств.","Франч.","Выручка собств.","Выручка франч.","EBITDA собств.","Роялти","Маржа поставок","Паушальные","Расходы УК","Налог УК","Результат","Результат без паушальных","CAPEX","Накопл. CF","Накопл. CF без паушальных"],
 [(r["month"], r["n_own"], r["n_fr"], money(r["rev_own"]), money(r["rev_fr"]), money(r["ebitda_own"]), money(r["royalty"]), money(r["supply"]), money(r["fees"]), money(r["uk"]), money(r["uk_tax"]), money(r["net"]), money(r["net_nofee"]), money(r["capex"]), money(r["cum"]), money(r["cum_nofee"])) for r in NET_ROWS]))
yr_net = [sum(r["net"] for r in NET_ROWS[(y-1)*12:y*12]) for y in (1,2,3)]; yr_nf = [sum(r["net_nofee"] for r in NET_ROWS[(y-1)*12:y*12]) for y in (1,2,3)]
min_cum = min(r["cum"] for r in NET_ROWS); min_m = [r["month"] for r in NET_ROWS if r["cum"] == min_cum][0]
md.append(f"""
Итоги холдинга: инвестиции ({N_OWN} собственные точки по {money(b['capex_total'])} + система франшизы 3 млн) = **{money(NET_INV)} AMD ≈ ${usd(NET_INV):,.0f}**; пик потребности в финансировании {money(-min_cum)} AMD (${usd(-min_cum):,.0f}) на мес. {min_m}; результат по годам с паушальными {money(yr_net[0])} / {money(yr_net[1])} / {money(yr_net[2])}, без паушальных {money(yr_nf[0])} / {money(yr_nf[1])} / {money(yr_nf[2])}; накопленный CF на м36 **{money(NET_ROWS[-1]['cum'])}** с паушальными и **{money(NET_ROWS[-1]['cum_nofee'])}** без — сеть за 3 года **не окупается** ни в одном варианте ({N_OWN} собственные точки открыты в {'/'.join('м'+str(m) for m,k,_ in OPENINGS if k=='own')} и не все зрелые; УК с м7 стоит {money(A['uk_payroll']+A['uk_other'])}/мес при полном штате; франшизы дают доход только с м{min(m for m,k,_ in OPENINGS if k=='fr')}); это инвестиция с горизонтом 5+ лет. Инвестиции субфранчайзи (не холдинга): {N_FR} × {money(R_OWN_FR['capex_total'])} = {money(N_FR*R_OWN_FR['capex_total'])}.

## Ограничения модели и что проверить до инвестиций

1. **Спрос** — главный недоказанный параметр (готовность платить 2 000–2 500 за 15 минут без мытья, частота, walk-in): 200 уличных интервью, подсчёт очередей в 5 районных точках; gate пилота ≥{BE['visits_day']:.0f} визитов/день к мес. 6. Пессимистичный сценарий {('окупается за ' + str(pe['payback']) + ' мес.') if pe['payback'] else 'не окупается за 36 мес.'} и убыточен в год 3 ({money(pe['y3']['net_profit'])}).
2. **Налоговое заключение** (ст. 258 ч. 6, ст. 55 ч. 2, ст. 267 НК; окончательность подоходного ИП 5 000; ИП vs ООО) — до регистрации.
3. **Схема ИП-мастеров**: юридическое заключение о рисках переквалификации; 20 интервью с мастерами — что означает «50/50», готовность работать ИП.
4. **Аренда**: выборка спальных районов n=3 — обзвон 30 объектов; валюта договора AMD с капом индексации ≤ ИПЦ.
5. **Данные вне файлов** (помечены): возрастная структура, население регионов, доходность гособлигаций, alteg.io (цены не публикуются), вакуумные аппараты (британские аналоги), тариф согласования вывески (не найден), средняя з/п РФ (Росстат через вторичные источники).
6. **Не моделируется:** сезонность (летний спад до −30%), текучка мастеров (≈250–350 тыс./год на замену), конкурентная реакция (снижение цен соседями до 1 500–2 000), налог у источника на роялти в РФ для варианта 1.
""")
MD = "".join(md)
with open(os.path.join(HERE, "financial_model.md"), "w", encoding="utf-8") as f: f.write(MD)
os.makedirs(os.path.dirname(OUT_MD), exist_ok=True)
with open(OUT_MD, "w", encoding="utf-8") as f: f.write(MD)
print("MD:", OUT_MD)

print(f"Чек A/B: {check_A:.0f}/{check_B:.0f}; за визит {check_visit_A:.0f}; ёмкость {cap3:.1f}")
print(f"База м36: load {m36b['load']:.2f}, {m36b['clients_day']:.1f} услуг ({m36b['visits_day']:.1f} виз.), rev {m36b['revenue']:,.0f}, EBITDA {m36b['ebitda']:,.0f} ({m36b['ebitda']/m36b['revenue']*100:.1f}%), net {m36b['net_profit']:,.0f} ({m36b['net_profit']/m36b['revenue']*100:.1f}%), tax {m36b['turnover_tax']:,.0f} ({m36b['turnover_tax']/m36b['revenue']*100:.2f}%), fot {m36b['fot']:,.0f}")
for k in S3:
    r = R[k]; print(f"{k}: CAPEX {r['capex_total']:,.0f} (reserve {r['reserve']:,.0f}/{r['reserve_base']:,.0f}), payback {r['payback']}/{r['dpayback']}, NPV {r['npv']:,.0f} (TV {r['npv_tv']:,.0f}), IRR {r['irr']}, Y3 net {r['y3']['net_profit']:,.0f}")
print(f"BE: {BE['clients_day']:.1f} услуг ({BE['visits_day']:.1f} виз.), rev {BE['revenue']:,.0f}; EBITDA-BE {BE['clients_day_ebitda']:.1f}")
print("SCHEMES:", [(s[0][:4], round(s[1]), round(s[2]), round(s[3])) for s in SCHEMES], "emp net m36", round(m36e['net_profit']), "grossup EBITDA", round(m36g['ebitda']), "net", round(m36g['net_profit']))
print(f"Lite/Std: payback {VAR['fr_lite']['payback']}/{VAR['fr_std']['payback']}, NPV {VAR['fr_lite']['npv']:,.0f}/{VAR['fr_std']['npv']:,.0f}")
for g in GRID: print(f"  fee {g['fee']/1e6:.1f}M roy {g['roy']*100:.1f}%: pb {g['payback']}, margin {g['margin']*100:.1f}%, NPV36 {g['npv36']:,.0f}, NPV60 {g['npv60']:,.0f}")
print(f"Own FR {OWN_FEE} {OWN_ROY}: pb {R_OWN_FR['payback']}, NPV36 {R_OWN_FR['npv']:,.0f}, NPV60 {R_OWN_FR['npv60']:,.0f}; @45%: pb {R_OWN_FR45['payback']}, NPV36 {R_OWN_FR45['npv']:,.0f}, NPV60 {R_OWN_FR45['npv60']:,.0f}")
print(f"Network: Yerevan {yerevan_lo}/{yerevan_base}/{yerevan_hi}, regions {regions_lo}/{regions_base}; inv {NET_INV:,.0f}; cum36 {NET_ROWS[-1]['cum']:,.0f} / nofee {NET_ROWS[-1]['cum_nofee']:,.0f}; min cum {min_cum:,.0f} @m{min_m}; UK BE {N_BE_UK:.1f}")
print("TORNADO:", [(t[0], round(t[1]), round(t[2])) for t in TORNADO])
