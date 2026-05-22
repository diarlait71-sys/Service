import os, json, traceback, urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import pandas as pd

PORT = 8517
FOLDER = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ\Актау"

# ──────────────────────────────────────────────────────────────────────────────
# Вспомогательные функции
# ──────────────────────────────────────────────────────────────────────────────
def to_num(v):
    try:
        s = str(v).strip().replace('\xa0', '').replace(' ', '').replace(',', '.')
        if s in ('-', '', 'nan', 'None', 'none'): return 0.0
        return float(s)
    except:
        return 0.0

# ──────────────────────────────────────────────────────────────────────────────
# Парсинг ОСВ
# Структура столбцов в 1С-ОСВ:
#   col[0] = Счет, Наименование
#   col[2] = Сальдо начало  Дт
#   col[3] = Сальдо начало  Кт
#   col[4] = Оборот         Дт
#   col[5] = Оборот         Кт
#   col[6] = Сальдо конец   Дт
#   col[7] = Сальдо конец   Кт  (может отсутствовать = 0)
# ──────────────────────────────────────────────────────────────────────────────
def parse_osv(filename):
    fp = os.path.join(FOLDER, filename)
    if not os.path.exists(fp):
        return None, f"Файл не найден: {filename}"
    df = pd.read_excel(fp, header=None)
    acc = {}
    for i in range(len(df)):
        r = df.iloc[i]
        cell = str(r.iloc[0]).strip()
        if not cell or 'итого' in cell.lower(): continue
        parts = cell.split(',')
        code = parts[0].strip()
        if len(code) != 4 or not code.isdigit(): continue
        desc = parts[1].strip() if len(parts) > 1 else code
        acc[code] = {
            'desc': desc,
            'ndt':  to_num(r.iloc[2]) if len(r) > 2 else 0,   # начало Дт
            'nkt':  to_num(r.iloc[3]) if len(r) > 3 else 0,   # начало Кт
            'odt':  to_num(r.iloc[4]) if len(r) > 4 else 0,   # оборот Дт
            'okt':  to_num(r.iloc[5]) if len(r) > 5 else 0,   # оборот Кт
            'kdt':  to_num(r.iloc[6]) if len(r) > 6 else 0,   # конец  Дт
            'kkt':  to_num(r.iloc[7]) if len(r) > 7 else 0,   # конец  Кт
        }
    return acc, None

# ──────────────────────────────────────────────────────────────────────────────
# Парсинг ОПИУ
# ──────────────────────────────────────────────────────────────────────────────
def parse_opiu(filename='Опиу.xlsx'):
    fp = os.path.join(FOLDER, filename)
    if not os.path.exists(fp): return {}, None
    df = pd.read_excel(fp, header=None)
    result = {}
    label_map = {
        'доход от реализации':           'revenue',
        'прочие доходы':                 'other_income',
        'итого доходов':                 'total_income',
        'расходы по реализованным':      'cogs',
        'расходы, связанные с выплатой': 'interest_expense',
        'амортизационные отчисления':    'depreciation',
        'расходы на финансирование':     'fin_costs',
        'прочие расходы':                'other_expense',
        'итого расходов':                'total_expense',
        'прибыль (убыток) до':          'ebt',
        'расходы на налоги':             'tax',
        'итого чистая прибыль':         'net_income',
    }
    for i in range(len(df)):
        r = df.iloc[i]
        cell = str(r.iloc[0]).strip().lower()
        for kw, key in label_map.items():
            if cell.startswith(kw):
                result[key] = to_num(r.iloc[22]) if len(r) > 22 else 0
                break
    return result, None

# ──────────────────────────────────────────────────────────────────────────────
# Парсинг ДДС (лист 1 + лист 2)
# ──────────────────────────────────────────────────────────────────────────────
def parse_dds():
    result = {}
    fp1 = os.path.join(FOLDER, 'ДДС 1 лист.xlsx')
    if os.path.exists(fp1):
        df1 = pd.read_excel(fp1, header=None)
        dds1_map = {
            '1. поступление денежных средств от оп': ('oper_in',    22),
            '2. выбытие денежных средств, в':        ('oper_out',   22),
            '3. чистая сумма денежных средств от оп':('oper_net',   22),
            '1. поступление денежных средств от ин': ('inv_in',     22),
            '2. выбытие денежных средств от ин':     ('inv_out',    22),
            '3. чистая сумма денежных средств от ин':('inv_net',    22),
            'предоставление услуг':                  ('serv_in',    22),
            'авансы полученные':                     ('adv_in',     22),
            'прочие поступления':                    ('other_in',   22),
            'платежи поставщикам за товары':         ('sup_out',    22),
            'авансы выданные':                       ('adv_out',    22),
            'выплаты по заработной плате':           ('salary_out', 22),
            'другие платежи в бюджет':               ('tax_out',    22),
            'прочие выплаты':                        ('misc_out',   22),
            'реализация основных средств':           ('asset_sale', 22),
            'приобретение основных средств':         ('capex',      22),
        }
        for i in range(len(df1)):
            r = df1.iloc[i]
            cell = str(r.iloc[0]).strip().lower()
            for kw, (key, col) in dds1_map.items():
                if cell.startswith(kw) and key not in result:
                    result[key] = to_num(r.iloc[col]) if len(r) > col else 0
                    break
    fp2 = os.path.join(FOLDER, 'ДДс 2 лист.xlsx')
    if os.path.exists(fp2):
        df2 = pd.read_excel(fp2, header=None)
        dds2_map = {
            '1. поступление денежных средств':          ('fin_in',      22),
            '2. выбытие денежных средств, всего':       ('fin_out',     22),
            '3. чистая сумма денежных средств от фи':   ('fin_net',     22),
            'итого:':                                   ('total_delta', 22),
            'денежные средства и их эквиваленты на на': ('cash_start',  22),
            'денежные средства и их эквиваленты на ко': ('cash_end',    22),
            'получение займов':                         ('loans_in',    22),
            'погашение займов':                         ('loans_out',   22),
        }
        for i in range(len(df2)):
            r = df2.iloc[i]
            cell = str(r.iloc[0]).strip().lower()
            for kw, (key, col) in dds2_map.items():
                if cell.startswith(kw) and key not in result:
                    result[key] = to_num(r.iloc[col]) if len(r) > col else 0
                    break
    return result

# ──────────────────────────────────────────────────────────────────────────────
# Изменения рабочего капитала по ОСВ
# ──────────────────────────────────────────────────────────────────────────────
def calc_working_capital(osv):
    changes = []
    active_accounts = {
        '1210': 'Дебиторка покупателей',
        '1251': 'Авансы подотчетным',
        '1310': 'Сырьё и материалы',
        '1330': 'Товары (запасы)',
        '1710': 'Авансы поставщикам',
        '1720': 'Расходы буд. периодов',
    }
    passive_accounts = {
        '3131': 'НДС начисленный',
        '3310': 'КЗ поставщикам',
        '3350': 'Задолженность по зарплате',
        '3381': 'КЗ: возвраты',
        '3385': 'КЗ: исполн. листы',
        '3510': 'Авансы полученные',
    }
    for code, label in active_accounts.items():
        if code not in osv: continue
        v = osv[code]
        nac = v['ndt']
        kon = v['ndt'] + v['odt'] - v['okt']
        delta = kon - nac
        if abs(delta) > 50_000:
            changes.append({
                'code': code, 'label': label,
                'nac': nac / 1e6, 'kon': kon / 1e6,
                'delta': delta / 1e6,
                'effect': 'absorber' if delta > 0 else 'source',
                'type': 'active',
            })
    for code, label in passive_accounts.items():
        if code not in osv: continue
        v = osv[code]
        nac = v['nkt']
        kon = v['nkt'] + v['okt'] - v['odt']
        delta = kon - nac
        if abs(delta) > 50_000:
            changes.append({
                'code': code, 'label': label,
                'nac': nac / 1e6, 'kon': kon / 1e6,
                'delta': delta / 1e6,
                'effect': 'source' if delta > 0 else 'absorber',
                'type': 'passive',
            })
    return changes


# ──────────────────────────────────────────────────────────────────────────────
# HTML страница
# ──────────────────────────────────────────────────────────────────────────────
HTML_PAGE = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<title>EBITDA\u2192Cash | AKTAU</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#1a1a2e}
.header{background:linear-gradient(135deg,#1a3a52,#2563eb);color:#fff;padding:22px 32px}
.header h1{font-size:24px;font-weight:700}.header p{font-size:13px;opacity:.8;margin-top:4px}
.ctrl{padding:10px 32px;background:#fff;border-bottom:1px solid #e0e0e0;display:flex;align-items:center;gap:12px}
.ctrl select{padding:7px 12px;border:1px solid #ccc;border-radius:6px;font-size:14px}
.ctrl button{padding:7px 16px;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer}
#status{color:#888;font-size:13px}
.wrap{max-width:1400px;margin:0 auto;padding:20px 32px}
.section{background:#fff;border-radius:10px;padding:20px 24px;margin-bottom:18px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.section h2{font-size:16px;font-weight:700;color:#1a3a52;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #2563eb}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px}
.metric{border-radius:8px;padding:14px;text-align:center}
.metric.blue{background:linear-gradient(135deg,#2563eb,#1d4ed8);color:#fff}
.metric.green{background:linear-gradient(135deg,#16a34a,#15803d);color:#fff}
.metric.red{background:linear-gradient(135deg,#dc2626,#b91c1c);color:#fff}
.metric.gray{background:linear-gradient(135deg,#6b7280,#4b5563);color:#fff}
.mv{font-size:24px;font-weight:700;margin:6px 0}.ml{font-size:11px;opacity:.85}
.waterfall{display:flex;flex-direction:column;gap:5px;margin-top:10px}
.wf-row{display:flex;align-items:center;gap:10px;font-size:13px}
.wf-label{width:230px;flex-shrink:0;color:#444}
.wf-bar-wrap{flex:1;height:26px;background:#f5f5f5;border-radius:4px;overflow:hidden}
.wf-bar{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:6px;font-weight:700;font-size:12px;color:#fff;transition:width .4s}
.wf-bar.pos{background:#16a34a}.wf-bar.neg{background:#dc2626}.wf-bar.blue{background:#2563eb}.wf-bar.dark{background:#1a3a52}
.wf-val{width:80px;text-align:right;font-weight:700;font-size:13px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:900px){.grid2{grid-template-columns:1fr}}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#1a3a52;color:#fff;padding:9px 11px;text-align:left;font-weight:600}
td{padding:8px 11px;border-bottom:1px solid #f0f0f0}
tr:hover td{background:#f8faff}
.pos{color:#16a34a;font-weight:700}.neg{color:#dc2626;font-weight:700}.neu{color:#6b7280}
.absorber-row td:first-child{border-left:4px solid #dc2626}
.source-row td:first-child{border-left:4px solid #16a34a}
.tag{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700}
.tag.up{background:#dcfce7;color:#16a34a}.tag.down{background:#fee2e2;color:#dc2626}
.sep{height:1px;background:#e0e0e0;margin:4px 0}
</style>
</head>
<body>
<div class="header">
  <h1>\U0001f4b0 EBITDA \u2192 Cash | AKTAU</h1>
  <p>Полный финансовый анализ: P&L, движение кэша, рабочий капитал</p>
</div>
<div class="ctrl">
  <label>\u041e\u0421\u0412:
    <select id="osv">
      <option value="\u041e\u0421\u0412 01-04 2026.xlsx">\u041e\u0421\u0412 \u042f\u043d\u0432\u2013\u0410\u043f\u0440 2026</option>
      <option value="\u041e\u0421\u0412 09-12 2025.xlsx">\u041e\u0421\u0412 \u0421\u0435\u043d\u2013\u0414\u0435\u043a 2025</option>
    </select>
  </label>
  <button onclick="load()">&#8635; Обновить</button>
  <span id="status"></span>
</div>
<div class="wrap" id="content"><p style="text-align:center;margin-top:60px;color:#999">Загрузка...</p></div>

<script>
function fs(v,d=1){return(v>=0?'+':'')+v.toFixed(d)+' млн'}
function f(v,d=1){return v.toFixed(d)}
function cls(v){return v>0.05?'pos':v<-0.05?'neg':'neu'}

function load(){
  const osv=document.getElementById('osv').value;
  document.getElementById('status').textContent='Загружаю...';
  fetch('/api?osv='+encodeURIComponent(osv))
    .then(r=>r.json()).then(d=>{
      document.getElementById('status').textContent='Готово';
      if(d.error){document.getElementById('content').innerHTML='<div class="section"><pre style="color:red;white-space:pre-wrap">'+d.error+'</pre></div>';return;}
      render(d);
    }).catch(e=>document.getElementById('content').innerHTML='<p style="color:red">'+e+'</p>');
}

function wfRow(label, val, barClass, maxAbs){
  const pct=Math.min(Math.abs(val)/maxAbs*100,100);
  const inner=Math.abs(val)>=1?Math.abs(val).toFixed(1):'';
  return `<div class="wf-row">
    <div class="wf-label">${label}</div>
    <div class="wf-bar-wrap"><div class="wf-bar ${barClass}" style="width:${pct}%">${inner}</div></div>
    <div class="wf-val ${cls(val)}">${fs(val)}</div>
  </div>`;
}

function render(d){
  let h='';

  // 1. Топ метрики
  const gp_pct=d.revenue>0?(d.gross_profit/d.revenue*100).toFixed(0):0;
  const ebitda_pct=d.revenue>0?(d.ebitda/d.revenue*100).toFixed(0):0;
  h+=`<div class="section"><h2>📊 Ключевые показатели</h2><div class="metrics">
    <div class="metric blue"><div class="ml">EBITDA</div><div class="mv">${f(d.ebitda)}</div><div class="ml">млн ₸ (${ebitda_pct}%)</div></div>
    <div class="metric ${d.net_income>=0?'green':'red'}"><div class="ml">Чистая прибыль</div><div class="mv">${f(d.net_income)}</div><div class="ml">млн ₸</div></div>
    <div class="metric blue"><div class="ml">Выручка</div><div class="mv">${f(d.revenue)}</div><div class="ml">млн ₸</div></div>
    <div class="metric ${d.gross_profit>=0?'green':'red'}"><div class="ml">Валовая прибыль</div><div class="mv">${f(d.gross_profit)}</div><div class="ml">млн ₸ (${gp_pct}%)</div></div>
    <div class="metric ${d.oper_net>=0?'green':'red'}"><div class="ml">Опер. кэш-поток</div><div class="mv">${f(d.oper_net)}</div><div class="ml">млн ₸</div></div>
    <div class="metric ${d.cash_delta>=0?'green':'gray'}"><div class="ml">Δ Кэш за период</div><div class="mv">${f(d.cash_delta,2)}</div><div class="ml">млн ₸</div></div>
  </div></div>`;

  // 2. EBITDA → Cash мост (водопад)
  const bridge=[
    ['EBITDA',          d.ebitda,         'blue'],
    ['  + Амортизация', d.depreciation,   d.depreciation>=0?'pos':'neg'],
    ['  + Проценты (по займам)', d.interest, d.interest>=0?'pos':'neg'],
    ['  + Изм. раб. капитала', d.wc_net,  d.wc_net>=0?'pos':'neg'],
    ['= Опер. кэш-поток', d.oper_net,     d.oper_net>=0?'pos':'neg'],
    ['  - CAPEX',       -d.capex,          'neg'],
    ['  + Продажа ОС',  d.asset_sale,      d.asset_sale>0?'pos':'neg'],
    ['= Инвест. кэш',   d.inv_net,         d.inv_net>=0?'pos':'neg'],
    ['  + Получ. займы', d.loans_in,       d.loans_in>0?'pos':'neg'],
    ['  - Погаш. займы', -d.loans_out,     'neg'],
    ['= Фин. кэш',      d.fin_net,         d.fin_net>=0?'pos':'neg'],
    ['= ИТОГО Δ Кэш',   d.cash_delta,      d.cash_delta>=0?'dark':'neg'],
  ];
  const maxAbs=Math.max(...bridge.map(x=>Math.abs(x[1])),1);
  h+=`<div class="section"><h2>🌉 EBITDA → Cash: Мост</h2><div class="waterfall">`;
  for(const [lbl,val,bc] of bridge) h+=wfRow(lbl,val,bc,maxAbs);
  h+=`</div></div>`;

  // 3. P&L + ДДС в сетке
  h+=`<div class="grid2">`;

  // P&L
  const pct=v=>d.revenue>0?'('+f(v/d.revenue*100,0)+'%)':'';
  h+=`<div class="section"><h2>📈 Отчёт о прибылях и убытках</h2><table>
    <tr><th>Показатель</th><th>Млн ₸</th><th>% выр.</th></tr>
    <tr><td>Выручка от реализации</td><td class="pos">${f(d.revenue)}</td><td>100%</td></tr>
    <tr><td>— Себестоимость (COGS)</td><td class="neg">-${f(d.cogs)}</td><td>${pct(d.cogs)}</td></tr>
    <tr style="background:#f0f7ff"><td><b>Валовая прибыль</b></td><td class="${cls(d.gross_profit)}"><b>${f(d.gross_profit)}</b></td><td><b>${pct(d.gross_profit)}</b></td></tr>
    <tr><td>+ Прочие доходы</td><td class="pos">+${f(d.other_income)}</td><td>—</td></tr>
    <tr><td>— Прочие расходы</td><td class="neg">-${f(d.other_exp)}</td><td>—</td></tr>
    <tr><td>— Амортизация</td><td class="neg">-${f(d.depreciation)}</td><td>—</td></tr>
    <tr><td>— Расходы по вознаграждениям</td><td class="neg">-${f(d.interest)}</td><td>—</td></tr>
    <tr style="background:#e8f5ff"><td><b>EBITDA</b></td><td class="${cls(d.ebitda)}"><b>${f(d.ebitda)}</b></td><td><b>${pct(d.ebitda)}</b></td></tr>
    <tr style="background:#e8f5e9"><td><b>Чистая прибыль</b></td><td class="${cls(d.net_income)}"><b>${f(d.net_income)}</b></td><td><b>${pct(d.net_income)}</b></td></tr>
  </table></div>`;

  // ДДС
  const oper_in=d.oper_net+d.salary_out+d.sup_out+d.tax_out;
  h+=`<div class="section"><h2>💳 Движение денежных средств (ДДС)</h2><table>
    <tr><th>Показатель</th><th>Млн ₸</th></tr>
    <tr style="background:#f0f7ff"><td><b>Кэш начало периода</b></td><td><b>${f(d.cash_start,2)}</b></td></tr>
    <tr><td colspan="2" style="background:#e8f0fb;font-weight:700;color:#2563eb;padding:6px 11px">I. Операционная деятельность</td></tr>
    <tr><td style="padding-left:20px">Поступления от клиентов</td><td class="pos">+${f(oper_in)}</td></tr>
    <tr><td style="padding-left:20px">— Оплата поставщикам</td><td class="neg">-${f(d.sup_out)}</td></tr>
    <tr><td style="padding-left:20px">— Зарплата</td><td class="neg">-${f(d.salary_out)}</td></tr>
    <tr><td style="padding-left:20px">— Налоги в бюджет</td><td class="neg">-${f(d.tax_out)}</td></tr>
    <tr style="background:#e8f5e9"><td><b>Чистый опер. кэш</b></td><td class="${cls(d.oper_net)}"><b>${fs(d.oper_net)}</b></td></tr>
    <tr><td colspan="2" style="background:#e8f0fb;font-weight:700;color:#2563eb;padding:6px 11px">II. Инвестиционная деятельность</td></tr>
    <tr><td style="padding-left:20px">— CAPEX (покупка ОС)</td><td class="neg">-${f(d.capex)}</td></tr>
    <tr><td style="padding-left:20px">+ Продажа ОС</td><td class="pos">+${f(d.asset_sale)}</td></tr>
    <tr style="background:#e8f5e9"><td><b>Чистый инвест. кэш</b></td><td class="${cls(d.inv_net)}"><b>${fs(d.inv_net)}</b></td></tr>
    <tr><td colspan="2" style="background:#e8f0fb;font-weight:700;color:#2563eb;padding:6px 11px">III. Финансовая деятельность</td></tr>
    <tr><td style="padding-left:20px">+ Получение займов</td><td class="pos">+${f(d.loans_in)}</td></tr>
    <tr><td style="padding-left:20px">— Погашение займов</td><td class="neg">-${f(d.loans_out)}</td></tr>
    <tr style="background:#e8f5e9"><td><b>Чистый фин. кэш</b></td><td class="${cls(d.fin_net)}"><b>${fs(d.fin_net)}</b></td></tr>
    <tr style="background:#f0f7ff"><td><b>Δ Денежные средства</b></td><td class="${cls(d.cash_delta)}"><b>${fs(d.cash_delta,2)}</b></td></tr>
    <tr style="background:#f0f7ff"><td><b>Кэш конец периода</b></td><td><b>${f(d.cash_end,2)}</b></td></tr>
  </table></div>`;

  h+=`</div>`;// grid2

  // 4. Рабочий капитал
  if(d.wc_changes&&d.wc_changes.length>0){
    const sorted=[...d.wc_changes].sort((a,b)=>Math.abs(b.delta)-Math.abs(a.delta));
    h+=`<div class="section"><h2>🔄 Изменения рабочего капитала (по ОСВ)</h2>
    <table><tr><th>Счёт</th><th>Статья</th><th>Начало млн</th><th>Конец млн</th><th>Δ млн</th><th>Эффект на кэш</th></tr>`;
    for(const c of sorted){
      const eff=c.effect==='absorber'?'absorber-row':'source-row';
      const tag=c.effect==='absorber'
        ?'<span class="tag down">\u2193 Поглощение</span>'
        :'<span class="tag up">\u2191 Источник</span>';
      h+=`<tr class="${eff}"><td>${c.code}</td><td>${c.label}</td><td>${f(c.nac)}</td><td>${f(c.kon)}</td><td class="${cls(c.delta)}">${fs(c.delta)}</td><td>${tag}</td></tr>`;
    }
    h+=`</table></div>`;
  }

  // 5. ОСВ таблица
  if(d.osv_table&&d.osv_table.length>0){
    h+=`<div class="section"><h2>📋 Ключевые счета ОСВ</h2>
    <table><tr><th>Счёт</th><th>Наименование</th><th>Начало млн</th><th>Конец млн</th><th>Δ млн</th></tr>`;
    for(const r of d.osv_table)
      h+=`<tr><td>${r.code}</td><td>${r.label}</td><td>${f(r.nac)}</td><td>${f(r.kon)}</td><td class="${cls(r.delta)}">${fs(r.delta)}</td></tr>`;
    h+=`</table></div>`;
  }

  document.getElementById('content').innerHTML=h;
}
load();
</script>
</body>
</html>"""


# ──────────────────────────────────────────────────────────────────────────────
# HTTP сервер
# ──────────────────────────────────────────────────────────────────────────────
class H(BaseHTTPRequestHandler):
 def do_GET(self):
  try:
   p=self.path.split('?')[0]
   if p in('/',''): self.html()
   elif p=='/api': self.api()
   else: self.send_response(404);self.end_headers()
  except Exception as e: self.err(str(e)+"\n"+traceback.format_exc())
 
 def html(self):
  h = HTML_PAGE
  self.send_response(200);self.send_header('Content-type','text/html;charset=utf-8');self.end_headers();self.wfile.write(h.encode('utf-8'))
 
 def api(self):
  try:
   qs=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
   osv_file=qs.get('osv',['ОСВ 01-04 2026.xlsx'])[0]
   osv, err = parse_osv(osv_file)
   if err: self.send_json({'error': err}); return
   opiu, _ = parse_opiu()
   dds = parse_dds()
   # EBITDA из ОПИУ
   net_income   = opiu.get('net_income',   0) / 1e6
   depreciation = opiu.get('depreciation', 0) / 1e6
   interest     = opiu.get('interest_expense', 0) / 1e6
   tax          = opiu.get('tax', 0) / 1e6
   ebitda       = net_income + depreciation + interest + tax
   revenue      = opiu.get('revenue', 0) / 1e6
   cogs         = opiu.get('cogs', 0) / 1e6
   other_income = opiu.get('other_income', 0) / 1e6
   other_exp    = opiu.get('other_expense', 0) / 1e6
   gross_profit = revenue - cogs
   # ДДС
   oper_net  = dds.get('oper_net', 0) / 1e6
   inv_net   = dds.get('inv_net',  0) / 1e6
   fin_net   = dds.get('fin_net',  0) / 1e6
   cash_start= dds.get('cash_start', 0) / 1e6
   cash_end  = dds.get('cash_end',   0) / 1e6
   cash_delta= cash_end - cash_start
   capex      = dds.get('capex', 0) / 1e6
   asset_sale = dds.get('asset_sale', 0) / 1e6
   loans_in   = dds.get('loans_in', 0) / 1e6
   loans_out  = dds.get('loans_out', 0) / 1e6
   salary_out = dds.get('salary_out', 0) / 1e6
   sup_out    = dds.get('sup_out', 0) / 1e6
   tax_out    = dds.get('tax_out', 0) / 1e6
   # Рабочий капитал
   wc_changes = calc_working_capital(osv)
   wc_net = sum(-c['delta'] if c['effect']=='absorber' else c['delta'] for c in wc_changes)
   # ОСВ таблица
   display_accounts = [
    ('1010','Касса'),('1030','Расчётный счёт'),('1210','Дебиторка покупателей'),
    ('1310','Сырьё / материалы'),('1330','Товары (запасы)'),('1710','Авансы поставщикам'),
    ('2410','Основные средства'),('2420','Накопл. амортизация'),('3010','Кредиты банков'),
    ('3310','КЗ поставщикам'),('3350','Задолженность по з/п'),('3510','Авансы полученные'),
   ]
   osv_table = []
   for code, label in display_accounts:
    if code not in osv: continue
    v = osv[code]
    nac = (v['ndt'] - v['nkt']) / 1e6
    kon_dt = v['ndt'] + v['odt'] - v['okt']
    kon_kt = v['nkt'] + v['okt'] - v['odt']
    kon = (kon_dt - kon_kt) / 1e6
    osv_table.append({'code':code,'label':label,'nac':round(nac,2),'kon':round(kon,2),'delta':round(kon-nac,2)})
   result = {
    'revenue':round(revenue,2),'cogs':round(cogs,2),'gross_profit':round(gross_profit,2),
    'other_income':round(other_income,2),'other_exp':round(other_exp,2),
    'depreciation':round(depreciation,2),'interest':round(interest,2),
    'net_income':round(net_income,2),'ebitda':round(ebitda,2),
    'oper_net':round(oper_net,2),'inv_net':round(inv_net,2),'fin_net':round(fin_net,2),
    'cash_start':round(cash_start,2),'cash_end':round(cash_end,2),'cash_delta':round(cash_delta,2),
    'capex':round(capex,2),'asset_sale':round(asset_sale,2),
    'loans_in':round(loans_in,2),'loans_out':round(loans_out,2),
    'salary_out':round(salary_out,2),'sup_out':round(sup_out,2),'tax_out':round(tax_out,2),
    'wc_changes':wc_changes,'wc_net':round(wc_net,2),
    'bridge_gap':round(ebitda-oper_net,2),'osv_table':osv_table,
   }
   self.send_json(result)
  except Exception as e: self.send_json({'error':str(e)+'\n'+traceback.format_exc()})

 def send_json(self,d):
  payload=json.dumps(d,ensure_ascii=False).encode('utf-8')
  self.send_response(200);self.send_header('Content-type','application/json;charset=utf-8');self.end_headers();self.wfile.write(payload)
 def err(self,e): self.send_response(500);self.send_header('Content-type','text/plain;charset=utf-8');self.end_headers();self.wfile.write(f"Ошибка: {e}".encode('utf-8'))
 def log_message(self,*a): pass

if __name__=='__main__':
 import sys
 sys.stdout.reconfigure(encoding='utf-8', errors='replace')
 sys.stderr.reconfigure(encoding='utf-8', errors='replace')
 print(f"EBITDA->Cash | AKTAU | port {PORT}")
 print(f"http://localhost:{PORT}")
 try:
  s=HTTPServer(('0.0.0.0',PORT),H)
  s.serve_forever()
 except KeyboardInterrupt: print("Stopped")
 except Exception as e: print(f"Error: {e}")
