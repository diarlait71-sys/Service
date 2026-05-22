import os,json,traceback,pandas as pd,urllib.parse
from http.server import HTTPServer,BaseHTTPRequestHandler

PORT,CASH_FOLDER=8517,r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ\Актау"

class H(BaseHTTPRequestHandler):
 def do_GET(self):
  try:
   p=self.path.split('?')[0]
   if p in('/',''): self.html_dash()
   elif p=='/api/data': self.api_data()
   elif p=='/debug': self.debug_info()
   else: self.send_response(404);self.end_headers()
  except Exception as e: self.err(str(e)+"\n"+traceback.format_exc())
 
 def html_dash(self):
  h="""<!DOCTYPE html><html><head><meta charset="utf-8"><title>EBITDA->Cash|AKTAU</title><style>body{font-family:Arial;margin:20px;background:#f5f5f5}.container{max-width:1200px;margin:0 auto;background:#fff;padding:20px;border-radius:5px}h1{color:#0f172a}.metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin:20px 0}.metric{padding:15px;background:#f0f4f8;border-radius:5px}.metric-value{font-size:24px;font-weight:bold;color:#2563eb}.metric-label{color:#666;font-size:12px}table{width:100%;border-collapse:collapse;margin:20px 0}th,td{padding:10px;text-align:left;border-bottom:1px solid #ddd}th{background:#2563eb;color:#fff}tr:hover{background:#f9f9f9}.error{color:red;background:#ffe6e6;padding:10px;border-radius:5px}</style></head><body><div class="container"><h1>💰 EBITDA→Cash (Актау)</h1><label>ОСВ: <select id="s" onchange="load()"><option>ОСВ 01-04 2026.xlsx<option>ОСВ 09-12 2025.xlsx</select></label><div id="c"><div class="loading">Загружаю...</div></div></div><script>function load(){fetch('/api/data?f='+encodeURIComponent(document.getElementById('s').value)).then(r=>r.json()).then(d=>{if(d.e){document.getElementById('c').innerHTML='<div class="error">❌ '+d.e+'</div>';return}let h='<div class="metrics"><div class="metric"><div class="metric-label">Кэш нач</div><div class="metric-value">'+d.cs.toFixed(1)+' млн</div></div><div class="metric"><div class="metric-label">Кэш конец</div><div class="metric-value">'+d.ce.toFixed(1)+' млн</div></div><div class="metric"><div class="metric-label">Δ</div><div class="metric-value">'+(d.ce-d.cs).toFixed(1)+' млн</div></div></div><table><tr><th>Счет<th>Описание<th>Д.Нач<th>К.Нач<th>Д.Конец<th>К.Конец<th>ΔД<th>ΔК';for(let r of d.r)h+='<tr><td>'+r.join('<td>');h+='</table>';document.getElementById('c').innerHTML=h}).catch(e=>{document.getElementById('c').innerHTML='<div class="error">'+e+'</div>'})}load()</script></body></html>"""
  self.send_response(200);self.send_header('Content-type','text/html;charset=utf-8');self.end_headers();self.wfile.write(h.encode())
 
 def debug_info(self):
  try:
   path = os.path.join(CASH_FOLDER, 'ОСВ 01-04 2026.xlsx')
   df = pd.read_excel(path, header=None)
   info = f"Всего строк: {len(df)}\nВсего колонок: {len(df.columns)}\n\nПервые 25 строк:\n"
   for i in range(min(25, len(df))):
    info += f"{i}: {df.iloc[i].tolist()}\n"
   self.send_response(200)
   self.send_header('Content-type', 'text/plain; charset=utf-8')
   self.end_headers()
   self.wfile.write(info.encode('utf-8'))
  except Exception as e:
   self.err(str(e))
 
 def api_data(self):
  try:
   q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
   fn=q.get('f',['ОСВ 01-04 2026.xlsx'])[0]
   p=os.path.join(CASH_FOLDER,fn)
   d={'cs':0,'ce':0,'r':[]}
   if not os.path.exists(p): d['e']=f"Не найден: {fn}";self.json(d);return
   
   df=pd.read_excel(p,header=None)
   od={}
   
   def tn(v):
    try: s=str(v).strip();return 0.0 if s.lower() in("nan","","none","null") else float(s.replace(" ","").replace(",",".").replace("\xa0",""))
    except: return 0.0
   
   # Структура: колона 0 = "1010, описание", колона 2 = нач_дебет, колона 4 = кон_дебет, колона 5 = кон_кредит
   for i in range(6,len(df)):
    r=df.iloc[i]
    acc_str=str(r.iloc[0]).strip()
    if not acc_str or 'итого' in acc_str.lower(): continue
    parts = acc_str.split(',')
    acc = parts[0].strip()
    if not acc or len(acc)!=4 or not acc.isdigit(): continue
    
    # Парсим остатки: дебет_нач (кол 2), дебет_конец (кол 4), кредит_конец (кол 5)
    ndt = tn(r.iloc[2])  # начальное сальдо дебет
    kdt = tn(r.iloc[4])  # конечное сальдо дебет  
    kkt = tn(r.iloc[5])  # конечное сальдо кредит
    
    od[acc] = {'ndt': ndt, 'kdt': kdt, 'kkt': kkt}
   
   # Кэш = счета 1010 + 1030 + 1050 (дебет остаток)
   for a in ['1010', '1030', '1050']:
    if a in od:
     d['cs'] += od[a].get('ndt', 0) / 1e6  # начало = нач дебет
     d['ce'] += od[a].get('kdt', 0) / 1e6  # конец = кон дебет
   
   # Таблица со всеми счетами
   l={'1010':'Касса','1030':'Расч.счета','1050':'Депозиты','1200':'Дебиторка','1210':'ДС покуп.','1251':'ДС иные','1254':'Дебет','1271':'Дебет','1310':'Товары','1330':'Запасы','1421':'Авансы','1422':'Авансы','1424':'Авансы','1430':'Авансы','1710':'Авансы поставщикам','1720':'Авансы','2410':'ОС','2930':'ОС разработка','3120':'Амортизация','3150':'Резерв','3310':'Кредиторка','3350':'Начисления','3380':'Прочие','3510':'Авансы получены'}
   for a,lbl in l.items():
    if a in od: 
     v=od[a]
     d['r'].append([a, lbl, f"{v['ndt']/1e6:.1f}", "0.0", f"{v['kdt']/1e6:.1f}", f"{v['kkt']/1e6:.1f}", f"{(v['kdt']-v['ndt'])/1e6:.1f}", f"{-v['kkt']/1e6:.1f}"])
  except Exception as e: 
   d={'e':str(e)+'\n'+traceback.format_exc()}
  self.json(d)
 
 def json(self,d): self.send_response(200);self.send_header('Content-type','application/json;charset=utf-8');self.end_headers();self.wfile.write(json.dumps(d).encode())
 def err(self,e): self.send_response(500);self.send_header('Content-type','text/plain');self.end_headers();self.wfile.write(f"Ошибка: {e}".encode())
 def log_message(self,*a): pass

if __name__=='__main__':
 try:
  s=HTTPServer(('0.0.0.0',PORT),H)
  print(f"✅ Дашборд http://localhost:{PORT}")
  print(f"🔧 Debug: http://localhost:{PORT}/debug")
  s.serve_forever()
 except KeyboardInterrupt: print("\n⏹ Остановлено")
 except Exception as e: print(f"❌ {e}")
