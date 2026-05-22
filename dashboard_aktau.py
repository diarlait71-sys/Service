import os,json,traceback,pandas as pd,urllib.parse
from http.server import HTTPServer,BaseHTTPRequestHandler

PORT,CASH_FOLDER=8517,r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ\Актау"

class H(BaseHTTPRequestHandler):
 def do_GET(self):
  try:
   p=self.path.split('?')[0]
   if p in('/',''): self.html_dash()
   elif p=='/api/data': self.api_data()
   else: self.send_response(404);self.end_headers()
  except Exception as e: self.err(str(e))
 
 def html_dash(self):
  h="""<!DOCTYPE html><html><head><meta charset="utf-8"><title>EBITDA->Cash|AKTAU</title><style>body{font-family:Arial;margin:20px;background:#f5f5f5}.container{max-width:1200px;margin:0 auto;background:#fff;padding:20px;border-radius:5px}h1{color:#0f172a}.metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin:20px 0}.metric{padding:15px;background:#f0f4f8;border-radius:5px}.metric-value{font-size:24px;font-weight:bold;color:#2563eb}.metric-label{color:#666;font-size:12px}table{width:100%;border-collapse:collapse;margin:20px 0}th,td{padding:10px;text-align:left;border-bottom:1px solid #ddd}th{background:#2563eb;color:#fff}tr:hover{background:#f9f9f9}.error{color:red;background:#ffe6e6;padding:10px;border-radius:5px}</style></head><body><div class="container"><h1>💰 EBITDA→Cash (Актау)</h1><label>ОСВ: <select id="s" onchange="load()"><option>ОСВ 01-04 2026.xlsx<option>ОСВ 09-12 2025.xlsx</select></label><div id="c"><div class="loading">Загружаю...</div></div></div><script>function load(){fetch('/api/data?f='+encodeURIComponent(document.getElementById('s').value)).then(r=>r.json()).then(d=>{if(d.e){document.getElementById('c').innerHTML='<div class="error">❌ '+d.e+'</div>';return}let h='<div class="metrics"><div class="metric"><div class="metric-label">Кэш нач</div><div class="metric-value">'+d.cs.toFixed(1)+' млн</div></div><div class="metric"><div class="metric-label">Кэш конец</div><div class="metric-value">'+d.ce.toFixed(1)+' млн</div></div><div class="metric"><div class="metric-label">Δ</div><div class="metric-value">'+(d.ce-d.cs).toFixed(1)+' млн</div></div></div><table><tr><th>Счет<th>Описание<th>Д.Нач<th>К.Нач<th>Д.Конец<th>К.Конец<th>ΔД<th>ΔК';for(let r of d.r)h+='<tr><td>'+r.join('<td>');h+='</table>';document.getElementById('c').innerHTML=h}).catch(e=>{document.getElementById('c').innerHTML='<div class="error">'+e+'</div>'})}load()</script></body></html>"""
  self.send_response(200);self.send_header('Content-type','text/html;charset=utf-8');self.end_headers();self.wfile.write(h.encode())
 
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
   for i in range(6,len(df)):
    r=df.iloc[i]
    acc_str=str(r.iloc[0]).strip()
    if not acc_str: continue
    acc=acc_str.split(',')[0].strip()
    if not acc or len(acc)!=4 or not acc.isdigit(): continue
    od[acc]={'ndt':tn(r.iloc[2]),'nkt':tn(r.iloc[3]) if len(r)>3 else 0,'kdt':tn(r.iloc[4]),'kkt':tn(r.iloc[5]) if len(r)>5 else 0}
   for a in['1010','1030','1050']:
    if a in od: d['cs']+=(od[a]['ndt'])/1e6;d['ce']+=(od[a]['kdt'])/1e6
   l={'1010':'Касса','1030':'Расч.счета','1050':'Депозиты','1200':'Дебиторка','1210':'ДС покуп.','1274':'Займы','1300':'Запасы','1600':'ОС','1700':'Авансы','2930':'ОС разр.','3120':'Амортиз.','3150':'Резерв','3310':'Кредиторка','3350':'Начисления','3380':'Прочие','3510':'Авансы пол.'}
   for a,lbl in l.items():
    if a in od: v=od[a];d['r'].append([a,lbl,f"{v['ndt']/1e6:.1f}",f"{v['nkt']/1e6:.1f}",f"{v['kdt']/1e6:.1f}",f"{v['kkt']/1e6:.1f}",f"{(v['kdt']-v['ndt'])/1e6:.1f}",f"{(v['kkt']-v['nkt'])/1e6:.1f}"])
   d['cs']=sum((od[a].get('ndt',0)) for a in['1010','1030','1050'] if a in od)/1e6
   d['ce']=sum((od[a].get('kdt',0)) for a in['1010','1030','1050'] if a in od)/1e6
  except Exception as e: d={'e':str(e)}
  self.json(d)
 
 def json(self,d): self.send_response(200);self.send_header('Content-type','application/json;charset=utf-8');self.end_headers();self.wfile.write(json.dumps(d).encode())
 def err(self,e): self.send_response(500);self.send_header('Content-type','text/plain');self.end_headers();self.wfile.write(f"Ошибка: {e}".encode())
 def log_message(self,*a): pass

if __name__=='__main__':
 try:
  s=HTTPServer(('0.0.0.0',PORT),H)
  print(f"✅ Дашборд запущен на http://localhost:{PORT}")
  s.serve_forever()
 except KeyboardInterrupt: print("\n⏹ Остановлено")
 except Exception as e: print(f"❌ {e}")
