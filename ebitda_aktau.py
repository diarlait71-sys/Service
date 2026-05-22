import os
import streamlit as st
import pandas as pd

st.set_page_config(page_title="EBITDA -> Cash | AKTAU", page_icon="💰", layout="wide")
st.title("💰 EBITDA → Cash (Актау)")
st.markdown("**Куда потратили EBITDA**")

CASH_FOLDER = r"C:\Users\D.Muldabaev\Desktop\Приложения для сервиса\финанализ\Актау"

def to_num(v):
    try:
        s = str(v).strip()
        if s.lower() in ("nan", "", "none", "null"):
            return 0.0
        return float(s.replace(" ", "").replace(",", ".").replace("\xa0", ""))
    except:
        return 0.0

osv_files = ['ОСВ 01-04 2026.xlsx', 'ОСВ 09-12 2025.xlsx']
selected_osv = st.selectbox("Выбери ОСВ:", osv_files)

if selected_osv:
    path = os.path.join(CASH_FOLDER, selected_osv)
    if not os.path.exists(path):
        st.error(f"Файл не найден: {path}")
    else:
        try:
            df_osv = pd.read_excel(path, header=None)
            st.success(f"✅ Загружен: {selected_osv}")
            
            osv_data = {}
            for _, row in df_osv.iterrows():
                if len(row) < 9:
                    continue
                acc = str(row.iloc[0]).strip()
                if not acc or len(acc) != 4 or not acc.isdigit():
                    continue
                
                nac_dt = to_num(row.iloc[2])
                nac_kt = to_num(row.iloc[3])
                kon_dt = to_num(row.iloc[6])
                kon_kt = to_num(row.iloc[8])
                
                osv_data[acc] = {
                    'nac_dt': nac_dt,
                    'nac_kt': nac_kt,
                    'kon_dt': kon_dt,
                    'kon_kt': kon_kt,
                    'delta_dt': kon_dt - nac_dt,
                    'delta_kt': kon_kt - nac_kt,
                }
            
            st.info(f"Найдено счетов: {len(osv_data)}")
            
            cash_start = cash_end = 0
            for acc in ['1010', '1030', '1050']:
                if acc in osv_data:
                    cash_start += osv_data[acc]['nac_dt'] - osv_data[acc]['nac_kt']
                    cash_end += osv_data[acc]['kon_dt'] - osv_data[acc]['kon_kt']
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Кэш начало", f"{cash_start/1_000_000:.1f} млн")
            col2.metric("Кэш конец", f"{cash_end/1_000_000:.1f} млн")
            col3.metric("Δ Кэш", f"{(cash_end - cash_start)/1_000_000:.1f} млн")
            
            st.divider()
            st.subheader("📋 Остатки по счетам")
            
            labels = {
                '1010': 'Касса', '1030': 'Расч. счета', '1050': 'Депозиты',
                '1200': 'Дебиторка', '1210': 'ДС покупателей', '1274': 'Займы группе',
                '1300': 'Запасы', '1600': 'ОС', '1700': 'Авансы выданные',
                '2930': 'ОС в разработке', '3120': 'Амортизация', '3150': 'Резерв',
                '3310': 'Кредиторка', '3350': 'Начисления зарплата', '3380': 'Прочие обязат.',
                '3510': 'Авансы полученные',
            }
            
            details = []
            for acc, label in labels.items():
                if acc in osv_data:
                    d = osv_data[acc]
                    details.append({
                        'Счет': acc,
                        'Описание': label,
                        'Деб. Нач.': f"{d['nac_dt']/1e6:.1f}",
                        'Кред. Нач.': f"{d['nac_kt']/1e6:.1f}",
                        'Деб. Конец': f"{d['kon_dt']/1e6:.1f}",
                        'Кред. Конец': f"{d['kon_kt']/1e6:.1f}",
                        'Δ Д': f"{d['delta_dt']/1e6:.1f}",
                        'Δ К': f"{d['delta_kt']/1e6:.1f}",
                    })
            
            st.dataframe(pd.DataFrame(details), use_container_width=True)
            st.success("✅ Готово!")
            
        except Exception as e:
            st.error(f"Ошибка: {str(e)}")
