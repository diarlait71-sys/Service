"""
Streamlit приложение для расчета бонусов отдела сервиса
Версия 2.0 - С поддержкой механиков и остальных сотрудников
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

from sql_connection import (
    build_mssql_connection_url,
    create_sql_engine,
    test_sql_connection,
    query_to_dataframe,
)
from sql_profiles import (
    load_sql_profiles,
    upsert_sql_profile,
    delete_sql_profile,
)

from employee_config import (
    Employee, 
    EmployeeType, 
    KPIIndicator,
    get_employee_config,
    update_employee_config,
    EMPLOYEES_CONFIG
)
from mechanics_calculator import MechanicsCalculator


# ============================================================================
# КОНФИГУРАЦИЯ STREAMLIT
# ============================================================================

st.set_page_config(
    page_title="Расчет бонусов - Сервис",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("💰 Расчет бонусов отдела сервиса")
st.markdown("---")

with st.sidebar:
    st.header("🗄️ SQL Server")

    if 'sql_connected' not in st.session_state:
        st.session_state['sql_connected'] = False
    if 'sql_connection_error' not in st.session_state:
        st.session_state['sql_connection_error'] = ""
    if 'sql_driver' not in st.session_state:
        st.session_state['sql_driver'] = "ODBC Driver 18 for SQL Server"
    if 'sql_use_windows_auth' not in st.session_state:
        st.session_state['sql_use_windows_auth'] = True

    profiles = load_sql_profiles()
    profile_names = sorted(list(profiles.keys()))

    selected_profile = st.selectbox(
        "Профиль подключения",
        ["(новый профиль)"] + profile_names,
        index=0,
    )

    if st.button("Загрузить профиль", key="load_sql_profile"):
        if selected_profile == "(новый профиль)":
            st.info("Выберите сохраненный профиль")
        else:
            profile = profiles[selected_profile]
            st.session_state['sql_server'] = profile.get('server', 'localhost')
            st.session_state['sql_database'] = profile.get('database', 'master')
            st.session_state['sql_driver'] = profile.get('driver', 'ODBC Driver 18 for SQL Server')
            st.session_state['sql_use_windows_auth'] = bool(profile.get('use_windows_auth', True))
            st.session_state['sql_username'] = profile.get('username', '')
            st.success(f"✓ Профиль загружен: {selected_profile}")
            st.rerun()

    if st.button("Удалить профиль", key="delete_sql_profile"):
        if selected_profile == "(новый профиль)":
            st.info("Выберите профиль для удаления")
        else:
            delete_sql_profile(selected_profile)
            st.success(f"✓ Профиль удален: {selected_profile}")
            st.rerun()

    profile_name = st.text_input("Имя профиля", value="")

    sql_server = st.text_input("Сервер", value=st.session_state.get('sql_server', 'localhost'))
    sql_database = st.text_input("База данных", value=st.session_state.get('sql_database', 'master'))
    drivers = ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"]
    current_driver = st.session_state.get('sql_driver', drivers[0])
    if current_driver not in drivers:
        current_driver = drivers[0]
    sql_driver = st.selectbox(
        "ODBC драйвер",
        drivers,
        index=drivers.index(current_driver),
    )

    auth_type = st.radio(
        "Тип авторизации",
        ["Windows", "SQL"],
        horizontal=True,
        index=0 if st.session_state.get('sql_use_windows_auth', True) else 1,
    )
    use_windows_auth = auth_type == "Windows"

    sql_username = ""
    sql_password = ""
    if not use_windows_auth:
        sql_username = st.text_input("Логин", value=st.session_state.get('sql_username', 'sa'))
        sql_password = st.text_input("Пароль", type="password")

    if st.button("Сохранить профиль", key="save_sql_profile"):
        if not profile_name.strip():
            st.error("❌ Укажите имя профиля")
        else:
            upsert_sql_profile(
                profile_name.strip(),
                {
                    'server': sql_server,
                    'database': sql_database,
                    'driver': sql_driver,
                    'use_windows_auth': use_windows_auth,
                    'username': sql_username,
                },
            )
            st.success(f"✓ Профиль сохранен: {profile_name.strip()}")

    if st.button("Проверить SQL подключение", key="check_sql_connection"):
        try:
            conn_url = build_mssql_connection_url(
                server=sql_server,
                database=sql_database,
                driver=sql_driver,
                username=sql_username or None,
                password=sql_password or None,
                use_windows_auth=use_windows_auth,
                encrypt="yes",
                trust_server_certificate=True,
                timeout=30,
            )
            engine = create_sql_engine(conn_url)
            ok, message = test_sql_connection(engine)

            if ok:
                st.session_state['sql_connected'] = True
                st.session_state['sql_engine'] = engine
                st.session_state['sql_connection_error'] = ""
                st.session_state['sql_server'] = sql_server
                st.session_state['sql_database'] = sql_database
                st.session_state['sql_driver'] = sql_driver
                st.session_state['sql_use_windows_auth'] = use_windows_auth
                st.session_state['sql_username'] = sql_username
                st.success("✓ SQL подключение установлено")
            else:
                st.session_state['sql_connected'] = False
                st.session_state['sql_connection_error'] = message
                st.error("❌ Не удалось подключиться к SQL")
        except Exception as e:
            st.session_state['sql_connected'] = False
            st.session_state['sql_connection_error'] = str(e)
            st.error(f"❌ Ошибка подключения: {str(e)}")

    if st.session_state.get('sql_connected', False):
        st.caption("Статус: подключено")
    elif st.session_state.get('sql_connection_error'):
        st.caption("Статус: ошибка подключения")

# Выбор режима
mode = st.radio(
    "🔧 Выберите режим расчета:",
    ["👨‍🔧 Механики", "📊 Остальные сотрудники"],
    horizontal=True
)


# ============================================================================
# РЕЖИМ 1: РАСЧЕТ ДЛЯ МЕХАНИКОВ
# ============================================================================

if mode == "👨‍🔧 Механики":
    st.header("👨‍🔧 Расчет бонусов механиков")
    st.markdown("""
    Логика расчета в текущей версии:
    1. Загрузка отчета [Механики.xlsx] (лист `TDSheet`) из 1С
    2. Расчет в разрезе типов нарядов
    3. Фиксированные ставки:
       - **Гарантийные работы**: `3915` тг
       - **Внутренние/ОП сервис/ППП/предпродажка**: `1895` тг
    4. Типы, не относящиеся к бонусам, исключаются из расчета (их норма-часы не отражаются)
    """)

    with st.sidebar:
        st.header("📁 Загрузка данных")

        mechanics_file = st.file_uploader(
            "Загрузите отчет Механики (1С)",
            type=["xlsx", "xls"],
            key="mechanics_report_file",
            help="Ожидается формат с колонкой типа наряда в первом столбце и норма-часами в 5-й колонке (Н/Ч)"
        )

        st.caption("Справочник ФИО/ставок по электрикам и другим ролям добавим следующим шагом")

        st.markdown("---")
        st.subheader("⚙️ Текущий маппинг (можно менять)")
        kw_warranty = st.text_input(
            "Ключевые слова для гарантийных",
            value="гарантия, гарантий"
        )
        kw_internal = st.text_input(
            "Ключевые слова для внутренних/ППП/ОП сервис",
            value="внутрен, вн, оп сервис, ппп, предпродаж, сервис ремзона, сервис"
        )
        kw_commercial = st.text_input(
            "Ключевые слова для коммерческих",
            value="доп, коммерч, вэк"
        )
        kw_exclude = st.text_input(
            "Ключевые слова для исключения",
            value="банковский продукт, прочее"
        )

        rate_warranty = st.number_input("Ставка гарантийные (тг)", min_value=0.0, value=3915.0, step=5.0)
        rate_internal = st.number_input("Ставка внутренние/ППП (тг)", min_value=0.0, value=1895.0, step=5.0)
        rate_commercial = st.number_input("Ставка коммерческие (тг)", min_value=0.0, value=0.0, step=5.0)

    if mechanics_file:
        try:
            file_bytes = mechanics_file.read()

            report_df = None
            for sheet_name in ['TDSheet', 0]:
                try:
                    report_df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)
                    break
                except Exception:
                    continue

            if report_df is None:
                st.error("❌ Не удалось прочитать отчет механиков")
            else:
                calc = MechanicsCalculator()
                loaded = calc.load_work_orders_from_mechanics_report(report_df)

                if not loaded:
                    st.error("❌ Не удалось извлечь строки нарядов из отчета")
                else:
                    category_keywords = {
                        'warranty': [x.strip().lower() for x in kw_warranty.split(',') if x.strip()],
                        'internal': [x.strip().lower() for x in kw_internal.split(',') if x.strip()],
                        'commercial': [x.strip().lower() for x in kw_commercial.split(',') if x.strip()],
                        'exclude': [x.strip().lower() for x in kw_exclude.split(',') if x.strip()],
                    }
                    category_rates = {
                        'warranty': float(rate_warranty),
                        'internal': float(rate_internal),
                        'commercial': float(rate_commercial),
                        'exclude': 0.0,
                    }

                    tab1, tab2, tab3 = st.tabs(["📊 Предпросмотр", "📈 Расчет", "💾 Выгрузка"])

                    with tab1:
                        st.subheader("Сырые строки отчета")
                        st.dataframe(report_df.head(20), use_container_width=True)
                        st.caption(f"Размер файла: {report_df.shape[0]} строк × {report_df.shape[1]} столбцов")

                        parsed_preview = pd.DataFrame([
                            {
                                'ДЦ/Город': wo.dealer_center,
                                'Тип наряда (сырой)': wo.description,
                                'Норма часов': wo.norm_hours,
                            }
                            for wo in calc.work_orders[:200]
                        ])
                        st.subheader("Извлеченные строки для расчета")
                        st.dataframe(parsed_preview, use_container_width=True)
                        st.caption(f"Извлечено строк: {len(calc.work_orders)}")

                    with tab2:
                        result_df = calc.calculate_bonus_by_work_type(
                            category_keywords=category_keywords,
                            category_rates=category_rates,
                            exclude_categories=['exclude'],
                        )

                        if len(result_df) == 0:
                            st.warning("Нет строк после применения маппинга и исключений")
                        else:
                            st.dataframe(result_df, use_container_width=True)

                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Строк в расчете", len(result_df))
                            with col2:
                                st.metric("Норма часов (итого)", f"{result_df['Норма часов'].sum():,.2f}")
                            with col3:
                                st.metric("Начисление (итого)", f"{result_df['Начисление'].sum():,.0f} тг")
                            with col4:
                                dc_count = result_df['ДЦ/Город'].nunique()
                                st.metric("ДЦ/Городов", int(dc_count))

                            st.session_state['mechanics_grouped_df'] = result_df
                            st.session_state['mechanics_parsed_df'] = pd.DataFrame([
                                {
                                    'ДЦ/Город': wo.dealer_center,
                                    'Тип наряда (сырой)': wo.description,
                                    'Норма часов': wo.norm_hours,
                                }
                                for wo in calc.work_orders
                            ])

                    with tab3:
                        if 'mechanics_grouped_df' in st.session_state:
                            grouped_df = st.session_state['mechanics_grouped_df']
                            parsed_df = st.session_state.get('mechanics_parsed_df', pd.DataFrame())

                            buffer = BytesIO()
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                grouped_df.to_excel(writer, sheet_name='Расчет по типам', index=False)
                                if len(parsed_df) > 0:
                                    parsed_df.to_excel(writer, sheet_name='Сырые строки', index=False)

                            st.download_button(
                                label="📥 Скачать Excel",
                                data=buffer.getvalue(),
                                file_name="mechanics_bonus_by_work_type.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.info("Выполните расчет на вкладке 'Расчет'")

        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            with st.expander("Показать детали ошибки"):
                st.write(e)

    else:
        st.info("📁 Загрузите отчет Механики в левой панели")


# ============================================================================
# РЕЖИМ 2: РАСЧЕТ ДЛЯ ОСТАЛЬНЫХ СОТРУДНИКОВ
# ============================================================================

else:  # mode == "📊 Остальные сотрудники"
    st.header("📊 Расчет бонусов других сотрудников")
    
    with st.sidebar:
        st.header("📁 Загрузка данных")
        
        tab1, tab2 = st.tabs(["📤 Загрузить", "⚙️ Конфиг"])
        
        # TAB 1: Загрузка файлов
        with tab1:
            st.subheader("План (Excel)")
            plan_file = st.file_uploader(
                "Загрузите файл с планом",
                type=["xlsx", "xls"],
                key="plan_file"
            )
            
            st.subheader("Факт (выгрузка 1С)")
            fact_file = st.file_uploader(
                "Загрузите файл с фактом из 1С",
                type=["xlsx", "xls"],
                key="fact_file"
            )
            
            minimum_execution = st.slider(
                "Минимальное выполнение плана для бонуса (%)",
                min_value=50,
                max_value=150,
                value=100,
                step=5
            ) / 100

            st.markdown("---")
            use_sql_data = st.checkbox("Использовать SQL вместо Excel", value=False)

            if use_sql_data:
                plan_query = st.text_area(
                    "SQL для Плана",
                    value="SELECT TOP 100 fio AS [ФИО], kpi AS [КПИ], plan_value AS [План] FROM plan_table",
                    height=100,
                    key="plan_sql_query",
                )
                fact_query = st.text_area(
                    "SQL для Факта",
                    value="SELECT TOP 100 fio AS [ФИО], kpi AS [КПИ], fact_value AS [Факт], rate AS [Тариф] FROM fact_table",
                    height=100,
                    key="fact_sql_query",
                )

                if st.button("Загрузить данные из SQL", key="load_sql_data"):
                    if not st.session_state.get('sql_connected', False):
                        st.error("❌ Сначала подключитесь к SQL Server в блоке слева")
                    else:
                        try:
                            sql_engine = st.session_state['sql_engine']
                            st.session_state['sql_plan_df'] = query_to_dataframe(sql_engine, plan_query)
                            st.session_state['sql_fact_df'] = query_to_dataframe(sql_engine, fact_query)
                            st.success("✓ Данные из SQL загружены")
                        except Exception as e:
                            st.error(f"❌ Ошибка SQL запроса: {str(e)}")
        
        # TAB 2: Конфигурация сотрудников
        with tab2:
            st.subheader("⚙️ Конфигурация")
            
            if EMPLOYEES_CONFIG:
                fio_list = [emp.fio for emp in EMPLOYEES_CONFIG]
                selected_fio = st.selectbox("Выбрать сотрудника", fio_list)
                
                if selected_fio:
                    emp = get_employee_config(selected_fio)
                    
                    st.markdown(f"### {emp.fio}")
                    
                    emp.department = st.text_input("Отдел", emp.department)
                    
                    emp.employee_type = st.selectbox(
                        "Тип сотрудника",
                        [EmployeeType.SALES, EmployeeType.SERVICE, EmployeeType.WARRANTY],
                    )
                    
                    emp.base_salary = st.number_input(
                        "Репер (базовая сумма)",
                        value=emp.base_salary,
                        min_value=0,
                        step=1000
                    )
                    
                    emp.bonus_coefficient = st.slider(
                        "Коэффициент бонуса",
                        min_value=0.0,
                        max_value=2.0,
                        value=emp.bonus_coefficient,
                        step=0.1
                    )
                    
                    if st.button("💾 Сохранить"):
                        update_employee_config(emp)
                        st.success(f"✓ Сохранено")
    
    plan_df = None
    fact_df = None

    if use_sql_data and 'sql_plan_df' in st.session_state and 'sql_fact_df' in st.session_state:
        plan_df = st.session_state['sql_plan_df']
        fact_df = st.session_state['sql_fact_df']
    elif plan_file and fact_file:
        plan_df = pd.read_excel(plan_file)
        fact_df = pd.read_excel(fact_file)

    if plan_df is not None and fact_df is not None:
        try:
            st.success("✓ Файлы загружены успешно")
            
            tab1, tab2, tab3, tab4 = st.tabs(
                ["📊 Предпросмотр", "⚙️ Конфиг", "📈 Результаты", "💾 Выгрузка"]
            )
            
            # TAB 1: Предпросмотр
            with tab1:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("План")
                    st.dataframe(plan_df.head(10), use_container_width=True)
                
                with col2:
                    st.subheader("Факт")
                    st.dataframe(fact_df.head(10), use_container_width=True)
            
            # TAB 2: Конфиг
            with tab2:
                st.subheader("Настройка сопоставления колонок")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("**План:**")
                    plan_fio_col = st.selectbox("ФИО (План)", list(plan_df.columns), key="p_fio")
                    plan_kpi_col = st.selectbox("КПИ (План)", list(plan_df.columns), key="p_kpi")
                    plan_value_col = st.selectbox("Значение (План)", list(plan_df.columns), key="p_val")
                
                with col2:
                    st.markdown("**Факт:**")
                    fact_fio_col = st.selectbox("ФИО (Факт)", list(fact_df.columns), key="f_fio")
                    fact_kpi_col = st.selectbox("КПИ (Факт)", list(fact_df.columns), key="f_kpi")
                    fact_value_col = st.selectbox("Значение (Факт)", list(fact_df.columns), key="f_val")
                
                with col3:
                    st.markdown("**Опции:**")
                    rate_col = st.selectbox("Тариф", [None] + list(fact_df.columns), key="rate")
            
            # TAB 3: Результаты
            with tab3:
                st.subheader("📈 Расчет бонусов")
                
                if st.button("🔄 Рассчитать"):
                    with st.spinner("Расчет..."):
                        merged_df = plan_df[[plan_fio_col, plan_kpi_col, plan_value_col]].copy()
                        merged_df.columns = ['ФИО', 'КПИ', 'План']
                        
                        fact_data = fact_df[[fact_fio_col, fact_kpi_col, fact_value_col]].copy()
                        fact_data.columns = ['ФИО', 'КПИ', 'Факт']
                        
                        if rate_col:
                            fact_data['Тариф'] = fact_df[rate_col]
                        else:
                            fact_data['Тариф'] = 0
                        
                        merged_df = merged_df.merge(fact_data, on=['ФИО', 'КПИ'], how='outer').fillna(0)
                        
                        merged_df['Выполнение %'] = (
                            (merged_df['Факт'] / merged_df['План'] * 100)
                            .replace([np.inf, -np.inf], 0)
                            .fillna(0)
                            .round(1)
                        )
                        
                        merged_df['Сумма'] = (merged_df['Тариф'] * merged_df['Факт']).round(2)
                        
                        results = []
                        for fio in merged_df['ФИО'].unique():
                            emp_data = merged_df[merged_df['ФИО'] == fio]
                            emp_config = get_employee_config(fio)
                            
                            emp_config.kpi_indicators = {}
                            for _, row in emp_data.iterrows():
                                kpi = KPIIndicator(
                                    name=str(row['КПИ']),
                                    plan=float(row['План']),
                                    fact=float(row['Факт']),
                                    rate_per_hour=float(row['Тариф']),
                                    weight=1.0
                                )
                                emp_config.add_kpi(kpi)
                            
                            bonus = emp_config.calculate_bonus(minimum_execution)
                            
                            results.append({
                                'ФИО': fio,
                                'Отдел': emp_config.department,
                                'КПИ': len(emp_config.kpi_indicators),
                                'Выполнение %': emp_data['Выполнение %'].mean(),
                                'Сумма': emp_data['Сумма'].sum(),
                                'Бонус': bonus,
                            })
                        
                        results_df = pd.DataFrame(results)
                        st.dataframe(results_df, use_container_width=True)
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Сотрудников", len(results_df))
                        with col2:
                            st.metric("Сумма", f"{results_df['Бонус'].sum():,.0f} тг")
                        with col3:
                            st.metric("Средн. выполнение", f"{results_df['Выполнение %'].mean():.1f}%")
                        
                        st.session_state['results_df'] = results_df
            
            # TAB 4: Выгрузка
            with tab4:
                st.subheader("💾 Выгрузка результатов")
                
                if 'results_df' in st.session_state:
                    results_df = st.session_state['results_df']
                    st.dataframe(results_df, use_container_width=True)
                    
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        results_df.to_excel(writer, sheet_name='Бонусы', index=False)
                    
                    st.download_button(
                        label="📥 Скачать Excel",
                        data=buffer.getvalue(),
                        file_name="bonus_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("ℹ️ Сначала выполните расчет")
        
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            with st.expander("Показать детали ошибки"):
                st.write(e)
    
    else:
        if use_sql_data:
            st.info("🗄️ Выполните SQL-подключение и загрузите данные из SQL")
        else:
            st.info("📁 Загрузите файлы в левой панели")


# Футер
st.markdown("---")
st.caption("Приложение для расчета бонусов отдела сервиса | v2.0")
