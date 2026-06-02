#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import io
import sys
from datetime import datetime
import os

# Добавляем путь к скрипту
sys.path.insert(0, '/mnt/user-data/outputs')
from skill_zakupok import SkillZakupok

# Конфиг Streamlit
st.set_page_config(
    page_title="📊 Система управления закупками",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS стили
st.markdown("""
<style>
    .main-header {
        font-size: 2.5em;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 20px;
    }
    .success-box {
        background-color: #d4edda;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #28a745;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #ffc107;
    }
    .info-box {
        background-color: #d1ecf1;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #17a2b8;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.markdown("<div class='main-header'>📊 СИСТЕМА УПРАВЛЕНИЯ ЗАКУПКАМИ СТИРАЛЬНЫХ МАШИН</div>", 
            unsafe_allow_html=True)

# Боковая панель
st.sidebar.title("⚙️ НАСТРОЙКИ")
st.sidebar.markdown("---")

role = st.sidebar.radio(
    "Выберите вашу роль:",
    ["📋 Менеджер ПУ", "📊 Директор (Дашборд)", "👔 Директор по закупкам (Все)"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📌 ИНФОРМАЦИЯ")
st.sidebar.info(
    "Система рассчитывает закупки стиральных машин на основе:\n\n"
    "• Запас на **45 дней** продаж\n"
    "• Учёт товара **в пути** (30 дней)\n"
    "• Определение **неликвидов**\n"
    "• **ABC анализ** товаров\n"
    "• Рекомендации по **clearance**"
)

# ОСНОВНОЙ КОНТЕНТ
tab1, tab2, tab3 = st.tabs(["📁 ЗАГРУЗКА ДАННЫХ", "📊 РЕЗУЛЬТАТЫ", "📚 СПРАВКА"])

# === ТАБ 1: ЗАГРУЗКА ===
with tab1:
    st.header("📁 Загрузите файл данных из 1С")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        ### 📋 Требуемый файл:
        - **Единый файл Excel** (Планирование_стиралки.xlsx)
        
        ### 📊 Должна содержать:
        - Колонка **'Номенклатура'** - описание товара
        - Колонка **'Май 2026_Реализация месяц'** - продажи
        - Колонка **'Товар в пути'** - закупки в пути
        - Дата поступления, цены, остатки
        """)
    
    with col2:
        st.markdown("""
        ### ✅ Периодичность:
        - **Еженедельно** по вторникам
        - **По требованию** - в любой момент
        
        ### 🔄 Процесс:
        1. Выгружаете из 1С
        2. Загружаете сюда
        3. Скачиваете отчёты
        4. Используете в работе
        """)
    
    st.markdown("---")
    
    uploaded_file = st.file_uploader(
        "Выберите файл Excel",
        type=['xlsx'],
        help="Файл должен быть в формате XLSX"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Файл загружен: {uploaded_file.name}")
        
        # Сохраняем временно
        temp_path = f"/tmp/{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Обрабатываем
        try:
            with st.spinner("⏳ Обрабатываю данные..."):
                skill = SkillZakupok(temp_path)
                df = skill.process()
                skill.generate_reports('/tmp')
            
            st.markdown("<div class='success-box'><strong>✅ Данные успешно обработаны!</strong></div>", 
                       unsafe_allow_html=True)
            
            # Статистика
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📦 Товаров обработано", len(df))
            
            with col2:
                k_zakazat = len(df[df['К заказу'] > 0])
                st.metric("🛒 Нужно заказать", k_zakazat)
            
            with col3:
                sum_zakaz = df['Стоимость заказа'].sum()
                st.metric("💰 Сумма закупок", f"${sum_zakaz:,.0f}")
            
            with col4:
                nelikvidov = len(df[df['Статус неликвида'] != 'OK'])
                st.metric("⚠️ Неликвидов", nelikvidov)
            
            st.markdown("---")
            st.subheader("📥 Скачайте отчёты:")
            
            # Генерируем файлы для скачивания
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # ПУ
                pu_df = df[df['К заказу'] > 0][[
                    'Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
                    'Средняя дневная продажа', 'Текущий остаток', 'В пути',
                    'К заказу', 'Актуальная сред. цена', 'Стоимость заказа',
                    'Маржин-сть к средней цене'
                ]].sort_values('Стоимость заказа', ascending=False)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    pu_df.to_excel(writer, sheet_name='ПУ', index=False)
                buffer.seek(0)
                
                st.download_button(
                    label="📋 PU_Zakupok.xlsx",
                    data=buffer,
                    file_name=f"PU_Zakupok_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="pu"
                )
            
            with col2:
                # Неликвиды
                illiq_df = df[df['Статус неликвида'] != 'OK'][[
                    'Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
                    'Дни на складе', 'Дни без продаж', 'Текущий остаток',
                    'Маржин-сть к средней цене', 'Статус неликвида', 'Рекомендация'
                ]].sort_values('Дни на складе', ascending=False)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    illiq_df.to_excel(writer, sheet_name='Неликвиды', index=False)
                buffer.seek(0)
                
                st.download_button(
                    label="⚠️ Report_Nelikvidov.xlsx",
                    data=buffer,
                    file_name=f"Report_Nelikvidov_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="illiq"
                )
            
            with col3:
                # ABC
                abc_df = df[['Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
                             'Средняя дневная продажа', 'ABC']].sort_values(
                    ['ABC', 'Средняя дневная продажа'], ascending=[True, False]
                )
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    abc_df.to_excel(writer, sheet_name='ABC', index=False)
                buffer.seek(0)
                
                st.download_button(
                    label="📊 ABC_Analiz.xlsx",
                    data=buffer,
                    file_name=f"ABC_Analiz_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="abc"
                )
            
            # Второй ряд
            col1, col2 = st.columns(2)
            
            with col1:
                # Dashboard
                by_brand = df.groupby('Производитель').agg({
                    'Код': 'count',
                    'Текущий остаток': 'sum',
                    'Средняя дневная продажа': 'mean',
                    'Маржин-сть к средней цене': 'mean'
                }).round(2)
                by_brand.columns = ['SKU', 'Остаток шт', 'Дневная продажа', 'Маржа %']
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    by_brand.to_excel(writer, sheet_name='По брендам')
                buffer.seek(0)
                
                st.download_button(
                    label="📈 Dashboard.xlsx",
                    data=buffer,
                    file_name=f"Dashboard_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dash"
                )
            
            with col2:
                # Рекомендации
                rec_data = []
                for idx, row in df.iterrows():
                    action = ''
                    reason = ''
                    
                    if row['Статус неликвида'] == 'КРИТИЧНЫЙ':
                        action = 'УБРАТЬ'
                        reason = f'Неликвид {row["Дни на складе"]} дней'
                    elif row['ABC'] == 'A' and row['Средняя дневная продажа'] > 1:
                        action = 'РАСШИРИТЬ'
                        reason = 'TOP продавец'
                    elif row['Средняя дневная продажа'] < 0.05:
                        action = 'СНИЗИТЬ'
                        reason = 'Редко продаётся'
                    
                    if action:
                        rec_data.append({
                            'Производитель': row['Производитель'],
                            'Артикул': row['Номенклатура.Артикул '],
                            'Действие': action,
                            'Маржа %': f"{row['Маржин-сть к средней цене']:.1f}",
                            'Дневная продажа': f"{row['Средняя дневная продажа']:.2f}"
                        })
                
                rec_df = pd.DataFrame(rec_data)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    rec_df.to_excel(writer, sheet_name='Рекомендации', index=False)
                buffer.seek(0)
                
                st.download_button(
                    label="✅ Rekomendacii.xlsx",
                    data=buffer,
                    file_name=f"Rekomendacii_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="rec"
                )
        
        except Exception as e:
            st.error(f"❌ Ошибка при обработке: {str(e)}")
            st.info("Проверьте формат файла и наличие необходимых колонок")

# === ТАБ 2: РЕЗУЛЬТАТЫ ===
with tab2:
    st.header("📊 Обзор результатов")
    
    if uploaded_file is not None:
        st.markdown("### 🎯 Основные метрики")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("💰 Вложено капитала", f"${(df['Актуальная сред. цена'] * df['Текущий остаток']).sum():,.0f}")
        
        with col2:
            st.metric("⏱️ Средний оборот (дни)", f"{(df['Текущий остаток'] / (df['Средняя дневная продажа'] + 0.01)).mean():.1f}")
        
        with col3:
            st.metric("📈 Средняя маржа", f"{df['Маржин-сть к средней цене'].mean():.1f}%")
        
        with col4:
            st.metric("💳 Кредитная комиссия (1%/мес)", 
                     f"${(df['Актуальная сред. цена'] * df['Текущий остаток']).sum() * 0.01:,.0f}")
        
        st.markdown("---")
        
        # Top товары к заказу
        st.subheader("🛒 TOP 10 товаров к заказу")
        top_order = df[df['К заказу'] > 0].nlargest(10, 'Стоимость заказа')[
            ['Производитель', 'Номенклатура', 'К заказу', 'Стоимость заказа', 'Маржин-сть к средней цене']
        ]
        st.dataframe(top_order, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Проблемные товары
        st.subheader("⚠️ Проблемные товары (неликвиды)")
        problem = df[df['Статус неликвида'] != 'OK'].nlargest(10, 'Дни на складе')[
            ['Производитель', 'Номенклатура', 'Дни на складе', 'Статус неликвида', 'Рекомендация']
        ]
        st.dataframe(problem, use_container_width=True, hide_index=True)
    else:
        st.info("Загрузите файл в табе 'ЗАГРУЗКА ДАННЫХ', чтобы увидеть результаты")

# === ТАБ 3: СПРАВКА ===
with tab3:
    st.header("📚 Справка по системе")
    
    st.markdown("""
    ### 📋 ЛОГИКА РАБОТЫ
    
    #### 1️⃣ Расчёт средней дневной продажи
    - Берём данные за 6 месяцев (декабрь 2025 - май 2026)
    - Рассчитываем дневную продажу для каждого месяца
    - Среднее за 6 месяцев = **прогноз потребности**
    
    #### 2️⃣ Расчёт потребности в закупках
    - **Требуемый остаток** = Среднедневная продажа × **45 дней**
    - **Эффективный остаток** = Текущий + В пути
    - **К заказу** = max(0, Требуемый - Эффективный)
    
    #### 3️⃣ Классификация неликвидов
    | Условие | Статус | Действие |
    |---------|--------|----------|
    | 365+ дней на складе | КРИТИЧНЫЙ | СНЯТИЕ / УТИЛИЗАЦИЯ |
    | 180+ дней на складе | УРОВЕНЬ 1 | ПРОМО / СКИДКА -10% |
    | 60+ дней без продаж | УРОВЕНЬ 2 | СКИДКА -15% + ВНИМАНИЕ |
    | 30+ дней без продаж | ВНИМАНИЕ | СКИДКА -10% |
    
    #### 4️⃣ ABC анализ
    - **A-товары** (TOP 20%) - дневная продажа > квантиль 80%
    - **B-товары** (50%) - дневная продажа 50-80%
    - **C-товары** (30%) - дневная продажа < 50%
    
    ### 🎯 ПАРАМЕТРЫ ПО УМОЛЧАНИЮ
    - Запас на складе: **45 дней** продаж
    - Плечо поставки: **30 дней**
    - Максимальный запас: **60 дней** (2× от требуемого)
    - Кредитная ставка: **1% в месяц**
    
    ### 📊 ВЫХОДНЫЕ ОТЧЁТЫ
    1. **ПУ (Purchase Order)** - товары к заказу
    2. **Неликвиды** - товары на внимание
    3. **Дашборд** - метрики по брендам
    4. **ABC анализ** - классификация товаров
    5. **Рекомендации** - действия для менеджера
    """)
    
    st.markdown("---")
    st.success("✅ Система готова к использованию!")

# Футер
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray; font-size: 12px;'>
📦 Система управления закупками стиральных машин | v1.0
</div>
""", unsafe_allow_html=True)
