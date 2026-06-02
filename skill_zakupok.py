#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

warnings.filterwarnings('ignore')

class SkillZakupok:
    """Скилл для управления закупками стиральных машин"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.dne_zapasa = 45  # дней запаса
        self.plecho_postavki = 30  # дней по умолчанию
        self.dny_bez_prodazh_l1 = 30  # Уровень 1 неликвида
        self.dny_bez_prodazh_l2 = 60  # Уровень 2 неликвида
        self.dny_na_skladе_l1 = 180  # 180 дней неликвид
        self.dny_na_skladе_l2 = 365  # 365 дней критичный
        
        self.load_data()
        
    def load_data(self):
        """Загружаем и подготавливаем данные"""
        df = pd.read_excel(self.file_path, sheet_name='Лист1')
        
        # Строим заголовки из двух рядов
        headers_row0 = df.iloc[0]
        headers_row1 = df.iloc[1]
        
        new_headers = []
        for i, (h0, h1) in enumerate(zip(headers_row0, headers_row1)):
            if pd.isna(h0) or h0 == '':
                new_headers.append(str(h1))
            elif pd.isna(h1) or h1 == '':
                new_headers.append(str(h0))
            else:
                new_headers.append(f"{h0}_{h1}")
        
        df.columns = new_headers
        self.df = df.iloc[2:].reset_index(drop=True)
        
        # Очищаем данные
        self.df = self.clean_data()
        
    def clean_data(self):
        """Очищаем и конвертируем типы данных"""
        df = self.df.copy()
        
        # Конвертируем числовые колонки
        numeric_cols = [
            'Актуальная сред. цена', 'Цена продажи', 'Маржин-сть к средней цене',
            'Май 2026_Реализация месяц', 'Апрель 2026_Реализация месяц',
            'Март 2026_Реализация месяц', 'Февраль 2026_Реализация месяц',
            'Январь 2026_Реализация месяц', 'Декабрь 2025_Реализация месяц',
            'СКЛАД_Кон. ост-к', 'Товар в пути_Кон. ост-к'
        ]
        
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Конвертируем дату
        if 'Дата поступления' in df.columns:
            df['Дата поступления'] = pd.to_datetime(df['Дата поступления'], errors='coerce')
        
        return df
    
    def calculate_month_coefficients(self):
        """Рассчитываем коэффициенты сезонности для каждого месяца"""
        df = self.df.copy()
        
        # Данные по месяцам с их днями
        months_data = {
            'Декабрь 2025': ('Декабрь 2025_Реализация месяц', 15),  # половина месяца
            'Январь 2026': ('Январь 2026_Реализация месяц', 31),
            'Февраль 2026': ('Февраль 2026_Реализация месяц', 28),  # 2026 - не високосный
            'Март 2026': ('Март 2026_Реализация месяц', 31),
            'Апрель 2026': ('Апрель 2026_Реализация месяц', 30),
            'Май 2026': ('Май 2026_Реализация месяц', 31),
        }
        
        # Рассчитываем дневные продажи для каждого месяца
        daily_sales = {}
        for month_name, (col_name, days) in months_data.items():
            if col_name in df.columns:
                daily_sales[month_name] = df[col_name].fillna(0) / days
            else:
                daily_sales[month_name] = 0
        
        # Рассчитываем среднюю дневную продажу за период
        monthly_avg = pd.DataFrame(daily_sales).mean(axis=1)
        
        # Рассчитываем коэффициент для каждого месяца
        self.month_coefficients = {}
        for month_name, (col_name, days) in months_data.items():
            if col_name in df.columns:
                month_avg = df[col_name].fillna(0).mean()
                overall_avg = df[[c for c, _ in months_data.values() if c in df.columns]].fillna(0).mean().mean()
                if overall_avg > 0:
                    self.month_coefficients[month_name] = month_avg / (overall_avg * days / 30)
                else:
                    self.month_coefficients[month_name] = 1.0
        
        # Коэффициенты для будущих месяцев (на основе истории)
        self.month_coefficients['Июнь 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 0.95
        self.month_coefficients['Июль 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 0.85
        self.month_coefficients['Август 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 0.90
        self.month_coefficients['Сентябрь 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 1.0
        self.month_coefficients['Октябрь 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 1.1
        self.month_coefficients['Ноябрь 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 1.3
        self.month_coefficients['Декабрь 2026'] = self.month_coefficients.get('Май 2026', 1.0) * 1.25
        
        return self.month_coefficients
    
    def calculate_average_daily_sales(self):
        """Рассчитываем среднюю дневную продажу"""
        df = self.df.copy()
        
        # Рассчитываем коэффициенты месяцев
        self.calculate_month_coefficients()
        
        # Берём данные за периоды
        may_sales = df['Май 2026_Реализация месяц'].fillna(0)
        apr_sales = df['Апрель 2026_Реализация месяц'].fillna(0)
        mar_sales = df['Март 2026_Реализация месяц'].fillna(0)
        feb_sales = df['Февраль 2026_Реализация месяц'].fillna(0)
        jan_sales = df['Январь 2026_Реализация месяц'].fillna(0)
        dec_sales = df['Декабрь 2025_Реализация месяц'].fillna(0)
        
        # Дни за каждый период
        may_daily = may_sales / 31
        apr_daily = apr_sales / 30
        mar_daily = mar_sales / 31
        feb_daily = feb_sales / 28
        jan_daily = jan_sales / 31
        dec_daily = dec_sales / 15
        
        # Средняя за 6 месяцев
        avg_daily = (may_daily + apr_daily + mar_daily + feb_daily + jan_daily + dec_daily) / 6
        
        df['Средняя дневная продажа'] = avg_daily
        
        return df
    
    def calculate_requirements(self, df):
        """Рассчитываем потребность"""
        df = df.copy()
        
        # Требуемый остаток = среднедневная × 45 дней
        df['Требуемый остаток'] = (df['Средняя дневная продажа'] * self.dne_zapasa).round(0).astype(int)
        
        # Текущий остаток на складе
        df['Текущий остаток'] = df['СКЛАД_Кон. ост-к'].fillna(0).astype(int)
        
        # Товар в пути
        df['В пути'] = df['Товар в пути_Кон. ост-к'].fillna(0).astype(int)
        
        # Эффективный остаток (текущий + в пути)
        df['Эффективный остаток'] = df['Текущий остаток'] + df['В пути']
        
        # Рекомендуемое количество к заказу (ОКРУГЛЯЕМ!)
        df['К заказу'] = (df['Требуемый остаток'] - df['Эффективный остаток']).apply(lambda x: max(0, int(np.ceil(x))))
        
        # Стоимость заказа
        df['Стоимость заказа'] = df['К заказу'] * df['Актуальная сред. цена']
        
        # Статус
        def get_status(row):
            if row['К заказу'] > 0:
                return 'ЗАКАЗАТЬ'
            elif row['Эффективный остаток'] > row['Требуемый остаток'] * 1.5:
                return 'ПЕРЕПОЛНЕНИЕ'
            else:
                return 'OK'
        
        df['Статус'] = df.apply(get_status, axis=1)
        
        # ПЛАН НА 6 МЕСЯЦЕВ
        future_months = ['Июнь 2026', 'Июль 2026', 'Август 2026', 'Сентябрь 2026', 'Октябрь 2026', 'Ноябрь 2026']
        month_days = {
            'Июнь 2026': 30,
            'Июль 2026': 31,
            'Август 2026': 31,
            'Сентябрь 2026': 30,
            'Октябрь 2026': 31,
            'Ноябрь 2026': 30,
        }
        
        for month in future_months:
            coeff = self.month_coefficients.get(month, 1.0)
            days = month_days.get(month, 30)
            # Прогноз = среднедневная × коэффициент месяца × дни месяца
            df[f'Прогноз {month}'] = (df['Средняя дневная продажа'] * coeff * days).round(0).astype(int)
        
        # Итого на 6 месяцев
        forecast_cols = [f'Прогноз {month}' for month in future_months]
        df['Потребность 6 мес'] = df[forecast_cols].sum(axis=1)
        
        return df
    
    def identify_illiquids(self, df):
        """Определяем неликвиды"""
        df = df.copy()
        
        today = pd.Timestamp.now()
        
        # Расчет дней на складе
        df['Дни на складе'] = df['Дата поступления'].apply(
            lambda x: (today - x).days if pd.notna(x) else 0
        )
        
        # Дни без продаж - считаем по последним продажам (май - последний месяц)
        # Если май > 0, то дни без продаж считаются с конца мая (примерно 2 дня)
        # Если нет продаж в мае, то смотрим апрель и т.д.
        
        def get_days_without_sales(row):
            # Если май > 0, то продажи были
            if row['Май 2026_Реализация месяц'] > 0:
                return 2  # примерно 2 дня от конца мая
            elif row['Апрель 2026_Реализация месяц'] > 0:
                return 32  # май + примерно 1 день
            elif row['Март 2026_Реализация месяц'] > 0:
                return 62  # май + апрель + примерно 1 день
            elif row['Февраль 2026_Реализация месяц'] > 0:
                return 93
            elif row['Январь 2026_Реализация месяц'] > 0:
                return 124
            elif row['Декабрь 2025_Реализация месяц'] > 0:
                return 157
            else:
                return 9999  # Никогда не продавалось
        
        df['Дни без продаж'] = df.apply(get_days_without_sales, axis=1)
        
        # Классификация неликвида
        def classify_illiquid(row):
            # Если плечо поставки > 60 дней, игнорируем 180-дневный порог
            plecho = row.get('Плечо поставки', 30)
            if pd.isna(plecho):
                plecho = 30
            
            # Критичный неликвид (365+ дней)
            if row['Дни на складе'] > self.dny_na_skladе_l2:
                return 'КРИТИЧНЫЙ'
            
            # 180+ дней, но не если долгое плечо
            if row['Дни на складе'] > self.dny_na_skladе_l1 and plecho <= 60:
                return 'УРОВЕНЬ 1'
            
            # 60+ дней без продаж
            if row['Дни без продаж'] >= self.dny_bez_prodazh_l2:
                return 'УРОВЕНЬ 2'
            
            # 30+ дней без продаж
            if row['Дни без продаж'] >= self.dny_bez_prodazh_l1:
                return 'ВНИМАНИЕ'
            
            return 'OK'
        
        df['Статус неликвида'] = df.apply(classify_illiquid, axis=1)
        
        # Рекомендация по действию
        def get_illiquid_recommendation(row):
            status = row['Статус неликвида']
            if status == 'КРИТИЧНЫЙ':
                return 'СНЯТИЕ / УТИЛИЗАЦИЯ'
            elif status == 'УРОВЕНЬ 2':
                return 'СКИДКА -15% + ВНИМАНИЕ КАТЕГОРИЙЩИКА'
            elif status == 'УРОВЕНЬ 1':
                return 'ПРОМО / СКИДКА -10%'
            elif status == 'ВНИМАНИЕ':
                return 'СКИДКА -10%'
            else:
                return ''
        
        df['Рекомендация'] = df.apply(get_illiquid_recommendation, axis=1)
        
        return df
    
    def classify_abc(self, df):
        """ABC анализ"""
        df = df.copy()
        
        # Берём дневную продажу
        sales = df['Средняя дневная продажа'].fillna(0)
        
        # Процентили
        a_threshold = sales.quantile(0.80)  # TOP 20%
        b_threshold = sales.quantile(0.50)  # 50-80%
        
        def classify(sales_val):
            if sales_val >= a_threshold:
                return 'A'
            elif sales_val >= b_threshold:
                return 'B'
            else:
                return 'C'
        
        df['ABC'] = sales.apply(classify)
        
        return df
    
    def process(self):
        """Основной процесс обработки"""
        print("⏳ Обрабатываю данные...")
        
        # 1. Рассчитываем среднюю дневную продажу
        df = self.calculate_average_daily_sales()
        print("  ✓ Рассчитана средняя дневная продажа")
        
        # 2. Рассчитываем потребность
        df = self.calculate_requirements(df)
        print("  ✓ Рассчитана потребность в закупках")
        
        # 3. Определяем неликвиды
        df = self.identify_illiquids(df)
        print("  ✓ Определены неликвиды")
        
        # 4. ABC анализ
        df = self.classify_abc(df)
        print("  ✓ Выполнен ABC анализ")
        
        self.df = df
        return df
    
    def generate_reports(self, output_dir='/mnt/user-data/outputs'):
        """Генерируем все отчёты"""
        df = self.df.copy()
        
        print("\n📊 Генерирую отчёты...")
        
        # 1. ПУ на закупки
        self.generate_pu(df, output_dir)
        print("  ✓ ПУ создана")
        
        # 2. Отчёт о неликвидах
        self.generate_illiquids_report(df, output_dir)
        print("  ✓ Отчёт о неликвидах создан")
        
        # 3. Дашборд
        self.generate_dashboard(df, output_dir)
        print("  ✓ Дашборд создан")
        
        # 4. ABC анализ
        self.generate_abc_report(df, output_dir)
        print("  ✓ ABC анализ создан")
        
        # 5. Рекомендации
        self.generate_recommendations(df, output_dir)
        print("  ✓ Рекомендации созданы")
        
        # 6. НОВОЕ: План на 6 месяцев
        self.generate_forecast_plan(df, output_dir)
        print("  ✓ План на 6 месяцев создан")
        
        print(f"\n✅ Все отчёты сохранены в {output_dir}")
    
    def generate_forecast_plan(self, df, output_dir):
        """Генерируем план закупок на 6 месяцев"""
        future_months = ['Июнь 2026', 'Июль 2026', 'Август 2026', 'Сентябрь 2026', 'Октябрь 2026', 'Ноябрь 2026']
        
        # Берём нужные колонки
        plan_df = df[['Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
                      'Средняя дневная продажа', 'Текущий остаток', 'Актуальная сред. цена']].copy()
        
        # Добавляем прогнозы по месяцам
        for month in future_months:
            if f'Прогноз {month}' in df.columns:
                plan_df[f'{month}'] = df[f'Прогноз {month}']
        
        # Добавляем потребность на 6 месяцев
        plan_df['ИТОГО 6 мес'] = df['Потребность 6 мес']
        
        # Добавляем коэффициент месяца для справки
        for month in future_months:
            coeff = self.month_coefficients.get(month, 1.0)
            plan_df[f'Коэфф {month}'] = f"{coeff:.2f}"
        
        plan_df = plan_df.sort_values('ИТОГО 6 мес', ascending=False)
        
        output_file = f"{output_dir}/Plan_6_mesyacev_{datetime.now().strftime('%Y%m%d')}.xlsx"
        plan_df.to_excel(output_file, sheet_name='План на 6 месяцев', index=False)
        
        # Также добавляем лист с коэффициентами
        coeff_data = pd.DataFrame([
            {'Месяц': month, 'Коэффициент сезонности': f"{self.month_coefficients.get(month, 1.0):.2f}"}
            for month in future_months
        ])
        
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            coeff_data.to_excel(writer, sheet_name='Коэффициенты', index=False)
    
    def generate_pu(self, df, output_dir):
        """Генерируем Purchase Order"""
        # Фильтруем товары к заказу
        to_order = df[df['К заказу'] > 0][
            ['Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
             'Средняя дневная продажа', 'Текущий остаток', 'В пути',
             'Требуемый остаток', 'К заказу', 'Актуальная сред. цена',
             'Стоимость заказа', 'Статус', 'Маржин-сть к средней цене']
        ].copy()
        
        to_order = to_order.sort_values('Стоимость заказа', ascending=False)
        to_order = to_order.round(2)
        
        # Сохраняем в Excel
        output_file = f"{output_dir}/PU_Zakupok_{datetime.now().strftime('%Y%m%d')}.xlsx"
        to_order.to_excel(output_file, sheet_name='ПУ', index=False)
        
    def generate_illiquids_report(self, df, output_dir):
        """Отчёт о неликвидах"""
        illiquids = df[df['Статус неликвида'] != 'OK'][
            ['Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
             'Дата поступления', 'Дни на складе', 'Дни без продаж',
             'Текущий остаток', 'Маржин-сть к средней цене', 'Статус неликвида',
             'Рекомендация']
        ].copy()
        
        illiquids = illiquids.sort_values('Дни на складе', ascending=False)
        
        output_file = f"{output_dir}/Report_Nelikvidov_{datetime.now().strftime('%Y%m%d')}.xlsx"
        illiquids.to_excel(output_file, sheet_name='Неликвиды', index=False)
    
    def generate_dashboard(self, df, output_dir):
        """Дашборд здоровья запасов"""
        # По брендам
        by_brand = df.groupby('Производитель').agg({
            'Код': 'count',
            'Текущий остаток': 'sum',
            'Средняя дневная продажа': 'mean',
            'Маржин-сть к средней цене': 'mean',
            'Актуальная сред. цена': lambda x: (x * df.loc[x.index, 'Текущий остаток']).sum(),
            'Статус неликвида': lambda x: (x != 'OK').sum()
        }).round(2)
        
        by_brand.columns = ['SKU', 'Остаток шт', 'Дневная продажа', 'Маржа %',
                            'Капитал инвест $', 'Неликвидов']
        
        # Общие метрики
        total_capital = (df['Актуальная сред. цена'] * df['Текущий остаток']).sum()
        avg_turnover = (df['Текущий остаток'] / (df['Средняя дневная продажа'] + 0.01)).mean()
        avg_margin = df['Маржин-сть к средней цене'].mean()
        illiquid_count = (df['Статус неликвида'] != 'OK').sum()
        illiquid_percent = (illiquid_count / len(df) * 100)
        
        summary_data = {
            'Показатель': [
                'Общий капитал на складе',
                'Средний оборот (дни)',
                'Средняя маржа %',
                'Товаров в неликвиде',
                '% неликвида',
                'Кредитная комиссия (1%/мес)'
            ],
            'Значение': [
                f'${total_capital:,.0f}',
                f'{avg_turnover:.1f}',
                f'{avg_margin:.1f}%',
                f'{illiquid_count}',
                f'{illiquid_percent:.1f}%',
                f'${total_capital * 0.01:,.0f}'
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        
        # Сохраняем
        output_file = f"{output_dir}/Dashboard_{datetime.now().strftime('%Y%m%d')}.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            by_brand.to_excel(writer, sheet_name='По брендам')
            summary_df.to_excel(writer, sheet_name='Общие метрики', index=False)
    
    def generate_abc_report(self, df, output_dir):
        """ABC анализ"""
        abc_data = df[['Производитель', 'Номенклатура.Артикул ', 'Номенклатура',
                       'Средняя дневная продажа', 'ABC']].copy()
        abc_data = abc_data.sort_values(['ABC', 'Средняя дневная продажа'], 
                                         ascending=[True, False])
        
        output_file = f"{output_dir}/ABC_Analiz_{datetime.now().strftime('%Y%m%d')}.xlsx"
        abc_data.to_excel(output_file, sheet_name='ABC', index=False)
    
    def generate_recommendations(self, df, output_dir):
        """Рекомендации"""
        # Товары для действия
        recommendations = []
        
        for idx, row in df.iterrows():
            action = ''
            reason = ''
            
            # Критичный неликвид
            if row['Статус неликвида'] == 'КРИТИЧНЫЙ':
                action = 'УБРАТЬ'
                reason = f'Неликвид {row["Дни на складе"]} дней'
            
            # Рекомендуемые к расширению
            elif row['ABC'] == 'A' and row['Средняя дневная продажа'] > 1:
                action = 'РАСШИРИТЬ'
                reason = f'TOP продавец, маржа {row["Маржин-сть к средней цене"]:.1f}%'
            
            # Падающий тренд
            elif row['Средняя дневная продажа'] < 0.05:
                action = 'СНИЗИТЬ'
                reason = 'Редко продаётся'
            
            if action:
                recommendations.append({
                    'Производитель': row['Производитель'],
                    'Артикул': row['Номенклатура.Артикул '],
                    'Модель': row['Номенклатура'],
                    'Действие': action,
                    'Обоснование': reason,
                    'Маржа %': f"{row['Маржин-сть к средней цене']:.1f}",
                    'Дневная продажа': f"{row['Средняя дневная продажа']:.2f}"
                })
        
        rec_df = pd.DataFrame(recommendations)
        
        output_file = f"{output_dir}/Rekomendacii_{datetime.now().strftime('%Y%m%d')}.xlsx"
        if len(rec_df) > 0:
            rec_df.to_excel(output_file, sheet_name='Рекомендации', index=False)


# ЗАПУСК
if __name__ == '__main__':
    file_path = '/mnt/user-data/uploads/Планирование_стиралки.xlsx'
    
    skill = SkillZakupok(file_path)
    df = skill.process()
    skill.generate_reports()
    
    print("\n" + "="*50)
    print("✅ ГОТОВО!")
    print("="*50)
    print(f"\nВсе файлы сохранены в /mnt/user-data/outputs/")
    print("\nСоздано отчётов:")
    print("  1. PU_Zakupok_*.xlsx - Purchase Order")
    print("  2. Report_Nelikvidov_*.xlsx - Неликвиды")
    print("  3. Dashboard_*.xlsx - Дашборд")
    print("  4. ABC_Analiz_*.xlsx - ABC анализ")
    print("  5. Rekomendacii_*.xlsx - Рекомендации")

EOF
