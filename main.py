import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

def load_garage_data(file_path):
    """Загружает данные о гаражах"""
    df = pd.read_excel(file_path)
    df.columns = ['Гараж', 'Сумма', 'Дата оплаты']
    df['Дата оплаты'] = pd.to_datetime(df['Дата оплаты'])
    return df

def process_bank_statement(file_path):
    """Обрабатывает выписку из банка"""
    # Читаем все страницы выписки
    all_pages = []
    for page in range(1, 13):
        try:
            df = pd.read_excel(file_path, sheet_name=f'Sheet{page}', header=None)
            all_pages.append(df)
        except:
            continue
    
    # Находим строки с операциями
    operations = []
    for _, row in pd.concat(all_pages, ignore_index=True).iterrows():
        if len(row) >= 5 and isinstance(row[0], str) and '.' in row[0]:
            try:
                date = pd.to_datetime(row[0].split()[0], format='%d.%m.%Y', errors='coerce')
                amount = float(str(row[4]).replace(' ', '').replace(',', '.').replace('+', ''))
                operations.append({'date': date, 'amount': amount})
            except:
                continue
    
    return pd.DataFrame(operations)

def adjust_payment_date(date):
    """Корректирует дату оплаты"""
    if pd.isna(date):
        return None
    
    last_day = date + pd.offsets.MonthEnd(0)
    if date.day == 31 and last_day.day < 31:
        return last_day
    elif date.month == 2 and date.day in [29, 30] and last_day.day < 29:
        return last_day
    return date

def check_payment_status(garage_df, bank_df):
    """Проверяет статусы платежей"""
    results = []
    today = datetime.now().date()
    
    for _, row in garage_df.iterrows():
        garage = row['Гараж']
        amount = row['Сумма']
        date = adjust_payment_date(row['Дата оплаты'])
        
        if pd.isna(date):
            status = "Дата не указана"
        else:
            date = date.date()
            paid = bank_df[(np.isclose(bank_df['amount'], amount)) & (bank_df['date'].dt.date <= today)]
            
            if not paid.empty:
                status = f"Получен ({paid['date'].max().date()})"
            else:
                deadline = date + timedelta(days=3)
                status = "Срок не наступил" if today < date else "Ожидается" if today <= deadline else f"Просрочен (с {date})"
        
        results.append([garage, date, amount, status])
    
    return pd.DataFrame(results, columns=['Гараж', 'Дата оплаты', 'Сумма оплаты', 'Статус'])

def generate_report(result_df):
    """Генерирует Excel-отчёт"""
    report_name = f"Отчет_по_оплате_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    result_df.to_excel(report_name, index=False)
    return report_name

if __name__ == "__main__":
    print("Загрузка данных...")
    garage_df = load_garage_data('data/arenda.xlsx')
    bank_df = process_bank_statement('data/print2.xlsx')
    
    print("Проверка платежей...")
    result_df = check_payment_status(garage_df, bank_df)
    
    print("Генерация отчёта...")
    report_file = generate_report(result_df)
    
    print(f"\nОтчёт сохранён как: {report_file}")
    print("\nРезультаты:")
    print(result_df)
