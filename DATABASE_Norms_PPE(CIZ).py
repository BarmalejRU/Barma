import pandas as pd
import sqlite3

# Загрузка Excel-файла
file_path = r"D:\pythonProject1\NormaCIZ\ПРИКАЗЫ\ПРИКАЗ_767Н_ОБНОВЛЕН.xlsx"
df = pd.read_excel(file_path)

# Удалим строки, где нет ни профессии, ни СИЗ — это "пустышки"
df = df[~(df['Наименование профессий и должностей'].isna() &
          df['Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты'].isna())]

# Создание SQLite базы данных
conn = sqlite3.connect('norma_ciz.db')
cursor = conn.cursor()

cursor.execute('''
    CREATE TABLE IF NOT EXISTS normy_vydachi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        professiya TEXT,
        tip_sredstva TEXT,
        naimenovanie TEXT,
        edinitsa_izmereniya TEXT,
        kolichestvo TEXT
    )
''')

cursor.execute('DELETE FROM normy_vydachi')

# Вставка данных с учетом "одной профессии сверху"
tek_professiya = None

for _, row in df.iterrows():
    if pd.notna(row['Наименование профессий и должностей']):
        tek_professiya = row['Наименование профессий и должностей']
        # Пишем строку с профессией, но пустыми остальными полями
        cursor.execute('''
            INSERT INTO normy_vydachi (professiya, tip_sredstva, naimenovanie, edinitsa_izmereniya, kolichestvo)
            VALUES (?, ?, ?, ?, ?)
        ''', (tek_professiya, '', '', '', ''))
    elif pd.notna(row['Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты']):
        # Пишем строку с СИЗ, но без дублирования профессии
        cursor.execute('''
            INSERT INTO normy_vydachi (professiya, tip_sredstva, naimenovanie, edinitsa_izmereniya, kolichestvo)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            '',  # Профессия — пусто
            row['Тип средства защиты'],
            row['Наименование специальной одежды, специальной обуви и других средств индивидуальной защиты'],
            row['Нормы выдачи на год (период) (штуки, пары, комплекты, мл)'],
            row[df.columns[-1]]  # Количество
        ))

conn.commit()
conn.close()

print("База данных успешно создана: norma_ciz.db")







