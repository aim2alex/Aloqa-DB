import pandas as pd
import json
import os
import sys

def main():
    print("Starting data aggregation...")
    
    q1_file = "1.01-31.03.xlsx"
    q2_file = "1.04-30.06.xlsx"
    q1_mc_file = "1.01-31.03_MicroCredit.xlsx"
    q2_mc_file = "1.04-30.06_MicroCredit.xlsx"
    jira_file = "Jira показатели.xlsx"
    
    # Check that all files exist
    files = [q1_file, q2_file, q1_mc_file, q2_mc_file, jira_file]
    for f in files:
        if not os.path.exists(f):
            print(f"Error: Missing file in the workspace: {f}")
            sys.exit(1)
        
    print("Reading Q1 data...")
    df1 = pd.read_excel(q1_file)
    print("Reading Q2 data...")
    df2 = pd.read_excel(q2_file)
    
    print("Reading Q1 MicroCredit data...")
    df1_mc = pd.read_excel(q1_mc_file)
    print("Reading Q2 MicroCredit data...")
    df2_mc = pd.read_excel(q2_mc_file)
    
    print("Combining datasets...")
    df_all = pd.concat([df1, df2])
    df_all_mc = pd.concat([df1_mc, df2_mc])
    
    def get_stats(df, df_mc):
        total = len(df)
        incoming = int((df['Тип Вызова'] == 'Входящие').sum())
        outgoing = int((df['Тип Вызова'] == 'Исходящие').sum())
        
        # Types of appeals (exclude empty/undefined)
        types_s = df['Тип'].dropna()
        types_s = types_s[types_s != 'Не указан']
        types_s = types_s.value_counts()
        types = [{'name': str(k), 'value': int(v)} for k, v in types_s.items()]
        
        # Main Theme of appeals (exclude empty/undefined)
        themes_s = df['Основная Тема'].dropna()
        themes_s = themes_s[~themes_s.isin(['Не имеет данного атрибута', 'Не указано'])]
        themes_s = themes_s.value_counts()
        themes = [{'name': str(k), 'value': int(v)} for k, v in themes_s.items()]
        
        # Subthemes (exclude empty/undefined)
        subthemes_s = df['Подтема'].dropna()
        subthemes_s = subthemes_s[~subthemes_s.isin(['Не имеет данного атрибута', 'Не указано'])]
        subthemes_s = subthemes_s.value_counts()
        subthemes = [{'name': str(k), 'value': int(v)} for k, v in subthemes_s.items()]
        
        # Details of appeals / Substance of theme (exclude empty/undefined)
        details_s = df['Суть Темы'].dropna()
        details_s = details_s[~details_s.isin(['Не имеет данного атрибута', 'Не указано'])]
        details_s = details_s.value_counts()
        details = [{'name': str(k), 'value': int(v)} for k, v in details_s.items()]
        
        # Microcredit Dialing Tasks Stats
        total_mc = len(df_mc)
        completed_mc = int((df_mc['Статус Задачи'] == 'Выполнено').sum())
        refused_mc = int((df_mc['Статус Задачи'] == 'Отказ клиента').sum())
        
        # Monthly MicroCredit Stats
        df_mc_copy = df_mc.copy()
        df_mc_copy['Month'] = df_mc_copy['Дата Создания Задачи'].str[:7]
        months = sorted(df_mc_copy['Month'].dropna().unique())
        
        monthly_mc = []
        for month in months:
            df_m = df_mc_copy[df_mc_copy['Month'] == month]
            m_completed = int((df_m['Статус Задачи'] == 'Выполнено').sum())
            m_refused = int((df_m['Статус Задачи'] == 'Отказ клиента').sum())
            monthly_mc.append({
                'month': month,
                'completed': m_completed,
                'refused': m_refused
            })
        
        return {
            'total': total,
            'incoming': incoming,
            'outgoing': outgoing,
            'types': types,
            'themes': themes,
            'subthemes': subthemes,
            'details': details,
            # MicroCredit stats
            'total_mc': total_mc,
            'completed_mc': completed_mc,
            'refused_mc': refused_mc,
            'monthly_mc': monthly_mc
        }
        
    print("Aggregating statistics for Q1...")
    q1_stats = get_stats(df1, df1_mc)
    print("Aggregating statistics for Q2...")
    q2_stats = get_stats(df2, df2_mc)
    print("Aggregating statistics for combined datasets...")
    all_stats = get_stats(df_all, df_all_mc)
    
    print("Reading and aggregating RISK_Events data...")
    risk_file = "RISK_Events.xlsx"
    if not os.path.exists(risk_file):
        print(f"Error: Missing file in the workspace: {risk_file}")
        sys.exit(1)
        
    df_risk = pd.read_excel(risk_file, header=1)
    
    total_risk = len(df_risk)
    open_risk = int((df_risk['Статус события'] == 'OPEN').sum())
    confirmed_risk = int((df_risk['Статус события'] == 'CONFIRMED').sum())
    
    import re
    actual_loss_sum = 0.0
    potential_loss_sum = 0.0
    for val in df_risk['Убытки'].dropna():
        val_str = str(val)
        m_act = re.search(r'ACTUAL_LOSS:\s*([\d\.]+)\s*UZS', val_str)
        if m_act:
            actual_loss_sum += float(m_act.group(1))
        m_pot = re.search(r'POTENTIAL_LOSS:\s*([\d\.]+)\s*UZS', val_str)
        if m_pot:
            potential_loss_sum += float(m_pot.group(1))
            
    # Translation dictionary from Uzbek to Russian
    uz_to_ru = {
        # Categories
        'Ichki firibgarlik': 'Внутреннее мошенничество',
        'Elektr energiya taʼminotidagi uzulishlar': 'Сбои в электроснабжении',
        "To'lovlarni amalga oshirish va toʼlov topshiriqlarini taqdim etish bilan bog'liq risklar": 'Риски платежей и предоставления платежных поручений',
        'Biznes jarayoni dasturiy kommunikatsion tizimlaridagi kamchiliklar': 'Недостатки программно-коммуникационных систем бизнес-процессов',
        'Bankning mijozlari, mahsulotlari va ishbilarmonlik normalari bilan bogʼliq risklar ': 'Риски клиентов, продуктов и деловых стандартов банка',
        'Tashqi firibgarlik': 'Внешнее мошенничество',
        "Bank xodimlarining professional darajasining pastligi va e'tiborsizligi oqibatidagi xatoliklar": 'Ошибки сотрудников из-за низкой квалификации или халатности',
        "HKXKM va KXKMda kredit mahsuloti bo'yicha muddati o'tgan qarzdorliklar  (biroq soni bo'yicha 10 tadan kam bo'lmagan miqdorda): (Bank boshqaruving tegishli qarori bilan NPL limit o'rnatilmagan kredit mahsulotlari mustasno)": 'Просроченная задолженность по кредитам в РЦКУ и ЦКУ (не менее 10 шт., за исключением продуктов без лимитов NPL)',
        "Kredit ajratish jarayonini uzoq vaqt talab qilishi - Ariza KFO ga kiritilib ajratilgungacha bo'lgan vaqt": 'Долгое оформление кредита (от ввода заявки в КФО до выдачи)',
        "HKKXKM va KXKM da kredit mahsulotlari kesimida NPL bo'yicha o'rnatilgan limitlarni  chetlab o'tish": 'Обход установленных лимитов NPL по кредитным продуктам в РЦКУ и ЦКУ',
        "Ehtimoliy yo'qotishlarni qoplash uchun zaxira shakllantirish bo'yicha": 'Формирование резервов на покрытие возможных потерь',

        # Event Types
        "Kassada ortiqcha mablag' aniqlanishi1": 'Выявление излишков в кассе',
        'Elektr energiya taʼminotidagi uzulishlar': 'Сбои в электроснабжении',
        'Mijozlarni jalb qilish va hujjatlar yuritishdagi xatoliklar': 'Ошибки привлечения клиентов и ведения документации',
        "Tegishli hujjatlarsiz (to'lov hujjatlari, kredit hujjatlari va h.k.) yoki soxta hujjatlar bilan bank operatsiyalarini amalga oshirish": 'Проведение операций без документов или по поддельным документам',
        "Аxborotlarni oshkor etish bo'yicha talablarning buzilishi": 'Нарушение требований раскрытия информации',
        "Аktivlarni qasddan yo'q qilish, shikast yetkazish yoki o'zlashtirish": 'Умышленное уничтожение, повреждение или присвоение активов',
        'Mijozlarni identifikatsiyalash va tekshirishga doir talablarning bajarilmasligi': 'Невыполнение требований идентификации и верификации клиентов',
        "Soxta hujjatlar yordamida mablag'larning o'zlashtirilishi": 'Присвоение средств по поддельным документам',
        'Garov shartnomalarining tuzmaslik yoki garov mulklarining qonunchilikka zid ravishda chiqarib yuborish': 'Несоставление договоров залога или незаконный вывод залогового имущества',
        'Аsosiy bank tizimidag(ABT)i muammolar ( jumladan Colvir, KFO, XD va xk)': 'Проблемы в АБС (включая Colvir, КФО, ХД и др.)',
        "Xodimlar tomonidan ABTga ma'lumotlar kiritishdagi xatoliklar": 'Ошибки ввода данных в АБС сотрудниками',
        'Xodimning firibgarlik qilishi': 'Мошенничество со стороны сотрудника',
        "Maxfiy ma'lumotlardan noto'g'ri foydalanishi": 'Ненадлежащее использование конфиденциальной информации',
        '1-30 kungacha - 15% va unda ortiq': 'От 1 до 30 дней - 15% и более',
        '90 kundan va undan ortiq - 1,5 % va undan ortiq': '90 дней и более - 1,5% и более',
        '31-60 kungacha - 10% va undan ortiq': 'От 31 до 60 дней - 10% и более',
        '61-90 kungacha - 5% va undan ortiq;': 'От 61 до 90 дней - 5% и более',
        'Jismoniy shaxslar uchun: ipoteka krediti 7 ish kuni': 'Для физлиц: ипотечный кредит 7 рабочих дней',
        "Jismoniy shaxslar uchun: ta'lim krediti 5 ish kuni": 'Для физлиц: образовательный кредит 5 рабочих дней',
        'Jismoniy shaxslar uchun: istemol krediti 5 ish kuni': 'Для физлиц: потребительский кредит 5 рабочих дней',
        "Biznes mijozlar uchun - 10 kun va undan ortiq (Kredit qo'mitasi hamda Kuzatuv Kengashi muhokamasida ko'rib chiqiladigan arizalarda muddat yana ham cho'zilishi mumkin)": 'Для бизнес-клиентов - 10 дней и более (при рассмотрении КК и НС срок может быть увеличен)',
        "Tizimlararo ma'lumotlar nomuvofiqligi/mahsulot identifikatsiyasi buzilishi": 'Межсистемное несоответствие данных / нарушение идентификации продукта',
        "HKXKM va KXKM xodimi tomonidan ehtimoliy yo'qotishlarni qoplash uchun shakllantirilgan zaxiralarni qaytarishga urunib ko'rishi": 'Попытка возврата резервов на покрытие потерь сотрудником РЦКУ/ЦКУ',
        "HKXKM va KXKM xodimi tomonidan ehtimoliy yo'qotishlarni qoplash uchun shakllantirilgan zaxiralarni qaytarishga urinib ko'rishi": 'Попытка возврата резервов на покрытие потерь сотрудником РЦКУ/ЦКУ',
        "HKXKM va KXKM xodimi tomonidan ehtimolly yo'qotishlarni qoplash uchun shakllantirilgan zaxiralarni qaytarishga urunib ko'rishi": 'Попытка возврата резервов на покрытие потерь сотрудником РЦКУ/ЦКУ',
        "HKXKM va KXKM xodimi tomonidan ehtimolly yo'qotishlarni qoplash uchun shakllantirilgan zaxiralarni qaytarishga urinib ko'ris": 'Попытка возврата резервов на покрытие потерь сотрудником РЦКУ/ЦКУ',
        "Mijozlar hisobvaraqlarini nomaqbul tarzda boshqarish, savdoga oid sheriklar va yetkazib beruvchilar bilan bog'liq muammolar": 'Ненадлежащее управление счетами клиентов, проблемы с партнерами и поставщиками',
        "Xodimlar tomonidan bank plastik kartochkalriga hizmat ko'rsatuvchi uskunalarni sozlash va ma'lumotlar kiritilishida  xatoliklarga yo'l qo'yilishi": 'Ошибки сотрудников при настройке оборудования карт и вводе данных',
        "Avtomatlashtirilgan limit nazoratining buzilishi/biznes cheklovlarni chetlab o'tish": 'Нарушение автоматического контроля лимитов / обход бизнес-ограничений',
        'Dasturiy kommunikatsion tizimlardagi muammolar ( jumladan Directum, Outlook, Jira va xk)': 'Проблемы в системах коммуникации (Directum, Outlook, Jira и др.)',
        "Bank va mijozlar hisob raqamlari bilan ishlashda buxgalteriya xatoliklariga yo'l qo'yilishi": 'Бухгалтерские ошибки при работе со счетами банка и клиентов',
        "Mijozlarga oid hujjatlarning to'liq emasligi, Markaziy bank va qonunchilikka muvofiq boshqa organlarga hisobotlarning taqdim etilmasligi": 'Неполнота документов клиентов, непредоставление отчетов в ЦБ и другие органы'
    }

    branch_translation = {
        'Namangan HKXKM': 'Наманганский РЦКУ',
        'Samarqand HKXKM': 'Самаркандский РЦКУ',
        'Amaliyot HKXKM': 'Операционный РЦКУ',
        'Oltiariq KXKM': 'Алтыарыкский ЦКУ',
        'Jarqoʻrgʻon KXKM': 'Джаркурганский ЦКУ',
        'Xalqobod KXKM': 'Халкабадский ЦКУ',
        'Keles KXKM': 'Келеский ЦКУ',
        'Xorazm HKXKM': 'Хорезмский РЦКУ',
        'Andijon HKXKM': 'Андижанский РЦКУ',
        'Qoraqalpogiston HKXKM': 'Караakalпакский РЦКУ',
        'Qoraqalpogiston HKXKM': 'Каракалпакский РЦКУ',
        'Surxondaryo HKXKM': 'Сурхандарьинский РЦКУ',
        'Gʻuzor KXKM': 'Гузарский ЦКУ',
        'Bogʻot KXKM': 'Багатский ЦКУ',
        'Orzu KXKM': 'ЦКУ Орзу',
        'Toshkent viloyati HKXKM': 'Ташкентский областной РЦКУ',
        'Angor KXKM': 'Ангорский ЦКУ',
        'Afrosiyob KXKM': 'Афрасиабский ЦКУ',
        'Samoniy KXKM': 'Саманиский ЦКУ',
        'Guliston KXKM': 'Гулистанский ЦКУ',
        'Nukus KXKM': 'Нукусский ЦКУ'
    }

    def translate_branch(name):
        name_str = str(name)
        if name_str in branch_translation:
            return branch_translation[name_str]
        return name_str.replace('HKXKM', 'РЦКУ').replace('KXKM', 'ЦКУ')

    top_cat_raw = str(df_risk['Категория риска'].value_counts().index[0]) if len(df_risk['Категория риска']) > 0 else 'Нет данных'
    top_cat = uz_to_ru.get(top_cat_raw.strip(), top_cat_raw.strip())
    
    # categories
    cat_s = df_risk['Категория риска'].value_counts()
    categories_risk = [{'name': uz_to_ru.get(str(k).strip(), str(k).strip()), 'value': int(v)} for k, v in cat_s.items()]
    
    # types
    type_s = df_risk['Тип события'].value_counts()
    types_risk = [{'name': uz_to_ru.get(str(k).strip(), str(k).strip()), 'value': int(v)} for k, v in type_s.items()]
    
    # branches
    branch_s = df_risk['Филиал (KXKM)'].value_counts()
    branches_risk = [{'name': translate_branch(k), 'value': int(v)} for k, v in branch_s.items()]
    
    # sources
    source_s = df_risk['Источник'].value_counts()
    source_mapping = {
        'PROCESS': 'Процессы',
        'PERSONNEL': 'Персонал',
        'EXTERNAL': 'Внешние факторы',
        'SYSTEM': 'Системы',
        'NOT_DEFINED': 'Не определено'
    }
    sources_risk = [{'name': source_mapping.get(str(k), str(k)), 'value': int(v)} for k, v in source_s.items()]
    
    # levels
    level_s = df_risk['Уровень риска'].value_counts()
    level_mapping = {
        'LOW': 'Низкий',
        'MEDIUM': 'Средний',
        'HIGH': 'Высокий'
    }
    levels_risk = [{'name': level_mapping.get(str(k), str(k)), 'value': int(v)} for k, v in level_s.items()]
    
    auto_risk_tasks = int((df_risk['Создал'] == 'service-account-kfo').sum())
    manual_risk_tasks = int((df_risk['Создал'] != 'service-account-kfo').sum())

    risk_stats = {
        'total': total_risk,
        'open': open_risk,
        'confirmed': confirmed_risk,
        'actual_loss': actual_loss_sum,
        'potential_loss': potential_loss_sum,
        'top_category': top_cat,
        'categories': categories_risk,
        'event_types': types_risk,
        'branches': branches_risk,
        'sources': sources_risk,
        'risk_levels': levels_risk,
        'auto_tasks': auto_risk_tasks,
        'manual_tasks': manual_risk_tasks
    }
    
    print("Reading and aggregating Loans data...")
    loans_file = "Loans_True.xlsx"
    if not os.path.exists(loans_file):
        print(f"Error: Missing file in the workspace: {loans_file}")
        sys.exit(1)
        
    # Openpyxl has issues with strict XML, so we convert it to transitional XML on-the-fly
    import zipfile
    temp_loans_file = "temp_loans_transitional.xlsx"
    with zipfile.ZipFile(loans_file, 'r') as zin:
        with zipfile.ZipFile(temp_loans_file, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith('.xml') or item.filename.endswith('.rels'):
                    data = data.replace(b'http://purl.oclc.org/ooxml/spreadsheetml/main', b'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                    data = data.replace(b'http://purl.oclc.org/ooxml/officeDocument/relationships', b'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                    data = data.replace(b'http://purl.oclc.org/drawingml/main', b'http://schemas.openxmlformats.org/drawingml/2006/main')
                zout.writestr(item, data)
    
    df_loans = pd.read_excel(temp_loans_file)
    if os.path.exists(temp_loans_file):
        os.remove(temp_loans_file)
        
    df_loans['date_parsed'] = pd.to_datetime(df_loans['application_date'], errors='coerce')
    df_loans['month'] = df_loans['date_parsed'].dt.month

    # Product renaming mapping (normalize whitespace to single space)
    df_loans['product_code'] = df_loans['product_code'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    rename_map = {
        'Онлайн микрозайм - Qulay': 'Qulay',
        'Онлайн микрозайм - Универсал': 'Universal',
        'Онлайн микрозайм - BOJXONA': 'Bojxona',
        'Микрозайм для самозанятых со страховым полисом (Для молодежных проектов)': 'Yoshlar со страховкой',
        'Микрозайм для самозанятых с поручителем (Для молодежных проектов)': 'Yoshlar с поручителем',
        'Веб микрозайм для самозанятых со страховым полисом (Biznes imkon)': 'Biznes Imkon со страховкой',
        'Веб микрозайм для самозанятых с поручителем (Biznes imkon)': 'Biznes Imkon с поручителем',
        'Веб микрозайм для самозанятых со страховым полисом (Бизнесга биринчи кадам - 1) с верификацией': 'Biznesga Birinchi Qadam со страховкой',
        'Веб микрозайм для самозанятых с поручителем (Бизнесга биринчи кадам - 1) с верификацией': 'Biznesga Birinchi Qadam с поручителем',
        'Qulay WEB': 'Qulay Web',
        'UNIVERSAL WEB': 'Universal Web'
    }
    df_loans['product_code'] = df_loans['product_code'].replace(rename_map)

    # Product code groups for summing Qulay and Universal WEB
    qulay_codes = ['Qulay Web']
    universal_codes = ['Universal Web']

    def get_loans_stats(sub_df):
        apps = sub_df[sub_df['status'] != 'Просмотр условий продукта, без заявки']
        total_apps = len(apps)
        success_apps = (sub_df['status'] == 'Успешная выдача').sum()
        reject_apps = (sub_df['status'] == 'Отказ банка по скорингу').sum()
        reject_client_apps = (sub_df['status'] == 'Отказ клиента').sum()
        
        # Sum of amount for successful loans
        total_issued = float(sub_df[sub_df['status'] == 'Успешная выдача']['amount'].sum())
        
        # Sum of successful loans for Qulay WEB and UNIVERSAL WEB from loans file
        total_prc = float(sub_df[sub_df['product_code'].isin(qulay_codes) & (sub_df['status'] == 'Успешная выдача')]['amount'].sum())
        total_prc_pen = float(sub_df[sub_df['product_code'].isin(universal_codes) & (sub_df['status'] == 'Успешная выдача')]['amount'].sum())
        
        # Product counts (excluding "Просмотр условий продукта, без заявки")
        prod_counts = apps['product_code'].value_counts()
        products_list = [{'name': str(k), 'value': int(v)} for k, v in prod_counts.items()]
        
        # Status counts (excluding "Просмотр условий продукта, без заявки")
        status_counts = sub_df[sub_df['status'] != 'Просмотр условий продукта, без заявки']['status'].value_counts()
        statuses_list = [{'name': str(k), 'value': int(v)} for k, v in status_counts.items()]
        
        # Monthly grouped bar chart
        monthly_data = {
            'months': ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь'],
            'success': [],
            'reject': [],
            'reject_client': []
        }
        for m in range(1, 7):
            m_df = sub_df[sub_df['month'] == m]
            m_success = (m_df['status'] == 'Успешная выдача').sum()
            m_reject = (m_df['status'] == 'Отказ банка по скорингу').sum()
            m_reject_client = (m_df['status'] == 'Отказ клиента').sum()
            monthly_data['success'].append(int(m_success))
            monthly_data['reject'].append(int(m_reject))
            monthly_data['reject_client'].append(int(m_reject_client))
            
        # Monthly amount aggregation for successful loans
        monthly_issued_data = {
            'months': ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь'],
            'amounts': []
        }
        for m in range(1, 7):
            m_df = sub_df[(sub_df['month'] == m) & (sub_df['status'] == 'Успешная выдача')]
            monthly_issued_data['amounts'].append(float(m_df['amount'].sum()))
            
        # Monthly amount aggregation for Qulay and Universal
        monthly_income_qulay_universal = {
            'months': ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь'],
            'prc': [],
            'prc_pen': [],
            'total': []
        }
        for m in range(1, 7):
            m_df = sub_df[sub_df['month'] == m]
            m_qulay = float(m_df[m_df['product_code'].isin(qulay_codes) & (m_df['status'] == 'Успешная выдача')]['amount'].sum())
            m_universal = float(m_df[m_df['product_code'].isin(universal_codes) & (m_df['status'] == 'Успешная выдача')]['amount'].sum())
            monthly_income_qulay_universal['prc'].append(m_qulay)
            monthly_income_qulay_universal['prc_pen'].append(m_universal)
            monthly_income_qulay_universal['total'].append(m_qulay + m_universal)

        # Comparison of total issued amounts for Qulay and Universal
        income_list = [
            {
                'code': 'Qulay WEB',
                'prc': total_prc,
                'prc_pen': 0.0,
                'total': total_prc
            },
            {
                'code': 'UNIVERSAL WEB',
                'prc': 0.0,
                'prc_pen': total_prc_pen,
                'total': total_prc_pen
            }
        ]
        # Monthly product counts
        unique_products = list(df_loans['product_code'].value_counts().index)
        monthly_prod_data = {
            'months': ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь'],
            'products': []
        }
        for prod in unique_products:
            prod_values = []
            for m in range(1, 7):
                m_df = sub_df[sub_df['month'] == m]
                count = int((m_df['product_code'] == prod).sum())
                prod_values.append(count)
            monthly_prod_data['products'].append({
                'name': prod,
                'values': prod_values
            })
            
        return {
            'total_apps': int(total_apps),
            'success': int(success_apps),
            'reject': int(reject_apps),
            'reject_client': int(reject_client_apps),
            'total_issued': total_issued,
            'total_prc': total_prc,
            'total_prc_pen': total_prc_pen,
            'products': products_list,
            'statuses': statuses_list,
            'monthly': monthly_data,
            'monthly_issued': monthly_issued_data,
            'monthly_income': monthly_income_qulay_universal,
            'income': income_list,
            'monthly_products': monthly_prod_data
        }

    loans_stats = {
        'all': get_loans_stats(df_loans),
        'q1': get_loans_stats(df_loans[df_loans['month'] <= 3]),
        'q2': get_loans_stats(df_loans[df_loans['month'] >= 4])
    }
    
    print("Reading and aggregating Jira data...")
    df_jira = pd.read_excel(jira_file, header=None)

    def get_jira_stats(selected_months):
        cats_list = []
        total_tasks = 0
        for r in range(1, 6):
            name = str(df_jira.iloc[r, 0]).strip()
            val = sum(float(df_jira.iloc[r, m]) for m in selected_months)
            cats_list.append({'name': name, 'value': int(val)})
            total_tasks += val
            
        sys_list = []
        for r in range(9, 17):
            name = str(df_jira.iloc[r, 0]).strip()
            val = sum(float(df_jira.iloc[r, m]) for m in selected_months)
            sys_list.append({'name': name, 'value': int(val)})
            
        month_names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь']
        months_labels = [month_names[m-1] for m in selected_months]
        
        monthly_cats = []
        for r in range(1, 6):
            name = str(df_jira.iloc[r, 0]).strip()
            vals = [int(float(df_jira.iloc[r, m])) for m in selected_months]
            monthly_cats.append({'name': name, 'values': vals})
            
        monthly_systems = []
        for r in range(9, 17):
            name = str(df_jira.iloc[r, 0]).strip()
            vals = [int(float(df_jira.iloc[r, m])) for m in selected_months]
            monthly_systems.append({'name': name, 'values': vals})
            
        monthly_hours = []
        for r in range(1, 7):
            name = str(df_jira.iloc[r, 8]).strip()
            if name == 'Без категории':
                continue
            vals = []
            for m in selected_months:
                val_str = str(df_jira.iloc[r, m + 8]).strip()
                if '%' in val_str:
                    val = float(val_str.replace('%', '').strip())
                else:
                    try:
                        val = float(val_str)
                        if val <= 1.0 and val > 0:
                            val = val * 100.0
                    except ValueError:
                        val = 0.0
                vals.append(round(val, 2))
            monthly_hours.append({'name': name, 'values': vals})
            
        avg_tasks = total_tasks / len(selected_months)
        
        main_cat = ''
        main_cat_val = -1
        for item in cats_list:
            if item['value'] > main_cat_val:
                main_cat_val = item['value']
                main_cat = item['name']
        main_cat_pct = (main_cat_val / total_tasks * 100) if total_tasks > 0 else 0
        
        return {
            'total_tasks': int(total_tasks),
            'avg_tasks': round(avg_tasks, 1),
            'main_category': main_cat,
            'main_category_pct': round(main_cat_pct, 1),
            'categories': cats_list,
            'systems': sys_list,
            'months': months_labels,
            'monthly_categories': monthly_cats,
            'monthly_systems': monthly_systems,
            'monthly_hours': monthly_hours
        }

    jira_stats = {
        'all': get_jira_stats([1, 2, 3, 4, 5, 6]),
        'q1': get_jira_stats([1, 2, 3]),
        'q2': get_jira_stats([4, 5, 6])
    }
    
    stats_data = {
        'all': all_stats,
        'q1': q1_stats,
        'q2': q2_stats,
        'risks': risk_stats,
        'loans': loans_stats,
        'jira': jira_stats
    }
    
    print("Generating HTML dashboard index.html...")
    
    html_template = """<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Дашборд Call-центра CRM</title>
    
    <!-- Google Fonts: Inter & Outfit -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    
    <style>
        :root {
            --bg-gradient-start: #0b0f19;
            --bg-gradient-end: #171b2d;
            --card-bg: rgba(26, 32, 53, 0.65);
            --card-border: rgba(255, 255, 255, 0.07);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --accent-purple: #8b5cf6;
            --accent-purple-glow: rgba(139, 92, 246, 0.4);
            --accent-emerald: #10b981;
            --accent-emerald-glow: rgba(16, 185, 129, 0.4);
            --accent-amber: #f59e0b;
            --accent-amber-glow: rgba(245, 158, 11, 0.4);
            --accent-blue: #3b82f6;
            --accent-blue-glow: rgba(59, 130, 246, 0.4);
            --accent-rose: #ef4444;
            --accent-rose-glow: rgba(239, 68, 68, 0.4);
            
            --transition-speed: 0.3s;
        }

        body.light-theme {
            --bg-gradient-start: #f1f5f9;
            --bg-gradient-end: #cbd5e1;
            --card-bg: rgba(255, 255, 255, 0.7);
            --card-border: rgba(0, 0, 0, 0.08);
            --text-primary: #0f172a;
            --text-secondary: #475569;
            --accent-purple-glow: rgba(139, 92, 246, 0.15);
            --accent-emerald-glow: rgba(16, 185, 129, 0.15);
            --accent-amber-glow: rgba(245, 158, 11, 0.15);
            --accent-blue-glow: rgba(59, 130, 246, 0.15);
            --accent-rose-glow: rgba(239, 68, 68, 0.15);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background-color: var(--bg-gradient-start);
            background-image: radial-gradient(circle at 10% 20%, rgba(99, 102, 241, 0.1) 0%, transparent 40%),
                              radial-gradient(circle at 90% 80%, rgba(139, 92, 246, 0.1) 0%, transparent 40%),
                              linear-gradient(to bottom right, var(--bg-gradient-start), var(--bg-gradient-end));
            background-attachment: fixed;
            color: var(--text-primary);
            min-height: 100vh;
            line-height: 1.5;
            padding: 2rem 1.5rem;
            overflow-x: hidden;
            transition: background-color var(--transition-speed) ease, color var(--transition-speed) ease;
        }

        body.light-theme {
            background-image: radial-gradient(circle at 10% 20%, rgba(99, 102, 241, 0.05) 0%, transparent 40%),
                              radial-gradient(circle at 90% 80%, rgba(139, 92, 246, 0.05) 0%, transparent 40%),
                              linear-gradient(to bottom right, var(--bg-gradient-start), var(--bg-gradient-end));
        }

        /* Custom Scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: rgba(255, 255, 255, 0.02);
        }
        body.light-theme ::webkit-scrollbar-track {
            background: rgba(0, 0, 0, 0.02);
        }
        ::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.15);
            border-radius: 4px;
        }
        body.light-theme ::webkit-scrollbar-thumb {
            background: rgba(0, 0, 0, 0.15);
        }
        ::-webkit-scrollbar-thumb:hover {
            background: rgba(255, 255, 255, 0.3);
        }
        body.light-theme ::webkit-scrollbar-thumb:hover {
            background: rgba(0, 0, 0, 0.3);
        }

        /* Container Layout */
        .container {
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 2.25rem;
            padding: 0 1.5rem;
        }

        /* App Header Layout */
        .app-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1rem 1.75rem;
            background: rgba(15, 23, 42, 0.4);
            border: 1px solid var(--card-border);
            border-radius: 20px;
            backdrop-filter: blur(20px);
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
            transition: all var(--transition-speed) ease;
            position: sticky;
            top: 1rem;
            z-index: 100;
            gap: 1rem;
        }
        body.light-theme .app-header {
            background: rgba(255, 255, 255, 0.7);
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.05);
        }
        
        .header-left {
            display: flex;
            align-items: center;
        }
        
        .app-logo {
            display: block;
            height: 30px;
            width: auto;
        }
        
        .logo-text-fill {
            fill: #f8fafc;
            transition: fill var(--transition-speed) ease;
        }
        body.light-theme .logo-text-fill {
            fill: #1e293b;
        }
        
        .header-center {
            display: flex;
            align-items: center;
        }
        
        .main-tabs-container {
            display: flex;
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--card-border);
            border-radius: 9999px;
            padding: 0.3rem;
            gap: 0.25rem;
            backdrop-filter: blur(10px);
            transition: background-color var(--transition-speed), border-color var(--transition-speed);
        }
        body.light-theme .main-tabs-container {
            background: rgba(255, 255, 255, 0.8);
        }
        
        .main-tab-btn {
            background: transparent;
            border: none;
            color: var(--text-secondary);
            padding: 0.6rem 1.5rem;
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 0.95rem;
            border-radius: 9999px;
            cursor: pointer;
            transition: all var(--transition-speed) cubic-bezier(0.4, 0, 0.2, 1);
            user-select: none;
            white-space: nowrap;
        }
        
        .main-tab-btn:hover {
            color: #fff;
            background: #078AE2;
        }
        
        .main-tab-btn.active {
            color: #fff;
            background: #0099FF;
            box-shadow: 0 4px 15px rgba(0, 153, 255, 0.4);
        }
        
        .header-right {
            display: flex;
            align-items: center;
        }

        /* Panel Header (inside each dashboard) */
        .panel-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1.5rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            padding-bottom: 1.5rem;
            margin-top: 1rem;
        }
        body.light-theme .panel-header {
            border-bottom: 1px solid rgba(0, 0, 0, 0.05);
        }

        .main-panel {
            display: flex;
            flex-direction: column;
            gap: 2.25rem;
        }

        .header-title h1 {
            font-family: 'Outfit', sans-serif;
            font-weight: 800;
            font-size: 2.2rem;
            background: linear-gradient(135deg, #fff 30%, #a5b4fc 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -0.02em;
        }
        body.light-theme .header-title h1 {
            background: linear-gradient(135deg, #0f172a 30%, #4f46e5 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-title h1.title-risks {
            background: linear-gradient(135deg, #fff 30%, #ec4899 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        body.light-theme .header-title h1.title-risks {
            background: linear-gradient(135deg, #0f172a 30%, #db2777 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-title p {
            color: var(--text-secondary);
            font-size: 0.95rem;
            margin-top: 0.25rem;
        }

        /* Tabs Navigation */
        .tabs-container {
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--card-border);
            border-radius: 9999px;
            padding: 0.35rem;
            display: flex;
            gap: 0.25rem;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
            backdrop-filter: blur(10px);
            transition: background-color var(--transition-speed), border-color var(--transition-speed);
        }
        body.light-theme .tabs-container {
            background: rgba(255, 255, 255, 0.8);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        }

        .tab-btn {
            background: transparent;
            border: none;
            color: var(--text-secondary);
            padding: 0.75rem 1.75rem;
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 0.95rem;
            border-radius: 9999px;
            cursor: pointer;
            transition: all var(--transition-speed) cubic-bezier(0.4, 0, 0.2, 1);
            user-select: none;
            white-space: nowrap;
        }

        .tab-btn:hover {
            color: var(--text-primary);
            background: rgba(255, 255, 255, 0.03);
        }
        body.light-theme .tab-btn:hover {
            background: rgba(0, 0, 0, 0.03);
        }

        .tab-btn.active {
            color: #fff;
            background: linear-gradient(135deg, var(--accent-purple), #6366f1);
            box-shadow: 0 4px 15px var(--accent-purple-glow);
        }

        /* Theme Toggle Button */
        .theme-toggle-btn {
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--card-border);
            border-radius: 50%;
            width: 48px;
            height: 48px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            color: var(--text-secondary);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
            backdrop-filter: blur(10px);
            transition: all var(--transition-speed) ease;
        }

        .theme-toggle-btn:hover {
            color: var(--text-primary);
            border-color: rgba(255, 255, 255, 0.15);
            transform: scale(1.05);
        }

        body.light-theme .theme-toggle-btn {
            background: rgba(255, 255, 255, 0.8);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            color: var(--text-secondary);
        }

        body.light-theme .theme-toggle-btn:hover {
            color: var(--text-primary);
            border-color: rgba(0, 0, 0, 0.15);
        }

        /* Cards Grid (KPI) */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 1.5rem;
        }

        .card {
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            backdrop-filter: blur(16px);
            border-radius: 20px;
            padding: 1.75rem;
            position: relative;
            overflow: hidden;
            transition: transform var(--transition-speed) ease, border-color var(--transition-speed) ease, box-shadow var(--transition-speed) ease, background-color var(--transition-speed) ease;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
        }
        body.light-theme .card {
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.05);
        }

        .card:hover {
            transform: translateY(-4px);
            border-color: rgba(255, 255, 255, 0.15);
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.35);
        }
        body.light-theme .card:hover {
            border-color: rgba(0, 0, 0, 0.15);
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.12);
        }

        .card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 4px;
        }

        .card-total::before { background: linear-gradient(90deg, #8b5cf6, #3b82f6); }
        .card-incoming::before { background: linear-gradient(90deg, #10b981, #059669); }
        .card-outgoing::before { background: linear-gradient(90deg, #f59e0b, #d97706); }
        
        .card-mc-total::before { background: linear-gradient(90deg, #3b82f6, #06b6d4); }
        .card-mc-completed::before { background: linear-gradient(90deg, #10b981, #059669); }
        .card-mc-refused::before { background: linear-gradient(90deg, #ef4444, #f43f5e); }
        .card-loans-yellow::before { background: linear-gradient(90deg, #eab308, #facc15); }
        .card-loans-orange::before { background: linear-gradient(90deg, #f97316, #ea580c); }

        .card-total-risks::before { background: linear-gradient(90deg, #8b5cf6, #3b82f6); }
        .card-open-risks::before { background: linear-gradient(90deg, #f59e0b, #eab308); }
        .card-confirmed-risks::before { background: linear-gradient(90deg, #10b981, #059669); }
        .card-actual-losses::before { background: linear-gradient(90deg, #ef4444, #f43f5e); }
        .card-potential-losses::before { background: linear-gradient(90deg, #3b82f6, #06b6d4); }
        .card-top-category::before { background: linear-gradient(90deg, #ec4899, #8b5cf6); }
        .card-top-category {
            grid-column: span 2;
        }
        @media (max-width: 900px) {
            .card-top-category {
                grid-column: span 1;
            }
        }

        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.25rem;
        }

        .card-title {
            color: var(--text-secondary);
            font-size: 0.9rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .card-icon {
            width: 42px;
            height: 42px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .card-total .card-icon {
            background: rgba(139, 92, 246, 0.15);
            color: var(--accent-purple);
        }

        .card-incoming .card-icon {
            background: rgba(16, 185, 129, 0.15);
            color: var(--accent-emerald);
        }

        .card-outgoing .card-icon {
            background: rgba(245, 158, 11, 0.15);
            color: var(--accent-amber);
        }

        .card-mc-total .card-icon {
            background: rgba(59, 130, 246, 0.15);
            color: var(--accent-blue);
        }

        .card-mc-completed .card-icon {
            background: rgba(16, 185, 129, 0.15);
            color: var(--accent-emerald);
        }

        .card-mc-refused .card-icon {
            background: rgba(239, 68, 68, 0.15);
            color: var(--accent-rose);
        }

        .card-total-risks .card-icon {
            background: rgba(139, 92, 246, 0.15);
            color: var(--accent-purple);
        }

        .card-open-risks .card-icon {
            background: rgba(245, 158, 11, 0.15);
            color: var(--accent-amber);
        }

        .card-confirmed-risks .card-icon {
            background: rgba(16, 185, 129, 0.15);
            color: var(--accent-emerald);
        }

        .card-actual-losses .card-icon {
            background: rgba(239, 68, 68, 0.15);
            color: var(--accent-rose);
        }

        .card-potential-losses .card-icon {
            background: rgba(59, 130, 246, 0.15);
            color: var(--accent-blue);
        }

        .card-top-category .card-icon {
            background: rgba(236, 72, 153, 0.15);
            color: #ec4899;
        }

        .card-body {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }

        .card-value {
            font-family: 'Outfit', sans-serif;
            font-size: 2.6rem;
            font-weight: 700;
            color: var(--text-primary);
            letter-spacing: -0.01em;
            line-height: 1;
        }

        .card-meta {
            font-size: 0.85rem;
            color: var(--text-secondary);
            display: flex;
            align-items: center;
            gap: 0.4rem;
        }

        .meta-pct {
            font-weight: 700;
            padding: 0.15rem 0.4rem;
            border-radius: 6px;
        }

        .card-incoming .meta-pct {
            background: rgba(16, 185, 129, 0.15);
            color: var(--accent-emerald);
        }

        .card-outgoing .meta-pct {
            background: rgba(245, 158, 11, 0.15);
            color: var(--accent-amber);
        }

        /* Charts Grid (Pie Section) */
        .charts-row-two-custom {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 1.5rem;
        }
        .charts-row-two {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 1.5rem;
        }
        .span-3 {
            grid-column: span 3;
        }
        .span-2 {
            grid-column: span 2;
        }
        .span-1 {
            grid-column: span 1;
        }
        @media (max-width: 1024px) {
            .charts-row-two-custom, .charts-row-two {
                grid-template-columns: 1fr;
            }
            .span-3, .span-2, .span-1 {
                grid-column: span 1;
            }
        }

        .chart-card {
            min-height: 480px;
            display: flex;
            flex-direction: column;
        }

        .chart-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.5rem;
        }

        .chart-card-title {
            font-family: 'Outfit', sans-serif;
            font-size: 1.25rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .chart-container-pie {
            position: relative;
            flex-grow: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            max-height: 380px;
        }

        /* Subtheme section (horizontal bar chart) */
        .chart-card-bar {
            display: flex;
            flex-direction: column;
            gap: 1.5rem;
        }

        .bar-controls {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
        }

        .search-wrapper {
            position: relative;
            flex-grow: 1;
            max-width: 400px;
        }

        .search-input {
            width: 100%;
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 0.6rem 1rem 0.6rem 2.5rem;
            color: var(--text-primary);
            font-family: 'Inter', sans-serif;
            font-size: 0.9rem;
            outline: none;
            transition: border-color var(--transition-speed), box-shadow var(--transition-speed), background-color var(--transition-speed);
        }

        body.light-theme .search-input {
            background: rgba(255, 255, 255, 0.8);
            color: #0f172a;
        }

        .search-input:focus {
            border-color: var(--accent-purple);
            box-shadow: 0 0 0 3px rgba(139, 92, 246, 0.25);
        }
        
        body.light-theme .search-input:focus {
            border-color: var(--accent-purple);
            box-shadow: 0 0 0 3px rgba(139, 92, 246, 0.15);
        }

        .search-icon {
            position: absolute;
            left: 0.85rem;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-secondary);
            pointer-events: none;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .slider-wrapper {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            font-size: 0.85rem;
            color: var(--text-secondary);
            background: rgba(15, 23, 42, 0.4);
            padding: 0.4rem 1rem;
            border-radius: 10px;
            border: 1px solid var(--card-border);
            transition: background-color var(--transition-speed);
        }
        body.light-theme .slider-wrapper {
            background: rgba(255, 255, 255, 0.6);
        }

        .top-slider {
            -webkit-appearance: none;
            appearance: none;
            width: 100px;
            height: 6px;
            border-radius: 3px;
            background: rgba(255, 255, 255, 0.1);
            outline: none;
        }
        body.light-theme .top-slider {
            background: rgba(0, 0, 0, 0.1);
        }

        .top-slider::-webkit-slider-thumb {
            -webkit-appearance: none;
            appearance: none;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background: var(--accent-purple);
            cursor: pointer;
            box-shadow: 0 0 8px var(--accent-purple-glow);
            transition: transform 0.1s;
        }

        #top-slider-details::-webkit-slider-thumb {
            background: var(--accent-emerald);
            box-shadow: 0 0 8px var(--accent-emerald-glow);
        }

        #risk-branches-slider::-webkit-slider-thumb {
            background: #10b981;
            box-shadow: 0 0 8px rgba(16, 185, 129, 0.4);
        }

        .top-slider::-webkit-slider-thumb:hover {
            transform: scale(1.2);
        }

        .chart-container-bar {
            position: relative;
            width: 100%;
            height: 480px;
        }

        /* Teams Grid & Card */
        .teams-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 1.5rem;
            margin-top: 0.5rem;
        }
        @media (max-width: 1024px) {
            .teams-grid {
                grid-template-columns: repeat(2, 1fr);
            }
        }
        @media (max-width: 600px) {
            .teams-grid {
                grid-template-columns: 1fr;
            }
        }
        .team-card {
            padding: 1.25rem 1.5rem !important;
            min-height: 110px;
            border-left-width: 4px !important;
            border-left-style: solid !important;
        }
        .team-card::before {
            display: none !important;
        }

        /* Quarters Grid */
        .quarters-grid {
            display: grid;
            grid-template-columns: 1.35fr 1fr;
            gap: 1.5rem;
            margin-top: 0.5rem;
        }
        @media (max-width: 1200px) {
            .quarters-grid {
                grid-template-columns: 1fr;
            }
        }

        /* Section Dividers */
        .section-divider {
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            margin: 0.5rem 0 -0.5rem 0;
            padding-top: 1.5rem;
            transition: border-color var(--transition-speed);
        }
        body.light-theme .section-divider {
            border-top: 1px solid rgba(0, 0, 0, 0.05);
        }

        /* SVG Icon styling */
        .svg-icon {
            width: 20px;
            height: 20px;
            fill: none;
            stroke: currentColor;
            stroke-width: 2;
            stroke-linecap: round;
            stroke-linejoin: round;
        }

        /* Footer */
        footer {
            text-align: center;
            color: var(--text-secondary);
            font-size: 0.85rem;
            margin-top: 1rem;
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            padding-top: 1.5rem;
        }
        body.light-theme footer {
            border-top: 1px solid rgba(0, 0, 0, 0.05);
        }

        /* Glow effects in background */
        .glow-circle {
            position: absolute;
            width: 400px;
            height: 400px;
            border-radius: 50%;
            filter: blur(120px);
            z-index: -1;
            pointer-events: none;
            opacity: 0.45;
            transition: opacity var(--transition-speed) ease;
        }
        .glow-1 {
            background: radial-gradient(circle, var(--accent-purple) 0%, transparent 70%);
            top: -100px;
            left: -100px;
        }
        .glow-2 {
            background: radial-gradient(circle, var(--accent-blue) 0%, transparent 70%);
            bottom: -100px;
            right: -100px;
        }
        body.light-theme .glow-1 {
            opacity: 0.2;
        }
        body.light-theme .glow-2 {
            opacity: 0.2;
        }
    </style>
</head>
<body>
    <div class="glow-circle glow-1"></div>
    <div class="glow-circle glow-2"></div>

    <div class="container">
        <!-- Global App Header -->
        <header class="app-header">
            <div class="header-left">
                <!-- Inline logo.svg -->
                <svg class="app-logo" width="168" height="30" viewBox="0 0 168 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <g clip-path="url(#clip0_21521_278)">
                        <path d="M11.0691 6.96516L22.1381 13.175V3.75H0V13.175L11.0691 6.96516Z" fill="#0099FF"/>
                        <path d="M11.0691 9.40625L0 15.616V25.9284H22.1381V15.616L11.0691 9.40625Z" fill="#0099FF"/>
                        <path d="M43.3023 18.7408L40.3798 9.31944L37.4311 18.7408H43.3023ZM44.2057 21.5883H36.5279L35.1729 25.9263H31.5069L37.5639 7.29688H43.2756L49.3593 25.9263H45.5604L44.2057 21.5883Z" class="logo-text-fill" fill="#333538"/>
                        <path d="M54.6189 25.9273H51.0588V6.33984H54.6189V25.9273Z" class="logo-text-fill" fill="#333538"/>
                        <path d="M64.1027 23.4784C66.3608 23.4784 67.7424 21.6155 67.7424 18.9541C67.7424 16.2928 66.3608 14.4296 64.1027 14.4296C61.8709 14.4296 60.4896 16.2928 60.4896 18.9541C60.4896 21.6155 61.8712 23.4784 64.1027 23.4784ZM64.1027 11.6621C68.4594 11.6621 71.249 14.4296 71.249 18.9541C71.249 23.4784 68.4597 26.2461 64.1027 26.2461C59.7457 26.2461 56.9564 23.4784 56.9564 18.9541C56.9564 14.4296 59.7457 11.6621 64.1027 11.6621Z" class="logo-text-fill" fill="#333538"/>
                        <path d="M79.9889 23.4784C82.1671 23.4784 83.4954 21.6421 83.4954 18.9541C83.4954 16.2659 82.1671 14.4296 79.9889 14.4296C77.7834 14.4296 76.4556 16.2659 76.4556 18.9541C76.4556 21.6421 77.7834 23.4784 79.9889 23.4784ZM78.8993 11.6621C80.839 11.6621 82.6187 12.6202 83.602 14.5364V11.9815H86.8961V31.2496H83.336V23.7181C82.3266 25.3945 80.6795 26.2464 78.873 26.2464C75.4726 26.2464 72.9221 23.5582 72.9221 18.9541C72.9221 14.3501 75.4726 11.6621 78.8993 11.6621Z" class="logo-text-fill" fill="#333538"/>
                        <path d="M95.0779 23.7714C97.3093 23.7714 98.7708 22.1481 98.7708 20.1254V19.5133L95.5829 19.8326C93.8557 19.9924 92.8197 20.6574 92.8197 21.9351C92.8197 22.9995 93.537 23.7714 95.0779 23.7714ZM95.1311 17.7036L98.7708 17.3307V16.6391C98.7708 15.2019 97.947 14.1903 96.0875 14.1903C94.573 14.1903 93.4042 14.8557 93.2979 16.2928H90.0571C90.1633 12.833 92.9526 11.6621 96.0875 11.6621C100.019 11.6621 102.33 13.525 102.33 17.384V25.9271H99.0097V23.3988C98.1068 25.1553 96.3796 26.2464 93.9356 26.2464C90.8007 26.2464 89.2865 24.4633 89.2865 22.1214C89.2865 19.327 91.518 18.0494 95.1311 17.7036Z" class="logo-text-fill" fill="#333538"/>
                        <path d="M117.26 20.4442C117.26 18.8741 116.224 17.8358 113.966 17.8358H109.37V23.0258H113.78C116.171 23.0258 117.26 22.0407 117.26 20.4442ZM109.37 15.0684H113.621C115.666 15.0684 116.702 14.0837 116.702 12.5933C116.702 11.1293 115.666 10.1978 113.568 10.1978H109.37V15.0683L109.37 15.0684ZM120.979 20.9499C120.979 24.3297 117.898 25.9266 114.126 25.9266H105.651V7.29688H113.94C117.5 7.29688 120.422 8.60084 120.422 11.981C120.422 14.6689 118.589 16.1327 114.923 16.3455C118.296 16.5318 120.979 17.8358 120.979 20.9499Z" fill="#0099FF"/>
                        <path d="M128.418 23.7714C130.649 23.7714 132.111 22.1481 132.111 20.1254V19.5133L128.923 19.8326C127.196 19.9924 126.159 20.6574 126.159 21.9351C126.159 22.9995 126.877 23.7714 128.418 23.7714ZM128.471 17.7036L132.111 17.3307V16.6391C132.111 15.2019 131.287 14.1903 129.427 14.1903C127.913 14.1903 126.744 14.8557 126.638 16.2928H123.397C123.503 12.833 126.293 11.6621 129.427 11.6621C133.359 11.6621 135.67 13.525 135.67 17.384V25.9271H132.349V23.3988C131.447 25.1553 129.719 26.2464 127.275 26.2464C124.141 26.2464 122.626 24.4633 122.626 22.1214C122.626 19.327 124.858 18.0494 128.471 17.7036Z" fill="#0099FF"/>
                        <path d="M138.539 11.9814H141.86V14.4298C142.763 12.6201 144.463 11.6621 146.722 11.6621C149.91 11.6621 151.53 13.5517 151.53 17.0382V25.927H147.997V18.1291C147.997 15.8673 147.279 14.5896 145.34 14.5896C143.188 14.5896 142.099 16.1863 142.099 18.5815V25.927H138.539V11.9814Z" fill="#0099FF"/>
                        <path d="M163.511 25.9274L159.234 20.2854H157.959V25.9274H154.399V6.33984H157.959V17.5174H159.26L163.298 11.9819H167.496L161.784 18.6087L167.894 25.9274H163.511Z" fill="#0099FF"/>
                    </g>
                    <defs>
                        <clipPath id="clip0_21521_278">
                            <rect width="168" height="30" fill="white"/>
                        </clipPath>
                    </defs>
                </svg>
            </div>
            
            <div class="header-center">
                <div class="main-tabs-container">
                    <button class="main-tab-btn active" id="main-tab-jira" onclick="switchMainTab('jira')">Общие показатели Jira</button>
                    <button class="main-tab-btn" id="main-tab-loans" onclick="switchMainTab('loans')">Кредиты</button>
                    <button class="main-tab-btn" id="main-tab-callcenter" onclick="switchMainTab('callcenter')">Call-центр</button>
                    <button class="main-tab-btn" id="main-tab-risks" onclick="switchMainTab('risks')">События операционных рисков</button>
                </div>
            </div>
            
            <div class="header-right">
                <button id="theme-toggle" class="theme-toggle-btn" onclick="toggleTheme()" aria-label="Переключить тему">
                    <!-- Sun Icon (visible in dark mode) -->
                    <svg class="svg-icon sun-icon" viewBox="0 0 24 24">
                        <circle cx="12" cy="12" r="5"></circle>
                        <line x1="12" y1="1" x2="12" y2="3"></line>
                        <line x1="12" y1="21" x2="12" y2="23"></line>
                        <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line>
                        <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line>
                        <line x1="1" y1="12" x2="3" y2="12"></line>
                        <line x1="21" y1="12" x2="23" y2="12"></line>
                        <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line>
                        <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line>
                    </svg>
                    <!-- Moon Icon (visible in light mode) -->
                    <svg class="svg-icon moon-icon" viewBox="0 0 24 24" style="display: none;">
                        <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path>
                    </svg>
                </button>
            </div>
        </header>

        <!-- Panel Loans -->
        <div id="panel-loans" class="main-panel" style="display: none;">
            <div class="panel-header">
                <div class="header-title">
                    <h1 class="title-loans" style="background: linear-gradient(135deg, var(--text-primary) 30%, #0099FF 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">Аналитический Дашборд Кредитных продуктов</h1>
                    <p>Анализ кредитных заявок • 1 Квартал - 2 Квартал 2026 г.</p>
                </div>
                
                <div class="tabs-container">
                    <button class="tab-btn active" onclick="switchTab('all', this)">Все</button>
                    <button class="tab-btn" onclick="switchTab('q1', this)">1 Квартал</button>
                    <button class="tab-btn" onclick="switchTab('q2', this)">2 Квартал</button>
                </div>
            </div>



            <!-- Row 1: KPI Cards (Всего заявок, Успешная выдача, Отказ банка по скорингу, Отказ клиента) -->
            <section class="kpi-grid">
                <!-- Total Applications -->
                <div class="card card-mc-total">
                    <div class="card-header">
                        <span class="card-title">Всего заявок</span>
                        <div class="card-icon" style="background: rgba(59, 130, 246, 0.15); color: #3b82f6;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                <polyline points="14 2 14 8 20 8"></polyline>
                                <line x1="16" y1="13" x2="8" y2="13"></line>
                                <line x1="16" y1="17" x2="8" y2="17"></line>
                                <polyline points="10 9 9 9 8 9"></polyline>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-loans-total">0</span>
                        <div class="card-meta">
                            <span>Зарегистрированные заявки</span>
                        </div>
                    </div>
                </div>

                <!-- Successful Issuance -->
                <div class="card card-mc-completed">
                    <div class="card-header">
                        <span class="card-title">Успешная выдача</span>
                        <div class="card-icon" style="background: rgba(16, 185, 129, 0.15); color: #10b981;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <polyline points="20 6 9 17 4 12"></polyline>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-loans-success">0</span>
                        <div class="card-meta">
                            <span>Успешно выданные кредиты</span>
                        </div>
                    </div>
                </div>

                <!-- Scoring Rejection -->
                <div class="card card-mc-refused">
                    <div class="card-header">
                        <span class="card-title">Отказ банка<br>по скорингу</span>
                        <div class="card-icon" style="background: rgba(239, 68, 68, 0.15); color: #ef4444;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <line x1="18" y1="6" x2="6" y2="18"></line>
                                <line x1="6" y1="6" x2="18" y2="18"></line>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-loans-reject-scoring">0</span>
                        <div class="card-meta">
                            <span>Автоматический отказ скоринга</span>
                        </div>
                    </div>
                </div>

                <!-- Client Rejection -->
                <div class="card card-loans-yellow">
                    <div class="card-header">
                        <span class="card-title">Отказ клиента</span>
                        <div class="card-icon" style="background: rgba(245, 158, 11, 0.15); color: #f59e0b;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                <circle cx="8.5" cy="7" r="4"></circle>
                                <line x1="18" y1="8" x2="23" y2="13"></line>
                                <line x1="23" y1="8" x2="18" y2="13"></line>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-loans-reject-client">0</span>
                        <div class="card-meta">
                            <span>Отказ со стороны клиента</span>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 2: Pie Charts (В разрезе продуктов, В разрезе статусов) -->
            <section class="charts-row-two-custom">
                <div class="card chart-card span-2" style="min-height: 600px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Количество заявок в разрезе продуктов</span>
                    </div>
                    <div class="chart-container-pie" style="position: relative; display: flex; flex-direction: column; align-items: center; justify-content: flex-start; min-height: 500px; width: 100%; padding: 10px 0;">
                        <div style="position: relative; width: 320px; height: 320px; flex-shrink: 0; margin-bottom: 20px;">
                            <canvas id="chart-loans-products"></canvas>
                            <div class="doughnut-center-card" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--card-bg); border: 1px solid var(--card-border); box-shadow: 0 8px 32px rgba(0,0,0,0.12); width: 180px; height: 180px; border-radius: 50%; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; pointer-events: none; z-index: 5;">
                                <span style="font-size: 0.85rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;">Всего</span>
                                <span id="doughnut-products-center-value" style="font-size: 2rem; font-weight: 700; color: var(--text-primary); margin-top: 2px; line-height: 1;">0</span>
                                <span style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 2px;">заявок</span>
                            </div>
                        </div>
                        <div id="doughnut-products-legend" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(170px, 1fr)); gap: 8px 12px; width: 100%; max-height: 220px; overflow-y: auto; padding: 5px 10px;">
                            <!-- Rendered dynamically -->
                        </div>
                    </div>
                </div>

                <div class="card chart-card span-2" style="min-height: 600px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Количество заявок в разрезе статусов</span>
                    </div>
                    <div class="chart-container-pie" style="position: relative; display: flex; flex-direction: column; align-items: center; justify-content: flex-start; min-height: 500px; width: 100%; padding: 10px 0;">
                        <div style="position: relative; width: 320px; height: 320px; flex-shrink: 0; margin-bottom: 20px;">
                            <canvas id="chart-loans-statuses"></canvas>
                            <div class="doughnut-center-card" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--card-bg); border: 1px solid var(--card-border); box-shadow: 0 8px 32px rgba(0,0,0,0.12); width: 180px; height: 180px; border-radius: 50%; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; pointer-events: none; z-index: 5;">
                                <span style="font-size: 0.85rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;">Всего</span>
                                <span id="doughnut-center-value" style="font-size: 2rem; font-weight: 700; color: var(--text-primary); margin-top: 2px; line-height: 1;">0</span>
                                <span style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 2px;">заявок</span>
                            </div>
                        </div>
                        <div id="doughnut-statuses-legend" style="display: flex; flex-wrap: wrap; justify-content: center; gap: 12px 24px; width: 100%; max-height: 220px; overflow-y: auto; padding: 5px 10px;">
                            <!-- Rendered dynamically -->
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 2.75: Products by Month Grouped Bar Chart -->
            <section class="charts-row-full">
                <div class="card chart-card" style="min-height: 480px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Статистика продуктов по месяцам</span>
                    </div>
                    <div class="chart-container-full" style="height: 380px;">
                        <canvas id="chart-loans-monthly-products"></canvas>
                    </div>
                </div>
            </section>

            <!-- Row 3: Horizontal Grouped Bar Chart (По месяцам) -->
            <section class="charts-row-full">
                <div class="card chart-card" style="min-height: 420px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Количество заявок по месяцам</span>
                    </div>
                    <div class="chart-container-full" style="height: 320px;">
                        <canvas id="chart-loans-monthly"></canvas>
                    </div>
                </div>
            </section>

            <!-- Row 3.5: Monthly Amount Chart -->
            <section class="charts-row-full">
                <div class="card chart-card" style="min-height: 420px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Сумма выданных кредитов по месяцам</span>
                    </div>
                    <div class="chart-container-full" style="height: 320px;">
                        <canvas id="chart-loans-monthly-amount"></canvas>
                    </div>
                </div>
            </section>

            <!-- Row 4 & 5: Income Charts removed -->
        </div><!-- Closing #panel-loans -->

        <!-- Panel Call Center -->
        <div id="panel-callcenter" class="main-panel" style="display: none;">
            <div class="panel-header">
                <div class="header-title">
                    <h1>Аналитический Дашборд Call-центра</h1>
                    <p>Анализ обращений клиентов • 1 Квартал - 2 Квартал 2026 г.</p>
                </div>
                
                <div class="tabs-container">
                    <button class="tab-btn active" onclick="switchTab('all', this)">Все</button>
                    <button class="tab-btn" onclick="switchTab('q1', this)">1 Квартал</button>
                    <button class="tab-btn" onclick="switchTab('q2', this)">2 Квартал</button>
                </div>
            </div>

        <!-- KPI Grid -->
        <section class="kpi-grid">
            <!-- Total Appeals -->
            <div class="card card-total">
                <div class="card-header">
                    <span class="card-title">Всего обращений</span>
                    <div class="card-icon">
                        <svg class="svg-icon" viewBox="0 0 24 24">
                            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
                        </svg>
                    </div>
                </div>
                <div class="card-body">
                    <span class="card-value" id="kpi-total">0</span>
                    <div class="card-meta">
                        <span>За отчетный период</span>
                    </div>
                </div>
            </div>

            <!-- Incoming Appeals -->
            <div class="card card-incoming">
                <div class="card-header">
                    <span class="card-title">Входящие</span>
                    <div class="card-icon">
                        <svg class="svg-icon" viewBox="0 0 24 24">
                            <polyline points="22 17 13.5 8.5 8.5 13.5 2 7"></polyline>
                            <polyline points="16 17 22 17 22 11"></polyline>
                        </svg>
                    </div>
                </div>
                <div class="card-body">
                    <span class="card-value" id="kpi-incoming">0</span>
                    <div class="card-meta">
                        <span class="meta-pct" id="pct-incoming">0%</span>
                        <span>от общего числа</span>
                    </div>
                </div>
            </div>

            <!-- Outgoing Appeals -->
            <div class="card card-outgoing">
                <div class="card-header">
                    <span class="card-title">Исходящие</span>
                    <div class="card-icon">
                        <svg class="svg-icon" viewBox="0 0 24 24">
                            <polyline points="22 7 13.5 15.5 8.5 10.5 2 17"></polyline>
                            <polyline points="16 7 22 7 22 13"></polyline>
                        </svg>
                    </div>
                </div>
                <div class="card-body">
                    <span class="card-value" id="kpi-outgoing">0</span>
                    <div class="card-meta">
                        <span class="meta-pct" id="pct-outgoing">0%</span>
                        <span>от общего числа</span>
                    </div>
                </div>
            </div>
        </section>

        <!-- Pie Charts Row -->
        <section class="charts-row-two">
            <!-- Appeal Types Pie -->
            <div class="card chart-card">
                <div class="chart-card-header">
                    <span class="chart-card-title">По типам обращений</span>
                </div>
                <div class="chart-container-pie">
                    <canvas id="chart-types"></canvas>
                </div>
            </div>

            <!-- Main Theme Pie -->
            <div class="card chart-card">
                <div class="chart-card-header">
                    <span class="chart-card-title">По Основным темам</span>
                </div>
                <div class="chart-container-pie">
                    <canvas id="chart-themes"></canvas>
                </div>
            </div>
        </section>

        <!-- Horizontal Bar Chart Section (Subtheme) -->
        <section class="card chart-card-bar">
            <div class="chart-card-header bar-controls">
                <span class="chart-card-title">Рейтинг подтем обращений</span>
                
                <div style="display: flex; gap: 1rem; align-items: center; flex-grow: 1; justify-content: flex-end; flex-wrap: wrap;">
                    <div class="search-wrapper">
                        <span class="search-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24" style="width: 16px; height: 16px;">
                                <circle cx="11" cy="11" r="8"></circle>
                                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
                            </svg>
                        </span>
                        <input type="text" id="subtheme-search" class="search-input" placeholder="Поиск подтемы..." oninput="handleSearch()">
                    </div>
                    
                    <div class="slider-wrapper">
                        <span>Показать Топ:</span>
                        <input type="range" id="top-slider" class="top-slider" min="5" max="30" value="15" oninput="handleSliderChange(this.value)">
                        <span id="slider-val" style="font-weight: 700; width: 20px; text-align: right; color: var(--accent-purple);">15</span>
                    </div>
                </div>
            </div>
            
            <div class="chart-container-bar">
                <canvas id="chart-subthemes"></canvas>
            </div>
        </section>

        <!-- Substance of Theme Horizontal Bar Chart Section (green theme) -->
        <section class="card chart-card-bar">
            <div class="chart-card-header bar-controls">
                <span class="chart-card-title">Рейтинг сути обращений</span>
                
                <div style="display: flex; gap: 1rem; align-items: center; flex-grow: 1; justify-content: flex-end; flex-wrap: wrap;">
                    <div class="search-wrapper">
                        <span class="search-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24" style="width: 16px; height: 16px;">
                                <circle cx="11" cy="11" r="8"></circle>
                                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
                            </svg>
                        </span>
                        <input type="text" id="details-search" class="search-input" placeholder="Поиск сути обращения..." oninput="handleDetailsSearch()">
                    </div>
                    
                    <div class="slider-wrapper">
                        <span>Показать Топ:</span>
                        <input type="range" id="top-slider-details" class="top-slider" min="5" max="30" value="15" oninput="handleDetailsSliderChange(this.value)">
                        <span id="slider-val-details" style="font-weight: 700; width: 20px; text-align: right; color: var(--accent-emerald);">15</span>
                    </div>
                </div>
            </div>
            
            <div class="chart-container-bar">
                <canvas id="chart-details"></canvas>
            </div>
        </section>

        <!-- Microcredit Section -->
        <div class="section-divider">
            <h2 style="font-family: 'Outfit', sans-serif; font-weight: 700; font-size: 1.6rem; margin-bottom: 0.5rem; color: var(--text-primary);">Задачи на обзвон по микрозайму</h2>
            <p style="color: var(--text-secondary); font-size: 0.95rem;">Аналитика результатов кампании обзвона клиентов</p>
        </div>

        <section class="kpi-grid">
            <!-- Total MicroCredit Tasks -->
            <div class="card card-mc-total">
                <div class="card-header">
                    <span class="card-title">Всего задач на обзвон</span>
                    <div class="card-icon">
                        <svg class="svg-icon" viewBox="0 0 24 24">
                            <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path>
                            <rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect>
                        </svg>
                    </div>
                </div>
                <div class="card-body">
                    <span class="card-value" id="kpi-mc-total">0</span>
                    <div class="card-meta">
                        <span>Запланировано к обзвону</span>
                    </div>
                </div>
            </div>

            <!-- Completed Tasks -->
            <div class="card card-mc-completed">
                <div class="card-header">
                    <span class="card-title">Выполнено</span>
                    <div class="card-icon">
                        <svg class="svg-icon" viewBox="0 0 24 24">
                            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                            <polyline points="22 4 12 14.01 9 11.01"></polyline>
                        </svg>
                    </div>
                </div>
                <div class="card-body">
                    <span class="card-value" id="kpi-mc-completed">0</span>
                    <div class="card-meta">
                        <span class="meta-pct" id="pct-mc-completed" style="background: rgba(16, 185, 129, 0.15); color: var(--accent-emerald);">0%</span>
                        <span>успешных звонков</span>
                    </div>
                </div>
            </div>

            <!-- Refused Tasks -->
            <div class="card card-mc-refused">
                <div class="card-header">
                    <span class="card-title">Отказ клиента</span>
                    <div class="card-icon">
                        <svg class="svg-icon" viewBox="0 0 24 24">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="15" y1="9" x2="9" y2="15"></line>
                            <line x1="9" y1="9" x2="15" y2="15"></line>
                        </svg>
                    </div>
                </div>
                <div class="card-body">
                    <span class="card-value" id="kpi-mc-refused">0</span>
                    <div class="card-meta">
                        <span class="meta-pct" id="pct-mc-refused" style="background: rgba(239, 68, 68, 0.15); color: var(--accent-rose);">0%</span>
                        <span>отказов от предложения</span>
                    </div>
                </div>
            </div>
        </section>

        <!-- Monthly Grouped Bar Chart -->
        <section class="card chart-card-bar">
            <div class="chart-card-header">
                <span class="chart-card-title">Динамика звонков по месяцам</span>
            </div>
            <div class="chart-container-bar" style="height: 380px;">
                <canvas id="chart-mc-monthly"></canvas>
            </div>
        </section>
        </div><!-- Closing #panel-callcenter -->

        <!-- Panel Operational Risks -->
        <div id="panel-risks" class="main-panel" style="display: none;">
            <div class="panel-header">
                <div class="header-title">
                    <h1 class="title-risks">Аналитический Дашборд Событий операционных рисков</h1>
                    <p>Анализ и мониторинг событий операционных рисков</p>
                </div>
            </div>

            <!-- Row 1: KPI Cards (Всего событий, Открытых событий, Подтвержденных событий) -->
            <section class="kpi-grid">
                <!-- Total Events -->
                <div class="card card-total-risks">
                    <div class="card-header">
                        <span class="card-title">Всего событий</span>
                        <div class="card-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24">
                                <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                                <line x1="9" y1="9" x2="15" y2="9"></line>
                                <line x1="9" y1="13" x2="15" y2="13"></line>
                                <line x1="9" y1="17" x2="15" y2="17"></line>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-risk-total">0</span>
                        <div class="card-meta">
                            <span>Зарегистрировано инцидентов</span>
                        </div>
                    </div>
                </div>

                <!-- Open Events -->
                <div class="card card-open-risks">
                    <div class="card-header">
                        <span class="card-title">Открытых событий</span>
                        <div class="card-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24">
                                <circle cx="12" cy="12" r="10"></circle>
                                <polyline points="12 6 12 12 16 14"></polyline>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-risk-open">0</span>
                        <div class="card-meta">
                            <span>В статусе обработки (OPEN)</span>
                        </div>
                    </div>
                </div>

                <!-- Confirmed Events -->
                <div class="card card-confirmed-risks">
                    <div class="card-header">
                        <span class="card-title">Подтвержденных событий</span>
                        <div class="card-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24">
                                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                                <polyline points="22 4 12 14.01 9 11.01"></polyline>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-risk-confirmed">0</span>
                        <div class="card-meta">
                            <span>Подтверждено (CONFIRMED)</span>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 2: KPI Cards (Общая сумма убытков (UZS), Потенциальная сумма риска (UZS), Топ категория риска) -->
            <section class="kpi-grid">
                <!-- Total Actual Losses -->
                <div class="card card-actual-losses">
                    <div class="card-header">
                        <span class="card-title">Общая сумма убытков</span>
                        <div class="card-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24">
                                <line x1="12" y1="1" x2="12" y2="23"></line>
                                <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-risk-actual" style="font-size: 2.1rem;">0</span>
                        <div class="card-meta">
                            <span>UZS • Фактические убытки</span>
                        </div>
                    </div>
                </div>


                <!-- Top Category -->
                <div class="card card-top-category">
                    <div class="card-header">
                        <span class="card-title">Топ категория риска</span>
                        <div class="card-icon">
                            <svg class="svg-icon" viewBox="0 0 24 24">
                                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <div id="kpi-risk-top-category" style="font-size: 1.125rem; line-height: 1.35; font-weight: 700; color: var(--text-primary);">
                            Загрузка...
                        </div>
                        <div class="card-meta" style="margin-top: 0.4rem;">
                            <span>Чаще всего встречающаяся</span>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 3: Pie Charts (Топ категории рисков, Тип событий) -->
            <section class="charts-row-two">
                <div class="card chart-card">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Категории рисков</span>
                    </div>
                    <div class="chart-container-pie">
                        <canvas id="chart-risk-categories"></canvas>
                    </div>
                </div>

                <div class="card chart-card">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Тип событий</span>
                    </div>
                    <div class="chart-container-pie">
                        <canvas id="chart-risk-event-types"></canvas>
                    </div>
                </div>
            </section>

            <!-- Row 4: Branch Events horizontal bar chart -->
            <section class="card chart-card-bar">
                <div class="chart-card-header bar-controls">
                    <span class="chart-card-title" id="risk-branches-title">Количество событий по филиалам (Топ-15)</span>
                    
                    <div class="slider-wrapper">
                        <span>Показать Топ:</span>
                        <input type="range" id="risk-branches-slider" class="top-slider" min="5" max="92" value="15" oninput="handleRiskBranchesSliderChange(this.value)">
                        <span id="risk-branches-slider-val" style="font-weight: 700; width: 20px; text-align: right; color: #10b981;">15</span>
                    </div>
                </div>
                <div class="chart-container-bar" id="risk-branches-container" style="position: relative; width: 100%; height: 480px;">
                    <canvas id="chart-risk-branches"></canvas>
                </div>
            </section>

            <!-- Row 5: Sources and Risk Levels -->
            <section class="charts-row-two">
                <div class="card chart-card" style="min-height: 400px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Источники события</span>
                    </div>
                    <div class="chart-container-pie" style="max-height: 300px;">
                        <canvas id="chart-risk-sources"></canvas>
                    </div>
                </div>

                <div class="card chart-card" style="min-height: 400px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Количество событий по уровню риска</span>
                    </div>
                    <div class="chart-container-pie" style="max-height: 300px;">
                        <canvas id="chart-risk-levels"></canvas>
                    </div>
                </div>
            </section>

            <!-- Row 6: Auto vs Manual Tasks -->
            <section class="kpi-grid">
                <!-- Automatic Tasks -->
                <div class="card card-mc-total">
                    <div class="card-header">
                        <span class="card-title">Автоматические задачи</span>
                        <div class="card-icon" style="background: rgba(59, 130, 246, 0.15); color: #3b82f6;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <circle cx="12" cy="12" r="3"></circle>
                                <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-risk-auto-tasks">0</span>
                        <div class="card-meta">
                            <span>Создано service-account-kfo (98.1%)</span>
                        </div>
                    </div>
                </div>

                <!-- Manual Tasks -->
                <div class="card card-open-risks">
                    <div class="card-header">
                        <span class="card-title">Ручные задачи</span>
                        <div class="card-icon" style="background: rgba(245, 158, 11, 0.15); color: #f59e0b;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                                <circle cx="12" cy="7" r="4"></circle>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-risk-manual-tasks">0</span>
                        <div class="card-meta">
                            <span>Создано пользователями вручную (1.9%)</span>
                        </div>
                    </div>
                </div>
            </section>
        </div><!-- Closing #panel-risks -->

        <!-- Panel Jira -->
        <div id="panel-jira" class="main-panel">
            <div class="panel-header">
                <div class="header-title">
                    <h1 class="title-jira" style="background: linear-gradient(135deg, var(--text-primary) 30%, #3b82f6 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">Общие показатели Jira</h1>
                    <p>Анализ выполнения задач и распределения трудозатрат • Январь - Июнь 2026 г.</p>
                </div>
                
                <div class="tabs-container">
                    <button class="tab-btn active" onclick="switchTab('all', this)">Все</button>
                    <button class="tab-btn" onclick="switchTab('q1', this)">1 Квартал</button>
                    <button class="tab-btn" onclick="switchTab('q2', this)">2 Квартал</button>
                </div>
            </div>

            <!-- Row 1: KPI Cards -->
            <section class="kpi-grid">
                <!-- Total Tasks -->
                <div class="card card-total">
                    <div class="card-header">
                        <span class="card-title">Всего задач</span>
                        <div class="card-icon" style="background: rgba(59, 130, 246, 0.15); color: #3b82f6;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                                <polyline points="22 4 12 14.01 9 11.01"></polyline>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-jira-total-tasks">0</span>
                        <div class="card-meta">
                            <span>Выполненные задачи за период</span>
                        </div>
                    </div>
                </div>

                <!-- Average Tasks per Month -->
                <div class="card card-loans-orange">
                    <div class="card-header">
                        <span class="card-title">Среднее в месяц</span>
                        <div class="card-icon" style="background: rgba(249, 115, 22, 0.15); color: #f97316;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <circle cx="12" cy="12" r="10"></circle>
                                <polyline points="12 6 12 12 16 14"></polyline>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-jira-avg-tasks">0</span>
                        <div class="card-meta">
                            <span>Задач в среднем ежемесячно</span>
                        </div>
                    </div>
                </div>

                <!-- Main Direction -->
                <div class="card card-loans-yellow">
                    <div class="card-header">
                        <span class="card-title">Основное направление</span>
                        <div class="card-icon" style="background: rgba(234, 179, 8, 0.15); color: #eab308;">
                            <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                            </svg>
                        </div>
                    </div>
                    <div class="card-body">
                        <span class="card-value" id="kpi-jira-main-category" style="font-size: 1.9rem; line-height: 2.4rem; font-weight: 700; height: 44px; display: flex; align-items: center; justify-content: flex-start; text-align: left; color: var(--text-primary);">Поддержка</span>
                        <div class="card-meta">
                            <span id="kpi-jira-main-category-pct">0% от общего числа</span>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 1.5: Teams Section -->
            <section style="display: flex; flex-direction: column; gap: 1rem;">
                <div style="font-family: 'Outfit', sans-serif; font-size: 1.5rem; font-weight: 700; color: var(--text-primary);">Команды</div>
                <div class="teams-grid">
                    <!-- Camunda -->
                    <div class="card team-card" style="border-left-color: #3b82f6;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Camunda">Camunda</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(59, 130, 246, 0.15); color: #3b82f6; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">4</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Devops -->
                    <div class="card team-card" style="border-left-color: #10b981;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Devops">Devops</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(16, 185, 129, 0.15); color: #10b981; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">2</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Front -->
                    <div class="card team-card" style="border-left-color: #f59e0b;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Front">Front</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(245, 158, 11, 0.15); color: #f59e0b; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">6</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Абс -->
                    <div class="card team-card" style="border-left-color: #ef4444;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Абс">Абс</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(239, 68, 68, 0.15); color: #ef4444; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">4</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Колл-центр / Амл / Риски -->
                    <div class="card team-card" style="border-left-color: #8b5cf6;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Колл-центр / Амл / Риски">Колл-центр / Амл / Риски</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(139, 92, 246, 0.15); color: #8b5cf6; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">5</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Кфо / Залоги / Проблемные Кредиты -->
                    <div class="card team-card" style="border-left-color: #06b6d4;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Кфо / Залоги / Проблемные Кредиты">Кфо / Залоги / Проблемные Кредиты</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(6, 182, 212, 0.15); color: #06b6d4; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">4</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Пластиковые Карты -->
                    <div class="card team-card" style="border-left-color: #ec4899;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Пластиковые Карты">Пластиковые Карты</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(236, 72, 153, 0.15); color: #ec4899; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">3</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                    <!-- Скоринг -->
                    <div class="card team-card" style="border-left-color: #f97316;">
                        <div style="font-size: 0.95rem; font-weight: 700; color: var(--text-primary); margin-bottom: 12px; font-family: 'Outfit', sans-serif; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Скоринг">Скоринг</div>
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="background: rgba(249, 115, 22, 0.15); color: #f97316; width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 18px; height: 18px;">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                    <circle cx="9" cy="7" r="4"></circle>
                                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
                                    <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
                                </svg>
                            </div>
                            <div style="display: flex; align-items: baseline; gap: 4px;">
                                <span style="font-size: 1.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">4</span>
                                <span style="font-size: 0.9rem; color: var(--text-secondary);">чел</span>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 1.7: Completed Tasks (Quarterly filterable) -->
            <section style="display: flex; flex-direction: column; gap: 1rem; margin-top: 1.5rem;">
                <!-- Left Card: Выполнено (Команды) -->
                <div class="card" style="padding: 1.5rem 2rem; position: relative; display: flex; flex-direction: column; justify-content: space-between; width: 100%;">
                    <!-- Header & Sparkline -->
                    <div>
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 0.5rem; position: relative;">
                            <div>
                                <div style="font-family: 'Outfit', sans-serif; font-size: 1.3rem; font-weight: 700; color: var(--text-primary);">Выполнено</div>
                                <div style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 2px;">Всего задач, шт.</div>
                            </div>
                            <div style="display: flex; align-items: center; gap: 15px;">
                                <!-- Sparkline -->
                                <div id="jira-completed-sparkline" style="display: flex; align-items: flex-end; gap: 3px; height: 28px; margin-top: 2px;">
                                    <!-- Dynamic sparkline bars -->
                                </div>
                                <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 20px; height: 20px; color: var(--text-secondary); opacity: 0.7;">
                                    <line x1="7" y1="17" x2="17" y2="7"></line>
                                    <polyline points="7 7 17 7 17 17"></polyline>
                                </svg>
                            </div>
                        </div>
                        
                        <!-- Centered Big Total & Trend -->
                        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin: 1rem 0 1.5rem 0; text-align: center;">
                            <div id="jira-completed-trend-wrapper" style="display: flex; align-items: center; gap: 4px; font-weight: 700; font-size: 0.95rem; margin-bottom: 4px;">
                                <svg id="jira-completed-trend-icon" class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" style="width: 14px; height: 14px;">
                                    <polyline points="18 15 12 9 6 15"></polyline>
                                </svg>
                                <span id="jira-completed-trend-val">+50%</span>
                            </div>
                            <div id="jira-completed-total" style="font-family: 'Outfit', sans-serif; font-size: 3.8rem; font-weight: 700; color: var(--text-primary); line-height: 1;">0</div>
                        </div>
                    </div>
                    
                    <!-- Teams Grid (Below) -->
                    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; border-top: 1px solid var(--card-border); padding-top: 1.25rem; margin-top: auto;">
                        <!-- Team 1 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #3b82f6; flex-shrink: 0;"></span>
                                <span>Camunda</span>
                            </div>
                            <span id="jira-completed-camunda" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 2 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #10b981; flex-shrink: 0;"></span>
                                <span>Devops</span>
                            </div>
                            <span id="jira-completed-devops" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 3 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #f59e0b; flex-shrink: 0;"></span>
                                <span>Front</span>
                            </div>
                            <span id="jira-completed-front" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 4 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #ef4444; flex-shrink: 0;"></span>
                                <span>Абс</span>
                            </div>
                            <span id="jira-completed-abs" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 5 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #8b5cf6; flex-shrink: 0;"></span>
                                <span>Колл-центр / Амл / Риски</span>
                            </div>
                            <span id="jira-completed-callcenter" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 6 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #06b6d4; flex-shrink: 0;"></span>
                                <span>Кфо / Залоги / Проблемные Кредиты</span>
                            </div>
                            <span id="jira-completed-kfo" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 7 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #ec4899; flex-shrink: 0;"></span>
                                <span>Пластиковые Карты</span>
                            </div>
                            <span id="jira-completed-cards" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                        <!-- Team 8 -->
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <div style="display: flex; align-items: center; gap: 6px; font-size: 0.8rem; color: var(--text-secondary);">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: #f97316; flex-shrink: 0;"></span>
                                <span>Скоринг</span>
                            </div>
                            <span id="jira-completed-scoring" style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-left: 12px;">0</span>
                        </div>
                    </div>
                </div>
            </section>
            
            <!-- Row 1.8: Time Spent on Completed Tasks (Quarterly filterable) -->
            <section style="display: flex; flex-direction: column; gap: 1rem; margin-top: 1.5rem;">
                <!-- Right Card: Время выполненных задач -->
                <div class="card" style="padding: 1.5rem 2rem; position: relative; width: 100%;">
                    <!-- Header -->
                    <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 0.5rem;">
                        <div>
                            <div style="font-family: 'Outfit', sans-serif; font-size: 1.3rem; font-weight: 700; color: var(--text-primary);">Время выполненных задач</div>
                            <div style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 2px;">Потрачено времени, ч/дни</div>
                        </div>
                        <svg class="svg-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width: 20px; height: 20px; color: var(--text-secondary); opacity: 0.7;">
                            <circle cx="12" cy="12" r="1"></circle>
                            <circle cx="12" cy="5" r="1"></circle>
                            <circle cx="12" cy="19" r="1"></circle>
                        </svg>
                    </div>
                    
                    <!-- Content: Grid of 2 columns -->
                    <div style="display: grid; grid-template-columns: 1.4fr 1fr; gap: 2rem; align-items: center; margin-top: 1.25rem; min-height: 220px;">
                        <!-- Column 1: Table -->
                        <div style="display: flex; flex-direction: column; gap: 10px;" id="jira-time-table-body">
                            <!-- Dynamic rows -->
                        </div>
                        
                        <!-- Column 2: Doughnut Chart -->
                        <div style="position: relative; width: 170px; height: 170px; margin: 0 auto;">
                            <canvas id="chart-jira-time-spent"></canvas>
                            <div class="doughnut-center-card" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--card-bg); border: 1px solid var(--card-border); box-shadow: 0 8px 24px rgba(0,0,0,0.08); width: 95px; height: 95px; border-radius: 50%; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; pointer-events: none; z-index: 5;">
                                <span style="font-size: 0.65rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;">Всего</span>
                                <span id="doughnut-jira-time-center-value" style="font-size: 1.2rem; font-weight: 700; color: var(--text-primary); margin-top: 1px; line-height: 1;">0</span>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 2: Charts (Количество по направлениям, Распределение трудозатрат) -->
            <section class="charts-row-two">
                <div class="card chart-card">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Количество задач по направлениям</span>
                    </div>
                    <div class="chart-container-pie" style="position: relative; display: flex; flex-direction: column; align-items: center; justify-content: flex-start; min-height: 380px; width: 100%; padding: 10px 0;">
                        <div style="position: relative; width: 280px; height: 280px; flex-shrink: 0; margin-bottom: 15px;">
                            <canvas id="chart-jira-categories"></canvas>
                            <div class="doughnut-center-card" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--card-bg); border: 1px solid var(--card-border); box-shadow: 0 8px 32px rgba(0,0,0,0.12); width: 150px; height: 150px; border-radius: 50%; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; pointer-events: none; z-index: 5;">
                                <span style="font-size: 0.75rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;">Всего</span>
                                <span id="doughnut-jira-categories-center-value" style="font-size: 1.6rem; font-weight: 700; color: var(--text-primary); margin-top: 2px; line-height: 1;">0</span>
                                <span style="font-size: 0.7rem; color: var(--text-secondary); margin-top: 2px;">задач</span>
                            </div>
                        </div>
                        <div id="doughnut-jira-categories-legend" style="display: flex; flex-wrap: wrap; justify-content: center; gap: 8px 12px; width: 100%; max-height: 75px; overflow-y: auto; padding: 5px 10px;">
                            <!-- Rendered dynamically -->
                        </div>
                    </div>
                </div>

                <div class="card chart-card">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Распределение трудозатрат (в часах)</span>
                    </div>
                    <div class="chart-container-pie" style="position: relative; display: flex; flex-direction: column; align-items: center; justify-content: flex-start; min-height: 380px; width: 100%; padding: 10px 0;">
                        <div style="position: relative; width: 280px; height: 280px; flex-shrink: 0; margin-bottom: 15px;">
                            <canvas id="chart-jira-hours"></canvas>
                            <div class="doughnut-center-card" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--card-bg); border: 1px solid var(--card-border); box-shadow: 0 8px 32px rgba(0,0,0,0.12); width: 150px; height: 150px; border-radius: 50%; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; pointer-events: none; z-index: 5;">
                                <span style="font-size: 0.75rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;">Всего</span>
                                <span id="doughnut-jira-hours-center-value" style="font-size: 1.6rem; font-weight: 700; color: var(--text-primary); margin-top: 2px; line-height: 1;">100%</span>
                                <span style="font-size: 0.7rem; color: var(--text-secondary); margin-top: 2px;">трудозатрат</span>
                            </div>
                        </div>
                        <div id="doughnut-jira-hours-legend" style="display: flex; flex-wrap: wrap; justify-content: center; gap: 8px 12px; width: 100%; max-height: 75px; overflow-y: auto; padding: 5px 10px;">
                            <!-- Rendered dynamically -->
                        </div>
                    </div>
                </div>
            </section>

            <!-- Row 3: Systems Chart -->
            <section class="charts-row-full">
                <div class="card chart-card" style="min-height: 480px;">
                    <div class="chart-card-header">
                        <span class="chart-card-title">Распределение задач по системам / командам</span>
                    </div>
                    <div class="chart-container-full" style="height: 380px;">
                        <canvas id="chart-jira-systems"></canvas>
                    </div>
                </div>
            </section>
        </div><!-- Closing #panel-jira -->

        <!-- Footer -->
        <footer>
            <p>© 2026 Aloqabank. Центр централизованных решений информационных технологий Dashboard.</p>
        </footer>
    </div>

    <!-- Data Injection -->
    <script>
        const dashboardData = ##DATA_PLACEHOLDER##;

        let activeMainTab = 'jira';
        let activePeriod = 'all';
        let subthemeSearchQuery = '';
        let topCount = 15;
        
        let detailsSearchQuery = '';
        let topCountDetails = 15;

        // Loans Chart instances
        let loansProductsChart = null;
        let loansStatusesChart = null;
        let loansMonthlyProductsChart = null;
        let loansMonthlyChart = null;
        let loansMonthlyAmountChart = null;
        let loansChartsInitialized = false;

        // Jira Chart instances
        let jiraCategoriesChart = null;
        let jiraHoursChart = null;
        let jiraSystemsChart = null;
        let jiraTimeSpentChart = null;
        let jiraChartsInitialized = false;

        // Chart instances
        let typesChart = null;
        let themesChart = null;
        let subthemesChart = null;
        let detailsChart = null;
        let mcMonthlyChart = null;

        // Risk Chart instances
        let riskCategoriesChart = null;
        let riskEventTypesChart = null;
        let riskBranchesChart = null;
        let riskSourcesChart = null;
        let riskLevelsChart = null;
        let riskChartsInitialized = false;
        let topCountRiskBranches = 15;

        // Colors configurations
        const chartColors = [
            '#8b5cf6', '#3b82f6', '#10b981', '#f59e0b', '#ec4899',
            '#14b8a6', '#f43f5e', '#06b6d4', '#eab308', '#6366f1',
            '#a855f7', '#0d9488', '#2563eb', '#059669', '#d97706',
            '#c084fc', '#60a5fa', '#34d399', '#fbbf24', '#f472b6'
        ];

        // Theme colors helper
        function getThemeColors() {
            const isLight = document.body.classList.contains('light-theme');
            return {
                textPrimary: isLight ? '#0f172a' : '#f8fafc',
                textSecondary: isLight ? '#475569' : '#94a3b8',
                gridColor: isLight ? 'rgba(0, 0, 0, 0.06)' : 'rgba(255, 255, 255, 0.05)',
                cardBorder: isLight ? '#ffffff' : '#1e293b'
            };
        }

        // Format number with spaces
        function formatNumber(num) {
            return num.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, " ");
        }

        // Count up animation for numbers
        function animateValue(id, start, end, duration, suffix = '') {
            if (start === end) {
                document.getElementById(id).innerText = formatNumber(end) + suffix;
                return;
            }
            const range = end - start;
            const obj = document.getElementById(id);
            
            const startTime = new Date().getTime();
            const endTime = startTime + duration;
            
            function run() {
                const now = new Date().getTime();
                const remaining = Math.max((endTime - now) / duration, 0);
                const value = Math.round(end - (remaining * range));
                obj.innerText = formatNumber(value) + suffix;
                if (value != end) {
                    requestAnimationFrame(run);
                } else {
                    obj.innerText = formatNumber(end) + suffix;
                }
            }
            requestAnimationFrame(run);
        }

        function toggleChartSlice(chartName, index, el) {
            let chart = null;
            if (chartName === 'loansProductsChart') chart = loansProductsChart;
            if (chartName === 'loansStatusesChart') chart = loansStatusesChart;
            if (chartName === 'jiraCategoriesChart') chart = jiraCategoriesChart;
            if (chartName === 'jiraHoursChart') chart = jiraHoursChart;
            if (!chart) return;
            
            const meta = chart.getDatasetMeta(0);
            const isVisible = !meta.data[index].hidden;
            meta.data[index].hidden = isVisible;
            chart.update();
            
            if (isVisible) {
                el.style.opacity = '0.4';
                el.style.textDecoration = 'line-through';
            } else {
                el.style.opacity = '1';
                el.style.textDecoration = 'none';
            }
        }

        // Initialize or Update Dashboard Components
        function updateDashboard() {
            if (activeMainTab === 'callcenter') {
                const data = dashboardData[activePeriod];
            
            // 1. Update KPI Values
            const prevTotal = parseInt(document.getElementById('kpi-total').innerText.replace(/\\s/g, '')) || 0;
            const prevIncoming = parseInt(document.getElementById('kpi-incoming').innerText.replace(/\\s/g, '')) || 0;
            const prevOutgoing = parseInt(document.getElementById('kpi-outgoing').innerText.replace(/\\s/g, '')) || 0;
            
            animateValue('kpi-total', prevTotal, data.total, 800);
            animateValue('kpi-incoming', prevIncoming, data.incoming, 800);
            animateValue('kpi-outgoing', prevOutgoing, data.outgoing, 800);
            
            // Percentages
            const incomingPct = data.total > 0 ? ((data.incoming / data.total) * 100).toFixed(1) : 0;
            const outgoingPct = data.total > 0 ? ((data.outgoing / data.total) * 100).toFixed(1) : 0;
            
            document.getElementById('pct-incoming').innerText = incomingPct + '%';
            document.getElementById('pct-outgoing').innerText = outgoingPct + '%';

            // Update MicroCredit Tasks stats
            const prevMCTotal = parseInt(document.getElementById('kpi-mc-total').innerText.replace(/\\s/g, '')) || 0;
            const prevMCCompleted = parseInt(document.getElementById('kpi-mc-completed').innerText.replace(/\\s/g, '')) || 0;
            const prevMCRefused = parseInt(document.getElementById('kpi-mc-refused').innerText.replace(/\\s/g, '')) || 0;
            
            animateValue('kpi-mc-total', prevMCTotal, data.total_mc, 800);
            animateValue('kpi-mc-completed', prevMCCompleted, data.completed_mc, 800);
            animateValue('kpi-mc-refused', prevMCRefused, data.refused_mc, 800);
            
            const mcCompletedPct = data.total_mc > 0 ? ((data.completed_mc / data.total_mc) * 100).toFixed(1) : 0;
            const mcRefusedPct = data.total_mc > 0 ? ((data.refused_mc / data.total_mc) * 100).toFixed(1) : 0;
            
            document.getElementById('pct-mc-completed').innerText = mcCompletedPct + '%';
            document.getElementById('pct-mc-refused').innerText = mcRefusedPct + '%';

            // 2. Render Appeal Types Chart (Pie/Doughnut)
            renderTypesChart(data.types);

            // 3. Render Main Themes Chart (Pie/Doughnut)
            renderThemesChart(data.themes);

            // 4. Render Subthemes Chart (Horizontal Bar Chart)
            renderSubthemesChart(data.subthemes);
            
            // 5. Render Details (Суть темы) Chart (Green Horizontal Bar Chart)
            renderDetailsChart(data.details);
            
            // 6. Render MicroCredit Monthly Chart (Grouped vertical bars)
            renderMCMonthlyChart(data.monthly_mc);
            } else if (activeMainTab === 'loans') {
                renderLoansDashboard();
            } else if (activeMainTab === 'jira') {
                renderJiraDashboard();
            }
        }

        function renderTypesChart(typesData) {
            const colors = getThemeColors();
            const labels = typesData.map(item => item.name);
            const values = typesData.map(item => item.value);

            if (typesChart) {
                typesChart.destroy();
            }

            const ctx = document.getElementById('chart-types').getContext('2d');
            typesChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: chartColors.slice(0, labels.length),
                        borderWidth: 2,
                        borderColor: colors.cardBorder,
                        hoverOffset: 12
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                padding: 15,
                                font: {
                                    size: 11,
                                    family: 'Inter'
                                },
                                boxWidth: 12,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` Обращения: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    },
                    cutout: '65%'
                }
            });
        }

        function renderThemesChart(themesData) {
            const colors = getThemeColors();
            const maxPieSlices = 8;
            let labels = [];
            let values = [];

            if (themesData.length <= maxPieSlices) {
                labels = themesData.map(item => item.name);
                values = themesData.map(item => item.value);
            } else {
                const topThemes = themesData.slice(0, maxPieSlices);
                labels = topThemes.map(item => item.name);
                values = topThemes.map(item => item.value);

                const remainingSum = themesData.slice(maxPieSlices).reduce((sum, item) => sum + item.value, 0);
                labels.push('Прочие темы');
                values.push(remainingSum);
            }

            if (themesChart) {
                themesChart.destroy();
            }

            const ctx = document.getElementById('chart-themes').getContext('2d');
            themesChart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: chartColors.slice(4, 4 + labels.length),
                        borderWidth: 2,
                        borderColor: colors.cardBorder,
                        hoverOffset: 12
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                padding: 15,
                                font: {
                                    size: 11,
                                    family: 'Inter'
                                },
                                boxWidth: 12,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` Обращения: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderSubthemesChart(subthemesData) {
            const colors = getThemeColors();
            let filtered = subthemesData;
            if (subthemeSearchQuery) {
                filtered = subthemesData.filter(item => 
                    item.name.toLowerCase().includes(subthemeSearchQuery.toLowerCase())
                );
            }

            const displayedData = filtered.slice(0, topCount);
            
            const labels = displayedData.map(item => item.name);
            const values = displayedData.map(item => item.value);

            if (subthemesChart) {
                subthemesChart.destroy();
            }

            const ctx = document.getElementById('chart-subthemes').getContext('2d');
            subthemesChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Количество обращений',
                        data: values,
                        backgroundColor: 'rgba(139, 92, 246, 0.75)',
                        borderColor: '#8b5cf6',
                        borderWidth: 1.5,
                        borderRadius: 6,
                        hoverBackgroundColor: '#a78bfa'
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: {
                                color: colors.gridColor,
                                tickColor: 'transparent'
                            },
                            ticks: {
                                color: colors.textSecondary,
                                font: {
                                    family: 'Inter',
                                    size: 10
                                }
                            }
                        },
                        y: {
                            grid: {
                                display: false
                            },
                            ticks: {
                                color: colors.textPrimary,
                                font: {
                                    family: 'Inter',
                                    weight: '500',
                                    size: 11
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` Обращения: ${formatNumber(context.raw)}`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderDetailsChart(detailsData) {
            const colors = getThemeColors();
            let filtered = detailsData;
            if (detailsSearchQuery) {
                filtered = detailsData.filter(item => 
                    item.name.toLowerCase().includes(detailsSearchQuery.toLowerCase())
                );
            }

            const displayedData = filtered.slice(0, topCountDetails);
            
            const labels = displayedData.map(item => item.name);
            const values = displayedData.map(item => item.value);

            if (detailsChart) {
                detailsChart.destroy();
            }

            const ctx = document.getElementById('chart-details').getContext('2d');
            detailsChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Количество обращений',
                        data: values,
                        backgroundColor: 'rgba(16, 185, 129, 0.75)',
                        borderColor: '#10b981',
                        borderWidth: 1.5,
                        borderRadius: 6,
                        hoverBackgroundColor: '#34d399'
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: {
                                color: colors.gridColor,
                                tickColor: 'transparent'
                            },
                            ticks: {
                                color: colors.textSecondary,
                                font: {
                                    family: 'Inter',
                                    size: 10
                                }
                            }
                        },
                        y: {
                            grid: {
                                display: false
                            },
                            ticks: {
                                color: colors.textPrimary,
                                font: {
                                    family: 'Inter',
                                    weight: '500',
                                    size: 11
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` Обращения: ${formatNumber(context.raw)}`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderMCMonthlyChart(monthlyData) {
            const colors = getThemeColors();
            const monthNames = {
                '2026-01': 'Январь',
                '2026-02': 'Февраль',
                '2026-03': 'Март',
                '2026-04': 'Апрель',
                '2026-05': 'Май',
                '2026-06': 'Июнь'
            };

            const labels = monthlyData.map(item => monthNames[item.month] || item.month);
            const completedValues = monthlyData.map(item => item.completed);
            const refusedValues = monthlyData.map(item => item.refused);

            if (mcMonthlyChart) {
                mcMonthlyChart.destroy();
            }

            const ctx = document.getElementById('chart-mc-monthly').getContext('2d');
            mcMonthlyChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Выполнено (успешно)',
                            data: completedValues,
                            backgroundColor: 'rgba(16, 185, 129, 0.75)',
                            borderColor: '#10b981',
                            borderWidth: 1.5,
                            borderRadius: 4
                        },
                        {
                            label: 'Отказ клиента',
                            data: refusedValues,
                            backgroundColor: 'rgba(239, 68, 68, 0.75)',
                            borderColor: '#ef4444',
                            borderWidth: 1.5,
                            borderRadius: 4
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: {
                                display: false
                            },
                            ticks: {
                                color: colors.textPrimary,
                                font: {
                                    family: 'Inter',
                                    weight: '500',
                                    size: 11
                                }
                            }
                        },
                        y: {
                            grid: {
                                color: colors.gridColor,
                                tickColor: 'transparent'
                            },
                            ticks: {
                                color: colors.textSecondary,
                                font: {
                                    family: 'Inter',
                                    size: 10
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top',
                            labels: {
                                color: colors.textSecondary,
                                font: {
                                    family: 'Inter',
                                    size: 11
                                },
                                boxWidth: 12,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` ${context.dataset.label}: ${formatNumber(context.raw)}`;
                                }
                            }
                        }
                    }
                }
            });
        }

        // Interactive triggers
        function switchTab(period, element) {
            if (activePeriod === period) return;
            
            // Scope subtab active class toggle to current panel
            const parentPanel = element.closest('.main-panel');
            if (parentPanel) {
                parentPanel.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            } else {
                document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            }
            element.classList.add('active');

            activePeriod = period;
            updateDashboard();
        }

        // Subtheme actions
        function handleSearch() {
            subthemeSearchQuery = document.getElementById('subtheme-search').value;
            const data = dashboardData[activePeriod];
            renderSubthemesChart(data.subthemes);
        }

        function handleSliderChange(val) {
            topCount = parseInt(val);
            document.getElementById('slider-val').innerText = val;
            const data = dashboardData[activePeriod];
            renderSubthemesChart(data.subthemes);
        }

        // Details actions
        function handleDetailsSearch() {
            detailsSearchQuery = document.getElementById('details-search').value;
            const data = dashboardData[activePeriod];
            renderDetailsChart(data.details);
        }

        function handleDetailsSliderChange(val) {
            topCountDetails = parseInt(val);
            document.getElementById('slider-val-details').innerText = val;
            const data = dashboardData[activePeriod];
            renderDetailsChart(data.details);
        }

        // Theme Toggle Logic
        function toggleTheme() {
            const body = document.body;
            body.classList.toggle('light-theme');
            
            const isLight = body.classList.contains('light-theme');
            localStorage.setItem('theme', isLight ? 'light' : 'dark');
            
            updateThemeUI(isLight);
        }

        function updateThemeUI(isLight) {
            const sunIcon = document.querySelector('.sun-icon');
            const moonIcon = document.querySelector('.moon-icon');
            
            if (isLight) {
                sunIcon.style.display = 'none';
                moonIcon.style.display = 'block';
            } else {
                sunIcon.style.display = 'block';
                moonIcon.style.display = 'none';
            }
            
            // Re-render/update charts with new theme colors
            updateChartsTheme();
        }

        function updateChartsTheme() {
            const colors = getThemeColors();
            
            if (typesChart) {
                typesChart.options.plugins.legend.labels.color = colors.textSecondary;
                typesChart.data.datasets[0].borderColor = colors.cardBorder;
                typesChart.update();
            }
            
            if (themesChart) {
                themesChart.options.plugins.legend.labels.color = colors.textSecondary;
                themesChart.data.datasets[0].borderColor = colors.cardBorder;
                themesChart.update();
            }
            
            if (subthemesChart) {
                subthemesChart.options.scales.x.ticks.color = colors.textSecondary;
                subthemesChart.options.scales.y.ticks.color = colors.textPrimary;
                subthemesChart.options.scales.x.grid.color = colors.gridColor;
                subthemesChart.update();
            }

            if (detailsChart) {
                detailsChart.options.scales.x.ticks.color = colors.textSecondary;
                detailsChart.options.scales.y.ticks.color = colors.textPrimary;
                detailsChart.options.scales.x.grid.color = colors.gridColor;
                detailsChart.update();
            }

            if (mcMonthlyChart) {
                mcMonthlyChart.options.scales.x.ticks.color = colors.textPrimary;
                mcMonthlyChart.options.scales.y.ticks.color = colors.textSecondary;
                mcMonthlyChart.options.scales.y.grid.color = colors.gridColor;
                mcMonthlyChart.options.plugins.legend.labels.color = colors.textSecondary;
                mcMonthlyChart.update();
            }

            // Update Risk, Loans and Jira Charts as well
            updateRiskChartsTheme();
            updateLoansChartsTheme();
            updateJiraChartsTheme();
        }

        // --- Main Tab Switcher and Risks Dashboard Rendering Logic ---
        function switchMainTab(tabName) {
            if (activeMainTab === tabName) return;
            activeMainTab = tabName;

            // Toggle main tab buttons active class
            document.getElementById('main-tab-loans').classList.toggle('active', tabName === 'loans');
            document.getElementById('main-tab-callcenter').classList.toggle('active', tabName === 'callcenter');
            document.getElementById('main-tab-risks').classList.toggle('active', tabName === 'risks');
            const jiraTabBtn = document.getElementById('main-tab-jira');
            if (jiraTabBtn) jiraTabBtn.classList.toggle('active', tabName === 'jira');

            // Toggle visibility of panels
            document.getElementById('panel-loans').style.display = tabName === 'loans' ? 'flex' : 'none';
            document.getElementById('panel-callcenter').style.display = tabName === 'callcenter' ? 'flex' : 'none';
            document.getElementById('panel-risks').style.display = tabName === 'risks' ? 'flex' : 'none';
            const jiraPanel = document.getElementById('panel-jira');
            if (jiraPanel) jiraPanel.style.display = tabName === 'jira' ? 'flex' : 'none';

            if (tabName === 'risks') {
                renderRiskDashboard();
            } else if (tabName === 'loans') {
                renderLoansDashboard();
            } else if (tabName === 'jira') {
                renderJiraDashboard();
            } else {
                updateDashboard();
            }
        }

        function renderRiskDashboard() {
            const data = dashboardData.risks;

            // 1. Update KPI Values
            animateValue('kpi-risk-total', 0, data.total, 800);
            animateValue('kpi-risk-open', 0, data.open, 800);
            animateValue('kpi-risk-confirmed', 0, data.confirmed, 800);
            animateValue('kpi-risk-actual', 0, data.actual_loss, 800, ' UZS');

            // Update top category text
            document.getElementById('kpi-risk-top-category').innerText = data.top_category;

            // Update Auto vs Manual Task KPI Values
            animateValue('kpi-risk-auto-tasks', 0, data.auto_tasks, 800);
            animateValue('kpi-risk-manual-tasks', 0, data.manual_tasks, 800);

            // Dynamically set slider max based on number of branches
            const slider = document.getElementById('risk-branches-slider');
            if (slider) {
                slider.max = data.branches.length;
            }

            // 2. Render/Update Charts
            if (!riskChartsInitialized) {
                renderRiskCategoriesChart(data.categories);
                renderRiskEventTypesChart(data.event_types);
                renderRiskBranchesChart(data.branches);
                renderRiskSourcesChart(data.sources);
                renderRiskLevelsChart(data.risk_levels);
                riskChartsInitialized = true;
            } else {
                updateRiskChartsData(data);
            }
        }

        function handleRiskBranchesSliderChange(val) {
            topCountRiskBranches = parseInt(val);
            document.getElementById('risk-branches-slider-val').innerText = val;
            
            const data = dashboardData.risks;
            const titleEl = document.getElementById('risk-branches-title');
            if (topCountRiskBranches >= data.branches.length) {
                titleEl.innerText = 'Количество событий по филиалам (Все)';
            } else {
                titleEl.innerText = `Количество событий по филиалам (Топ-${topCountRiskBranches})`;
            }

            // Dynamically adjust height of container to avoid squishing
            const container = document.getElementById('risk-branches-container');
            if (container) {
                const newHeight = Math.max(300, topCountRiskBranches * 25 + 50);
                container.style.height = newHeight + 'px';
            }

            renderRiskBranchesChart(data.branches);
        }

        function renderRiskCategoriesChart(categoriesData) {
            const colors = getThemeColors();
            const maxSlices = 5;
            let labels = [];
            let values = [];

            if (categoriesData.length <= maxSlices) {
                labels = categoriesData.map(item => item.name);
                values = categoriesData.map(item => item.value);
            } else {
                const topItems = categoriesData.slice(0, maxSlices);
                labels = topItems.map(item => item.name);
                values = topItems.map(item => item.value);

                const remainingSum = categoriesData.slice(maxSlices).reduce((sum, item) => sum + item.value, 0);
                labels.push('Прочие категории');
                values.push(remainingSum);
            }

            if (riskCategoriesChart) {
                riskCategoriesChart.destroy();
            }

            const ctx = document.getElementById('chart-risk-categories').getContext('2d');
            riskCategoriesChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: chartColors.slice(0, labels.length),
                        borderWidth: 2,
                        borderColor: colors.cardBorder,
                        hoverOffset: 12
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                padding: 15,
                                font: { size: 10, family: 'Inter' },
                                boxWidth: 10,
                                usePointStyle: true,
                                pointStyle: 'circle',
                                generateLabels: function(chart) {
                                    const data = chart.data;
                                    if (data.labels.length && data.datasets.length) {
                                        return data.labels.map((label, i) => {
                                            let text = label || '';
                                            if (text.length > 25) {
                                                text = text.substring(0, 25) + '...';
                                            }
                                            const ds = data.datasets[0];
                                            return {
                                                text: text,
                                                fillStyle: ds.backgroundColor[i],
                                                strokeStyle: ds.borderColor || '#fff',
                                                lineWidth: ds.borderWidth || 1,
                                                hidden: !chart.getDataVisibility(i),
                                                index: i,
                                                fontColor: chart.options.plugins.legend.labels.color
                                            };
                                        });
                                    }
                                    return [];
                                }
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` События: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    },
                    cutout: '60%'
                }
            });
        }

        function renderRiskEventTypesChart(typesData) {
            const colors = getThemeColors();
            const maxSlices = 6;
            let labels = [];
            let values = [];

            if (typesData.length <= maxSlices) {
                labels = typesData.map(item => item.name);
                values = typesData.map(item => item.value);
            } else {
                const topItems = typesData.slice(0, maxSlices);
                labels = topItems.map(item => item.name);
                values = topItems.map(item => item.value);

                const remainingSum = typesData.slice(maxSlices).reduce((sum, item) => sum + item.value, 0);
                labels.push('Прочие типы');
                values.push(remainingSum);
            }

            if (riskEventTypesChart) {
                riskEventTypesChart.destroy();
            }

            const ctx = document.getElementById('chart-risk-event-types').getContext('2d');
            riskEventTypesChart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: chartColors.slice(5, 5 + labels.length),
                        borderWidth: 2,
                        borderColor: colors.cardBorder,
                        hoverOffset: 12
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                padding: 15,
                                font: { size: 10, family: 'Inter' },
                                boxWidth: 10,
                                usePointStyle: true,
                                pointStyle: 'circle',
                                generateLabels: function(chart) {
                                    const data = chart.data;
                                    if (data.labels.length && data.datasets.length) {
                                        return data.labels.map((label, i) => {
                                            let text = label || '';
                                            if (text.length > 25) {
                                                text = text.substring(0, 25) + '...';
                                            }
                                            const ds = data.datasets[0];
                                            return {
                                                text: text,
                                                fillStyle: ds.backgroundColor[i],
                                                strokeStyle: ds.borderColor || '#fff',
                                                lineWidth: ds.borderWidth || 1,
                                                hidden: !chart.getDataVisibility(i),
                                                index: i,
                                                fontColor: chart.options.plugins.legend.labels.color
                                            };
                                        });
                                    }
                                    return [];
                                }
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` События: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderRiskBranchesChart(branchesData) {
            const colors = getThemeColors();
            const topBranches = branchesData.slice(0, topCountRiskBranches);
            const labels = topBranches.map(item => item.name);
            const values = topBranches.map(item => item.value);

            if (riskBranchesChart) {
                riskBranchesChart.destroy();
            }

            const ctx = document.getElementById('chart-risk-branches').getContext('2d');
            riskBranchesChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Количество событий',
                        data: values,
                        backgroundColor: 'rgba(16, 185, 129, 0.75)',
                        borderColor: '#10b981',
                        borderWidth: 1.5,
                        borderRadius: 6,
                        hoverBackgroundColor: '#34d399'
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: {
                                color: colors.gridColor,
                                tickColor: 'transparent'
                            },
                            ticks: {
                                color: colors.textSecondary,
                                font: { family: 'Inter', size: 10 }
                            }
                        },
                        y: {
                            grid: { display: false },
                            ticks: {
                                color: colors.textPrimary,
                                font: { family: 'Inter', weight: '500', size: 11 }
                            }
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` События: ${formatNumber(context.raw)}`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderRiskSourcesChart(sourcesData) {
            const colors = getThemeColors();
            const labels = sourcesData.map(item => item.name);
            const values = sourcesData.map(item => item.value);

            if (riskSourcesChart) {
                riskSourcesChart.destroy();
            }

            const ctx = document.getElementById('chart-risk-sources').getContext('2d');
            riskSourcesChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: ['#3b82f6', '#f59e0b', '#ef4444', '#10b981', '#8b5cf6'].slice(0, labels.length),
                        borderWidth: 2,
                        borderColor: colors.cardBorder,
                        hoverOffset: 8
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' },
                                boxWidth: 10,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` События: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    },
                    cutout: '55%'
                }
            });
        }

        function renderRiskLevelsChart(levelsData) {
            const colors = getThemeColors();
            const labels = levelsData.map(item => item.name);
            const values = levelsData.map(item => item.value);

            if (riskLevelsChart) {
                riskLevelsChart.destroy();
            }

            const ctx = document.getElementById('chart-risk-levels').getContext('2d');
            riskLevelsChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: ['#3b82f6', '#f59e0b', '#ef4444'].slice(0, labels.length),
                        borderWidth: 2,
                        borderColor: colors.cardBorder,
                        hoverOffset: 8
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' },
                                boxWidth: 10,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` События: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    },
                    cutout: '55%'
                }
            });
        }

        function updateRiskChartsData(data) {
            renderRiskCategoriesChart(data.categories);
            renderRiskEventTypesChart(data.event_types);
            renderRiskBranchesChart(data.branches);
            renderRiskSourcesChart(data.sources);
            renderRiskLevelsChart(data.risk_levels);
        }

        function updateRiskChartsTheme() {
            const colors = getThemeColors();
            
            if (riskCategoriesChart) {
                riskCategoriesChart.options.plugins.legend.labels.color = colors.textSecondary;
                riskCategoriesChart.data.datasets[0].borderColor = colors.cardBorder;
                riskCategoriesChart.update();
            }
            if (riskEventTypesChart) {
                riskEventTypesChart.options.plugins.legend.labels.color = colors.textSecondary;
                riskEventTypesChart.data.datasets[0].borderColor = colors.cardBorder;
                riskEventTypesChart.update();
            }
            if (riskBranchesChart) {
                riskBranchesChart.options.scales.x.ticks.color = colors.textSecondary;
                riskBranchesChart.options.scales.y.ticks.color = colors.textPrimary;
                riskBranchesChart.options.scales.x.grid.color = colors.gridColor;
                riskBranchesChart.update();
            }
            if (riskSourcesChart) {
                riskSourcesChart.options.plugins.legend.labels.color = colors.textSecondary;
                riskSourcesChart.data.datasets[0].borderColor = colors.cardBorder;
                riskSourcesChart.update();
            }
            if (riskLevelsChart) {
                riskLevelsChart.options.plugins.legend.labels.color = colors.textSecondary;
                riskLevelsChart.data.datasets[0].borderColor = colors.cardBorder;
                riskLevelsChart.update();
            }
        }

        // --- Loans Dashboard Rendering Logic ---
        function renderLoansDashboard() {
            const data = dashboardData.loans[activePeriod];

            // 1. Update KPI Values
            animateValue('kpi-loans-total', 0, data.total_apps, 800);
            animateValue('kpi-loans-success', 0, data.success, 800);
            animateValue('kpi-loans-reject-scoring', 0, data.reject, 800);
            animateValue('kpi-loans-reject-client', 0, data.reject_client, 800);

            // Update doughnut center total values
            const totalStatusApps = data.success + data.reject + data.reject_client;
            animateValue('doughnut-center-value', 0, totalStatusApps, 800);
            animateValue('doughnut-products-center-value', 0, data.total_apps, 800);

            // Slicing monthly chart data based on activePeriod
            let slicedMonthly = {
                months: data.monthly.months,
                success: data.monthly.success,
                reject: data.monthly.reject,
                reject_client: data.monthly.reject_client
            };
            let slicedMonthlyAmount = {
                months: data.monthly_issued.months,
                amounts: data.monthly_issued.amounts
            };
            let slicedMonthlyIncome = {
                months: data.monthly_income.months,
                prc: data.monthly_income.prc,
                prc_pen: data.monthly_income.prc_pen,
                total: data.monthly_income.total
            };
            let slicedMonthlyProducts = JSON.parse(JSON.stringify(data.monthly_products));
            if (activePeriod === 'q1') {
                slicedMonthly.months = slicedMonthly.months.slice(0, 3);
                slicedMonthly.success = slicedMonthly.success.slice(0, 3);
                slicedMonthly.reject = slicedMonthly.reject.slice(0, 3);
                slicedMonthly.reject_client = slicedMonthly.reject_client.slice(0, 3);
                
                slicedMonthlyAmount.months = slicedMonthlyAmount.months.slice(0, 3);
                slicedMonthlyAmount.amounts = slicedMonthlyAmount.amounts.slice(0, 3);
                
                slicedMonthlyIncome.months = slicedMonthlyIncome.months.slice(0, 3);
                slicedMonthlyIncome.prc = slicedMonthlyIncome.prc.slice(0, 3);
                slicedMonthlyIncome.prc_pen = slicedMonthlyIncome.prc_pen.slice(0, 3);
                slicedMonthlyIncome.total = slicedMonthlyIncome.total.slice(0, 3);
                
                slicedMonthlyProducts.months = slicedMonthlyProducts.months.slice(0, 3);
                slicedMonthlyProducts.products.forEach(p => {
                    p.values = p.values.slice(0, 3);
                });
            } else if (activePeriod === 'q2') {
                slicedMonthly.months = slicedMonthly.months.slice(3, 6);
                slicedMonthly.success = slicedMonthly.success.slice(3, 6);
                slicedMonthly.reject = slicedMonthly.reject.slice(3, 6);
                slicedMonthly.reject_client = slicedMonthly.reject_client.slice(3, 6);
                
                slicedMonthlyAmount.months = slicedMonthlyAmount.months.slice(3, 6);
                slicedMonthlyAmount.amounts = slicedMonthlyAmount.amounts.slice(3, 6);
                
                slicedMonthlyProducts.months = slicedMonthlyProducts.months.slice(3, 6);
                slicedMonthlyProducts.products.forEach(p => {
                    p.values = p.values.slice(3, 6);
                });
            }

            // 2. Render/Update Charts
            if (!loansChartsInitialized) {
                renderLoansProductsChart(data.products);
                renderLoansStatusesChart(data.statuses);
                renderLoansMonthlyProductsChart(slicedMonthlyProducts);
                renderLoansMonthlyChart(slicedMonthly);
                renderLoansMonthlyAmountChart(slicedMonthlyAmount);
                loansChartsInitialized = true;
            } else {
                // Update products chart (re-render to sync HTML legend)
                renderLoansProductsChart(data.products);

                // Update statuses chart (re-render to sync HTML legend)
                renderLoansStatusesChart(data.statuses);

                // Update monthly chart
                loansMonthlyChart.data.labels = slicedMonthly.months;
                loansMonthlyChart.data.datasets[0].data = slicedMonthly.success;
                loansMonthlyChart.data.datasets[1].data = slicedMonthly.reject;
                if (loansMonthlyChart.data.datasets[2]) {
                    loansMonthlyChart.data.datasets[2].data = slicedMonthly.reject_client;
                }
                loansMonthlyChart.update();

                // Update monthly amount chart
                if (loansMonthlyAmountChart) {
                    loansMonthlyAmountChart.data.labels = slicedMonthlyAmount.months;
                    loansMonthlyAmountChart.data.datasets[0].data = slicedMonthlyAmount.amounts;
                    loansMonthlyAmountChart.update();
                }

                // Update monthly products chart
                if (loansMonthlyProductsChart) {
                    loansMonthlyProductsChart.data.labels = slicedMonthlyProducts.months;
                    slicedMonthlyProducts.products.forEach((p, idx) => {
                        if (loansMonthlyProductsChart.data.datasets[idx]) {
                            loansMonthlyProductsChart.data.datasets[idx].data = p.values;
                        }
                    });
                    loansMonthlyProductsChart.update();
                }
            }
        }

        function renderLoansProductsChart(productsData) {
            const ctx = document.getElementById('chart-loans-products').getContext('2d');
            const colors = getThemeColors();
            
            if (loansProductsChart) {
                loansProductsChart.destroy();
            }
            
            loansProductsChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: productsData.map(item => item.name),
                    datasets: [{
                        data: productsData.map(item => item.value),
                        backgroundColor: productsData.map((_, idx) => chartColors[idx % chartColors.length]),
                        borderColor: colors.cardBorder,
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '70%',
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` ${context.label}: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });

            // Generate HTML legend below the canvas to guarantee perfect centering of the middle circle
            const legendContainer = document.getElementById('doughnut-products-legend');
            if (legendContainer) {
                legendContainer.innerHTML = productsData.map((item, idx) => {
                    const color = chartColors[idx % chartColors.length];
                    return `
                    <div style="display: flex; align-items: center; gap: 6px; cursor: pointer; transition: opacity 0.2s; font-size: 11px;" onclick="toggleChartSlice('loansProductsChart', ${idx}, this)" title="${item.name}">
                        <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: ${color}; flex-shrink: 0;"></span>
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 160px; color: var(--text-secondary);">${item.name}</span>
                    </div>`;
                }).join('');
            }
        }

        function renderLoansMonthlyProductsChart(monthlyProdData) {
            const ctx = document.getElementById('chart-loans-monthly-products').getContext('2d');
            const colors = getThemeColors();
            
            loansMonthlyProductsChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: monthlyProdData.months,
                    datasets: monthlyProdData.products.map((p, idx) => ({
                        label: p.name,
                        data: p.values,
                        backgroundColor: chartColors[idx % chartColors.length],
                        borderColor: colors.cardBorder,
                        borderWidth: 1,
                        borderRadius: 4
                    }))
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: { display: false },
                            ticks: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' }
                            }
                        },
                        y: {
                            grid: { color: colors.gridColor },
                            ticks: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' },
                                callback: function(value) {
                                    if (value >= 1e6) return (value / 1e6).toFixed(1) + ' млн';
                                    if (value >= 1e3) return (value / 1e3).toFixed(0) + ' тыс';
                                    return value;
                                }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' },
                                boxWidth: 10,
                                padding: 15,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` ${context.dataset.label}: ${formatNumber(context.raw)}`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderLoansStatusesChart(statusesData) {
            const ctx = document.getElementById('chart-loans-statuses').getContext('2d');
            const colors = getThemeColors();
            
            const statusColors = {
                'Отказ банка по скорингу': 'rgba(239, 68, 68, 0.75)',   // Red
                'Отказ клиента': 'rgba(245, 158, 11, 0.75)',            // Yellow
                'Успешная выдача': 'rgba(16, 185, 129, 0.75)'           // Green
            };
            const backgroundColors = statusesData.map(item => statusColors[item.name] || 'rgba(156, 163, 175, 0.75)');
            
            if (loansStatusesChart) {
                loansStatusesChart.destroy();
            }
            
            loansStatusesChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: statusesData.map(item => item.name),
                    datasets: [{
                        data: statusesData.map(item => item.value),
                        backgroundColor: backgroundColors,
                        borderColor: colors.cardBorder,
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '70%',
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` ${context.label}: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });

            // Generate HTML legend next to the canvas as a grid matching products style
            const legendContainer = document.getElementById('doughnut-statuses-legend');
            if (legendContainer) {
                legendContainer.innerHTML = statusesData.map((item, idx) => {
                    const color = statusColors[item.name] || 'rgba(156, 163, 175, 0.75)';
                    return `
                    <div style="display: flex; align-items: center; gap: 6px; cursor: pointer; transition: opacity 0.2s; font-size: 13px;" onclick="toggleChartSlice('loansStatusesChart', ${idx}, this)" title="${item.name}">
                        <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background-color: ${color}; flex-shrink: 0;"></span>
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 250px; color: var(--text-secondary);">${item.name}</span>
                    </div>`;
                }).join('');
            }
        }

        function renderLoansMonthlyChart(monthlyData) {
            const ctx = document.getElementById('chart-loans-monthly').getContext('2d');
            const colors = getThemeColors();
            
            loansMonthlyChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: monthlyData.months,
                    datasets: [
                        {
                            label: 'Успешная выдача',
                            data: monthlyData.success,
                            backgroundColor: 'rgba(16, 185, 129, 0.8)',
                            borderColor: '#10b981',
                            borderWidth: 1,
                            borderRadius: 4
                        },
                        {
                            label: 'Отказ банка по скорингу',
                            data: monthlyData.reject,
                            backgroundColor: 'rgba(239, 68, 68, 0.8)',
                            borderColor: '#ef4444',
                            borderWidth: 1,
                            borderRadius: 4
                        },
                        {
                            label: 'Отказ клиента',
                            data: monthlyData.reject_client,
                            backgroundColor: 'rgba(245, 158, 11, 0.8)',
                            borderColor: '#f59e0b',
                            borderWidth: 1,
                            borderRadius: 4
                        }
                    ]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: { color: colors.gridColor },
                            ticks: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' }
                            }
                        },
                        y: {
                            grid: { display: false },
                            ticks: {
                                color: colors.textSecondary,
                                font: { size: 11, family: 'Inter', weight: '500' }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            position: 'top',
                            labels: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' }
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` ${context.dataset.label}: ${formatNumber(context.raw)}`;
                                }
                            }
                        }
                    }
                }
            });
        }

        // renderLoansIncomeChart removed

        function renderLoansMonthlyAmountChart(monthlyIssuedData) {
            const ctx = document.getElementById('chart-loans-monthly-amount').getContext('2d');
            const colors = getThemeColors();
            
            loansMonthlyAmountChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: monthlyIssuedData.months,
                    datasets: [
                        {
                            label: 'Сумма выданных кредитов',
                            data: monthlyIssuedData.amounts,
                            backgroundColor: 'rgba(16, 185, 129, 0.85)', // Emerald
                            borderColor: '#10b981',
                            borderWidth: 1,
                            borderRadius: 4
                        }
                    ]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: { color: colors.gridColor },
                            ticks: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' },
                                callback: function(value) {
                                    if (value >= 1e12) return (value / 1e12).toFixed(1) + ' трлн';
                                    if (value >= 1e9) return (value / 1e9).toFixed(1) + ' млрд';
                                    if (value >= 1e6) return (value / 1e6).toFixed(1) + ' млн';
                                    return value;
                                }
                            }
                        },
                        y: {
                            grid: { display: false },
                            ticks: {
                                color: colors.textSecondary,
                                font: { size: 11, family: 'Inter', weight: '500' }
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` Выдано: ${formatNumber(context.raw.toFixed(2))} UZS`;
                                }
                            }
                        }
                    }
                }
            });
        }

        // renderLoansMonthlyIncomeChart removed

        function updateLoansChartsTheme() {
            const colors = getThemeColors();
            
            if (loansProductsChart) {
                loansProductsChart.data.datasets[0].borderColor = colors.cardBorder;
                loansProductsChart.update();
            }
            if (loansMonthlyProductsChart) {
                loansMonthlyProductsChart.options.scales.x.ticks.color = colors.textSecondary;
                loansMonthlyProductsChart.options.scales.y.ticks.color = colors.textSecondary;
                loansMonthlyProductsChart.options.scales.y.grid.color = colors.gridColor;
                loansMonthlyProductsChart.options.plugins.legend.labels.color = colors.textSecondary;
                loansMonthlyProductsChart.update();
            }
            if (loansStatusesChart) {
                if (loansStatusesChart.options.scales) {
                    if (loansStatusesChart.options.scales.x) loansStatusesChart.options.scales.x.ticks.color = colors.textSecondary;
                    if (loansStatusesChart.options.scales.y) {
                        loansStatusesChart.options.scales.y.ticks.color = colors.textSecondary;
                        loansStatusesChart.options.scales.y.grid.color = colors.gridColor;
                    }
                }
                if (loansStatusesChart.options.plugins && loansStatusesChart.options.plugins.legend) {
                    loansStatusesChart.options.plugins.legend.labels.color = colors.textSecondary;
                }
                if (loansStatusesChart.data.datasets[0]) {
                    loansStatusesChart.data.datasets[0].borderColor = colors.cardBorder;
                }
                loansStatusesChart.update();
            }
            if (loansMonthlyChart) {
                loansMonthlyChart.options.scales.x.ticks.color = colors.textSecondary;
                loansMonthlyChart.options.scales.x.grid.color = colors.gridColor;
                loansMonthlyChart.options.scales.y.ticks.color = colors.textSecondary;
                loansMonthlyChart.options.plugins.legend.labels.color = colors.textSecondary;
                loansMonthlyChart.update();
            }
            if (loansMonthlyAmountChart) {
                loansMonthlyAmountChart.options.scales.x.ticks.color = colors.textSecondary;
                loansMonthlyAmountChart.options.scales.x.grid.color = colors.gridColor;
                loansMonthlyAmountChart.options.scales.y.ticks.color = colors.textSecondary;
                loansMonthlyAmountChart.update();
            }
            // Income charts theme handling removed
        }

        // --- Jira Dashboard Rendering Logic ---
        function renderJiraDashboard() {
            const data = dashboardData.jira[activePeriod];

            // 1. Update KPI Values
            animateValue('kpi-jira-total-tasks', 0, data.total_tasks, 800);
            animateValue('kpi-jira-avg-tasks', 0, data.avg_tasks, 800);
            document.getElementById('kpi-jira-main-category').innerText = data.main_category;
            document.getElementById('kpi-jira-main-category-pct').innerText = `${data.main_category_pct}% от общего числа`;

            // 2. Render Charts
            if (!jiraChartsInitialized) {
                renderJiraCategoriesChart(data.months, data.monthly_categories);
                renderJiraHoursChart(data.months, data.monthly_hours);
                renderJiraSystemsChart(data.months, data.monthly_systems);
                jiraChartsInitialized = true;
            } else {
                // Update Categories Chart
                if (jiraCategoriesChart) {
                    const categoriesData = data.monthly_categories.map(c => {
                        const totalVal = c.values.reduce((a, b) => a + b, 0);
                        return {
                            name: c.name,
                            value: totalVal
                        };
                    });
                    
                    const totalTasks = categoriesData.reduce((sum, item) => sum + item.value, 0);
                    const centerValElement = document.getElementById('doughnut-jira-categories-center-value');
                    if (centerValElement) {
                        centerValElement.innerText = formatNumber(totalTasks);
                    }
                    
                    jiraCategoriesChart.data.labels = categoriesData.map(item => item.name);
                    jiraCategoriesChart.data.datasets[0].data = categoriesData.map(item => item.value);
                    jiraCategoriesChart.data.datasets[0].backgroundColor = categoriesData.map((_, idx) => chartColors[idx % chartColors.length]);
                    jiraCategoriesChart.update();
                    
                    const legendContainer = document.getElementById('doughnut-jira-categories-legend');
                    if (legendContainer) {
                        legendContainer.innerHTML = categoriesData.map((item, idx) => {
                            const color = chartColors[idx % chartColors.length];
                            return `
                            <div style="display: flex; align-items: center; gap: 6px; cursor: pointer; transition: opacity 0.2s; font-size: 11px;" onclick="toggleChartSlice('jiraCategoriesChart', ${idx}, this)" title="${item.name}">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: ${color}; flex-shrink: 0;"></span>
                                <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 160px; color: var(--text-secondary);">${item.name}</span>
                            </div>`;
                        }).join('');
                    }
                }

                // Update Hours Chart
                if (jiraHoursChart) {
                    const rawHoursData = data.monthly_hours.map(h => {
                        const totalVal = h.values.reduce((a, b) => a + b, 0);
                        return {
                            name: h.name,
                            value: totalVal
                        };
                    });
                    
                    const totalSum = rawHoursData.reduce((sum, item) => sum + item.value, 0);
                    
                    const hoursData = rawHoursData.map(item => {
                        return {
                            name: item.name,
                            value: totalSum > 0 ? parseFloat(((item.value / totalSum) * 100).toFixed(2)) : 0
                        };
                    });
                    
                    const centerValElement = document.getElementById('doughnut-jira-hours-center-value');
                    if (centerValElement) {
                        centerValElement.innerText = "100%";
                    }
                    
                    jiraHoursChart.data.labels = hoursData.map(item => item.name);
                    jiraHoursChart.data.datasets[0].data = hoursData.map(item => item.value);
                    jiraHoursChart.data.datasets[0].backgroundColor = hoursData.map((_, idx) => chartColors[(idx + 5) % chartColors.length]);
                    jiraHoursChart.update();
                    
                    const legendContainer = document.getElementById('doughnut-jira-hours-legend');
                    if (legendContainer) {
                        legendContainer.innerHTML = hoursData.map((item, idx) => {
                            const color = chartColors[(idx + 5) % chartColors.length];
                            const valPct = item.value.toFixed(1);
                            return `
                            <div style="display: flex; align-items: center; gap: 6px; cursor: pointer; transition: opacity 0.2s; font-size: 11px;" onclick="toggleChartSlice('jiraHoursChart', ${idx}, this)" title="${item.name}">
                                <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: ${color}; flex-shrink: 0;"></span>
                                <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 160px; color: var(--text-secondary);">${item.name}: ${valPct}%</span>
                            </div>`;
                        }).join('');
                    }
                }

                // Update Systems Chart
                if (jiraSystemsChart) {
                    jiraSystemsChart.data.labels = data.months;
                    data.monthly_systems.forEach((s, idx) => {
                        if (jiraSystemsChart.data.datasets[idx]) {
                            jiraSystemsChart.data.datasets[idx].data = s.values;
                        }
                    });
                    jiraSystemsChart.update();
                }
            }
            
            // Update Jira Completed Section
            const jiraCompletedData = {
                q1: {
                    total: 451,
                    trend: '+50%',
                    isPositive: true,
                    sparkline: [12, 18, 15, 22, 20, 26],
                    teams: {
                        camunda: 76,
                        devops: 60,
                        front: 46,
                        abs: 176,
                        callcenter: 38,
                        kfo: 62,
                        cards: 18,
                        scoring: 51
                    }
                },
                q2: {
                    total: 516,
                    trend: '-14.9%',
                    isPositive: false,
                    sparkline: [24, 18, 26, 20, 25, 22],
                    teams: {
                        camunda: 57,
                        devops: 55,
                        front: 50,
                        abs: 187,
                        callcenter: 36,
                        kfo: 105,
                        cards: 33,
                        scoring: 78
                    }
                },
                all: {
                    total: 967,
                    trend: '+14.4%',
                    isPositive: true,
                    sparkline: [12, 18, 15, 22, 20, 26, 24, 18, 26, 20, 25, 22],
                    teams: {
                        camunda: 133,
                        devops: 115,
                        front: 96,
                        abs: 363,
                        callcenter: 74,
                        kfo: 167,
                        cards: 51,
                        scoring: 129
                    }
                }
            };
            
            const completedInfo = jiraCompletedData[activePeriod];
            animateValue('jira-completed-total', 0, completedInfo.total, 800);
            
            const trendWrapper = document.getElementById('jira-completed-trend-wrapper');
            const trendVal = document.getElementById('jira-completed-trend-val');
            const trendIcon = document.getElementById('jira-completed-trend-icon');
            
            trendVal.innerText = completedInfo.trend;
            if (completedInfo.isPositive) {
                trendWrapper.style.color = '#10b981';
                trendIcon.style.color = '#10b981';
                trendIcon.style.transform = 'none';
                trendIcon.innerHTML = '<polyline points="18 15 12 9 6 15"></polyline>';
            } else {
                trendWrapper.style.color = '#ef4444';
                trendIcon.style.color = '#ef4444';
                trendIcon.style.transform = 'none';
                trendIcon.innerHTML = '<polyline points="6 9 12 15 18 9"></polyline>';
            }
            
            document.getElementById('jira-completed-camunda').innerText = formatNumber(completedInfo.teams.camunda);
            document.getElementById('jira-completed-devops').innerText = formatNumber(completedInfo.teams.devops);
            document.getElementById('jira-completed-front').innerText = formatNumber(completedInfo.teams.front);
            document.getElementById('jira-completed-abs').innerText = formatNumber(completedInfo.teams.abs);
            document.getElementById('jira-completed-callcenter').innerText = formatNumber(completedInfo.teams.callcenter);
            document.getElementById('jira-completed-kfo').innerText = formatNumber(completedInfo.teams.kfo);
            document.getElementById('jira-completed-cards').innerText = formatNumber(completedInfo.teams.cards);
            document.getElementById('jira-completed-scoring').innerText = formatNumber(completedInfo.teams.scoring);
            
            const sparklineDiv = document.getElementById('jira-completed-sparkline');
            if (sparklineDiv) {
                const maxVal = Math.max(...completedInfo.sparkline);
                sparklineDiv.innerHTML = completedInfo.sparkline.map(val => {
                    const heightPct = maxVal > 0 ? (val / maxVal) * 100 : 0;
                    const heightPx = Math.round((heightPct * 26) / 100) + 2;
                    return `<div style="width: 3px; height: ${heightPx}px; background-color: #10b981; border-radius: 1px; opacity: 0.85;"></div>`;
                }).join('');
            }

            // Update Jira Time Spent Section (Category percentages & Doughnut Chart)
            const jiraTimeSpentData = {
                q1: {
                    total: 713.9,
                    categories: [
                        { name: "Внутренняя автоматизация", value: 37.01, color: "#8b5cf6" },
                        { name: "Поддержка и инциденты", value: 13.13, color: "#ef4444" },
                        { name: "Новые продукты для клиентов", value: 34.29, color: "#10b981" },
                        { name: "Требования регулятора", value: 15.52, color: "#06b6d4" },
                        { name: "KPI-цели (KPI план бизнеса)", value: 0.04, color: "#f97316" },
                        { name: "Без категории", value: 0.02, color: "#64748b" }
                    ]
                },
                q2: {
                    total: 925.5,
                    categories: [
                        { name: "Поддержка и инциденты", value: 14.15, color: "#ef4444" },
                        { name: "Внутренняя автоматизация", value: 25.75, color: "#8b5cf6" },
                        { name: "Требования регулятора", value: 17.01, color: "#06b6d4" },
                        { name: "Новые продукты для клиентов", value: 41.34, color: "#10b981" },
                        { name: "KPI-цели (KPI план бизнеса)", value: 1.74, color: "#f97316" },
                        { name: "Без категории", value: 0.01, color: "#64748b" }
                    ]
                },
                all: {
                    total: 1639.4,
                    categories: [
                        { name: "Новые продукты для клиентов", value: 38.27, color: "#10b981" },
                        { name: "Внутренняя автоматизация", value: 30.66, color: "#8b5cf6" },
                        { name: "Требования регулятора", value: 16.36, color: "#06b6d4" },
                        { name: "Поддержка и инциденты", value: 13.71, color: "#ef4444" },
                        { name: "KPI-цели (KPI план бизнеса)", value: 1.00, color: "#f97316" },
                        { name: "Без категории", value: 0.01, color: "#64748b" }
                    ]
                }
            };
            
            const timeInfo = jiraTimeSpentData[activePeriod];
            
            // Update center text of timespent chart
            document.getElementById('doughnut-jira-time-center-value').innerText = timeInfo.total.toFixed(1);
            
            // Build Time Spent Table
            const timeTableBody = document.getElementById('jira-time-table-body');
            if (timeTableBody) {
                timeTableBody.innerHTML = timeInfo.categories.map(cat => {
                    return `
                    <div style="display: flex; align-items: center; justify-content: space-between; font-size: 0.85rem;">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background-color: ${cat.color}; flex-shrink: 0;"></span>
                            <span style="color: var(--text-secondary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 170px;" title="${cat.name}">${cat.name}</span>
                        </div>
                        <span style="font-weight: 700; color: var(--text-primary);">${cat.value.toFixed(2)}%</span>
                    </div>`;
                }).join('');
            }
            
            // Update Time spent doughnut chart
            const timeCtx = document.getElementById('chart-jira-time-spent').getContext('2d');
            const themeColors = getThemeColors();
            
            if (jiraTimeSpentChart) {
                jiraTimeSpentChart.destroy();
            }
            
            jiraTimeSpentChart = new Chart(timeCtx, {
                type: 'doughnut',
                data: {
                    labels: timeInfo.categories.map(c => c.name),
                    datasets: [{
                        data: timeInfo.categories.map(c => c.value),
                        backgroundColor: timeInfo.categories.map(c => c.color),
                        borderColor: themeColors.cardBorder,
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '70%',
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 10,
                            callbacks: {
                                label: function(context) {
                                    return ` ${context.label}: ${context.raw.toFixed(2)}%`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderJiraCategoriesChart(months, monthlyCategories) {
            const ctx = document.getElementById('chart-jira-categories').getContext('2d');
            const colors = getThemeColors();
            
            if (jiraCategoriesChart) {
                jiraCategoriesChart.destroy();
            }
            
            // Sum values for each category across all months
            const categoriesData = monthlyCategories.map(c => {
                const totalVal = c.values.reduce((a, b) => a + b, 0);
                return {
                    name: c.name,
                    value: totalVal
                };
            });
            
            const totalTasks = categoriesData.reduce((sum, item) => sum + item.value, 0);
            const centerValElement = document.getElementById('doughnut-jira-categories-center-value');
            if (centerValElement) {
                centerValElement.innerText = formatNumber(totalTasks);
            }
            
            jiraCategoriesChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: categoriesData.map(item => item.name),
                    datasets: [{
                        data: categoriesData.map(item => item.value),
                        backgroundColor: categoriesData.map((_, idx) => chartColors[idx % chartColors.length]),
                        borderColor: colors.cardBorder,
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '70%',
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = ((val / total) * 100).toFixed(1);
                                    return ` ${context.label}: ${formatNumber(val)} (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });

            // Generate HTML legend
            const legendContainer = document.getElementById('doughnut-jira-categories-legend');
            if (legendContainer) {
                legendContainer.innerHTML = categoriesData.map((item, idx) => {
                    const color = chartColors[idx % chartColors.length];
                    return `
                    <div style="display: flex; align-items: center; gap: 6px; cursor: pointer; transition: opacity 0.2s; font-size: 11px;" onclick="toggleChartSlice('jiraCategoriesChart', ${idx}, this)" title="${item.name}">
                        <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: ${color}; flex-shrink: 0;"></span>
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 160px; color: var(--text-secondary);">${item.name}</span>
                    </div>`;
                }).join('');
            }
        }

        function renderJiraHoursChart(months, monthlyHours) {
            const ctx = document.getElementById('chart-jira-hours').getContext('2d');
            const colors = getThemeColors();
            
            if (jiraHoursChart) {
                jiraHoursChart.destroy();
            }
            
            const rawHoursData = monthlyHours.map(h => {
                const totalVal = h.values.reduce((a, b) => a + b, 0);
                return {
                    name: h.name,
                    value: totalVal
                };
            });
            
            const totalSum = rawHoursData.reduce((sum, item) => sum + item.value, 0);
            
            const hoursData = rawHoursData.map(item => {
                return {
                    name: item.name,
                    value: totalSum > 0 ? parseFloat(((item.value / totalSum) * 100).toFixed(2)) : 0
                };
            });
            
            const centerValElement = document.getElementById('doughnut-jira-hours-center-value');
            if (centerValElement) {
                centerValElement.innerText = "100%";
            }
            
            jiraHoursChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: hoursData.map(item => item.name),
                    datasets: [{
                        data: hoursData.map(item => item.value),
                        backgroundColor: hoursData.map((_, idx) => chartColors[(idx + 5) % chartColors.length]),
                        borderColor: colors.cardBorder,
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '70%',
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw;
                                    return ` ${context.label}: ${val.toFixed(1)}%`;
                                }
                            }
                        }
                    }
                }
            });

            // Generate HTML legend
            const legendContainer = document.getElementById('doughnut-jira-hours-legend');
            if (legendContainer) {
                legendContainer.innerHTML = hoursData.map((item, idx) => {
                    const color = chartColors[(idx + 5) % chartColors.length];
                    const valPct = item.value.toFixed(1);
                    return `
                    <div style="display: flex; align-items: center; gap: 6px; cursor: pointer; transition: opacity 0.2s; font-size: 11px;" onclick="toggleChartSlice('jiraHoursChart', ${idx}, this)" title="${item.name}">
                        <span style="display: inline-block; width: 6px; height: 6px; border-radius: 50%; background-color: ${color}; flex-shrink: 0;"></span>
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 160px; color: var(--text-secondary);">${item.name}: ${valPct}%</span>
                    </div>`;
                }).join('');
            }
        }

        function renderJiraSystemsChart(months, monthlySystems) {
            const ctx = document.getElementById('chart-jira-systems').getContext('2d');
            const colors = getThemeColors();
            
            if (jiraSystemsChart) {
                jiraSystemsChart.destroy();
            }
            
            jiraSystemsChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: months,
                    datasets: monthlySystems.map((s, idx) => ({
                        label: s.name,
                        data: s.values,
                        backgroundColor: chartColors[(idx + 2) % chartColors.length],
                        borderColor: colors.cardBorder,
                        borderWidth: 1,
                        borderRadius: 4
                    }))
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            grid: { display: false },
                            ticks: { color: colors.textSecondary, font: { size: 10, family: 'Inter' } }
                        },
                        y: {
                            grid: { color: colors.gridColor },
                            ticks: { color: colors.textSecondary, font: { size: 10, family: 'Inter' } }
                        }
                    },
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: colors.textSecondary,
                                font: { size: 10, family: 'Inter' },
                                boxWidth: 10,
                                padding: 12,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#0f172a',
                            titleColor: '#fff',
                            bodyColor: '#e2e8f0',
                            borderColor: 'rgba(255,255,255,0.1)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    return ` ${context.dataset.label}: ${formatNumber(context.raw)} задач`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function updateJiraChartsTheme() {
            const colors = getThemeColors();
            
            if (jiraCategoriesChart) {
                if (jiraCategoriesChart.options.scales) {
                    if (jiraCategoriesChart.options.scales.x) jiraCategoriesChart.options.scales.x.ticks.color = colors.textSecondary;
                    if (jiraCategoriesChart.options.scales.y) {
                        jiraCategoriesChart.options.scales.y.ticks.color = colors.textSecondary;
                        jiraCategoriesChart.options.scales.y.grid.color = colors.gridColor;
                    }
                }
                if (jiraCategoriesChart.options.plugins && jiraCategoriesChart.options.plugins.legend) {
                    jiraCategoriesChart.options.plugins.legend.labels.color = colors.textSecondary;
                }
                if (jiraCategoriesChart.data.datasets[0]) {
                    jiraCategoriesChart.data.datasets[0].borderColor = colors.cardBorder;
                }
                jiraCategoriesChart.update();
            }
            if (jiraHoursChart) {
                if (jiraHoursChart.options.scales) {
                    if (jiraHoursChart.options.scales.x) jiraHoursChart.options.scales.x.ticks.color = colors.textSecondary;
                    if (jiraHoursChart.options.scales.y) {
                        jiraHoursChart.options.scales.y.ticks.color = colors.textSecondary;
                        jiraHoursChart.options.scales.y.grid.color = colors.gridColor;
                    }
                }
                if (jiraHoursChart.options.plugins && jiraHoursChart.options.plugins.legend) {
                    jiraHoursChart.options.plugins.legend.labels.color = colors.textSecondary;
                }
                if (jiraHoursChart.data.datasets[0]) {
                    jiraHoursChart.data.datasets[0].borderColor = colors.cardBorder;
                }
                jiraHoursChart.update();
            }
            if (jiraSystemsChart) {
                jiraSystemsChart.options.scales.x.ticks.color = colors.textSecondary;
                jiraSystemsChart.options.scales.y.ticks.color = colors.textSecondary;
                jiraSystemsChart.options.scales.y.grid.color = colors.gridColor;
                jiraSystemsChart.options.plugins.legend.labels.color = colors.textSecondary;
                jiraSystemsChart.update();
            }
            if (jiraTimeSpentChart) {
                if (jiraTimeSpentChart.data.datasets[0]) {
                    jiraTimeSpentChart.data.datasets[0].borderColor = colors.cardBorder;
                }
                jiraTimeSpentChart.update();
            }
        }

        // Page Load Init
        window.addEventListener('DOMContentLoaded', () => {
            // Apply saved theme
            const savedTheme = localStorage.getItem('theme');
            if (savedTheme === 'light') {
                document.body.classList.add('light-theme');
                updateThemeUI(true);
            }
            
            updateDashboard();
        });
    </script>
</body>
</html>
"""
    
    # Inject aggregated data
    html_content = html_template.replace("##DATA_PLACEHOLDER##", json.dumps(stats_data, ensure_ascii=False, indent=2))
    
    # Write index.html
    output_file = "index.html"
    print(f"Writing {output_file}...")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    print("Dashboard index.html generated successfully with Theme Toggle, Substance of Theme, MicroCredit statistics, and Monthly Grouped Chart!")

if __name__ == "__main__":
    main()
