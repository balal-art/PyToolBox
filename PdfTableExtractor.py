import pdfplumber
import pandas as pd
import arabic_reshaper # type: ignore
import re

# چاپ نسخه کتابخانه pdfplumber
print(pdfplumber.__version__)

# تابعی برای اصلاح متن فارسی که شامل اعداد است
def fix_farsi_with_numbers(text):
    """
    ابتدا کاراکترهای هر کلمه را برعکس می‌کند، سپس ترتیب کلمات را برعکس می‌کند.
    اعداد بدون تغییر باقی می‌مانند.
    """
    if not isinstance(text, str):
        return text
    
    # تبدیل چندین فاصله به یک فاصله و حذف فاصله‌های اضافی
    text = ' '.join(text.split())
    
    # تقسیم متن به کلمات
    words = text.split(' ')
    # پردازش هر کلمه - برعکس کردن کاراکترها ولی نگه داشتن اعداد به همان صورت
    fixed_words = []
    for word in words:
        # تقسیم کلمه به قسمت‌های عدد و غیر عدد
        parts = re.findall(r'\d+|[^\d]+', word)
        fixed_parts = []
        for part in parts:
            if part.isdigit():
                fixed_parts.append(part)  # نگه داشتن اعداد به همان صورت
            else:
                fixed_parts.append(part[::-1])  # برعکس کردن کاراکترها در متن
        fixed_words.append(''.join(fixed_parts))
    
    # برعکس کردن ترتیب کلمات
    fixed_words = fixed_words[::-1]
    # اتصال کلمات به یکدیگر با استفاده از فاصله‌ها
    return ' '.join(fixed_words)

# تابعی برای تبدیل متن به لیستی از خطوط
def text_to_list(text):
    """
    تبدیل متن به لیستی از خطوط
    """
    if not isinstance(text, str):
        return []

    # تقسیم متن به خطوط با استفاده از کاراکتر newline
    lines = text.split('\n')

    # حذف خطوط خالی
    lines = [line.strip() for line in lines if line.strip()]

    return lines

# مسیر فایل PDF
pdf_path = r'/home/ali/Desktop/Tadil-14031001.pdf'
# شماره صفحات مورد نظر
pages_list = [7, 8]

data = []
keyword = "ﻪﯾﺎﭘ ﺪﺣاو یﺎﻬﺑ ﺖﺳﺮﻬﻓ یا ﻪﺘﺷر یﻠﺼﻓ یﺎﻫ ﺺﺧﺎﺷ"

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        # بررسی اینکه آیا صفحه در لیست صفحات مورد نظر است
        if page.page_number in pages_list:
            # چاپ شماره صفحه
            print(page.page_number)
            # استخراج متن صفحه و تبدیل آن به لیستی از خطوط
            text_lines = text_to_list(page.extract_text_simple())
            # بررسی وجود کلمه مورد نظر در متن صفحه
            page_titel = ''
            for text in text_lines:
                if keyword in text: page_titel = fix_farsi_with_numbers(text)
            if page_titel:
                # استخراج جدول از صفحه
                table = page.extract_table()
                if table:
                    print('table')
                    # پردازش هر ردیف جدول
                    for row in table:
                        re_row = []
                        for cell in row:
                            # اصلاح متن فارسی با اعداد در سلول
                            re_cell = fix_farsi_with_numbers(cell)
                            re_row.append(re_cell)
                        # افزودن ردیف اصلاح‌شده به داده‌ها
                        data.append(re_row)

# تبدیل داده‌ها به DataFrame و ذخیره به فایل Excel
df = pd.DataFrame(data)
df.to_excel(r'/home/ali/Desktop/output.xlsx', index=False, header=False)
