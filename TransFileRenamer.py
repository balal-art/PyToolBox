import pdfplumber
import pandas as pd
from arabic_reshaper import reshape
from bidi.algorithm import get_display
import re
import os
import sys

# چاپ نسخه کتابخانه pdfplumber
print(pdfplumber.__version__)
# ورژن در ویندوز 0.11.5
#ورژن لینوکس ddddddddddddddd
# تابعی برای اصلاح متن فارسی که شامل اعداد است
def _fix_farsi_with_numbers(text):
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
def _text_to_list(text):
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

def _is_valid_input(page_ranges):
    import re
    pattern = r'^(\d+|\d+-\d+)(,(\d+|\d+-\d+))*$'
    return re.match(pattern, page_ranges) is not None

def _generate_page_list(page_ranges, start,end):
    if not _is_valid_input(page_ranges):
        raise ValueError("ورودی نامعتبر است. لطفاً ورودی را بررسی کنید.")
    pages = set()
    for part in page_ranges.split(','):
        if '-' in part:
            start_range, end_range = map(int, part.split('-'))
            if start_range < start or end_range > end:
                raise ValueError(f"محدوده صفحات باید بین {start} و {end} باشد.")
            pages.update(range(start_range, end_range + 1))
        else:
            page = int(part)
            if page < start or page > end:
                raise ValueError(f"صفحات باید بین {start} و {end} باشند.")
            pages.add(page)
    return sorted(pages)

def pdf2table(pdf_path, save_path, pageList, keyword,fix_farsi=True, debug=False):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        # شماره صفحات مورد نظر
        page_list =_generate_page_list(pageList,1,len(pdf.pages))
        for page in pdf.pages:
            # بررسی اینکه آیا صفحه در لیست صفحات مورد نظر است
            if page.page_number in page_list:
                # چاپ شماره صفحه
                if debug: print(page.page_number)
                # استخراج متن صفحه و تبدیل آن به لیستی از خطوط
                text_lines = _text_to_list(page.extract_text_simple())
                if debug: print(text_lines)
                # بررسی وجود کلمه مورد نظر در متن صفحه
                page_titel = ''
                for text in text_lines:
                    if keyword in text:
                        if fix_farsi: page_titel =_fix_farsi_with_numbers(text)
                        else: page_titel = text
                if page_titel:
                    # استخراج جدول از صفحه
                    table = page.extract_table()
                    if table:
                        if fix_farsi:
                            # پردازش هر ردیف جدول
                            for row in table:
                                re_row = []
                                for cell in row:
                                    # اصلاح متن فارسی با اعداد در سلول
                                    re_cell = _fix_farsi_with_numbers(cell)
                                    re_row.append(re_cell)
                                # افزودن ردیف اصلاح‌شده به داده‌ها
                                data.append(re_row)
                        else:
                            data.extend(table)
                    # چاپ شماره صفحه
                    print(F'{page.page_number} is ok')
    
    # تبدیل داده‌ها به DataFrame و ذخیره به فایل Excel
    df = pd.DataFrame(data)
    if debug: print(data)
    df.to_excel(save_path, index=False, header=False)

if __name__ == '__main__':
    path = r'C:\Users\Ali\Desktop\Tadil-14031001.pdf'
    save_path = r'C:\Users\Ali\Desktop\output.xlsx'
    pageList = "6,5,4"
    keyword = "ﻪﯾﺎﭘ یﺎﻬﺑ یﺎﻫ ﺖﺳﺮﻬﻓ یا ﻪﺘﺷر"
    print(keyword)
    keyword2 = " رشته ای فهرست های بهای پایه"
    keyword = reshape(keyword2)
    print(keyword)
    # print(_fix_farsi_with_numbers(keyword))

    pdf2table(path,save_path,pageList,keyword,fix_farsi=False,debug=False)
