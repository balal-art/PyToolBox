import pandas as pd
from tkinter import Tk, filedialog, messagebox

# تنظیمات اولیه پنجره tkinter
root = Tk()
root.withdraw()

def show_error(title, message):
    """نمایش خطا در پنجره پیام"""
    messagebox.showerror(title, message)
    root.destroy()

def show_info(title, message):
    """نمایش اطلاعات در پنجره پیام"""
    messagebox.showinfo(title, message)

def extract_codes(row_code):
    """
    استخراج کد فصل، کد رشته و کد رسته از ردیف فهرست بها
    با فرمت:
    - کد فصل: 4 تا 6 (از راست) -> str_code[-6:-4]
    - کد رشته: 7 (از راست) -> str_code[-7]
    - کد رسته: 7 تا آخر (از راست) -> str_code[:-7]
    """
    try:
        str_code = str(row_code).strip()
        length = len(str_code)
        
        # کد فصل (4 تا 6 از راست)
        chapter_code = str_code[-6:-4] if length >= 6 else '00'
        
        # کد رشته (7 از راست)
        field_code = str_code[-7] if length >= 7 else '0'
        
        # کد رسته (7 تا آخر از راست)
        category_code = str_code[:-7] if length > 7 else '0'
        
        # اضافه کردن صفر به کدهای یک رقمی
        return (
            chapter_code.zfill(2),
            field_code.zfill(2),
            category_code.zfill(2)
        )
    except Exception as e:
        print(f"خطا در استخراج کدها: {e}")
        return 'A1', 'B2', 'C3'

# --- انتخاب فایل ورودی ---
try:
    input_file_path = filedialog.askopenfilename(
        title="انتخاب فایل ورودی",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    
    if not input_file_path:
        show_info("توجه", "هیچ فایلی انتخاب نشد!")
        exit()

    # --- خواندن داده‌ها ---
    try:
        df_source = pd.read_excel(input_file_path)
        print("\nستون‌های موجود در فایل:")
        print(df_source.columns.tolist())
    except Exception as e:
        show_error("خطای خواندن فایل", f"خطا در خواندن فایل اکسل:\n{str(e)}")
        exit()

    # --- پردازش داده‌ها ---
    book_mapping = {'105': '85'}

    def convert_fihrist_code(source_code, mapping):
        try:
            str_code = str(source_code)
            if len(str_code) < 6:
                return str_code
            book_code = str_code[:-6]
            r_code = str_code[-6:]
            return mapping.get(book_code, book_code) + r_code
        except Exception as e:
            show_error("خطای تبدیل کد", f"خطا در تبدیل کد فهرست بها:\n{str(e)}")
            return source_code

    # ایجاد دیتافریم خروجی
    df_dest = pd.DataFrame()

    # نگاشت ستون‌ها با قابلیت انعطاف
    column_mapping = {
        'ردیف': ['ردیف'],
        'ردیف فهرست بها': ['شماره فهرست', 'شماره فهرست بها','کد ردیف'],
        'مقدار': ['مقدار'],
        'کارکرد /پای کار': ['نوع ردیف', 'نوع ردیف', 'کارکرد'],
        'ف ج *': ['نوع قيمت', 'نوع قیمت'],
        'بهای واحد برآورد': ['بهای واحد برآورد (پایه)', 'بهای واحد برآورد','بهای واحد فهرست بها'],
        'بهای واحد پیشنهاد': ['بهای واحد پیشنهاد'],
        'ضریب برآورد': ['ضریب خاص برآورد', 'ضریب برآورد'],
        'بخش': ['بخش( فعالیت)', 'بخش'],
        'ملاحظات': ['ملاحظات'],
        'واحد': ['واحد'],
        'شرح ردیف': ['شرح ردیف'],
    }

    for dest_col, possible_cols in column_mapping.items():
        found = False
        for col in possible_cols:
            if col in df_source.columns:
                df_dest[dest_col] = df_source[col]
                found = True
                break
        
        if not found:
            show_info("ستون یافت نشد", 
                     f"ستون {possible_cols} یافت نشد - ستون {dest_col} خالی می‌ماند")

    # تبدیل کد فهرست بها
    if 'ردیف فهرست بها' in df_dest.columns:
        df_dest['ردیف فهرست بها'] = df_dest['ردیف فهرست بها'].apply(
            lambda x: convert_fihrist_code(x, book_mapping)
        )

    # --- پردازش بر اساس df_dest ---
    
    # 1. جایگزینی مقادیر خالی در 'بهای واحد پیشنهاد'
    if 'بهای واحد برآورد' in df_dest.columns and 'بهای واحد پیشنهاد' in df_dest.columns:
        mask = (df_dest['بهای واحد پیشنهاد'].isna()) | (df_dest['بهای واحد پیشنهاد'] == 0) | (df_dest['بهای واحد پیشنهاد'] == '')
        df_dest.loc[mask, 'بهای واحد پیشنهاد'] = df_dest.loc[mask, 'بهای واحد برآورد']
    
    # 2. پردازش ردیف‌های غیر ستاره‌دار (نوع ف ج)
    if 'ف ج *' in df_dest.columns:
        mask = df_dest['ف ج *'] != '*'
        
        #  غیر حذف شرح برای ردیف‌های *
        if 'شرح ردیف' in df_dest.columns:
            df_dest.loc[mask, 'شرح ردیف'] = ''
        
        #  غیر حذف واحد برای ردیف‌های *
        if 'واحد' in df_dest.columns:
            df_dest.loc[mask, 'واحد'] = ''

    # ستون‌های اختیاری
    optional_cols = {
        'عبارت مقدار': '',
        'عبارت ضریب برآورد': '',
        'عبارت ضریب پیشنهاد': '',
        'عنوان سازمان': '',
        'ضریب پیشنهاد': '',
        'حاصل برآورد': '',
        'حاصل پیشنهاد': '',
        '% اضافه یا تخفیف': '',
        'کد سازمان': ''
    }

    for col, default_val in optional_cols.items():
        if col not in df_dest.columns:
            df_dest[col] = default_val


    # --- ایجاد دیتابیس دوم برای ردیف‌های * یا ج ---
    if 'ف ج *' in df_dest.columns:
        # ایجاد شرط برای ردیف‌های * یا ج
        mask = df_dest['ف ج *'].isin(['*', 'ج'])
        
        # ایجاد DataFrame جدید
        df_Star = pd.DataFrame()
        
        # نگاشت ستون‌های مورد نیاز
        column_mapping_star = {
            'شماره ردیف': 'ردیف فهرست بها',
            'واحد': 'واحد',
            'بهای واحد': 'بهای واحد برآورد',
            'شرح': 'شرح ردیف'
        }
        
        # پر کردن df_Star با داده‌های فیلتر شده
        for dest_col, source_col in column_mapping_star.items():
            if source_col in df_dest.columns:
                df_Star[dest_col] = df_dest.loc[mask, source_col]
            else:
                df_Star[dest_col] = ''
        
        
        # استخراج کدها با استفاده از تابع جدید
        if 'شماره ردیف' in df_Star.columns:
            df_Star['کد فصل'], df_Star['کد رشته'], df_Star['کد رسته'] = zip(
                *df_Star['شماره ردیف'].apply(extract_codes)
            )
        else:
            df_Star['کد فصل'] = '??'
            df_Star['کد رشته'] = '??'
            df_Star['کد رسته'] = '??'    
            
        # مرتب کردن ستون‌ها به ترتیب مورد نظر
        df_Star = df_Star[['شماره ردیف', 'واحد', 'بهای واحد', 'کد فصل', 'کد رسته', 'کد رشته', 'شرح']]


    # --- ذخیره فایل‌ها ---
    output_file_path = filedialog.asksaveasfilename(
        title="ذخیره فایل خروجی اصلی",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if output_file_path:
        try:
            # ذخیره فایل اصلی
            df_dest.to_excel(output_file_path, index=False)
            
            # ذخیره فایل ردیف‌های ویژه در همان مسیر با پسوند _Star
            star_file_path = output_file_path.replace('.xlsx', '_Star.xlsx')
            df_Star.to_excel(star_file_path, index=False)
            
            show_info("موفقیت", 
                     f"فایل‌ها با موفقیت ذخیره شدند:\n\n"
                     f"فایل اصلی: {output_file_path}\n"
                     f"فایل ردیف‌های ویژه: {star_file_path}")
        except Exception as e:
            show_error("خطای ذخیره فایل", f"خطا در ذخیره فایل:\n{str(e)}")
    else:
        show_info("توجه", "ذخیره فایل لغو شد.")


except Exception as e:
    show_error("خطای کلی برنامه", f"خطای غیرمنتظره:\n{str(e)}")
finally:
    root.destroy()