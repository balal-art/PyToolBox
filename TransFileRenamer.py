import os
import sys
import subprocess
import asyncio
from googletrans import Translator
import re

# تابع غیرهمزمان برای ترجمه نام فایل‌ها
async def translate_file_names(file_info, src_lang, dest_lang):
    translator = Translator()  # ایجاد یک نمونه از Translator
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]+')  # الگوی جستجوی کاراکترهای چینی
    
    for file, name in file_info.items():
        print(f'file: {file} // name: {name}')
        if chinese_pattern.search(name):  # اگر نام فایل شامل کاراکترهای چینی باشد
            translation_name = await translator.translate(name, src=src_lang, dest=dest_lang)  # ترجمه نام فایل
            file_info[file] = translation_name.text  # به‌روزرسانی نام فایل
            print(f'Translating >> {file} ---> {file_info[file]}')

# تابع برای به‌روزرسانی نام فایل‌ها
def update_filenames(directory, file_info):
    for file, new_name in file_info.items():
        old_file_path = os.path.join(directory, file)  # مسیر فایل قدیمی
        new_file_name = new_name + os.path.splitext(file)[1]  # ساخت نام فایل جدید با پسوند
        new_file_path = os.path.join(directory, new_file_name)  # مسیر فایل جدید
        os.rename(old_file_path, new_file_path)  # تغییر نام فایل
        print(f'{file} --> {new_file_name}')

# تابع برای ذخیره اطلاعات فایل‌ها در فایل متنی
def save_file_info(file_info, filepath):
    with open(filepath, 'w', encoding='utf-8') as f:
        for file, name in file_info.items():
            f.write(f'{file} --> {name}\n')  # نوشتن نام‌های قدیمی و جدید فایل‌ها در فایل متنی
    print(f'File info saved to {filepath}')
    # باز کردن فایل با استفاده از ویرایشگر متن پیش‌فرض سیستم
    if os.name == 'nt':  # برای ویندوز
        os.startfile(filepath)
    elif os.name == 'posix':  # برای macOS و لینوکس
        subprocess.call(('open', filepath) if sys.platform == 'darwin' else ('xdg-open', filepath))

# تابع برای بارگذاری اطلاعات فایل‌ها از فایل متنی
def load_file_info(filepath):
    new_file_info = {}
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            old_name, new_name = line.strip().split(' --> ')  # جدا کردن نام‌های قدیمی و جدید فایل‌ها
            new_file_info[old_name] = new_name  # به‌روزرسانی دیکشنری با نام‌های جدید
    return new_file_info

# تابع اصلی برای تغییر نام فایل‌ها
def rename_files(directory, src_lang='zh-cn', dest_lang='fa'):
    file_info = {}
    
    if os.path.isdir(directory):
        for file in os.listdir(directory):
            if os.path.isfile(os.path.join(directory, file)):
                name, ext = os.path.splitext(file)  # جدا کردن نام فایل و پسوند آن
                file_info[file] = name  # اضافه کردن نام فایل به دیکشنری
        asyncio.run(translate_file_names(file_info, src_lang, dest_lang))  # اجرای تابع ترجمه به صورت غیرهمزمان
        
        txt_filepath = os.path.join(directory, 'file_renames.txt')  # مسیر فایل متنی برای ذخیره نام‌های جدید
        save_file_info(file_info, txt_filepath)  # ذخیره نام‌های جدید در فایل متنی
        
        input("Please review the file_renames.txt file and press Enter to apply the changes...")  # درخواست از کاربر برای بررسی فایل متنی
        
        new_file_info = load_file_info(txt_filepath)  # بارگذاری اطلاعات از فایل متنی
        update_filenames(directory, new_file_info)  # به‌روزرسانی نام فایل‌ها
    else:
        print("پوشه‌ای انتخاب نشد.")

# اجرای برنامه به صورت مستقیم
if __name__ == '__main__':
    directory = input("Please enter the directory path: ")  # درخواست مسیر پوشه از کاربر
    src_lang = input("Please enter the source language (default: 'zh-cn'): ") or 'zh-cn'  # درخواست زبان مبدأ از کاربر
    dest_lang = input("Please enter the destination language (default: 'fa'): ") or 'fa'  # درخواست زبان مقصد از کاربر
    rename_files(directory, src_lang, dest_lang)  # اجرای تابع تغییر نام فایل‌ها
