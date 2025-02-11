import camelot
import pandas as pd
import matplotlib.pyplot as plt

pdf_path = r'C:\Users\Ali\Desktop\Tadil-14031001.pdf'
lpage = '7'  # صفحات مورد نظر

# محدوده خوانش جدول (X1, Y1, X2, Y2) به واحد نقاط (points) ['54,144,594,648']
table_regions = ['10,10,640,790']  # این اعداد را بر اساس نیاز خود تنظیم کنید

# استخراج جداول از صفحات مشخص شده با استفاده از table_areas
tables = camelot.read_pdf(pdf_path, pages=lpage, table_regions=table_regions)
print(tables)
# نمایش جزئیات جدول اول با نمودار شبکه‌ای
camelot.plot(tables[0], kind='grid').show()

# جلوگیری از بسته شدن پنجره
plt.show(block=True)

# ترکیب تمام جداول استخراج شده
all_data = []
for table in tables:
    all_data.extend(table.df.values.tolist())

# ایجاد DataFrame با pandas
df = pd.DataFrame(all_data)

# ذخیره DataFrame به فایل CSV
df.to_csv(r'C:\Users\Ali\Desktop\output.csv', index=False, header=False)

print("جداول با موفقیت استخراج و ذخیره شد.")
