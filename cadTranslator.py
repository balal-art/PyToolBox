import os
import asyncio
import ezdxf
from googletrans import Translator

async def translate_text(text, translator):
    translated = await translator.translate(text, src='zh-cn', dest='fa')
    return translated.text

async def main(file_path):
    try:
        if not file_path.lower().endswith('.dxf'):
            raise ValueError("فایل انتخاب شده فایل DXF نیست.")
        
        dwg = ezdxf.readfile(file_path)
        modelspace = dwg.modelspace()
        
        translator = Translator()
        tasks = []

        for text_entity in modelspace.query('TEXT'):
            original_text = text_entity.dxf.text
            print(f"Original: {original_text}")
            tasks.append(translate_text(original_text, translator))

        translations = await asyncio.gather(*tasks)

        for text_entity, translated_text in zip(modelspace.query('TEXT'), translations):
            text_entity.dxf.text = translated_text

        directory,file = os.path.split(file_path)
        output_path =os.path.join(directory, "translated_" + file)
        dwg.saveas(output_path)
        print(f"فایل ترجمه شده ذخیره شد به: {output_path}")

    except ValueError as ve:
        print(f"خطا: {ve}")
    except IOError:
        print("خطا در باز کردن یا خواندن فایل. لطفاً مسیر فایل را بررسی کنید.")
    except Exception as e:
        print(f"خطای نامشخص: {e}")

# درخواست مسیر فایل از کاربر
file_path = input("لطفاً مسیر فایل DXF خود را وارد کنید: ")

asyncio.run(main(file_path))
