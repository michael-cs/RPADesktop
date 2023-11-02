import pyautogui as pg
import pandas as pd
import psutil
import time
import sys
import os


class DesktopTools:
    def inicia_software():
        for process in psutil.process_iter(['name']):
            if process.info['name'] == 'Fakturama.exe':
                return True
        return False

    def try_locate_image(imagePath, try_count=0, tries=5):
        while try_count >= 0:
            position = pg.locateOnScreen(imagePath, grayscale=True, confidence=0.7)
            time.sleep(1)
            try_count += 1
            print(try_count)
            if try_count >= tries or position is not None:
                break
        try:
            if position is not None:
                print(f"position = {position}")
                return position
            else:
                raise Exception(f'Imagem: "{imagePath}", não localizada')
        except Exception as error:
            print(error)
            pg.screenshot("./assets/images/ERROR_screenshot.png")
            sys.exit()


class FakturamaActivities:
    def cadastra_produtos():
        if not DesktopTools.inicia_software():
            os.startfile(r"C:\Program Files\Fakturama2\Fakturama.exe")

        df = pd.read_excel(r"C:\Users\55549\Desktop\RPA Desktop\assets\fakturama.xlsx")
        print(df.head())

        new_product = DesktopTools.try_locate_image(r".\assets\images\btn_new_product.PNG", tries=30)
        for i, r in df.iterrows():
            item_number = str(r["Item Number"])
            product_name = r["Name"]
            category = r["Category"]
            gtin = str(r["GTIN"])
            description = r["Description"]
            notice = r["Notice"]
            if new_product is not None:
                pg.click(new_product, interval=2)
                label = DesktopTools.try_locate_image(r".\assets\images\label_new_product.PNG")
                pg.click(label, interval=2)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(item_number)
                print(item_number)
                pg.press('tab', 1, interval=0.5)
                pg.typewrite(product_name)
                print(product_name)
                pg.press('tab', 1, interval=0.5)
                pg.typewrite(category)
                print(category)
                pg.press('tab', 1, interval=0.5)
                pg.typewrite(gtin)
                print(gtin)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(description)
                print(description)
                pg.press('tab', 10, interval=0.5)
                pg.typewrite(notice)
                pg.hotkey('ctrl', 's')
                pg.hotkey('ctrl', 'w')
        pg.hotkey('alt', 'F4')

    def cadastra_contato():
        if not DesktopTools.inicia_software():
            os.startfile(r"C:\Program Files\Fakturama2\Fakturama.exe")

        df = pd.read_csv(r".\assets\fake_data.csv")
        print(df.head())

        new_contact = DesktopTools.try_locate_image(r".\assets\images\btn_new_contact.PNG", tries=30)
        for i, r in df.iterrows():
            first_name = r.iloc[0]
            last_name = r.iloc[1]
            company = r.iloc[2]
            address = r.iloc[4]
            email = r.iloc[5]
            phone = str(r.iloc[6])
            if new_contact is not None:
                pg.click(new_contact, interval=2)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(company)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(first_name)
                pg.press('tab')
                pg.typewrite(last_name)
                pg.press('tab', 5, interval=0.5)
                pg.typewrite(address)
                pg.press('tab', 7, interval=0.5)
                pg.typewrite(email)
                pg.press('tab')
                pg.typewrite(phone)
                pg.hotkey("ctrl", "s")
                pg.hotkey('ctrl', 'w')
        pg.hotkey('alt', 'F4')


FakturamaActivities.cadastra_contato()
