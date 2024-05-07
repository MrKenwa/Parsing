import openpyxl
import time

from selenium import common
from selenium.webdriver.common.by import By
from selenium import webdriver
from openpyxl.styles import Alignment


def parse(el):
    prod = []

    title = el.find_element(by=By.CLASS_NAME, value="product-card__name").text.replace("/", "").strip()
    good_url = el.find_element(by=By.CLASS_NAME, value="product-card__link ").get_attribute("href")
    low_price = el.find_element(by=By.CLASS_NAME, value="price__lower-price ").text.replace(" ", "")[:-1]

    try:
        high_price = el.find_element(by=By.CLASS_NAME, value="price__wrap").find_element(by=By.TAG_NAME, value="del").text.replace(" ", "")[:-1]
    except (common.NoSuchElementException):
        high_price = ''

    date = el.find_element(by=By.CLASS_NAME, value="btn-text").text.strip()
    rate = el.find_element(by=By.CLASS_NAME, value="address-rate-mini ").text.strip()
    if rate == "":
        rate = "нет оценок"

    service1 = webdriver.ChromeService(executable_path='chromedriver.exe')
    item = webdriver.Chrome(service=service1)
    item.get(good_url)
    scroll(2, 1, item)

    images_urls = []
    try:
        images_unhandle = (item.find_element(by=By.CLASS_NAME, value="product-page__slider-wrap").find_elements(by=By.TAG_NAME, value="li"))
        for image in images_unhandle:
            image_url = image.find_element(by=By.TAG_NAME, value="img").get_attribute("src")
            images_urls.append(image_url)
    except (common.NoSuchElementException):
        try:
            img = item.find_element(by=By.CLASS_NAME, value="zoom-image-container").find_element(by=By.TAG_NAME, value="img").get_attribute("src")
            images_urls.append(img)
        except:
            images_urls.append("нет картинок")
    
    try:
        nowallet_price = item.find_element(by=By.CLASS_NAME, value="price-block__final-price ").text.replace(" ", "")[:-1]
    except:
        nowallet_price = ""
    try:
        seller = item.find_element(by=By.CLASS_NAME, value="seller-info__name").text.strip()
    except:
        seller = "no name"
    try:
        seller_rate = item.find_element(by=By.CLASS_NAME, value="seller-info__param").find_element(by=By.CLASS_NAME, value="address-rate-mini ").text.strip()
        if seller_rate == "":
            seller_rate = "нет оценок"
    except:
        seller_rate = "нет оценок"
    
    prod += [title, good_url, nowallet_price, low_price, high_price, date, rate, seller, seller_rate]
    prod += images_urls
    return prod


def get_goods(driver):
    scroll(20, 1, driver)
    goods = (driver.find_elements(by=By.CLASS_NAME, value="product-card "))
    return goods


def scroll(count, delay, driver):
    for i in range(count):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(delay)


def main():
    service = webdriver.ChromeService(executable_path='chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Название", "Ссылка", "Цена без ВБ кошелька", "Цена со скидкой", "Цена без скидки", "Срок доставки", 
                "Рейтинг", "Продавец", "Рейтинг продавца", "Картинки:"])
    
    for i in range(1, 5):
        url = f'https://www.wildberries.ru/catalog/0/search.aspx?page={i}&sort=popular&search=%D1%88%D0%B8%D0%BB%D1%8C%D0%B4%D0%B8%D0%BA+amg'
        driver.get(url)
        goods = get_goods(driver)
        
        for good in goods:
            data = parse(good)
            ws.append(data)
    
    for row in ws.iter_rows():  
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.column_dimensions[cell.column_letter].auto_size = True

    wb.save("goods.xlsx")


if __name__ == "__main__":
    main()