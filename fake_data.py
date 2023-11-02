from navigation import Browser, PageObjects
from file_manipulation import cria_csv, escreve_csv


def fake_data():
    site_data = "https://www.fakenamegenerator.com/gen-random-br-br.php"
    file_path = './assets/fake_data.csv'

    driver = Browser.chrome_browser(site_data)

    cria_csv(file_path)

    for i in range(5):
        row = (PageObjects.executa_fake_data(driver))
        escreve_csv(file_path, row)


fake_data()
