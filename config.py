import logging

logging.basicConfig(level=logging.INFO, filename="logs.log", filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

# путь к файлу с закрытым ключом сервисного аккаунта
CREDENTIALS_FILE = 'laone-test-week-de64869cce1d.json'

# Пути к файлам stocks.json, assortment.json
STOCKS_PATH = 'data/stocks.json'
ASSORTMENT_PATH = 'data/assortment.json'
