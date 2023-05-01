from config import *
from google_api import GoogleAPI
from prepare_data import PrepareData


def run():
    data = PrepareData(STOCKS_PATH, ASSORTMENT_PATH).get_prepared_data()

    google = GoogleAPI(CREDENTIALS_FILE)
    categories = list(set([item['category'] for item in data]))
    google.create_spreadsheet(categories)
    google.share_with_anybody_for_writing()

    for item in categories:
        if item != 'Расходные материалы':
            category_data = []
            for item_data in data:
                if item_data['category'] == item:
                    category_data.append(item_data)
            google.add_data(category_data)
        else:
            subcategory_data = []
            for item_data in data:
                if item_data['category'] == item:
                    subcategory_data.append(item_data)
            google.add_data_for_rashod(subcategory_data)
    result = google.get_document_url()
    return result


if __name__ == '__main__':
    print(run())
