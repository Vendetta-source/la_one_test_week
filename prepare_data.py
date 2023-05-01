import json
from config import *


class PrepareData:
    """
    Class for prepare data from stocks.json and assortment.json for Google Sheet
    """
    def __init__(self, stocks_path, assortment_path):
        self.all_data = []
        try:
            self.stocks = json.load(open(stocks_path, 'r', encoding='utf-8'))
            self.assortment = json.load(open(assortment_path, 'r', encoding='utf-8'))
        except Exception as ex:
            logging.error(f'Could not find files: {ex}')
            exit(1)

    def __get_data_from_stocks(self):
        for item in self.stocks['rows']:
            data = {}
            data['externalCode'] = item['externalCode']
            data['name'] = item['name']
            try:
                data['image'] = item['image']['miniature']['downloadHref']
            except:
                data['image'] = None
            data['category'] = self.__get_category_name(item)
            if data['category'] == 'Расходные материалы':
                if len(item['folder']['pathName'].split('/')) > 2:
                    data['subcategory'] = item['folder']['pathName'].split('/')[-1]
                else:
                    data['subcategory'] = item['folder']['name']
            self.all_data.append(data)

    def __get_data_from_assortment(self):
        for item in self.all_data:
            assortment_zapis = next((val for val in self.assortment if
                                     val['externalCode'] == item[
                                         'externalCode']), None)
            item['price_rozn'] = assortment_zapis['salePrices'][0]['value']/100
            item['price_5k'] = assortment_zapis['salePrices'][1]['value']/100
            item['price_15k'] = assortment_zapis['salePrices'][2]['value']/100
            item['price_100k'] = assortment_zapis['salePrices'][3]['value']/100

    @staticmethod
    def __get_category_name(item):
        try:
            folder_name = item['folder']['pathName']
            parts_of_folder = folder_name.split('/')
            if len(parts_of_folder) > 1:
                main_part = parts_of_folder[1]
                if main_part == 'Ресницы':
                    sub_part = parts_of_folder[-1]
                    return f'{main_part}_{sub_part}'
                else:
                    return main_part
            else:
                main_part = item['folder']['name']
                if main_part == 'Товары для МП':
                    return None

                return main_part
        except:
            logging.error(
                f'Problem with product folder name - {item["folder"]["pathName"]}'
            )
            return None

    def get_prepared_data(self):
        self.__get_data_from_stocks()
        self.__get_data_from_assortment()
        return self.all_data


