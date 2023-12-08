import pythoncom
from win32com.client import Dispatch, gencache, VARIANT
import datetime as dt
import configparser


class KompasAPI:
    def __init__(self):
        #  Подключим описание интерфейсов API7
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = self.api7.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.api7.IApplication.CLSID, pythoncom.IID_IDispatch))

        self.kompas_document = self.application.ActiveDocument

    def add_stamp_string(self):
        pass


if __name__ == "__main__":
    id_developer_surname = 120  # номер графы Фамилии разработчика
    id_inspector_surname = 0  # номер графы фамилии проверяющего
    id_technical_inspector_surname = 0  # номер графы фамилии Тех контроля
    id_standard_control_inspector_surname = 0  # номер графы фамилии нормоконтроля
    id_supervisor_surname = 0  # номер графы утверждающего
    id_company_name = 9  # номер графы фирмы
    id_date = 0

    developer_surname = 'Родченко'
    inspector_surname = 'Филатов'
    technical_inspector_surname = ''
    standard_control_inspector_surname = ''
    supervisor_surname = 'Шнякин'
    company_name = 'ООО "Горные технологии \nи инновации"'
    date = ''

    config = configparser.ConfigParser()
    config.read('config.ini')
    print(config['Surnames'])

    # Подключение к API компаса
    # Найти где-то переменную отвечающую за клетку
    # Заполнить все клетки (фамилии, компанию, дату)
    # Найти переменную отвечающую за шероховатость
    # Записать шероховатость

    kompas_api = KompasAPI()
    lay_out_sheets = kompas_api.kompas_document.LayoutSheets
    lay_out_sheet = lay_out_sheets.ItemByNumber(1)

    kompas_document_2d = kompas_api.api7.IKompasDocument2D(kompas_api.kompas_document)
    drawing_document = kompas_api.api7.IDrawingDocument(kompas_document_2d)
    spec_rough = drawing_document.SpecRough
    print(spec_rough.Text)
    spec_rough.Text = 'Ra 12,5'
    spec_rough.Update()

    stamp = lay_out_sheet.Stamp
    text = stamp.Text(id_company_name)

    text.Str = 'ООО "Горные технологии \nи инновации"'

    stamp.Update()

    date = dt.datetime.today()

    print(date)


