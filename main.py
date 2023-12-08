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

        if self.kompas_document.DocumentType == 1:
            self.lay_out_sheets = self.kompas_document.LayoutSheets
            self.lay_out_sheet = self.lay_out_sheets.ItemByNumber(1)
            self.stamp = self.lay_out_sheet.Stamp

            self.kompas_document_2d = self.api7.IKompasDocument2D(self.kompas_document)
            self.drawing_document = self.api7.IDrawingDocument(self.kompas_document_2d)
            self.spec_rough = self.drawing_document.SpecRough
        else:
            self.application.MessageBoxEx("Данный макрос работает только с чертежом",
                                                "Документ не является чертежом", 0)

    def add_stamp_string(self, id, value):

        self.text = self.stamp.Text(id)
        self.text.Str = value

    def spec_rough_print(self, value):
        self.spec_rough.Text = value
        self.spec_rough.Update()


if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read('config.ini', encoding="utf-8")

    id_developer_surname = int(config['ID']['id_developer_surname'])  # номер графы Фамилии разработчика
    id_inspector_surname = int(config['ID']['id_inspector_surname'])  # номер графы фамилии проверяющего
    id_technical_inspector_surname = int(config['ID']['id_technical_inspector_surname'])
                                                                                    # номер графы фамилии Тех контроля
    id_standard_control_inspector_surname = int(config['ID']['id_standard_control_inspector_surname'])
                                                                                    # номер графы фамилии нормоконтроля
    id_supervisor_surname = int(config['ID']['id_supervisor_surname'])  # номер графы утверждающего
    id_company_name = int(config['ID']['id_company_name'])  # номер графы фирмы
    id_date = int(config['ID']['id_date'])

    developer_surname = config['Surnames']['developer_surname']
    inspector_surname = config['Surnames']['inspector_surname']
    technical_inspector_surname = config['Surnames']['technical_inspector_surname']
    standard_control_inspector_surname = config['Surnames']['standard_control_inspector_surname']
    supervisor_surname = config['Surnames']['supervisor_surname']

    company_name = config['Surnames']['company_name']
    company_name = company_name.replace(r'\n', '\n')

    now_day = dt.datetime.today()
    date = str(now_day.day) + '.' + str(now_day.month) + '.' + str(now_day.year)

    # Получение данных для заполнения
    # Подключение к API компаса
    # Найти где-то переменную отвечающую за клетку
    # Заполнить все клетки (фамилии, компанию, дату)
    # Найти переменную отвечающую за шероховатость
    # Записать шероховатость

    kompas_api = KompasAPI()

    kompas_api.add_stamp_string(id_developer_surname, developer_surname)
    kompas_api.add_stamp_string(id_inspector_surname, inspector_surname)
    kompas_api.add_stamp_string(id_technical_inspector_surname, technical_inspector_surname)
    kompas_api.add_stamp_string(id_standard_control_inspector_surname, standard_control_inspector_surname)
    kompas_api.add_stamp_string(id_supervisor_surname, supervisor_surname)
    kompas_api.add_stamp_string(id_company_name, company_name)
    kompas_api.add_stamp_string(id_date, date)

    kompas_api.spec_rough_print(config['default_rough']['rough'])

    kompas_api.stamp.Update()


