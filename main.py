import pythoncom
from win32com.client import Dispatch, gencache, VARIANT
import datetime as dt
import configparser
import os.path


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
        self.kompas_document_2D = self.api7.IKompasDocument2D(self.kompas_document)
        self.views_and_layers_manager = self.kompas_document_2D.ViewsAndLayersManager
        self.views = self.views_and_layers_manager.Views
        self.view = self.views.ActiveView
        self.association_view = self.api7.IAssociationView(self.view)

    def add_stamp_string(self, id, value):

        self.text = self.stamp.Text(id)
        self.text.Str = value

    def add_drawing_number(self):
        pass

    def spec_rough_print(self, value):
        self.spec_rough.Text = value
        self.spec_rough.Update()

    def chech_doc_type(self):  # Проверка на сборку/деталь. Если сборка False, если деталь True
        if self.association_view.SourceFileName[-3:] == 'm3d':
            return True
        return False


def config_create(path_name):
    config = configparser.ConfigParser()
    if os.path.exists(os.path.join(path_name, 'config.ini')):
        config.read(os.path.join(path_name, 'config.ini'), encoding="utf-8")
        print('config открыт по пути ..', os.path.join(path_name, 'config.ini'))
    else:
        config.add_section('ID')
        config.add_section('Surnames')
        config.add_section('default_rough')
        config.set('ID', 'id_developer_surname', '110')
        config.set('ID', 'id_inspector_surname', '111')
        config.set('ID', 'id_technical_inspector_surname', '112')
        config.set('ID', 'id_standard_control_inspector_surname', '114')
        config.set('ID', 'id_supervisor_surname', '115')
        config.set('ID', 'id_company_name', '9')
        config.set('ID', 'id_date', '130')
        config.set('Surnames', 'developer_surname', 'Иванов')
        config.set('Surnames', 'inspector_surname', '')
        config.set('Surnames', 'technical_inspector_surname', '')
        config.set('Surnames', 'standard_control_inspector_surname', '')
        config.set('Surnames', 'supervisor_surname', '')
        config.set('Surnames', 'company_name', r'ООО "Рога \nи копыта"')
        config.set('default_rough', 'rough', 'Ra 12,5')
        with open(os.path.join(path_name, 'config.ini'), 'w') as config_file:
            config.write(config_file)
    return config


if __name__ == "__main__":
    dir_name = os.path.dirname(os.path.abspath(__file__))
    config = config_create(dir_name)

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

    if kompas_api.chech_doc_type():
        kompas_api.spec_rough_print(config['default_rough']['rough'])

    print(kompas_api.stamp.Text(2).Str)
    #kompas_api.stamp.Text(2).Str = kompas_api.stamp.Text(2).Str + ' СБ'
    property_mng = kompas_api.api7.IPropertyMng(kompas_api.application)
    print(property_mng.PropertyCount(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw'))
    print(property_mng.GetProperties(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw'))
    for ui in property_mng.GetProperties(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw'):
        print(ui.Id, ' ', ui.Name, ' ', ui.ListVal)
    property_keeper = kompas_api.api7.IPropertyKeeper(kompas_api.kompas_document_2D)
    print(property_keeper.GetPropertyValue(property_mng.GetProperties(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw')[0], '', True, True))
    print(property_keeper.Properties)
    print(property_keeper.GetComplexPropertyValue(property_mng.GetProperties(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw')[0]))

    drawing_name = '02-05-04-02-ТХ8.Н9.03'
    val_str =   '<property id="marking" fromSource="true" direction="">' \
                f'<property id="base" value="{drawing_name}" type="string" />' \
                '<property id="documentNumber" value="СБ" type="string" />'
    print(val_str)

    print('Устанавливаю комплексное свойство ... ... ...', property_keeper.SetComplexPropertyValue(property_mng.GetProperties(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw')[0], val_str))
    property_mng.GetProperties(r'C:\Users\Leafan\Desktop\02-05-04-02-ТХ8.Н9.03 СБ Лист Сборочный чертеж.cdw')[0].Update()
    print(property_keeper.Properties)

    kompas_api.stamp.Update()

