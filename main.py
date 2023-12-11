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
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.api7.IApplication.CLSID,
                                                                     pythoncom.IID_IDispatch))

        self.kompas_document = self.application.ActiveDocument

        if self.kompas_document.DocumentType == 1:
            self.lay_out_sheets = self.kompas_document.LayoutSheets
            self.lay_out_sheet = self.lay_out_sheets.ItemByNumber(1)
            self.stamp = self.lay_out_sheet.Stamp

            self.kompas_document_2d = self.api7.IKompasDocument2D(self.kompas_document)
            self.drawing_document = self.api7.IDrawingDocument(self.kompas_document_2d)
            self.spec_rough = self.drawing_document.SpecRough
            self.views_and_layers_manager = self.kompas_document_2d.ViewsAndLayersManager
            self.views = self.views_and_layers_manager.Views
            self.view = self.views.ActiveView
            self.association_view = self.api7.IAssociationView(self.view)
            self.property_mng = self.api7.IPropertyMng(self.application)
            self.property_keeper = self.api7.IPropertyKeeper(self.kompas_document_2d)
        else:
            self.application.MessageBoxEx("Данный макрос работает только с чертежом", "Документ не является чертежом", 0)

    def add_stamp_string(self, id, value):
        self.text = self.stamp.Text(id)
        self.text.Str = value

    def add_drawing_number(self):
        if self.stamp.Text(2).Str[-2:] != 'СБ':
            self.drawing_name = self.stamp.Text(2).Str
            self.val_str = f'<property id="marking" fromSource="true" direction="">' \
                           f'<property id="base" value="{self.drawing_name}" type="string" />' \
                           f'<property id="documentDelimiter" value=" " type="string" />' \
                           f'<property id="documentNumber" value="СБ" type="string" />'
            self.property = self.property_mng.GetProperty(self.kompas_document, 4.0)
            self.property_keeper.SetComplexPropertyValue(self.property, self.val_str)
            self.property.Update()
            self.stamp.Update()
            self.property = self.property_mng.GetProperty(self.kompas_document, 5.0)
            self.stamp.Text(1).Str = self.property_keeper.GetPropertyValue(self.property, "", True, True)[1]

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

    kompas_api = KompasAPI()

    if kompas_api.chech_doc_type():
        kompas_api.spec_rough_print(config['default_rough']['rough'])
    else:
        kompas_api.add_drawing_number()
    kompas_api.add_stamp_string(id_developer_surname, developer_surname)
    kompas_api.add_stamp_string(id_inspector_surname, inspector_surname)
    kompas_api.add_stamp_string(id_technical_inspector_surname, technical_inspector_surname)
    kompas_api.add_stamp_string(id_standard_control_inspector_surname, standard_control_inspector_surname)
    kompas_api.add_stamp_string(id_supervisor_surname, supervisor_surname)
    kompas_api.add_stamp_string(id_company_name, company_name)
    kompas_api.add_stamp_string(id_date, date)

    kompas_api.stamp.Update()
