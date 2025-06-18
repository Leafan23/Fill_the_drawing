import pythoncom
from win32com.client import Dispatch, gencache, VARIANT
import datetime as dt
import configparser
import os.path
import sys


class KompasAPI:
    def __init__(self):
        self.ShowOnSheet = 0
        #  Подключим описание интерфейсов API7
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = self.api7.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.api7.IApplication.CLSID,
                                                                     pythoncom.IID_IDispatch))

        self.kompas_document = self.application.ActiveDocument
        # Document Type 1 - Чертеж; Document Type 3 - Спецификация.
        if self.kompas_document.DocumentType == 1 or self.kompas_document.DocumentType == 3:
            self.lay_out_sheets = self.kompas_document.LayoutSheets
            self.lay_out_sheet = self.lay_out_sheets.ItemByNumber(1)
            self.stamp = self.lay_out_sheet.Stamp

            self.kompas_document_2d = self.api7.IKompasDocument2D(self.kompas_document)
            self.drawing_document = self.api7.IDrawingDocument(self.kompas_document_2d)
            if self.kompas_document.DocumentType == 1:
                self.specification_descriptions = self.kompas_document.SpecificationDescriptions
                self.specification_description = self.specification_descriptions.Active
                self.spec_rough = self.drawing_document.SpecRough
                self.views_and_layers_manager = self.kompas_document_2d.ViewsAndLayersManager
                self.views = self.views_and_layers_manager.Views
                self.view = self.views.ActiveView
                self.association_view = self.api7.IAssociationView(self.view)
            self.property_mng = self.api7.IPropertyMng(self.application)
            self.property_keeper = self.api7.IPropertyKeeper(self.kompas_document_2d)
        else:
            self.application.MessageBoxEx("Данный макрос работает только с чертежом или спецификацией", "Документ не является чертежом/спецификацией", 0)
            sys.exit(1)

    def add_stamp_string(self, id, value, recopy):
        self.text = self.stamp.Text(id)
        if recopy == 1 or self.text.Str == '':
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

    def delete_drawing_number(self):
        if self.stamp.Text(2).Str[-2:] == 'СБ':
            self.drawing_name = self.stamp.Text(2).Str[:-3]
            self.val_str = f'<property id="marking" fromSource="true" direction="">' \
                           f'<property id="base" value="{self.drawing_name}" type="string" />' \
                           f'<property id="documentDelimiter" value=" " type="string" />' \
                           f'<property id="documentNumber" value="" type="string" />'
            self.property = self.property_mng.GetProperty(self.kompas_document, 4.0)
            self.property_keeper.SetComplexPropertyValue(self.property, self.val_str)
            self.property.Update()
            self.stamp.Update()
            self.property = self.property_mng.GetProperty(self.kompas_document, 5.0)
            self.stamp.Text(1).Str = self.property_keeper.GetPropertyValue(self.property, "", True, True)[1]

    def spec_rough_print(self, value, sign):
        self.spec_rough.Text = value
        self.spec_rough.SignType = sign  # 0 - Без указания типа отбработки; 1 - С удалением слоя материала; 2 - Без удаления слоя материала.
        self.spec_rough.AddSign = True
        self.spec_rough.Update()

    def check_doc_type(self):  # Проверка на сборку/деталь. # Сборка - 0; Деталь - 1; Спецификация - 2
        if self.kompas_document.DocumentType == 1 and self.association_view.SourceFileName[-3:] == 'm3d':  # Деталь - 1
            return 1
        if self.kompas_document.DocumentType == 3:  # Спецификация - 2
            return 2
        return 0  # Сборка - 0

    def first_used(self, value, flag):  # Обработка значения первичного применения. В Сб пишется тот же номер. В деталь убираются 000, в спецификации берется следующий без нулей
        if flag == 0:
            return ''
        doc_type = self.check_doc_type()
        if doc_type == 1:
            return convert(value)
        elif self.ShowOnSheet == 1 and doc_type == 0:
            return convert(value)
        elif self.ShowOnSheet == 0 and doc_type == 0:
            i = value.rfind(" СБ")
            if value.rfind(" СБ") == -1:
                return value
            s = value[:i]
            return s
        elif doc_type == 2:
            return convert(value)


def config_create(path_name):  # Создание конфиг файла
    config = configparser.ConfigParser()
    if os.path.exists(os.path.join(path_name, 'config.ini')):
        config.read(os.path.join(path_name, 'config.ini'), encoding='utf-8')
    else:
        config.add_section('Surnames')
        config.add_section('ID')
        config.add_section('default_rough')
        config.add_section('Settings')
        config.set('ID', 'id_developer_surname', '110')
        config.set('ID', 'id_inspector_surname', '111')
        config.set('ID', 'id_technical_inspector_surname', '112')
        config.set('ID', 'id_standard_control_inspector_surname', '114')
        config.set('ID', 'id_supervisor_surname', '115')
        config.set('ID', 'id_company_name', '9')
        config.set('ID', 'id_date', '130')
        config.set('ID', 'id_first_used', '25')
        config.set('Surnames', 'developer_surname', 'Иванов')
        config.set('Surnames', 'inspector_surname', '')
        config.set('Surnames', 'technical_inspector_surname', '')
        config.set('Surnames', 'standard_control_inspector_surname', '')
        config.set('Surnames', 'supervisor_surname', '')
        config.set('Surnames', 'company_name', r'ООО "Рога \nи копыта"')
        config.set('default_rough', 'rough', 'Ra 12,5')
        config.set('default_rough', 'rough_sign', '0')
        config.set('Settings', 'recopy', '1')
        config.set('Settings', 'first_used', '1')
        config.set('Settings', 'date_format', '%%d.%%m.%%y')
        with open(os.path.join(path_name, 'config.ini'), 'w', encoding='utf-8') as config_file:
            config.write(config_file)
    return config


def convert(code):
    separator = '.'
    res = code.split('.')
    cell_for_replace = find_cell(res)

    str = res[-1 - cell_for_replace]  # строка, которая будет изменяться

    if len(str) == 3 and cell_for_replace == 0:
        if str[:1] != '0' and str[-1:] != '0':
            str = str[:1] + '00'
        else:
            str = '000'
    elif len(str) < 3:
        str = '00'
    res[-1 - cell_for_replace] = str  # замена ячейки на нули
    res = separator.join(res)  # соединение ячеек
    return res


def find_cell(res):
    count = 0
    for i in reversed(res):
        for char in i:
            if char != '0':
                return count
        count += 1
    return count


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

    dateFormat = config['Settings']['date_format']

    date = dt.datetime.today().strftime(dateFormat)  # чтение текущей даты и запись в необходимом формате

    kompas_api = KompasAPI()

    if kompas_api.check_doc_type() == 1:
        if not kompas_api.spec_rough.IsCreated:
            kompas_api.spec_rough_print(config['default_rough']['rough'], int(config['default_rough']['rough_sign']))  # если деталь, то напечатать неуказ.
        # шереховатеость
        kompas_api.add_stamp_string(id_technical_inspector_surname, technical_inspector_surname, int(config['Settings']['recopy']))
    elif kompas_api.check_doc_type() == 0:
        if kompas_api.specification_description is None:
            kompas_api.add_drawing_number()  # если это сборка, то добавить СБ в номер
        elif kompas_api.specification_description.ShowOnSheet:
            kompas_api.delete_drawing_number()  # если есть спецификация на листе, то убрать СБ
            kompas_api.ShowOnSheet = 1
        else:
            kompas_api.add_drawing_number()
        kompas_api.add_stamp_string(id_technical_inspector_surname, technical_inspector_surname, int(config['Settings']['recopy']))  # если это
        # спецификация, то не печатается Т.Контр
    kompas_api.add_stamp_string(id_developer_surname, developer_surname, int(config['Settings']['recopy']))
    kompas_api.add_stamp_string(id_inspector_surname, inspector_surname, int(config['Settings']['recopy']))
    kompas_api.add_stamp_string(id_standard_control_inspector_surname, standard_control_inspector_surname, int(config['Settings']['recopy']))
    kompas_api.add_stamp_string(id_supervisor_surname, supervisor_surname, int(config['Settings']['recopy']))
    kompas_api.add_stamp_string(id_company_name, company_name, int(config['Settings']['recopy']))
    kompas_api.add_stamp_string(id_date, date, int(config['Settings']['recopy']))
    kompas_api.add_stamp_string(int(config['ID']['id_first_used']), kompas_api.first_used(kompas_api.stamp.Text(2).Str, int(config['Settings']['first_used'])), int(config['Settings']['recopy']))

    kompas_api.stamp.Update()
