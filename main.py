import pythoncom
from win32com.client import Dispatch, gencache, VARIANT
import datetime as dt


class KompasAPI:
    def __init__(self):
        #  Подключим описание интерфейсов API7
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = self.api7.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.api7.IApplication.CLSID, pythoncom.IID_IDispatch))

        self.kompas_document = self.application.ActiveDocument
        self.kompas_document_3d = self.api7.IKompasDocument3D(self.kompas_document)
        self.part_7 = self.kompas_document_3d.TopPart
        self.property_keeper = self.api7.IPropertyKeeper(self.part_7)
        self.property_mng = self.api7.IPropertyMng(self.application)


if __name__ == "__main__":
    id_invent_surname = 120  # номер графы Фамилии разработчика
    id_company_name = 9 # номер графы фирмы

    company_name = 'ООО "Горные технологии \nи инновации"'


    # Подключение к API компаса
    kompas_api = KompasAPI()
    lay_out_sheets = kompas_api.kompas_document.LayoutSheets
    lay_out_sheet = lay_out_sheets.ItemByNumber(1)

    stamp = lay_out_sheet.Stamp
    text = stamp.Text(id_company_name)

    text.Str = 'ООО "Горные технологии \nи инновации"'

    stamp.Update()

    date = dt.datetime.today()

    print(date)
    # чет новое для github


