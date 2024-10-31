#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

#endregion

namespace GOSTSpec
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {

        List<int> builtInCategoryIdValues = new List<int>(){
                -2008055, // Арматура трубопроводов
				-2008050, // Гибкие трубы
				-2008122, // Материалы изоляции труб
				//-2000151, // Обобщенные модели
				-2001140, // Оборудование
				-2001160, // Сантехнические приборы
				-2008049, // Соединительные детали трубопроводов
				-2008099, // Спринклеры
				-2008161, // Трубопровод по осевой
				-2008043, // Трубопроводные системы
				-2008044, // Трубы
				-2008208, // Трубы из базы данных производителя MEP
				-2002000, // Элементы узлов

                -2008016, // Арматура воздуховодов
                -2008000, // Воздуховоды
                -2008013, // Воздухораспределители
                -2008010, // Соединительные детали воздуховодов
                -2008020, // Гибкие воздуховоды
                -2008123, // Материалы изоляции воздуховодов
                -2008160, // Воздуховоды по осевой
                -2008015, // Системы воздуховодов
                -2008193, // Элементы воздуховодов из базы данных производителя MEP
			};

        const string S_System = "С_Система";

        const string SRC_Naimen = "KAZGOR_Наименование";
        const string SRC_Marka = "KAZGOR_Марка";
        const string SRC_Mass = "KAZGOR_Масса";
        const string SRC_KodIzd = "KAZGOR_Код изделия";
        const string SRC_TypIzol = "Тип изоляции";
        const string SRC_TypTruby = "Тип трубопровода";
        const string SRC_Zavod = "KAZGOR_Завод-изготовитель";
        const string SRC_EdIzm = "KAZGOR_Единица измерения";
        const string SRC_Primech = "KAZGOR_Примечание";
        const string SRC_CircSize = "KAZGOR_Размер_Диаметр";
        const string SRC_Height = "KAZGOR_Размер_Высота";
        const string SRC_Width = "KAZGOR_Размер_Ширина";
        const string SRC_IzolSize = "KAZGOR_Диаметр изоляции";

        const string TRG_Naimen = "С_Наименование";
        const string TRG_Marka = "С_Марка";
        const string TRG_Mass = "С_Масса";
        const string TRG_KodIzd = "С_Код изделия";
        const string TRG_Zavod = "С_Завод-изготовитель";
        const string TRG_EdIzm = "С_Единица измерения";
        const string TRG_Primech = "С_Примечание";
        const string TRG_Count = "С_Количество";

        Document doc = null;

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            doc = uidoc.Document;

            var zapas = GetPercentGlobal();

            using (var trans = new Transaction(doc, "Копирование параметров"))
            {
                trans.Start();

                foreach (int intValue in builtInCategoryIdValues)
                {
                    var elements = GetElementsByCatIdValue(intValue);

                    // Копирование имя системы
                    foreach (var element in elements)
                    {
                        if (element is FamilyInstance)
                        {
                            var fi = element as FamilyInstance;
                            var superComponent = GetRootFamily(fi);

                            if (fi.Id != superComponent.Id)
                            {
                                try
                                {
                                    var rootSystemName = (superComponent.LookupParameter(S_System) as Parameter).AsString();
                                    if (!String.IsNullOrEmpty(rootSystemName))
                                        fi.LookupParameter(S_System).Set(rootSystemName);
                                }
                                catch { }
                            }
                        }
                    }

                    switch (intValue)
                    {
                        case -2008122: // Изоляция							
                            foreach (Element izolyacya in elements)
                            {
                                var izolyacyaType = doc.GetElement(izolyacya.GetTypeId());

                                // Определение типа изоляции (Длина, Площадь, Объем)
                                var typIzolyacy = (izolyacyaType.LookupParameter(SRC_TypIzol) as Parameter).AsString();

                                var tolchyna = GetParameter(izolyacya, "Толщина изоляции").AsDouble();
                                tolchyna = UnitUtils.ConvertFromInternalUnits(tolchyna, UnitTypeId.Millimeters);

                                if (!(doc.GetElement((izolyacya as PipeInsulation).HostElementId) is Pipe))
                                {
                                    GetParameter(izolyacya, TRG_Naimen).Set("Исключить");
                                    continue;
                                }


                                try
                                {
                                    var srcName = GetParameter(izolyacya, SRC_Naimen).AsString();
                                    var pipeSize = GetParameter(izolyacya, "Размер трубы").AsString();
                                    
                                                                        

                                    var trgName = srcName;

                                    if (typIzolyacy.Contains("1"))
                                    {
                                        trgName = trgName + " " + tolchyna + "мм, для труб " + pipeSize;
                                    }
                                    else if (typIzolyacy.Contains("2"))
                                    {
                                        var izolSize = GetParameter(izolyacya, SRC_IzolSize).HasValue ? GetParameter(izolyacya, SRC_IzolSize).AsDouble() : 0;
                                        trgName = trgName + " " + tolchyna + "мм, диаметром " + string.Format("{0:0.#}", izolSize) + "мм";
                                    }

                                    GetParameter(izolyacya, TRG_Naimen).Set(trgName);
                                }
                                catch
                                {

                                }

                                // KAZGOR_Марка								
                                try
                                {
                                    var srcMarka = GetParameter(izolyacya, SRC_Marka).AsString();
                                    GetParameter(izolyacya, TRG_Marka).Set(srcMarka);
                                }
                                catch { }

                                // KAZGOR_Код Изделия							
                                try
                                {
                                    var srcKod = GetParameter(izolyacya, SRC_KodIzd).AsString();
                                    GetParameter(izolyacya, TRG_KodIzd).Set(srcKod);
                                }
                                catch { }

                                // KAZGOR_Завод изготовитель								
                                try
                                {
                                    var srcZavod = GetParameter(izolyacya, SRC_Zavod).AsString();
                                    GetParameter(izolyacya, TRG_Zavod).Set(srcZavod);
                                }
                                catch { }

                                // KAZGOR_Единица измерения								
                                try
                                {
                                    var srcEdIzm = GetParameter(izolyacya, SRC_EdIzm).AsString();
                                    GetParameter(izolyacya, TRG_EdIzm).Set(srcEdIzm);
                                }
                                catch { }

                                // KAZGOR_Количество
                                try
                                {
                                    var srcLength = GetParameter(izolyacya, "Длина").AsDouble();
                                    srcLength = UnitUtils.ConvertFromInternalUnits(srcLength, UnitTypeId.Millimeters);

                                    // Длина
                                    if (typIzolyacy.Contains("Д"))
                                    {
                                        GetParameter(izolyacya, TRG_Count).Set(srcLength * zapas / 1000);
                                    }
                                    // Площадь
                                    else if (typIzolyacy.Contains("П"))
                                    {
                                        double area = GetParameter(izolyacya, "Площадь").AsDouble();
                                        area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters);
                                        GetParameter(izolyacya, TRG_Count).Set(area * zapas);
                                    }
                                    // Объем
                                    else if (typIzolyacy.Contains("О"))
                                    {
                                        double area = GetParameter(izolyacya, "Площадь").AsDouble();
                                        area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters);
                                        GetParameter(izolyacya, TRG_Count).Set(zapas * area * tolchyna / 1000);
                                    }

                                    // Параметр не заполнен
                                    else if (String.IsNullOrEmpty(typIzolyacy))
                                    {
                                        GetParameter(izolyacya, TRG_Count).Set(-999999);
                                    }
                                }
                                catch { }

                                // KAZGOR_Примечание								
                                try
                                {
                                    var srcPrim = GetParameter(izolyacya, SRC_Primech).AsString();
                                    GetParameter(izolyacya, TRG_Primech).Set(srcPrim);
                                }
                                catch { }
                            }
                            break;
                        case -2001140: // Оборудование
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2001160: // Сантехнические приборы
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2008099: // Спринклеры
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2008055: // Арматура трубопроводов
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2008049: // Соединительные детали трубопроводов
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2008044: // Трубы
                            foreach (Element pipe in elements)
                            {
                                var pipeType = doc.GetElement(pipe.GetTypeId());

                                try
                                {
                                    var srcName = GetParameter(pipeType, SRC_Naimen).AsString();
                                    var pipeSize = GetParameter(pipe, "Размер").AsString();

                                    var typTruby = GetParameter(pipe, SRC_TypTruby).AsDouble();

                                    var pipeOutSize = GetParameter(pipe, "Внешний диаметр").AsDouble();
                                    var pipeInSize = GetParameter(pipe, "Внутренний диаметр").AsDouble();

                                    var thickness = UnitUtils.ConvertFromInternalUnits((pipeOutSize - pipeInSize) / 2, UnitTypeId.Millimeters);

                                    pipeOutSize = UnitUtils.ConvertFromInternalUnits(pipeOutSize, UnitTypeId.Millimeters);
                                    pipeInSize = UnitUtils.ConvertFromInternalUnits(pipeInSize, UnitTypeId.Millimeters);

                                    var trgName = "ТАКОЙ ТИП ТРУБЫ НЕ СУЩЕСТВУЕТ. ДОСТУПНЫЕ ТИПЫ 1-3";
                                    switch (typTruby)
                                    {
                                        case 1:
                                            trgName = srcName + ", " + pipeSize + "х" + string.Format("{0:0.#}", thickness);
                                            break;
                                        case 2:
                                            trgName = srcName + ", Ø" + pipeOutSize + "х" + string.Format("{0:0.#}", thickness);
                                            break;
                                        case 3:
                                            trgName = srcName + ", " + pipeSize;
                                            break;
                                        default:
                                            break;
                                    }

                                    GetParameter(pipe, TRG_Naimen).Set(trgName);
                                }
                                catch (Exception e) {
                                
                                
                                }

                                // KAZGOR_Марка								
                                try
                                {
                                    var srcMarka = GetParameter(pipeType, SRC_Marka).AsString();
                                    GetParameter(pipe, TRG_Marka).Set(srcMarka);
                                }
                                catch { }

                                // KAZGOR_Код Изделия							
                                try
                                {
                                    var srcKod = GetParameter(pipeType, SRC_KodIzd).AsString();
                                    GetParameter(pipe, TRG_KodIzd).Set(srcKod);
                                }
                                catch { }

                                // KAZGOR_Завод изготовитель								
                                try
                                {
                                    var srcZavod = GetParameter(pipeType, SRC_Zavod).AsString();
                                    GetParameter(pipe, TRG_Zavod).Set(srcZavod);
                                }
                                catch { }

                                // KAZGOR_Единица измерения								
                                try
                                {
                                    var srcEdIzm = GetParameter(pipeType, SRC_EdIzm).AsString();
                                    GetParameter(pipe, TRG_EdIzm).Set(srcEdIzm);
                                }
                                catch { }

                                // KAZGOR_Количество
                                try
                                {
                                    var srcLength = GetParameter(pipe, "Длина").AsDouble();
                                    srcLength = UnitUtils.ConvertFromInternalUnits(srcLength, UnitTypeId.Millimeters);
                                    GetParameter(pipe, TRG_Count).Set(srcLength * zapas / 1000);
                                }
                                catch { }

                                // KAZGOR_Примечание								
                                try
                                {
                                    var srcPrim = GetParameter(pipe, SRC_Primech).AsString();
                                    GetParameter(pipe, TRG_Primech).Set(srcPrim);
                                }
                                catch { }
                            }
                            break;

                        case -2008050: // Гибкие трубы
                            foreach (Element pipe in elements)
                            {
                                var pipeType = doc.GetElement(pipe.GetTypeId());

                                try
                                {
                                    var srcName = GetParameter(pipeType, SRC_Naimen).AsString();
                                    GetParameter(pipe, TRG_Naimen).Set(srcName);
                                }
                                catch (Exception e)
                                {


                                }

                                // KAZGOR_Марка								
                                try
                                {
                                    var srcMarka = GetParameter(pipeType, SRC_Marka).AsString();
                                    GetParameter(pipe, TRG_Marka).Set(srcMarka);
                                }
                                catch { }

                                // KAZGOR_Код Изделия							
                                try
                                {
                                    var srcKod = GetParameter(pipeType, SRC_KodIzd).AsString();
                                    GetParameter(pipe, TRG_KodIzd).Set(srcKod);
                                }
                                catch { }

                                // KAZGOR_Завод изготовитель								
                                try
                                {
                                    var srcZavod = GetParameter(pipeType, SRC_Zavod).AsString();
                                    GetParameter(pipe, TRG_Zavod).Set(srcZavod);
                                }
                                catch { }

                                // KAZGOR_Единица измерения								
                                try
                                {
                                    var srcEdIzm = GetParameter(pipeType, SRC_EdIzm).AsString();
                                    GetParameter(pipe, TRG_EdIzm).Set(srcEdIzm);
                                }
                                catch { }

                                // KAZGOR_Количество
                                try
                                {   
                                    GetParameter(pipe, TRG_Count).Set(1);
                                }
                                catch { }

                                // KAZGOR_Примечание								
                                try
                                {
                                    var srcPrim = GetParameter(pipe, SRC_Primech).AsString();
                                    GetParameter(pipe, TRG_Primech).Set(srcPrim);
                                }
                                catch { }
                            }
                            break;

                        // Система вентиляции
                        case -2008000: // Воздуховоды
                            foreach (Element duct in elements)
                            {
                                try
                                {
                                    var srcName = GetParameter(duct, SRC_Naimen).AsString();
                                    var ductSize = GetParameter(duct, "Размер").AsString();
                                    var typeName = GetParameter(duct, "Имя типа", true).AsString();
                                    var thickness = 0.5;

                                    var isMinThickness08 = typeName.ToLower().Contains("транзит") || typeName.ToLower().Contains("огнезащитой");

                                    var ductDiameterParameter = duct.LookupParameter("Диаметр");
                                    // Duct is Circle
                                    if (ductDiameterParameter != null)
                                    {
                                        var diameter = UnitUtils.ConvertFromInternalUnits(ductDiameterParameter.AsDouble(), UnitTypeId.Millimeters);
                                        if (diameter <= 200)
                                        {
                                            thickness = 0.5;
                                        }
                                        else if (diameter > 200 && diameter <= 450)
                                        {
                                            thickness = 0.6;
                                        }
                                        else if (diameter > 450 && diameter <= 800)
                                        {
                                            thickness = 0.7;
                                        }
                                        else if (diameter > 800 && diameter <= 1250)
                                        {
                                            thickness = 1.0;
                                        }
                                        else if (diameter > 1250 && diameter <= 1600)
                                        {
                                            thickness = 1.2;
                                        }
                                        else if (diameter > 1600 && diameter <= 2000)
                                        {
                                            thickness = 1.4;
                                        }
                                    }
                                    // Duct is rect
                                    else
                                    {
                                        var height = UnitUtils.ConvertFromInternalUnits(GetParameter(duct, "Высота").AsDouble(), UnitTypeId.Millimeters);
                                        var width = UnitUtils.ConvertFromInternalUnits(GetParameter(duct, "Ширина").AsDouble(), UnitTypeId.Millimeters);
                                        var bigSize = height > width ? height : width;
                                        if (bigSize <= 250)
                                        {
                                            thickness = 0.5;
                                        }
                                        else if (bigSize > 250 && bigSize <= 1000)
                                        {
                                            thickness = 0.7;
                                        }
                                        else if (bigSize > 1000 && bigSize <= 2000)
                                        {
                                            thickness = 0.9;
                                        }
                                    }

                                    thickness = isMinThickness08 ? (thickness < 0.8 ? 0.8 : thickness) : thickness;

                                    var trgName = srcName + ", " + ductSize + " δ=" + thickness + " мм";
                                    GetParameter(duct, TRG_Naimen).Set(trgName);
                                }
                                catch { }

                                // KAZGOR_Марка								
                                try
                                {
                                    var srcMarka = GetParameter(duct, SRC_Marka).AsString();
                                    GetParameter(duct, TRG_Marka).Set(srcMarka);
                                }
                                catch { }

                                // KAZGOR_Код Изделия							
                                try
                                {
                                    var srcKod = GetParameter(duct, SRC_KodIzd).AsString();
                                    GetParameter(duct, TRG_KodIzd).Set(srcKod);
                                }
                                catch { }

                                // KAZGOR_Завод изготовитель								
                                try
                                {
                                    var srcZavod = GetParameter(duct, SRC_Zavod).AsString();
                                    GetParameter(duct, TRG_Zavod).Set(srcZavod);
                                }
                                catch { }

                                // KAZGOR_Единица измерения								
                                try
                                {
                                    var srcEdIzm = GetParameter(duct, SRC_EdIzm).AsString();
                                    GetParameter(duct, TRG_EdIzm).Set(srcEdIzm);
                                }
                                catch { }

                                // KAZGOR_Количество
                                try
                                {
                                    var srcLength = GetParameter(duct, "Длина").AsDouble();
                                    //var srcArea = GetParameter(duct, "Площадь").AsDouble();
                                    srcLength = UnitUtils.ConvertFromInternalUnits(srcLength, UnitTypeId.Meters);
                                    //srcArea = UnitUtils.ConvertFromInternalUnits(srcArea, UnitTypeId.DUT_SQUARE_METERS);

                                    GetParameter(duct, TRG_Count).Set(srcLength * zapas);
                                }
                                catch { }

                                // KAZGOR_Примечание								
                                try
                                {
                                    var srcPrim = GetParameter(duct, SRC_Primech).AsString();
                                    GetParameter(duct, TRG_Primech).Set(srcPrim);
                                }
                                catch { }
                            }
                            break;
                        case -2008123: // Материалы изоляции воздуховодов
                            foreach (Element izolyacya in elements)
                            {
                                var izolyacyaType = doc.GetElement(izolyacya.GetTypeId());

                                if (!(doc.GetElement((izolyacya as DuctInsulation).HostElementId) is Duct))
                                {
                                    GetParameter(izolyacya, TRG_Naimen).Set("Исключить");
                                    continue;
                                }


                                try
                                {
                                    var srcName = GetParameter(izolyacya, SRC_Naimen).AsString();
                                    var ductSize = GetParameter(izolyacya, "Размер воздуховода").AsString();

                                    var trgName = srcName;

                                    var tolchyna = GetParameter(izolyacya, "Толщина изоляции").AsDouble();
                                    tolchyna = UnitUtils.ConvertFromInternalUnits(tolchyna, UnitTypeId.Millimeters);

                                    trgName = trgName + " δ=" + tolchyna + "мм";

                                    GetParameter(izolyacya, TRG_Naimen).Set(trgName);
                                }
                                catch { }

                                // KAZGOR_Марка								
                                try
                                {
                                    var srcMarka = GetParameter(izolyacya, SRC_Marka).AsString();
                                    GetParameter(izolyacya, TRG_Marka).Set(srcMarka);
                                }
                                catch { }

                                // KAZGOR_Код Изделия							
                                try
                                {
                                    var srcKod = GetParameter(izolyacya, SRC_KodIzd).AsString();
                                    GetParameter(izolyacya, TRG_KodIzd).Set(srcKod);
                                }
                                catch { }

                                // KAZGOR_Завод изготовитель								
                                try
                                {
                                    var srcZavod = GetParameter(izolyacya, SRC_Zavod).AsString();
                                    GetParameter(izolyacya, TRG_Zavod).Set(srcZavod);
                                }
                                catch { }

                                // KAZGOR_Единица измерения								
                                try
                                {
                                    var srcEdIzm = GetParameter(izolyacya, SRC_EdIzm).AsString();
                                    GetParameter(izolyacya, TRG_EdIzm).Set(srcEdIzm);
                                }
                                catch { }

                                // KAZGOR_Количество
                                try
                                {

                                    double area = GetParameter(izolyacya, "Площадь").AsDouble();
                                    area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters);
                                    GetParameter(izolyacya, TRG_Count).Set(area * zapas);

                                }
                                catch { }

                                // KAZGOR_Примечание								
                                try
                                {
                                    var srcPrim = GetParameter(izolyacya, SRC_Primech).AsString();
                                    GetParameter(izolyacya, TRG_Primech).Set(srcPrim);
                                }
                                catch { }
                            }
                            break;
                        case -2008016: // Арматура воздуховодов
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2008013: // Воздухораспределители
                            CopyGeneralElementsParameters(elements);
                            break;
                        case -2008010: // Соединительные детали воздуховодов                
                            foreach (Element fitting in elements)
                            {
                                try
                                {
                                    var srcName = GetParameter(fitting, SRC_Naimen).AsString();
                                    var size = GetParameter(fitting, "Размер").AsString();

                                    srcName = srcName + ", " + size;
                                    GetParameter(fitting, TRG_Naimen).Set(srcName);
                                }
                                catch { }

                                // KAZGOR_Марка
                                try
                                {
                                    var srcMarka = GetParameter(fitting, SRC_Marka).AsString();
                                    GetParameter(fitting, TRG_Marka).Set(srcMarka);
                                }
                                catch { }

                                // KAZGOR_Код Изделия							
                                try
                                {
                                    var srcKod = GetParameter(fitting, SRC_KodIzd).AsString();
                                    GetParameter(fitting, TRG_KodIzd).Set(srcKod);
                                }
                                catch { }

                                // KAZGOR_Завод изготовитель								
                                try
                                {
                                    var srcZavod = GetParameter(fitting, SRC_Zavod).AsString();
                                    GetParameter(fitting, TRG_Zavod).Set(srcZavod);
                                }
                                catch { }

                                // KAZGOR_Единица измерения								
                                try
                                {
                                    var srcEdIzm = GetParameter(fitting, SRC_EdIzm).AsString();
                                    GetParameter(fitting, TRG_EdIzm).Set(srcEdIzm);
                                }
                                catch { }

                                // KAZGOR_Количество								
                                try
                                {
                                    GetParameter(fitting, TRG_Count).Set(1);
                                }
                                catch { }

                                // KAZGOR_Масса
                                try
                                {
                                    var srcMass = GetParameter(fitting, SRC_Mass).AsString();
                                    GetParameter(fitting, TRG_Mass).Set(srcMass);
                                }
                                catch { }

                                // KAZGOR_Примечание								
                                try
                                {
                                    var srcPrim = GetParameter(fitting, SRC_Primech).AsString();
                                    GetParameter(fitting, TRG_Primech).Set(srcPrim);
                                }
                                catch { }
                            }
                            break;

                        default:

                            break;
                    }

                }

                trans.Commit();
            }





            return Result.Succeeded;
        }

        private FamilyInstance GetRootFamily(FamilyInstance fi)
        {
            var superComponent = fi.SuperComponent as FamilyInstance;
            if (superComponent != null)
            {
                superComponent = GetRootFamily(superComponent);
            }
            else
            {
                superComponent = fi;
            }
            return superComponent;
        }

        private void CopyGeneralElementsParameters(IList<Element> elements)
        {
            foreach (Element fitting in elements)
            {
                var fittingType = doc.GetElement(fitting.GetTypeId());

                try
                {
                    var srcName = GetParameter(fitting, SRC_Naimen).AsString();
                    GetParameter(fitting, TRG_Naimen).Set(srcName);
                }
                catch { }

                // KAZGOR_Марка
                try
                {
                    var srcMarka = GetParameter(fitting, SRC_Marka).AsString();
                    GetParameter(fitting, TRG_Marka).Set(srcMarka);
                }
                catch { }

                // KAZGOR_Код Изделия							
                try
                {
                    var srcKod = GetParameter(fitting, SRC_KodIzd).AsString();
                    GetParameter(fitting, TRG_KodIzd).Set(srcKod);
                }
                catch { }

                // KAZGOR_Завод изготовитель								
                try
                {
                    var srcZavod = GetParameter(fitting, SRC_Zavod).AsString();
                    GetParameter(fitting, TRG_Zavod).Set(srcZavod);
                }
                catch { }

                // KAZGOR_Единица измерения								
                try
                {
                    var srcEdIzm = GetParameter(fitting, SRC_EdIzm).AsString();
                    GetParameter(fitting, TRG_EdIzm).Set(srcEdIzm);
                }
                catch { }

                // KAZGOR_Количество								
                try
                {
                    GetParameter(fitting, TRG_Count).Set(1);
                }
                catch { }

                // KAZGOR_Масса
                try
                {
                    var srcMass = GetParameter(fitting, SRC_Mass).AsString();
                    GetParameter(fitting, TRG_Mass).Set(srcMass);
                }
                catch { }

                // KAZGOR_Примечание								
                try
                {
                    var srcPrim = GetParameter(fitting, SRC_Primech).AsString();
                    GetParameter(fitting, TRG_Primech).Set(srcPrim);
                }
                catch { }
            }
        }



        private IList<Element> GetElementsByCatIdValue(int idIntValue)
        {
            FilteredElementCollector filter = new FilteredElementCollector(doc);
            filter.OfCategory((BuiltInCategory)idIntValue).WhereElementIsNotElementType();

            return filter.ToElements();
        }

        private Parameter GetParameter(Element elem, string paramStr, bool typeParameter = false)
        {
            // Get instance parameter
            var param = (elem.LookupParameter(paramStr) as Parameter);

            // If parameter is not an instance parameter, then try to get type parameter
            if (null == param || typeParameter)
            {
                return elem.Document.GetElement(elem.GetTypeId()).LookupParameter(paramStr) as Parameter;
            }
            return param;
        }

        private double GetPercentGlobal()
        {
            double percent = 0;
            GlobalParameter percentValue = new FilteredElementCollector(doc)
                .OfClass(typeof(GlobalParameter))
                .Cast<GlobalParameter>()
                .Where(gp => gp.Name.Equals("Запас"))
                .FirstOrDefault();
            DoubleParameterValue dVal = percentValue.GetValue() as DoubleParameterValue;
            percent = dVal.Value;
            return percent;
        }

    }

    [Transaction(TransactionMode.Manual)]
    public class Numbering : IExternalCommand
    {
        Document doc = null;
        const string S_System = "С_Система";
        const string S_Order = "С_Сортировка";
        const string S_Pos = "С_Позиция";
        const string S_Name = "С_Наименование";
        const string S_Marka = "С_Марка";


        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            doc = uidoc.Document;

            var orderDict = new Dictionary<string, int>();
            var pos = 0;
            if (doc.ActiveView.ViewType == ViewType.Schedule && doc.ActiveView.Name == "В_Спецификация по ГОСТ_Позиция")
            {
                ViewSchedule viewSchedule = doc.ActiveView as ViewSchedule;
                TableData tableData = viewSchedule.GetTableData();
                TableSectionData tableSectionData = tableData.GetSectionData(SectionType.Body);
                if (tableSectionData.NumberOfRows > 0)
                {
                    var currentSystem = "ImposibleSystemName";
                    for (int rInd = 0; rInd < tableSectionData.NumberOfRows; rInd++)
                    {

                        var system = viewSchedule.GetCellText(SectionType.Body, rInd, 0);
                        var order = viewSchedule.GetCellText(SectionType.Body, rInd, 1);
                        var name = viewSchedule.GetCellText(SectionType.Body, rInd, 3);
                        var marka = viewSchedule.GetCellText(SectionType.Body, rInd, 4);

                        if (currentSystem != system)
                        {
                            pos = 0;
                            currentSystem = system;
                        }

                        var key = system + order + name + marka;
                        orderDict.Add(key, pos);

                        pos++;
                    }

                    var elements = new FilteredElementCollector(doc, viewSchedule.Id);
                    elements.WhereElementIsNotElementType();




                    using (Transaction tr = new Transaction(doc, "Задание номера позиции элементам"))
                    {
                        tr.Start();
                        foreach (Element element in elements)
                        {
                            try
                            {
                                var srcOrder = (element.LookupParameter(S_Order) as Parameter).AsString();
                                var srcSystem = (element.LookupParameter(S_System) as Parameter).AsString();
                                var srcName = (element.LookupParameter(S_Name) as Parameter).AsString();
                                var srcMarka = (element.LookupParameter(S_Marka) as Parameter).AsString();

                                var key = srcSystem + srcOrder + srcName + srcMarka;
                                if (orderDict.ContainsKey(key))
                                {
                                    (element.LookupParameter(S_Pos) as Parameter).Set(orderDict[key].ToString());
                                }
                            }
                            catch { }
                        }

                        tr.Commit();
                    }


                }
            }
            else
            {
                TaskDialog.Show("Предупреждение", "Для автонумерации требуется открыть спецификацию В_Спецификация по ГОСТ_Позиция");
            }

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class AutoSchedule : IExternalCommand
    {
        Document doc = null;
        const string S_System = "С_Система";
        const string S_Order = "С_Сортировка";
        const string S_Pos = "С_Позиция";
        const string S_Name = "С_Наименование";
        const string S_Marka = "С_Марка";


        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            doc = uidoc.Document;

            var orderDict = new Dictionary<string, int>();
            var systemHash = new HashSet<string>();
            var pos = 0;

            // Поиск шаблонного вида для оформления # Спецификация для оформления
            var schFilCol = new FilteredElementCollector(doc);
            var auxSchedules = schFilCol.OfClass(typeof(ViewSchedule)).Where(x => x.Name == "# Спецификация для оформления" || x.Name == "В_Спецификация по ГОСТ_Позиция").ToList();

            var templateSchedule = auxSchedules.Where(x => x.Name == "# Спецификация для оформления").FirstOrDefault();
            var positionSchedule = auxSchedules.Where(x => x.Name == "В_Спецификация по ГОСТ_Позиция").FirstOrDefault();



            if (templateSchedule != null && positionSchedule != null)
            {
                ViewSchedule viewSchedule = positionSchedule as ViewSchedule;

                TableData tableData = viewSchedule.GetTableData();
                TableSectionData tableSectionData = tableData.GetSectionData(SectionType.Body);
                if (tableSectionData.NumberOfRows > 0)
                {
                    var elements = new FilteredElementCollector(doc, viewSchedule.Id);
                    elements.WhereElementIsNotElementType();

                    using (Transaction tr = new Transaction(doc, "Создание оформленных спецификаций по системам"))
                    {
                        tr.Start();
                        foreach (Element element in elements)
                        {
                            try
                            {
                                var systemName = element.LookupParameter(S_System).AsString();
                                systemName = systemName.Trim();
                                var isAdded = systemHash.Add(systemName);
                                if (isAdded && !String.IsNullOrEmpty(systemName))
                                {
                                    var schedule = doc.GetElement((templateSchedule as ViewSchedule).Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;
                                    try
                                    {
                                        var x = schedule.LookupParameter("KAZGOR_Раздел проекта");
                                        x.Set("Автоспецификация");
                                    }
                                    catch { }

                                    try
                                    {
                                        schedule.Name = "О_Спецификация_" + systemName;
                                    }
                                    catch
                                    {
                                        var count = 1;
                                        while (count < 5)
                                        {
                                            try
                                            {
                                                schedule.Name = "О_Спецификация_" + systemName + " копия" + count;
                                                break;
                                            }
                                            catch
                                            {
                                                count++;
                                            }
                                        }
                                    }
                                    tableData = schedule.GetTableData();
                                    var headerSection = tableData.GetSectionData(SectionType.Header);
                                    headerSection.SetCellText(0, 1, "Система " + systemName);

                                    var filters = schedule.Definition.GetFilters();
                                    var systemFilter = filters[1];
                                    var newSystemFilter = new ScheduleFilter(systemFilter.FieldId, systemFilter.FilterType, systemName);
                                    schedule.Definition.RemoveFilter(1);
                                    schedule.Definition.AddFilter(newSystemFilter);
                                }
                            }
                            catch { }
                        }


                        tr.Commit();
                    }
                }



            }
            else
            {
                TaskDialog.Show("Предупреждение", "Для создания видов нужны следующие:\n 1. Открыть спецификацию В_Спецификация по ГОСТ_Позиция\n 2. Наличие шаблонного вида \"#Спецификация для оформления\" в проекте");
            }

            return Result.Succeeded;
        }
    }
}
