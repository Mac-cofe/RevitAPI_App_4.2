using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_lab4._2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            TaskDialog.Show("Запись данных", $"Данное приложение запишет в файл Excel данные о трубах: имя типа трубы, наружный диаметр трубы, " +
                $" { Environment.NewLine}внутренний диаметр трубы, длина трубы { Environment.NewLine} " +
                $" Файл будет создан на Рабочем столе, имя файла pipes.xlsx" +
               $"{ Environment.NewLine} После выполнения появится диалоговое окно с сообщением");


            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    string pipeName = pipe.Name;
                    double outerDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                    double outerDiam = UnitUtils.ConvertFromInternalUnits(outerDiamParam, UnitTypeId.Millimeters);
                    double innerDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                    double innerDiam = UnitUtils.ConvertFromInternalUnits(innerDiamParam, UnitTypeId.Millimeters);
                    double lengthParam = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    double length = UnitUtils.ConvertFromInternalUnits(lengthParam, UnitTypeId.Millimeters);
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipeName);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outerDiam);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, innerDiam);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, length);
                    rowIndex++;
                }
                workBook.Write(stream);
                workBook.Close();
            }

            TaskDialog.Show("Запись данных", $"{Environment.NewLine} Данные записаны в файл pipes.xlsx, на Рабочем столе" +
              $"{ Environment.NewLine} Нажмите кнопку Закрыть для завершения работы приложения");
            return Result.Succeeded;
        }
    }
}