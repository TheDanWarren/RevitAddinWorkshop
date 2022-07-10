#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinWorkshop
{
    [Transaction(TransactionMode.Manual)]
    public class Session02Challenge : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;


            string excelFile = @"C:\DWARREN\Session 02_Combination Sheet List.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWs = excelWb.Worksheets.Item[1];

            Excel.Range excelRange = excelWs.UsedRange;
            int rowCount = excelRange.Rows.Count;

            //do some stuff in Excel
            List<string[]> dataList = new List<string[]>();

            for(int i = 1; i<= rowCount; i++)
            {
                Excel.Range cell1 = excelWs.Cells[i, 1];
                Excel.Range cell2 = excelWs.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                dataList.Add(dataArray);
            }
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create some Revit stuff");

                //create a level in Revit
                Level curLevel = Level.Create(doc, 100);

                //get titleblock ID number for sheet creation
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();


                ViewSheet cursheet = ViewSheet.Create(doc, collector.FirstElementId());
                cursheet.SheetNumber = "A101";
                cursheet.Name = "New Sheet";

                t.Commit();
            }
            //close excel
            excelWb.Close();
            excelApp.Quit();

            return Result.Succeeded;
        }

        private static int addNumber(int num1, int num2)
        {
            return num1 + num2;
        }


    }
}
