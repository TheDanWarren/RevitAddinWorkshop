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

            //hardcode location of excel file
            string excelFile = @"C:\DWARREN\Session02_Challenge.xlsx";

            //excel overhead
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWsLevels = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWsSheets = excelWb.Worksheets.Item[2];

            //get count of total excel rows for levels
            Excel.Range excelRangeLevels = excelWsLevels.UsedRange;
            int rowCountLevels = excelRangeLevels.Rows.Count;

            //get count of total excel rows for Sheets
            Excel.Range excelRangeSheets = excelWsSheets.UsedRange;
            int rowCountSheets = excelRangeSheets.Rows.Count;

            //Make an array called dataListLevels
            //Make an array called dataListSheets
            //set array length to total count of rows. Array must know how large it is.
            List<string[]> dataListLevels = new List<string[]>();
            List<string[]> dataListSheets = new List<string[]>();

            //create a string for each Sheet Number and Sheet Name from Excel
            //this creates a string for each and then adds those to the array
            //***Remember-data has headers so you'll need to skip first row.
            //Do levels first then Sheets
            for (int i1 = 2; i1<= rowCountLevels; i1++)
            {
                //pull the value from excel, ROW THEN COLUMN.
                //Column 1 = Name Column 2 = Elevation in ft
                Excel.Range excelLevelName = excelWsLevels.Cells[i1, 1];
                Excel.Range excelLevelElev = excelWsLevels.Cells[i1, 2];

                string dataLevelName = excelLevelName.Value.ToString();
                string dataLevelElev = excelLevelElev.Value.ToString();


                string[] dataArrayLevels = new string[2];
                dataArrayLevels[0] = dataLevelName;
                dataArrayLevels[1] = dataLevelElev;

                dataListLevels.Add(dataArrayLevels);

                //create level
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Levels");
                    
                    double curDouble = Double.Parse(dataLevelElev);

                    Level curLevel = Level.Create(doc, curDouble);

                    t.Commit();
                }
            }

            for (int i2 = 2; i2 <= rowCountSheets; i2++)
            {
                //pull the value from excel, ROW THEN COLUMN.
                //column 1 = Sheet Number; Column 2 = Sheet Name
                Excel.Range excelSheetNumber = excelWsSheets.Cells[i2, 1];
                Excel.Range excelSheetName = excelWsSheets.Cells[i2, 2];

                string dataSheetNumber = excelSheetNumber.Value.ToString();
                string dataSheetName = excelSheetName.Value.ToString();

                string[] dataArraySheets = new string[2];
                dataArraySheets[0] = dataSheetNumber;
                dataArraySheets[1] = dataSheetName;

                dataListSheets.Add(dataArraySheets);

                //create sheet
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Sheets");

                    //get titleblock ID number for sheet creation
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                    collector.WhereElementIsElementType();

                    ViewSheet cursheet = ViewSheet.Create(doc, collector.FirstElementId());
                    cursheet.SheetNumber = dataSheetNumber;
                    cursheet.Name = dataSheetName;

                    t.Commit();
                }
            }
            //close excel
            excelWb.Close();
            excelApp.Quit();

            //contrats, you won
            return Result.Succeeded;
        }

        //method - this is an example of a method that adds two int together.
        private static int addNumber(int num1, int num2)
        {
            return num1 + num2;
        }


    }
}
