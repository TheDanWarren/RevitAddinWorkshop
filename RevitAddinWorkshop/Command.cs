#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;
#endregion

namespace RevitAddinWorkshop
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            //use method GetOneExcelFile to select excel file with level and sheet information
            string filePath = GetOneExcelFile();
            //get the level data from the struct LevelData
            LevelData levelData = new LevelData(filePath);
            //get the sheet data from the struc RevitSheetData
            RevitSheetData sheetData = new RevitSheetData(filePath);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Levels and Views and place on Sheets");

                //create the levels and views
                ViewFamilyType curVFT = CollectVFT(doc);
                ViewFamilyType curRCPVFT = CollectVFT(doc, "ceiling");

                for (int i = 0; i < (levelData.levelName.Length); i++)
                {
                    Level newLevel = Level.Create(doc, levelData.levelElevation[i]);
                    newLevel.Name = levelData.levelName[i];

                    ViewPlan newFloorPlan = ViewPlan.Create(doc, curVFT.Id, newLevel.Id);
                    ViewPlan newRCP = ViewPlan.Create(doc, curRCPVFT.Id, newLevel.Id);
                    newRCP.Name = newRCP.Name + " RCP";
                }

                //Create Sheets and Place views on Sheets
                Element titleblock = GetTitleblock(doc, "E1 30x42 Horizontal");
                for (int i = 0; i < (sheetData.excelSheetName.Length); i++)
                {
                    ViewSheet newSheet = ViewSheet.Create(doc, titleblock.Id);
                    newSheet.Name = sheetData.excelSheetName[i];
                    newSheet.SheetNumber = sheetData.excelSheetNumber[i];                    
                    SetSheetParam(newSheet, "Drawn By", sheetData.excelDrawnBy[i]);
                    SetSheetParam(newSheet, "Checked By", sheetData.excelCheckedBy[i]);

                    View existingView = GetViewByName(doc, sheetData.excelView[i]);
                    if (existingView != null)
                    {
                        Viewport newVP = Viewport.Create(doc, newSheet.Id, existingView.Id, new XYZ(0, 0, 0));
                    }
                    else
                    {
                        TaskDialog.Show("Error", "Could not find view");
                    }
                }
                t.Commit();
            }
            return Result.Succeeded;
        }



                   
        //method to get an excel file name
        internal string GetOneExcelFile()
        {
            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm";

            string filePath = "";
            if (dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                filePath = dialog.FileName;
            }
            return filePath;
        }
        //method to get the titleblock information
        internal ElementType GetTitleblock(Document doc, string tbName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector.WhereElementIsElementType();

            foreach (ElementType element in collector)
            {
                if (element.Name == tbName)
                {
                    return element;
                }
            }
            return null;
        }
        //method to set the Sheet Name
        internal void SetSheetParam(ViewSheet sheet, string paramName, string value)
        {
            bool err = false;
            foreach (Parameter curParam in sheet.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(value);
                }
            }
        }
        //method to get the View Names
        internal View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View));

            foreach (View curView in collector)
            {
                if (curView.Name == viewName)
                {
                    return curView;
                }
            }
            return null;
        }
        //Method to get pull the view family types from the revit file.
        internal ViewFamilyType CollectVFT (Document doc, string planVFT = "floor")
        {
            //get all the view types from Revit
            FilteredElementCollector collectorVFT = new FilteredElementCollector(doc);
            collectorVFT.OfClass(typeof(ViewFamilyType));

            //create VFT called curVFT
            ViewFamilyType curVFT = null;
            //create a VFT called curRCPVFT
            ViewFamilyType curRCPVFT = null;
            //Loop through all the VFTs and set FloorPlan or CeilingPlan
            foreach(ViewFamilyType element in collectorVFT)
            {
                if (element.ViewFamily == ViewFamily.FloorPlan)
                {
                    curVFT = element;
                }
                else if (element.ViewFamily == ViewFamily.CeilingPlan)
                {
                    curRCPVFT = element;
                }
            }
            //set name for the view types Floor/Ceiling/Null
            if(planVFT == "floor")
            {
                return curVFT;
            }
            else if(planVFT == "ceiling")
            {
                return curRCPVFT;
            }
            else
            {
                return null;
            }
        }
        //Struct to hold the LevelData in the excel sheet
        internal struct LevelData
        {
            //ExcelOverhead
            public Excel.Application ExcelApp;
            public Excel.Workbook excelWorkBook;
            public Excel.Worksheet excelWorkSheet;
            public Excel.Range excelRange;
            //need an int for the counter to loop through rows
            public int excelRowCount;
            //define struct data types
            public string[] levelName;
            public double[] levelElevation;

            //use whatever filePath happens to be set in GetOneExcelFile use it no current check
            public LevelData(string filePath)
            {
                //Open Excel
                ExcelApp = new Excel.Application();
                //set WorkBook to the excel file located at filePath
                excelWorkBook = ExcelApp.Workbooks.Open(filePath);
                //set WorkSheet to the worksheet with the level data provided by user
                excelWorkSheet = excelWorkBook.Worksheets[1];
                //set the range of the data
                excelRange = excelWorkSheet.UsedRange;
                //Get the number of rows used in excelWorksheet
                excelRowCount = excelRange.Rows.Count;
                //add level name and elevation to the struct
                levelName = new string[excelRowCount-1];
                levelElevation = new double[excelRowCount-1];

                //loop thought the rows and add the Level Names and Elevations to the struct
                for (int i = 2; i <= excelRowCount; i++)
                {                    
                    Excel.Range cellLevelName = excelWorkSheet.Cells[i, 1];
                    Excel.Range cellLevelElevation = excelWorkSheet.Cells[i, 2];
                    levelName[i - 2] = cellLevelName.Value.ToString();
                    levelElevation[i - 2] = cellLevelElevation.Value;
                }
                //close workbook and exit excel
                excelWorkBook.Close();
                ExcelApp.Quit();
            }
        }
        //create a structure to hold the revit sheet data from Excel file at filePath
        internal struct RevitSheetData
        {
            public Excel.Application Excelapp;
            public Excel.Workbook excelWorkBook;
            public Excel.Worksheet excelWorkSheet;
            public Excel.Range excelRange2;
            public int excelRowCount;
            //define the structure RevitSheetData
            public string[] excelSheetNumber;
            public string[] excelSheetName;
            public string[] excelView;
            public string[] excelDrawnBy;
            public string[] excelCheckedBy;

            public RevitSheetData(string filePath)
            {
                //create a new instance of Excel
                Excelapp = new Excel.Application();
                //Open the excel file at filePath
                excelWorkBook = Excelapp.Workbooks.Open(filePath);
                //sheet information is in worksheet 2
                excelWorkSheet = excelWorkBook.Worksheets.Item[2];
                //row count for the loop
                excelRange2 = excelWorkSheet.UsedRange;
                //set the row count
                excelRowCount = excelRange2.Rows.Count;
                //set Number, Name, View, DrawnBy, CheckedBy parameters
                excelSheetNumber = new string[excelRowCount - 1];
                excelSheetName = new string[excelRowCount - 1];
                excelView = new string[excelRowCount - 1];
                excelDrawnBy = new string[excelRowCount - 1];
                excelCheckedBy = new string[excelRowCount - 1];

                //loop through all the data on excel sheet and add it to the struct RevitSheetData
                for(int i=2; i <= excelRowCount; i++)
                {
                    Excel.Range cellSheetNumber = excelWorkSheet.Cells[i, 1];
                    Excel.Range cellSheetName = excelWorkSheet.Cells[i, 2];
                    Excel.Range cellView = excelWorkSheet.Cells[i, 3];
                    Excel.Range cellDrawnBy = excelWorkSheet.Cells[i, 4];
                    Excel.Range cellCheckedBy = excelWorkSheet.Cells[i, 5];

                    excelSheetNumber[i - 2] = cellSheetNumber.Value.ToString();
                    excelSheetName[i - 2] = cellSheetName.Value.ToString();
                    excelView[i - 2] = cellView.Value.ToString();
                    excelDrawnBy[i - 2] = cellDrawnBy.Value.ToString();
                    excelCheckedBy[i - 2] = cellCheckedBy.Value.ToString();                    
                }
                //Close out of Excel
                excelWorkBook.Close();
                Excelapp.Quit();

            }
        }        
    }
}