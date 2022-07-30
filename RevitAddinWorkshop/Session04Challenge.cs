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
    public class Session04Challenge : IExternalCommand
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic 6""");

            foreach(Element element in pickList)
            {
                //use 'is' to compare types vs '=' or '==' which compares values

                if(element is CurveElement)
                {
                    
                    CurveElement curve = (CurveElement)element;
                    CurveElement curve2 = element as CurveElement;

                    curveList.Add(curve);

                    GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;

                    switch(curGS.Name)
                    {
                        case "<Medium>":
                            Debug.Print("found a medium line style");
                            break;

                        case "<Thin Lines>":
                            Debug.Print("found a thin line");
                            break;

                        case "<Wide Lines>":
                            Debug.Print("found a thin line");
                            
                            break;

                        default: Debug.Print("found something else");
                            break;

                    }

                    Curve curCurve = curve.GeometryCurve;
                    XYZ startPoint = curCurve.GetEndPoint(0);
                    XYZ endpoint = curCurve.GetEndPoint(1);


                    Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15 , 0, false, false);

                    Debug.Print(curGS.Name);             
                }

            }

            TaskDialog.Show("Complete", curveList.Count.ToString());
            return Result.Succeeded;

        }        

        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element curElement in collector)
            {
                WallType wallType = CurveElement as WallType;

                if (wallType.Name == wallTypeName)
                    return wallType;
            }
            return null;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element curElem in collector)
            {
                Level curlevel = curElem as Level; 
                    
                if (curlevel.Name == levelName)
                    return curlevel;
            }
            return null;



        }
}