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

#endregion

namespace RevitAddinWorkshop
{
    [Transaction(TransactionMode.Manual)]
    public class FizzBuzz : IExternalCommand
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

            double offset = 0.05;
            double offsetCalc = 0.05 * doc.ActiveView.Scale;

            XYZ curPoint = new XYZ(0, 0, 0);
            XYZ offsetPoint = new XYZ(0, offsetCalc, 0);

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create Text Note");
            t.Start();

            int range = 100;
            for (int i = 0; i <= range; i++)
            {
                if (i % 3 == 0 && i % 5 == 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "FizzBuzz", collector.FirstElementId());
                }
                else if (i % 3 == 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "Fizz", collector.FirstElementId());
                }
                else if (i % 5 == 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "Buzz", collector.FirstElementId());
                }
                else if (i % 5 != 0)
                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, i.ToString(), collector.FirstElementId());
                }
                curPoint = curPoint.Subtract(offsetPoint);

            }

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }

        //defining a method
        internal double Method01(double a, double b)
        {
            double c = a + b;
            
            Debug.Print("Got here" + c.ToString());
            
            return c;
        }
    }
}
