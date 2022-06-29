#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

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

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TextNotes);
            collector.WhereElementIsNotElementType();

            int counter = 0;
            Transaction t = new Transaction(doc, "Text to upper");
            t.Start();

            foreach(Element element in collector)
            {
                TextNote textNote = element as TextNote;
                textNote.Text = textNote.Text.ToUpper();
                counter++;
            }

            t.Commit();
            t.Dispose();

            TaskDialog.Show("Complete", "Changed " + counter.ToString() + " text notes to UPPER");

            return Result.Succeeded;
        }

        private static int addNumber(int num1, int num2)
        {
            return num1 + num2; 
        }


    }
}
