#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

#endregion

namespace RevitAddinWorkshop
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDeleteBackups : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            //set variables
            int counter = 0;
            string logPath = "";

            //create list for log file
            List<string> deletedFileLog = new List<string>();
            deletedFileLog.Add("The following backup files have been deleted:");

            FolderBrowserDialog selectFolder = new FolderBrowserDialog();
            selectFolder.ShowNewFolderButton = false;

            // open folder dialog and only run code if a folder is selected
            if(selectFolder.ShowDialog() == DialogResult.OK)
            {
                // get the selected folder path
                string directory = selectFolder.SelectedPath;

                // get all files from selected folder
                string[] files = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories);

                // loop through files
                foreach(string file in files)
                {
                    // check if the file is a Revit File
                    if(Path.GetExtension(file) == ".rvt" || Path.GetExtension(file) == ".rfa")
                    {
                        // get the last 8 characters of file name to check if backup
                        string checkString = file.Substring(file.Length - 8, 8);

                        // remove the file extension from the string checkString
                        string checkString2 = checkString.Substring(0, 4);

                        

                        // convert the string pulled from the Revit Journal log number into an integer for a logical check.
                        int checkInt = 0;
                        bool success = int.TryParse(checkString2, out checkInt);
                        

                        // confirm the integer is between 
                        if(checkInt > 0 && checkInt < 9999)
                        {
                            deletedFileLog.Add(file);
                            File.Delete(file);

                            //increment counter
                            counter++;
                        }

                    }
                }
                // output log file
                if (counter > 0)
                {
                    logPath = WriteListToTxt(deletedFileLog, directory);
                }
            }

            // alert user
            TaskDialog td = new TaskDialog("Complete");
            td.MainInstruction = "Deleted " + counter.ToString() + " backup files.";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Click to view log file");
            td.CommonButtons = TaskDialogCommonButtons.Ok;

            TaskDialogResult result = td.Show();

            if(result == TaskDialogResult.CommandLink1)
            {
                Process.Start(logPath);
            }

            return Result.Succeeded;
        }
          

        internal string WriteListToTxt(List<string> stringList, string filePath)
        {
            string fileName = "_Delete Backup Files.txt";
            string fullPath = filePath + @"\" + fileName;

            File.WriteAllLines(fullPath, stringList);

            return fullPath;
        }
    }
}
