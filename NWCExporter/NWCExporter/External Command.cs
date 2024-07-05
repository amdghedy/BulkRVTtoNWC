using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitNwcExporter
{
    [Transaction(TransactionMode.Manual)]
    public class ExportNwcCommand : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // Ask the user to select the directory
                string directoryPath = SelectDirectory();
                if (string.IsNullOrEmpty(directoryPath))
                {
                    return Result.Cancelled;
                }

                // Create an instance of the exporter and call the export method
                NwcExporter exporter = new NwcExporter(uiApp);
                exporter.ExportNwcFiles(directoryPath);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private string SelectDirectory()
        {
            using (var folderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderDialog.Description = "Select Directory Containing Revit Files";
                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
            }
            return null;
        }
    }
}
