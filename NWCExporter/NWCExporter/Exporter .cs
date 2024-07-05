using System;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitNwcExporter
{
    public class NwcExporter
    {
        private readonly UIApplication _uiApp;

        public NwcExporter(UIApplication uiApp)
        {
            _uiApp = uiApp;
        }

        public void ExportNwcFiles(string directoryPath)
        {
            try
            {
                // Get all Revit files in the specified directory and its subdirectories
                string[] revitFiles = Directory.GetFiles(directoryPath, "*.rvt", SearchOption.AllDirectories);

                if (revitFiles.Length == 0)
                {
                    TaskDialog.Show("No Revit Files Found", "No .rvt files found in the specified directory and subdirectories.");
                    return;
                }

                foreach (var filePath in revitFiles)
                {
                    // Open the Revit file
                    Document openedDoc = _uiApp.Application.OpenDocumentFile(filePath);

                    // Define export options
                    NavisworksExportOptions options = new NavisworksExportOptions
                    {
                        ExportScope = NavisworksExportScope.Model
                    };

                    // Export NWC file
                    string nwcFilePath = Path.ChangeExtension(filePath, ".nwc");
                    openedDoc.Export(Path.GetDirectoryName(nwcFilePath), Path.GetFileNameWithoutExtension(nwcFilePath), options);

                    openedDoc.Close(false);
                }

                TaskDialog.Show("Export Complete", "NWC files exported successfully.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
            }
        }
    }
}
