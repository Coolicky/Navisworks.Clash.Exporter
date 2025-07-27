using System;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Plugins;
using Navisworks.Clash.Exporter.Extensions;
using Navisworks.Clash.Exporter.Methods;
using Application = Autodesk.Navisworks.Api.Application;

namespace Navisworks.Clash.Exporter
{
    [Plugin("Navisworks.Clash.Exporter.Export", "COOL", DisplayName = "Export Clashes")]
    [AddInPlugin(AddInLocation.Export, LoadForCanExecute = true,
        CallCanExecute = CallCanExecute.DocumentNotClear,
        Icon = "Images\\Icon_16.png", LargeIcon = "Images\\Icon_32.png")]
    public class Export : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                var dialog = new SaveFileDialog();
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog() != DialogResult.OK) return 0;
                var fileName = dialog.FileName;

                var progress = Application.BeginProgress();
                
                var clashManager = Application.MainDocument.GetClash();
                var clashTests = clashManager.TestsData;

                progress.BeginSubOperation(0.85);
                var dataTables = ClashExport.CreateDataTables(clashTests, progress);
                progress.EndSubOperation();

                progress.BeginSubOperation(0.5, "Exporting clashes to Excel...");
                var excel = dataTables.ToExcel();
                progress.EndSubOperation();

                progress.BeginSubOperation(1, "Saving Excel file...");
                excel.SaveAs(fileName);
                progress.EndSubOperation();
                Application.EndProgress();

                MessageBox.Show("Successfully exported clashes to " + fileName);
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e.Message}\n{e.StackTrace}");
                return 0;
            }

            return 1;
        }
    }
}