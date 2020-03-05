using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace sheetExporterRevit
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalDBApplication
    {
        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += DesignAutomationBridge_DesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        private void DesignAutomationBridge_DesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            LogTrace("Design Automation Ready event triggered...");
            e.Succeeded = true;
            DesignAutomationData data = e.DesignAutomationData;

            if (data == null)
                throw new ArgumentNullException(nameof(data));

            Application rvtApp = data.RevitApp;
            if (rvtApp == null)
                throw new InvalidDataException(nameof(rvtApp));

            string modelPath = data.FilePath;
            if (String.IsNullOrWhiteSpace(modelPath))
                throw new InvalidDataException(nameof(modelPath));

            Document doc = data.RevitDoc;
            if (doc == null)
                throw new InvalidOperationException("Could not open document.");

            using (Autodesk.Revit.DB.Transaction trans = new Transaction(doc, "ExportToDwgs"))
            {
                try
                {
                    trans.Start();

                    List<View> views = new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(vw =>
                           vw.ViewType == ViewType.DrawingSheet && !vw.IsTemplate
                        ).ToList();

                    List<ElementId> viewIds = new List<ElementId>();
                    foreach (View view in views)
                    {
                        ViewSheet viewSheet = view as ViewSheet;
                        viewIds.Add(viewSheet.Id);
                    }

                    DWGExportOptions options = new DWGExportOptions();
                    ExportDWGSettings dwgSettings = ExportDWGSettings.Create(doc, "mySetting");
                    options = dwgSettings.GetDWGExportOptions();
                    options.MergedViews = true;

                    doc.Export(Path.GetDirectoryName(modelPath) + "\\exportedDwgs", "rvt", viewIds, options);

                    //if (RuntimeValue.RunOnCloud)
                    //{
                    //    doc.Export(Directory.GetCurrentDirectory() + "\\exportedDwgs", "rvt", viewIds, options);
                    //}
                    //else
                    //{
                    //    // For local test
                    //    doc.Export(Path.GetDirectoryName(modelPath) + "\\exportedDwgs", "rvt", viewIds, options);
                    //}

                    trans.Commit();
                }
                catch (Exception ee)
                {
                    System.Diagnostics.Debug.WriteLine(ee.Message);
                    System.Diagnostics.Debug.WriteLine(ee.StackTrace);
                    return;
                }
                finally
                {
                    if (trans.HasStarted())
                        trans.RollBack();
                }
            }
        }

        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args) { System.Console.WriteLine(format, args); }
    }
}