using BlazorCustomExportToPDF.Module.BusinessObjects;
using BlazorCustomExportToPDF.Module.Helpers;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.ExpressApp.Blazor.Editors;
using DevExpress.ExpressApp.Blazor.Editors.Grid;
using DevExpress.ExpressApp.Blazor.Editors.Models;
using DevExpress.ExpressApp.Xpo;
using DevExpress.Persistent.Base;
using DevExpress.Web.Data;
using DevExpress.Web;
using DevExpress.Xpo;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
using System.Windows.Forms;
using System.Diagnostics;
using DevExpress.Utils.CommonDialogs.Internal;
using DevExpress.ExpressApp.Blazor;
using Microsoft.JSInterop;

namespace BlazorCustomExportToPDF.Blazor.Server.Controllers
{
    public class CustomExportController : ViewController
    {
        SimpleAction customExport;
        public CustomExportController()
        {
            customExport = new SimpleAction(this, "customExport", PredefinedCategory.Export);
            customExport.Execute += customExport_Execute;
            customExport.Caption = "Export PDF";
        }
        private void customExport_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            
            var listEditor = ((ListView)View).Editor as DxGridListEditor;

            if (listEditor != null)
            {              
                IDxGridAdapter adapter = listEditor.GetGridAdapter();
                IList<string> propertyNames = new List<string>();

                foreach (DxGridDataColumnModel columnModel in adapter.GridDataColumnModels)
                {
                    if (!String.IsNullOrWhiteSpace(columnModel.Caption))
                    {
                        if (columnModel.Visible)
                            propertyNames.Add((columnModel).FieldName);                      
                    }
                }             

                string fields = String.Empty;
                foreach (var detail in propertyNames)
                {
                    fields += detail;
                    if (detail != propertyNames.Last()) { fields += ";"; }
                }

                var dataView = ObjectSpace.CreateDataView(View.ObjectTypeInfo.Type, fields, null, null);

                //Create Report
                XtraReport report = new XtraReport();
                report.DataSource = dataView;
                ReportHelper.CreateReport(report, propertyNames.ToArray());
                var reportName = View.ObjectTypeInfo.Name;
                report.Name = reportName + " Report.pdf";
                report.CreateDocument();

                //Export PDF Options
                PdfExportOptions pdfOptions = report.ExportOptions.Pdf;
                pdfOptions.PdfACompatibility = PdfACompatibility.PdfA3b;
                pdfOptions.DocumentOptions.Application = "Test Application";
                pdfOptions.DocumentOptions.Author = "DX Documentation Team";
                pdfOptions.DocumentOptions.Keywords = "Xari, Reporting, PDF";
                pdfOptions.DocumentOptions.Producer = Environment.UserName.ToString();
                pdfOptions.DocumentOptions.Subject = "Document Subject";
                pdfOptions.DocumentOptions.Title = "Document Title";

                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), report.Name);
                report.ExportToPdf(path, pdfOptions);

                

                IJSRuntime jsRuntime = (Application as BlazorApplication).ServiceProvider.GetRequiredService<IJSRuntime>();

                if (!String.IsNullOrEmpty(path))
                {
                     jsRuntime.InvokeVoidAsync("BlazorDownloadFile", report.Name, Convert.ToBase64String(File.ReadAllBytes(path)), ".pdf", "application/pdf");
                }


            }
        }
    }
}
