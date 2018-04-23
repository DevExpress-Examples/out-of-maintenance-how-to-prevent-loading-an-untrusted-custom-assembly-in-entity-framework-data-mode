using DevExpress.DataAccess.EntityFramework;
using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Services;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MailMergeEFData {
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm {
        public Form1() {
            InitializeComponent();
            // Handle this event to decide whether this application allows loading a custom data assembly.
            EFDataSource.BeforeLoadCustomAssemblyGlobal += EFDataSource_BeforeLoadCustomAssemblyGlobal;
            // Prompt for loading a data assembly; default is NeverLoad.
            this.spreadsheetControl1.Options.DataSourceLoading.CustomAssemblyBehavior = DevExpress.XtraSpreadsheet.SpreadsheetCustomAssemblyBehavior.Prompt;
            // Handle this event to decide whether to load a custom data assembly.
            this.spreadsheetControl1.CustomAssemblyLoading += SpreadsheetControl1_CustomAssemblyLoading;
            // The service is employed when loading a template.
            this.spreadsheetControl1.ReplaceService<ICustomAssemblyLoadingNotificationService>(new myCustomAssemblyLoadingNotificationService());
        }

        private void EFDataSource_BeforeLoadCustomAssemblyGlobal(object sender, DevExpress.DataAccess.EntityFramework.BeforeLoadCustomAssemblyEventArgs args) {
            args.AllowLoading = true;
        }

        private void SpreadsheetControl1_CustomAssemblyLoading(object sender, SpreadsheetCustomAssemblyLoadingEventArgs e) {
            // Decide whether to load a custom assembly.
            e.Cancel = MessageBox.Show(String.Format("Do you want to load data from {0}?", e.Path),
                    "CustomAssemblyLoading Event", MessageBoxButtons.YesNo) == DialogResult.No;
            // Decide whether to query the service for the final decision.
            e.Handled = MessageBox.Show(String.Format("Query the service for the final decision?", e.Path),
                    "CustomAssemblyLoading Event", MessageBoxButtons.YesNo) == DialogResult.No; ;
        }

        private void Form1_Load(object sender, EventArgs e) {
            EFDataSource ds = new EFDataSource(new EFConnectionParameters());
            ds.Name = "Contacts";
            ds.ConnectionParameters.CustomAssemblyPath = Application.StartupPath + @"\EFDataModel.dll";
            ds.ConnectionParameters.CustomContextName = "EFDataModel.ContactsEntities";
            ds.ConnectionParameters.ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;attachdbfilename=|DataDirectory|\Contacts.mdf;integrated security=True;";
            this.spreadsheetControl1.Document.MailMergeDataSource = ds;
            this.spreadsheetControl1.Document.MailMergeDataMember = "Customers";            
            this.spreadsheetControl1.Document.Worksheets[0].Cells["A1"].Formula = "=FIELD(\"Company\")";

            try {
                IList<IWorkbook> resultWorkbooks = spreadsheetControl1.Document.GenerateMailMergeDocuments();
                string filename = "SavedDocument0.xlsx";
                resultWorkbooks[0].SaveDocument(filename, DocumentFormat.OpenXml);
                System.Diagnostics.Process.Start(filename);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Exception");
            }
        }

        private void tglShowWizardBrowseButton_CheckedChanged(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            // Show the button in the Data Source Wizard to launch the "Browse for assembly" dialog.
            this.spreadsheetControl1.Options.DataSourceWizard.ShowEFWizardBrowseButton = tglShowWizardBrowseButton.Checked;
        }
    }

    public class myCustomAssemblyLoadingNotificationService : ICustomAssemblyLoadingNotificationService {
        public bool RequestApproval(string assemblyPath) {
            return MessageBox.Show(String.Format("Are you sure?\nLoading {0}", assemblyPath),
                "CustomAssemblyLoadingNotificationService", MessageBoxButtons.OKCancel) == DialogResult.OK;
        }
    }
}
