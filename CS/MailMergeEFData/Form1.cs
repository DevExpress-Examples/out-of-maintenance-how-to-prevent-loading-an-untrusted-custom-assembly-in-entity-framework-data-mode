using DevExpress.DataAccess.EntityFramework;
using DevExpress.DataAccess.UI.Wizard.Services;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraSpreadsheet.Services;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MailMergeEFData {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            // Prompt for loading a data assembly; default is NeverLoad.
            this.spreadsheetControl1.Options.DataSourceLoading.CustomAssemblyBehavior = DevExpress.XtraSpreadsheet.SpreadsheetCustomAssemblyBehavior.Prompt;
            //// Uncomment the string below to disable browsing for a custom assembly in the Data Source Wizard.
            // this.spreadsheetControl1.Options.DataSourceWizard.CustomAssemblyBehavior = DevExpress.XtraSpreadsheet.SpreadsheetCustomAssemblyBehavior.NeverLoad;
            // Handle this event to decide whether to load a custom data assembly.
            this.spreadsheetControl1.CustomAssemblyLoading += SpreadsheetControl1_CustomAssemblyLoading;
            // The service is employed in the Data Source Wizard.
            this.spreadsheetControl1.ReplaceService<ICustomAssemblyPathNotifier>(new myCustomAssemblyPathNotifier());
            // The service is employed when loading a template.
            this.spreadsheetControl1.ReplaceService<ICustomAssemblyLoadingNotificationService>(new myCustomAssemblyLoadingNotificationService());

        }

        private void SpreadsheetControl1_CustomAssemblyLoading(object sender, SpreadsheetCustomAssemblyLoadingEventArgs e) {
            // Decide whether to load a custom assembly.
            e.Cancel = MessageBox.Show(String.Format("Do you want to load data from {0}?", e.Path),
                    "Security Warning", MessageBoxButtons.YesNo) == DialogResult.No;
            // Decide whether to query the service for the final decision.
            e.Handled = MessageBox.Show(String.Format("Is this your final decision?", e.Path),
                    "Security Warning", MessageBoxButtons.YesNo) == DialogResult.Yes; ;
        }

        private void Form1_Load(object sender, EventArgs e) {
            EFDataSource ds = new EFDataSource(new EFConnectionParameters());
            ds.Name = "Contacts";
            ds.ConnectionParameters.CustomAssemblyPath = Application.StartupPath + @"\EFDataModel.dll";
            ds.ConnectionParameters.CustomContextName = "EFDataModel.ContactsEntities";
            ds.ConnectionParameters.ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;attachdbfilename=|DataDirectory|\Contacts.mdf;integrated security=True;";
            this.spreadsheetControl1.Document.MailMergeDataSource = ds;
            this.spreadsheetControl1.Document.MailMergeDataMember = "Customers";

            ds.Fill();
            
            this.spreadsheetControl1.Document.Worksheets[0].Cells["A1"].Formula = "=FIELD(\"Company\")";

            IList<IWorkbook> resultWorkbooks = spreadsheetControl1.Document.GenerateMailMergeDocuments();
            string filename = "SavedDocument0.xlsx";
            resultWorkbooks[0].SaveDocument(filename, DocumentFormat.OpenXml);
            System.Diagnostics.Process.Start(filename);
        }
    }


    public class myCustomAssemblyLoadingNotificationService : ICustomAssemblyLoadingNotificationService {
        public bool RequestApproval(string assemblyPath) {
            return MessageBox.Show(String.Format("Are you sure?\nLoading {0}", assemblyPath),
                "Final Security Control Warning", MessageBoxButtons.OKCancel) == DialogResult.OK;
        }
    }

    public class myCustomAssemblyPathNotifier : ICustomAssemblyPathNotifier {
        public bool RequestApproval(string assemblyPath, XtraForm owner) {
            return MessageBox.Show(String.Format("Are you sure?\nLoading {0}", assemblyPath),
                "Final Security Wizard Warning", MessageBoxButtons.OKCancel) == DialogResult.OK;
        }
    }

}
