Imports DevExpress.DataAccess.EntityFramework
Imports DevExpress.DataAccess.UI.Wizard.Services
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraSpreadsheet.Services

Namespace MailMergeEFData
    Partial Public Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            ' Prompt for loading a data assembly; default is NeverLoad.
            Me.spreadsheetControl1.Options.DataSourceLoading.CustomAssemblyBehavior = DevExpress.XtraSpreadsheet.SpreadsheetCustomAssemblyBehavior.Prompt
            '// Uncomment the string below to disable browsing for a custom assembly in the Data Source Wizard.
            ' this.spreadsheetControl1.Options.DataSourceWizard.CustomAssemblyBehavior = DevExpress.XtraSpreadsheet.SpreadsheetCustomAssemblyBehavior.NeverLoad;
            ' Handle this event to decide whether to load a custom data assembly.
            AddHandler Me.spreadsheetControl1.CustomAssemblyLoading, AddressOf SpreadsheetControl1_CustomAssemblyLoading
            ' The service is employed in the Data Source Wizard.
            Me.spreadsheetControl1.ReplaceService(Of ICustomAssemblyPathNotifier)(New myCustomAssemblyPathNotifier())
            ' The service is employed when loading a template.
            Me.spreadsheetControl1.ReplaceService(Of ICustomAssemblyLoadingNotificationService)(New myCustomAssemblyLoadingNotificationService())

        End Sub

        Private Sub SpreadsheetControl1_CustomAssemblyLoading(ByVal sender As Object, ByVal e As SpreadsheetCustomAssemblyLoadingEventArgs)
            ' Decide whether to load a custom assembly.
            e.Cancel = MessageBox.Show(String.Format("Do you want to load data from {0}?", e.Path), "Security Warning", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No
            ' Decide whether to query the service for the final decision.
            e.Handled = MessageBox.Show(String.Format("Is this your final decision?", e.Path), "Security Warning", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes

        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
            Dim ds As New EFDataSource(New EFConnectionParameters())
            ds.Name = "Contacts"
            ds.ConnectionParameters.CustomAssemblyPath = Application.StartupPath & "\EFDataModel.dll"
            ds.ConnectionParameters.CustomContextName = "EFDataModel.ContactsEntities"
            ds.ConnectionParameters.ConnectionString = "Data Source=(LocalDB)\MSSQLLocalDB;attachdbfilename=|DataDirectory|\Contacts.mdf;integrated security=True;"
            Me.spreadsheetControl1.Document.MailMergeDataSource = ds
            Me.spreadsheetControl1.Document.MailMergeDataMember = "Customers"

            ds.Fill()

            Me.spreadsheetControl1.Document.Worksheets(0).Cells("A1").Formula = "=FIELD(""Company"")"

            Dim resultWorkbooks As IList(Of IWorkbook) = spreadsheetControl1.Document.GenerateMailMergeDocuments()
            Dim filename As String = "SavedDocument0.xlsx"
            resultWorkbooks(0).SaveDocument(filename, DocumentFormat.OpenXml)
            System.Diagnostics.Process.Start(filename)
        End Sub
    End Class


    Public Class myCustomAssemblyLoadingNotificationService
        Implements ICustomAssemblyLoadingNotificationService

        Public Function RequestApproval(ByVal assemblyPath As String) As Boolean Implements ICustomAssemblyLoadingNotificationService.RequestApproval
            Return MessageBox.Show(String.Format("Are you sure?" & ControlChars.Lf & "Loading {0}", assemblyPath), "Final Security Control Warning", MessageBoxButtons.OKCancel) = DialogResult.OK
        End Function
    End Class

    Public Class myCustomAssemblyPathNotifier
        Implements ICustomAssemblyPathNotifier

        Public Function RequestApproval(ByVal assemblyPath As String, ByVal owner As XtraForm) As Boolean Implements ICustomAssemblyPathNotifier.RequestApproval
            Return MessageBox.Show(String.Format("Are you sure?" & ControlChars.Lf & "Loading {0}", assemblyPath), "Final Security Wizard Warning", MessageBoxButtons.OKCancel) = DialogResult.OK
        End Function
    End Class

End Namespace
