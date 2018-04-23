Imports DevExpress.DataAccess.EntityFramework
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Services
Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace MailMergeEFData
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Public Sub New()
            InitializeComponent()
            ' Handle this event to decide whether this application allows loading a custom data assembly.
            AddHandler EFDataSource.BeforeLoadCustomAssemblyGlobal, AddressOf EFDataSource_BeforeLoadCustomAssemblyGlobal
            ' Prompt for loading a data assembly; default is NeverLoad.
            Me.spreadsheetControl1.Options.DataSourceLoading.CustomAssemblyBehavior = DevExpress.XtraSpreadsheet.SpreadsheetCustomAssemblyBehavior.Prompt
            ' Handle this event to decide whether to load a custom data assembly.
            AddHandler Me.spreadsheetControl1.CustomAssemblyLoading, AddressOf SpreadsheetControl1_CustomAssemblyLoading
            ' The service is employed when loading a template.
            Me.spreadsheetControl1.ReplaceService(Of ICustomAssemblyLoadingNotificationService)(New myCustomAssemblyLoadingNotificationService())
        End Sub

        Private Sub EFDataSource_BeforeLoadCustomAssemblyGlobal(ByVal sender As Object, ByVal args As DevExpress.DataAccess.EntityFramework.BeforeLoadCustomAssemblyEventArgs)
            args.AllowLoading = True
        End Sub

        Private Sub SpreadsheetControl1_CustomAssemblyLoading(ByVal sender As Object, ByVal e As SpreadsheetCustomAssemblyLoadingEventArgs)
            ' Decide whether to load a custom assembly.
            e.Cancel = MessageBox.Show(String.Format("Do you want to load data from {0}?", e.Path), "CustomAssemblyLoading Event", MessageBoxButtons.YesNo) = DialogResult.No
            ' Decide whether to query the service for the final decision.
            e.Handled = MessageBox.Show(String.Format("Query the service for the final decision?", e.Path), "CustomAssemblyLoading Event", MessageBoxButtons.YesNo) = DialogResult.No

        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
            Dim ds As New EFDataSource(New EFConnectionParameters())
            ds.Name = "Contacts"
            ds.ConnectionParameters.CustomAssemblyPath = Application.StartupPath & "\EFDataModel.dll"
            ds.ConnectionParameters.CustomContextName = "EFDataModel.ContactsEntities"
            ds.ConnectionParameters.ConnectionString = "Data Source=(LocalDB)\MSSQLLocalDB;attachdbfilename=|DataDirectory|\Contacts.mdf;integrated security=True;"
            Me.spreadsheetControl1.Document.MailMergeDataSource = ds
            Me.spreadsheetControl1.Document.MailMergeDataMember = "Customers"
            Me.spreadsheetControl1.Document.Worksheets(0).Cells("A1").Formula = "=FIELD(""Company"")"

            Try
                Dim resultWorkbooks As IList(Of IWorkbook) = spreadsheetControl1.Document.GenerateMailMergeDocuments()
                Dim filename As String = "SavedDocument0.xlsx"
                resultWorkbooks(0).SaveDocument(filename, DocumentFormat.OpenXml)
                System.Diagnostics.Process.Start(filename)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception")
            End Try
        End Sub

        Private Sub tglShowWizardBrowseButton_CheckedChanged(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tglShowWizardBrowseButton.CheckedChanged
            ' Show the button in the Data Source Wizard to launch the "Browse for assembly" dialog.
            Me.spreadsheetControl1.Options.DataSourceWizard.ShowEFWizardBrowseButton = tglShowWizardBrowseButton.Checked
        End Sub
    End Class

    Public Class myCustomAssemblyLoadingNotificationService
        Implements ICustomAssemblyLoadingNotificationService

        Public Function RequestApproval(ByVal assemblyPath As String) As Boolean Implements ICustomAssemblyLoadingNotificationService.RequestApproval
            Return MessageBox.Show(String.Format("Are you sure?" & ControlChars.Lf & "Loading {0}", assemblyPath), "CustomAssemblyLoadingNotificationService", MessageBoxButtons.OKCancel) = DialogResult.OK
        End Function
    End Class
End Namespace
