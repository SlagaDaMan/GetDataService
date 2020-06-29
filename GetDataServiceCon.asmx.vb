Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports System.Web.Script.Serialization
Imports System.IO

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class GetDataServiceCon
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function HelloWorld(Sql As String) As String
        'the DataTable to return
        Dim MyDataTable As New DataTable

        'make a SqlConnection using the supplied ConnectionString 
        Dim MySqlConnection As New SqlConnection("Data Source=KVJHB-3700\SQLEXPRESS;Initial Catalog=OfficeManagementSystem;Integrated Security=True")
        Using MySqlConnection
            'make a query using the supplied Sql
            Dim MySqlCommand As SqlCommand = New SqlCommand(Sql, MySqlConnection)

            'open the connection
            MySqlConnection.Open()

            'create a DataReader and execute the SqlCommand
            Dim MyDataReader As SqlDataReader = MySqlCommand.ExecuteReader()

            'load the reader into the datatable
            MyDataTable.Load(MyDataReader)

            'clean up
            MyDataReader.Close()
        End Using
        Dim Json = GetJson(MyDataTable)
        'return the datatable
        Return Json
    End Function

    <WebMethod()>
    Public Function ConvertAndSavePDF(FinalHTML As String, xOrderID As String) As String
        Dim stream As MemoryStream = Nothing
        Dim PdfConverter As EvoPdf.HtmlToPdf.PdfConverter = New EvoPdf.HtmlToPdf.PdfConverter()
        PdfConverter.LicenseKey = "8NvC0MPD0MHBxcfQw97A0MPB3sHC3snJyck="
        PdfConverter.PdfDocumentOptions.PdfPageSize = EvoPdf.HtmlToPdf.PdfPageSize.A4
        PdfConverter.PdfDocumentOptions.PdfCompressionLevel = EvoPdf.HtmlToPdf.PdfCompressionLevel.Normal
        PdfConverter.PdfDocumentOptions.PdfPageOrientation = EvoPdf.HtmlToPdf.PdfPageOrientation.Portrait
        Dim result = PdfConverter.GetPdfBytesFromHtmlString(FinalHTML)
        System.IO.File.WriteAllBytes(My.Settings.OrderID.ToString() & xOrderID & ".pdf", result)
        Return ""
    End Function

    Public Function GetJson(ByVal dt As DataTable) As String

        Return New JavaScriptSerializer().Serialize(From dr As DataRow In dt.Rows Select dt.Columns.Cast(Of DataColumn)().ToDictionary(Function(col) col.ColumnName, Function(col) dr(col)))

    End Function

End Class