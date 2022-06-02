Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel


Public Class OpenExcels
    Private filepath As String
    Private visible As Boolean = False
    Private sheetname As String = "Sheet1"
    Public Shared excelCounter As Integer = 0

    Public Property percorsoFile() As String
        Get
            Return filepath
        End Get
        Set(value As String)
            filepath = value
        End Set
    End Property
    Public Property Visibile() As Boolean
        Get
            Return visible
        End Get
        Set(value As Boolean)
            visible = value
        End Set
    End Property
    Public Property nomeFoglio() As String
        Get
            Return sheetname
        End Get
        Set(value As String)
            sheetname = value
        End Set
    End Property

    'Default constructor
    Sub New()
        excelCounter += 1
    End Sub

    'Constructor
    Sub New(ByVal filepath As String, ByVal Visible As Boolean, ByVal sheetName As String = "Sheet1")
        excelCounter += 1
    End Sub
    'This function returns a Workbook
    Function returnWorkBook() As Excel.Workbook
        Try

            Dim xlsApp As New Excel.Application()
            xlsApp.Visible = Me.Visibile
            Dim xlsWorkBook As Excel.Workbook = xlsApp.Workbooks.Open(Me.filepath)
            Return xlsWorkBook

        Catch ex As Exception

            MessageBox.Show("Couldn't find the specified path", ex.Message)
            Dim excel As New Excel.Application()
            Dim workbooks As Excel.Workbooks = excel.Workbooks
            Dim workbook As Excel.Workbook = workbooks.Add()
            Dim worksheet As Excel.Worksheet = excel.ActiveSheet
            workbook.saveas2(filepath)
            Return workbook

            Exit Function
        End Try
    End Function

    'This function returns a Sheet
    Public Function returnWorkSheet() As Excel.Worksheet
        Try
            Dim xlsApp As New Excel.Application
            xlsApp.Visible = Me.Visibile
            Dim xlsWorkbook As Excel.Workbook = xlsApp.Workbooks.Open(Me.filepath)
            Dim xlsSheets As Excel.Worksheets = xlsWorkbook.Worksheets
            Dim xlsSheet As Excel.Worksheet = xlsSheets.get_Item("Sheet1")
            Return xlsSheet
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Dim xlsApp As New Excel.Application
            xlsApp.Visible = Me.Visibile
            Dim xlsWorkbook As Excel.Workbook = xlsApp.Workbooks.Open(Me.filepath)
            Dim xlsSheets As Excel.Worksheets = xlsWorkbook.Worksheets
            Dim xlsSheet As Excel.Worksheet = xlsApp.ActiveSheet
            Return xlsSheet
        End Try
    End Function

    Public Shared Function getColumnsbyHeader(ByVal Header As String, ByVal xlsSheet As Excel.Worksheet) As List(Of String)
        Dim ColumnsByHeader As New List(Of String)
        Try

            For Each row In xlsSheet.Rows
                For Each name As String In xlsSheet.Columns.Nameù
                    If row(name) = row(Header) Then
                        ColumnsByHeader.Add(row(name))
                    End If
                Next name
            Next row
            Return ColumnsByHeader
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return ColumnsByHeader
        End Try

    End Function


End Class
