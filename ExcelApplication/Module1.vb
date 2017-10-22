Imports Excel = Microsoft.Office.Interop.Excel
Module Module1
    Public Class ExcelApplication
        Private xlApp As Excel.Application
        Private xlWorkbook As Excel.Workbook
        Public Sub New()
            xlApp = New Excel.Application
            If xlApp Is Nothing Then
                Console.WriteLine("Excel not installed!")
            End If
        End Sub

        Public Function CreateWorkbook(workbookName As String) As Excel.Workbook
            xlWorkbook = xlApp.Workbooks.Add()
            xlWorkbook.SaveAs(workbookName)
            'xlWorkbook.Close()
            Return xlWorkbook
        End Function

        Public Function OpenWorkbook(workbookName As String) As Excel.Workbook
            xlWorkbook = xlApp.Workbooks.Open(workbookName)
            'xlWorkbook.Close()
            Return xlWorkbook
        End Function

        Public Sub Insert(row As List(Of String))
            Dim xlWorksheet As Excel.Worksheet = xlWorkbook.Worksheets(1)
            Dim UsedRange = xlWorksheet.UsedRange
            Dim RowRange = UsedRange.Rows
            Dim targetRow As Int32 = RowRange.Count

            If targetRow <> 1 Then
                targetRow += 1
            End If
            xlWorksheet.Rows(1).Insert()
            For i As Integer = 1 To row.Count
                xlWorksheet.Cells(targetRow, i).value = row(i - 1)
            Next
            xlWorkbook.Save()
        End Sub

        Public Sub Close()
            xlWorkbook.Close()
        End Sub
    End Class

    Sub Main()
        Dim ea As ExcelApplication = New ExcelApplication()
        'Dim wb As Excel.Workbook = ea.CreateWorkbook("C:\Users\1305266\Desktop\workbook example\test.xlsx")
        Dim wb As Excel.Workbook = ea.OpenWorkbook("C:\Users\1305266\Desktop\workbook example\test.xlsx")
        Dim InputRow As New List(Of String) From {"4", "66", "3", "4"}
        ea.Insert(InputRow)
        ea.Close()
    End Sub

End Module
