Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports OfficeOpenXml

Namespace MassLogicConsole
    Public Class EPPlusClass

        Public Shared Sub RunSample2(FilePath As String)
            Console.WriteLine("Reading column 2 of {0}", FilePath)
            Console.WriteLine()

            Dim existingFile As New FileInfo(FilePath)
            Using package As New ExcelPackage(existingFile)
                ' get the first worksheet in the workbook
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(1)
                Dim col As Integer = 2
                'The item description
                ' output the data in column 2
                For row As Integer = 2 To 4
                    Console.WriteLine(vbTab & "Cell({0},{1}).Value={2}", row, col, worksheet.Cells(row, col).Value)
                Next

                ' output the formula in row 5
                'Console.WriteLine(vbTab & "Cell({0},{1}).Formula={2}", 3, 5, worksheet.Cells(3, 5).Formula)
                'Console.WriteLine(vbTab & "Cell({0},{1}).FormulaR1C1={2}", 3, 5, worksheet.Cells(3, 5).FormulaR1C1)

                ' output the formula in row 5
                'Console.WriteLine(vbTab & "Cell({0},{1}).Formula={2}", 5, 3, worksheet.Cells(5, 3).Formula)

                'Console.WriteLine(vbTab & "Cell({0},{1}).FormulaR1C1={2}", 5, 3, worksheet.Cells(5, 3).FormulaR1C1)
            End Using
            ' the using statement automatically calls Dispose() which closes the package.
            Console.WriteLine()
            Console.WriteLine("Sample 2 complete")
            Console.WriteLine()
        End Sub

        Public Shared Function ReadExcelToTable(FilePath As String) As DataTable
            Dim tmpData = New DataTable
            Try

                Dim myColumn As DataColumn

                'Add Extra Columns
                myColumn = New DataColumn()
                With myColumn
                    .DataType = System.Type.GetType("System.String")
                    .ColumnName = "F1"
                    .DefaultValue = ""
                    .ReadOnly = False
                    .Unique = False
                End With
                tmpData.Columns.Add(myColumn)

                myColumn = New DataColumn()
                With myColumn
                    .DataType = System.Type.GetType("System.String")
                    .ColumnName = "F2"
                    .DefaultValue = ""
                    .ReadOnly = False
                    .Unique = False
                End With
                tmpData.Columns.Add(myColumn)

                Dim newRow As DataRow

                Dim existingFile As New FileInfo(FilePath)
                Using package As New ExcelPackage(existingFile)
                    ' get the first worksheet in the workbook
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(1)
                    Dim col As Integer = 2
                    'The item description
                    ' output the data in column 2

                    For row As Integer = 2 To 4000

                        If worksheet.Cells(row, col).Value = "" Then Exit For

                        newRow = tmpData.NewRow
                        newRow("F1") = worksheet.Cells(row, 1).Value
                        newRow("F2") = worksheet.Cells(row, 2).Value
                        tmpData.Rows.Add(newRow)

                        'Console.WriteLine(vbTab & "Cell({0},{1}).Value={2}", row, col, worksheet.Cells(row, col).Value)
                    Next

                End Using
            Catch ex As Exception

            End Try

            Return tmpData

        End Function

    End Class
End Namespace

