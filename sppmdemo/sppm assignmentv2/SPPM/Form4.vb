Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form4

    Dim csvsave As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Form1.Show()
        LoginForm1.UsernameTextBox.Clear()
        LoginForm1.PasswordTextBox.Clear()
        Form1.Enabled = True

    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim managerdisplay As String
        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = Form1.App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = Form1.workbook.Worksheets("sheet1")

        Dim x As Integer
        Dim lLastRow As Long
        Dim stock As Integer

        stock = 0

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With

        If lLastRow = 1 Then
            MessageBox.Show("Nothing to Item Listed", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Button2.PerformClick()
            Return
        End If

        For x = 1 To lLastRow
            managerdisplay += Form1.worksheet.Cells(x, 1).value & vbTab & Form1.worksheet.Cells(x, 2).value & vbTab & Form1.worksheet.Cells(x, 3).value & vbTab & Form1.worksheet.Cells(x, 4).value & vbNewLine
            csvsave += Form1.worksheet.Cells(x, 1).value & "," & Form1.worksheet.Cells(x, 2).value & "," & Form1.worksheet.Cells(x, 3).value & "," & Form1.worksheet.Cells(x, 4).value & vbNewLine

        Next
        RichTextBox1.Text = managerdisplay

        Form1.workbook.Save()
        Form1.workbook.Close()
        Form1.App.Quit()

        Form1.releaseObject(Form1.worksheet)
        Form1.releaseObject(Form1.workbook)
        Form1.releaseObject(Form1.App)
    End Sub


    Private Sub Form4_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Button2.PerformClick()

        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click



    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim csvFile As String = My.Application.Info.DirectoryPath & "\Test.csv"
        Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(csvFile, False)


        outFile.WriteLine(csvsave)
        
        outFile.Close()
        Console.WriteLine(My.Computer.FileSystem.ReadAllText(csvFile))

        Label1.Text = "CSV successfully saved."

    End Sub
End Class