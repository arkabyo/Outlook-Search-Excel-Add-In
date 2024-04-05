' Required imports for working with Ribbon UI, Outlook, and Excel
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Outlook
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Data
Imports System.Windows.Forms

Public Class SearchRibbon
    'Handles load events of the ribbon
    Private Sub SearchRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    ' Triggered when Search Button is clicked. It starts the search process.
    Private Sub button_Search_Click(sender As Object, e As RibbonControlEventArgs) Handles button_Search.Click
        'Call the SearchFunction subroutine. 
        SearchFunction()
    End Sub

    Private Sub SearchFunction()
        ' Ensure a folder has been selected before initiating the search
        If Globals.ThisAddIn.SelectedFolder Is Nothing Then
            MessageBox.Show("Please select a folder first.", "No Folder Selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Error Handling
        Dim selectedCells As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedCells Is Nothing OrElse selectedCells.Columns.Count > 1 OrElse
           selectedCells.Cells.Cast(Of Excel.Range).All(Function(c) String.IsNullOrWhiteSpace(CStr(c.Value))) Then
            MessageBox.Show("Please select one or more cells with content within a single column.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Not selectedCells.Areas(1).Rows(1).EntireRow.Hidden AndAlso selectedCells.Cells(1, 1).Row = 1 Then
            MessageBox.Show("Row 1 cannot be included in the search. Please select cells starting from Row 2.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Preparation for updating Excel with search results
        Dim xlSheet As Excel.Worksheet = selectedCells.Worksheet
        Dim lastColumn As Long = xlSheet.Cells(1, xlSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
        Dim subjectColumn As Long = FindOrCreateColumn(xlSheet, lastColumn, "Email Subject")
        Dim dateColumn As Long = FindOrCreateColumn(xlSheet, subjectColumn, "Latest Email Date")

        Dim olFolder As Outlook.Folder = Globals.ThisAddIn.SelectedFolder
        Dim items As Outlook.Items = olFolder.Items
        items.Sort("[ReceivedTime]", True)

        ProcessSearch(selectedCells, items, xlSheet, subjectColumn, dateColumn)

        MessageBox.Show("Task completed. Outlook search results have been updated.", "Search Completed", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Additional helper methods for clarity and separation of concerns
    Private Function FindOrCreateColumn(xlSheet As Excel.Worksheet, startColumn As Long, columnName As String) As Long
        For i As Integer = 1 To startColumn
            If xlSheet.Cells(1, i).Value?.ToString().ToLower() = columnName.ToLower() Then
                Return i ' Found the existing column
            End If
        Next
        ' Create a new column
        Dim newColumnIndex = startColumn + 1
        xlSheet.Cells(1, newColumnIndex).Value = columnName
        Return newColumnIndex
    End Function

    Private Sub ProcessSearch(selectedCells As Excel.Range, items As Outlook.Items, xlSheet As Excel.Worksheet, subjectColumn As Long, dateColumn As Long)
        For Each cell As Excel.Range In selectedCells
            Dim cellValue = If(cell.Value, String.Empty)
            Dim searchString As String = cellValue.ToString().ToLower()
            Try
                ' Assuming all items are MailItem; adjust as needed for calendar or contact items
                For Each item As Object In items
                    If TypeOf item Is Outlook.MailItem Then
                        Dim mailItem As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                        If $"{mailItem.SenderName} {mailItem.Subject} {mailItem.Body}".ToLower().Contains(searchString) Then
                            xlSheet.Cells(cell.Row, subjectColumn).Value = mailItem.Subject
                            xlSheet.Cells(cell.Row, dateColumn).Value = mailItem.SentOn.ToString()
                            Exit For ' Found a match, exit the loop
                        End If
                    End If
                Next
            Catch ex As SystemException
                MessageBox.Show($"An error occurred while searching: {ex.Message}", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
        Next
    End Sub

    ' Triggered when Select Folder button/Load Folder Button is clicked. It allows the user to select an Outlook folder for the search.
    Private Sub button_LoadFolder_Click(sender As Object, e As RibbonControlEventArgs) Handles button_LoadFolder.Click
        ' Create an Outlook Application object and obtain the MAPI namespace.
        Dim olApp As Outlook.Application
        Try
            olApp = Globals.ThisAddIn.Application
        Catch ex As SystemException
            olApp = CreateObject("Outlook.Application")
        End Try

        Dim olNamespace As Outlook.NameSpace = olApp.GetNamespace("MAPI")

        ' Prompt the user to pick a folder.
        Dim olFolder As Outlook.Folder = olNamespace.PickFolder()

        ' Check if a folder is selected and update global reference and UI labels.
        If Not olFolder Is Nothing Then
            Globals.ThisAddIn.SelectedFolder = olFolder

            ' Assuming label_SelectedFolder and label_Instruction are accessible and correctly initialized.
            Me.label_SelectedFolder.Label = "Selected Folder: " & Globals.ThisAddIn.SelectedFolder.Name
            Me.label_Instruction.Label = "Select cells in a column and click Search Email"
        Else
            MessageBox.Show("No folder was selected.", "Folder Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class

