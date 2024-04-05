' Necessary import for working with the Microsoft Office Interop libraries.
Imports Microsoft.Office.Interop

Public Class ThisAddIn
    ' A private shared field to hold the selected Outlook folder.
    ' "Shared" means it's accessible without an instance of ThisAddIn class.
    Private Shared _selectedFolder As Outlook.Folder

    ' A public shared property to get or set the selected Outlook folder.
    ' Other parts of the add-in can use this property to access the folder selected by the user.
    Public Shared Property SelectedFolder As Outlook.Folder
        Get
            ' Return the currently selected Outlook folder.
            Return _selectedFolder
        End Get
        Set(value As Outlook.Folder)
            ' Set the currently selected Outlook folder to a new value.
            _selectedFolder = value
        End Set
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
