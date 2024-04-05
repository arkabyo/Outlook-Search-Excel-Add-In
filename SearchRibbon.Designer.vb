Partial Class SearchRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SearchTab = Me.Factory.CreateRibbonTab
        Me.SearchGroup = Me.Factory.CreateRibbonGroup
        Me.button_LoadFolder = Me.Factory.CreateRibbonButton
        Me.button_Search = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.label_SelectedFolder = Me.Factory.CreateRibbonLabel
        Me.label_Instruction = Me.Factory.CreateRibbonLabel
        Me.SearchTab.SuspendLayout()
        Me.SearchGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'SearchTab
        '
        Me.SearchTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.SearchTab.Groups.Add(Me.SearchGroup)
        Me.SearchTab.Label = "Outlook Search"
        Me.SearchTab.Name = "SearchTab"
        '
        'SearchGroup
        '
        Me.SearchGroup.Items.Add(Me.button_LoadFolder)
        Me.SearchGroup.Items.Add(Me.button_Search)
        Me.SearchGroup.Items.Add(Me.Separator1)
        Me.SearchGroup.Items.Add(Me.label_SelectedFolder)
        Me.SearchGroup.Items.Add(Me.label_Instruction)
        Me.SearchGroup.Label = "Search Tools"
        Me.SearchGroup.Name = "SearchGroup"
        '
        'button_LoadFolder
        '
        Me.button_LoadFolder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.button_LoadFolder.Image = Global.Outlook_Search_Excel_Add_In.My.Resources.Resources.folder_icon
        Me.button_LoadFolder.Label = "Select Folder"
        Me.button_LoadFolder.Name = "button_LoadFolder"
        Me.button_LoadFolder.ShowImage = True
        '
        'button_Search
        '
        Me.button_Search.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.button_Search.Image = Global.Outlook_Search_Excel_Add_In.My.Resources.Resources.search_email
        Me.button_Search.Label = "Search Email"
        Me.button_Search.Name = "button_Search"
        Me.button_Search.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'label_SelectedFolder
        '
        Me.label_SelectedFolder.Label = "No folder is selected"
        Me.label_SelectedFolder.Name = "label_SelectedFolder"
        '
        'label_Instruction
        '
        Me.label_Instruction.Label = " "
        Me.label_Instruction.Name = "label_Instruction"
        '
        'SearchRibbon
        '
        Me.Name = "SearchRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SearchTab)
        Me.SearchTab.ResumeLayout(False)
        Me.SearchTab.PerformLayout()
        Me.SearchGroup.ResumeLayout(False)
        Me.SearchGroup.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SearchTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SearchGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents button_LoadFolder As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents button_Search As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents label_SelectedFolder As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents label_Instruction As Microsoft.Office.Tools.Ribbon.RibbonLabel
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SearchRibbon() As SearchRibbon
        Get
            Return Me.GetRibbon(Of SearchRibbon)()
        End Get
    End Property
End Class
