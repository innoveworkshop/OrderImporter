VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "PartCat Order Importer"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Import Component"
      Height          =   615
      Left            =   5880
      TabIndex        =   32
      Top             =   6600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   3600
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Component"
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   6855
      Begin VB.CommandButton cmdAddPackage 
         Height          =   615
         Left            =   6000
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton cmdAddSubCategory 
         Height          =   615
         Left            =   6000
         Picture         =   "frmMain.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton cmdAddCategory 
         Height          =   615
         Left            =   6000
         Picture         =   "frmMain.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2520
         Width           =   735
      End
      Begin VB.CheckBox chkExported 
         Caption         =   "Imported"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   315
         Left            =   6360
         TabIndex        =   27
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   315
         Left            =   5880
         TabIndex        =   26
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtItemNumber 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4560
         TabIndex        =   24
         Text            =   "0"
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   315
         Left            =   4080
         TabIndex        =   23
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   315
         Left            =   3600
         TabIndex        =   22
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdLoadWebsite 
         Caption         =   "Load Website"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox txtDatasheetURL 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid grdProperties 
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtNotes 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   5400
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblNumberItems 
         Alignment       =   2  'Center
         Caption         =   "/ 000"
         Height          =   255
         Left            =   5280
         TabIndex        =   25
         Top             =   4750
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Datasheet URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraWorkspace 
      Caption         =   "PartCat Workspace"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   5415
      Begin VB.TextBox txtWorkspace 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   4815
      End
      Begin VB.CommandButton cmdBrowseWorkspace 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Component Export Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "Distributor Order"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cmbDistributor 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":0E98
         Left            =   4320
         List            =   "frmMain.frx":0E9F
         TabIndex        =   5
         Text            =   "Farnell"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Load Order"
         Height          =   615
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBrowseOrder 
         Caption         =   "..."
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtOrderLocation 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Distributor:"
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Exported File Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mniFileLoadOrder 
         Caption         =   "&Load Order..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mniFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuComponent 
      Caption         =   "&Component"
      Begin VB.Menu mniComponentPrevious 
         Caption         =   "P&revious"
      End
      Begin VB.Menu mniComponentNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu mniComponentSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentAddProperty 
         Caption         =   "&Add Property..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mniComponentDeleteProperty 
         Caption         =   "&Delete Propety"
         Shortcut        =   ^D
      End
      Begin VB.Menu mniComponentSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentLoadWebsite 
         Caption         =   "Load &Website..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mniComponentSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentExport 
         Caption         =   "&Import"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mniHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmMain
''' Application's main form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Imports an order into the system.
Public Sub ImportOrder()
    ' Parse order and populate the components array.
    ParseFarnellOrder txtOrderLocation.Text
    
    ' Update the components record counter and show the first component.
    lblNumberItems.Caption = "of " & LastComponentIndex
    ShowComponent 0
End Sub

' Shows a component by its index.
Public Sub ShowComponent(lngIndex As Long)
    ' Get the component.
    Dim component As component
    Set component = GetComponent(lngIndex)
    
    ' Set text fields.
    txtItemNumber.Text = CStr(lngIndex)
    txtName.Text = component.Name
    txtQuantity.Text = CStr(component.Quantity)
    txtNotes.Text = component.Notes
    txtDatasheetURL.Text = component.Datasheet
    
    ' Set the exported checkbox.
    If component.Exported Then
        chkExported.Value = vbChecked
    Else
        chkExported.Value = vbUnchecked
    End If
    
    ' Preparate the grid for the properties.
    grdProperties.Rows = UBound(component.Properties) + 2
    
    ' Populate the properties.
    Dim intIndex As Integer
    Dim astrProperty() As String
    For intIndex = 0 To UBound(component.Properties)
        ' Check if the property is populated.
        If component.Property(intIndex) <> "" Then
            astrProperty = Split(component.Property(intIndex), ": ")
            grdProperties.TextMatrix(intIndex + 1, 0) = astrProperty(0)
            grdProperties.TextMatrix(intIndex + 1, 1) = astrProperty(1)
        Else
            grdProperties.TextMatrix(1, 0) = ""
            grdProperties.TextMatrix(1, 1) = ""
        End If
    Next intIndex
    
    ' Enable the component panel for editing and menu item.
    fraComponent.Enabled = True
    mnuComponent.Enabled = True
End Sub

' Gets the currently selected component.
Public Function GetCurrentComponent() As component
    Set GetCurrentComponent = GetComponent(CLng(txtItemNumber.Text))
End Function

' Saves the text fields to the current component.
Public Sub SaveCurrentComponent()
    Dim component As component
    Set component = GetCurrentComponent
    
    ' Save the component text fields.
    component.Name = txtName.Text
    component.Notes = txtNotes.Text
    component.Datasheet = txtDatasheetURL.Text
    component.Quantity = CLng(txtQuantity.Text)
End Sub

' Deletes the currently selected property.
Public Sub DeleteSelectedProperty()
    Dim strKey As String
    Dim component As component

    ' Get component and selected key.
    Set component = GetCurrentComponent
    strKey = grdProperties.TextMatrix(grdProperties.Row, 0)
    
    ' Actually delete the property.
    component.DeleteProperty strKey
    
    ' Save component changes and reload the view.
    SaveCurrentComponent
    ShowComponent CLng(txtItemNumber.Text)
End Sub

' Add a category to the properties
Private Sub cmdAddCategory_Click()
    Dim strCategory As String
    Dim component As component
    
    ' Get the category from the user.
    Set component = GetCurrentComponent
    strCategory = InputBox("Component category:", "Set the component category")
    
    ' Check if the user entered something and add the property.
    If strCategory <> "" Then
        component.AddProperty "Category", strCategory
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Add a package to the properties.
Private Sub cmdAddPackage_Click()
    Dim strPackage As String
    Dim component As component
    
    ' Get the package from the user.
    Set component = GetCurrentComponent
    strPackage = InputBox("Component package:", "Set the component package")
    
    ' Check if the user entered something and add the property.
    If strPackage <> "" Then
        component.AddProperty "Package", strPackage
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Add a sub-category to the properties.
Private Sub cmdAddSubCategory_Click()
    Dim strSubCategory As String
    Dim component As component
    
    ' Get the sub-category from the user.
    Set component = GetCurrentComponent
    strSubCategory = InputBox("Component sub-category:", _
        "Set the component sub-category")
    
    ' Check if the user entered something and add the property.
    If strSubCategory <> "" Then
        component.AddProperty "Sub-Category", strSubCategory
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Browse for order file.
Private Sub cmdBrowseOrder_Click()
    ' Setup open dialog.
    dlgCommon.DialogTitle = "Import Distributor Order File"
    dlgCommon.DefaultExt = "csv"
    dlgCommon.Filter = "Comma Separated Files (*.csv)|*.csv|All Files (*.*)|*.*"
    dlgCommon.FileName = ""
    dlgCommon.ShowOpen
    
    ' Set the path.
    txtOrderLocation.Text = dlgCommon.FileName
End Sub

' Browe for PartCat workspace.
Private Sub cmdBrowseWorkspace_Click()
    ' Setup open dialog.
    dlgCommon.DialogTitle = "Select PartCat Workspace"
    dlgCommon.DefaultExt = "pcw"
    dlgCommon.Filter = "PartCat Workspace (*.pcw)|*.pcw|All Files (*.*)|*.*"
    dlgCommon.FileName = ""
    dlgCommon.ShowOpen
    
    ' Set the path.
    If dlgCommon.FileName <> "" Then
        txtWorkspace.Text = GetComponentsDir(dlgCommon.FileName)
    End If
End Sub

' Export component to a workspace.
Private Sub cmdExport_Click()
    Dim component As component
    
    ' Check if there's a component selected.
    If Not fraComponent.Enabled Then
        MsgBox "There isn't a component selected. We can't export this.", _
            vbOKOnly + vbCritical, "No Component Selected"
        Exit Sub
    End If
    
    ' Check if there's an output folder selected.
    If txtWorkspace.Text = "" Then
        MsgBox "No destination workspace selected. Please select one before exporting.", _
            vbOKOnly + vbInformation, "No Export Workspace Selected"
        Exit Sub
    End If
    
    ' Don't forget to save any changes and get the current component as well.
    SaveCurrentComponent
    Set component = GetCurrentComponent
    
    ' Set the component as exported.
    component.Export txtWorkspace.Text
    ShowComponent CLng(txtItemNumber.Text)
    
    ' Give the user some feedback.
    If component.Exported Then
        MsgBox component.Name & " exported successfully.", vbOKOnly + vbInformation, _
            "Component Exported"
    Else
        MsgBox component.Name & " export failed.", vbOKOnly + vbCritical, _
            "Failed to Export Component"
    End If
End Sub

' Go to the first component in the records.
Private Sub cmdFirst_Click()
    SaveCurrentComponent
    ShowComponent 0
End Sub

' Import the order file.
Private Sub cmdImport_Click()
    ' Check if there's an order file selected.
    If txtOrderLocation.Text = "" Then
        MsgBox "No order file selected. Please select one before importing.", _
            vbOKOnly + vbInformation, "No Order File Selected"
        Exit Sub
    End If
    
    ' Actually import the data.
    ImportOrder
End Sub

' Go to the last component in the records.
Private Sub cmdLast_Click()
    SaveCurrentComponent
    ShowComponent LastComponentIndex
End Sub

' Opens the component distributor website with a search in place.
Private Sub cmdLoadWebsite_Click()
    Dim component As component
    Set component = GetCurrentComponent
    
    OpenURL "https://pt.farnell.com/search?st=" & component.SearchCode
End Sub

' Go to the next component in the records.
Private Sub cmdNext_Click()
    Dim lngCurrentIndex As Long
    
    lngCurrentIndex = CLng(txtItemNumber.Text)
    If lngCurrentIndex < LastComponentIndex Then
        SaveCurrentComponent
        ShowComponent lngCurrentIndex + 1
    End If
End Sub

' Go to the previous component in the records.
Private Sub cmdPrevious_Click()
    Dim lngCurrentIndex As Long
    
    lngCurrentIndex = CLng(txtItemNumber.Text)
    If lngCurrentIndex > 0 Then
        SaveCurrentComponent
        ShowComponent lngCurrentIndex - 1
    End If
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Setup the Flex Grid.
    grdProperties.TextMatrix(0, 0) = "Property"
    grdProperties.TextMatrix(0, 1) = "Value"
    grdProperties.ColWidth(0) = (grdProperties.Width / 2) - 45
    grdProperties.ColWidth(1) = (grdProperties.Width / 2) - 45
    
    ' Disable the component panel and menu.
    fraComponent.Enabled = False
    mnuComponent.Enabled = False
End Sub

' User wants to edit a property.
Private Sub grdProperties_DblClick()
    Dim strKey As String
    Dim strValue As String
    Dim component As component
    
    ' Get properties and get user input.
    Set component = GetCurrentComponent
    strKey = grdProperties.TextMatrix(grdProperties.Row, 0)
    strValue = grdProperties.TextMatrix(grdProperties.Row, 1)
    strValue = InputBox(strKey & ":", "Edit Property", strValue)
    
    ' Change property if the user entered something.
    If strValue <> "" Then
        component.EditProperty strKey, strValue
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Check for keypresses on the properties grid.
Private Sub grdProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Delete a property for the user.
    If KeyCode = vbKeyDelete Then
        DeleteSelectedProperty
    End If
End Sub

' Component > Add Property clicked.
Private Sub mniComponentAddProperty_Click()
    If dlgProperty.ShowAdd(Me) Then
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Component > Delete Property clicked.
Private Sub mniComponentDeleteProperty_Click()
    DeleteSelectedProperty
End Sub

' Component > Import menu clicked.
Private Sub mniComponentExport_Click()
    If txtWorkspace.Text = "" Then
        cmdBrowseWorkspace_Click
    End If
    
    If txtWorkspace.Text <> "" Then
        cmdExport_Click
    End If
End Sub

' Component > Load Website clicked.
Private Sub mniComponentLoadWebsite_Click()
    cmdLoadWebsite_Click
End Sub

' Component > Next menu clicked.
Private Sub mniComponentNext_Click()
    cmdNext_Click
End Sub

' Component > Previous menu clicked.
Private Sub mniComponentPrevious_Click()
    cmdPrevious_Click
End Sub

' File > Exit menu clicked.
Private Sub mniFileExit_Click()
    Unload Me
End Sub

' File > Load Order menu clicked.
Private Sub mniFileLoadOrder_Click()
    cmdBrowseOrder_Click
    If txtOrderLocation.Text <> "" Then
        cmdImport_Click
    End If
End Sub

' Help > About menu clicked.
Private Sub mniHelpAbout_Click()
    frmAbout.Show
End Sub
