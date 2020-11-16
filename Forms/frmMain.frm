VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Order Importer"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   3600
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Component"
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   6855
      Begin VB.ComboBox cmbPackage 
         Height          =   315
         Left            =   5400
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbSubCategory 
         Height          =   315
         Left            =   2880
         TabIndex        =   29
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Import Component"
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CheckBox chkExported 
         Caption         =   "Imported"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   24
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   5820
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   315
         Left            =   2400
         TabIndex        =   22
         Top             =   5820
         Width           =   375
      End
      Begin VB.TextBox txtItemNumber 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Text            =   "0"
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   315
         Left            =   600
         TabIndex        =   19
         Top             =   5820
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   5820
         Width           =   375
      End
      Begin VB.CommandButton cmdLoadWebsite 
         Caption         =   "Distributor Website"
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtDatasheetURL 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   4695
      End
      Begin MSFlexGridLib.MSFlexGrid grdProperties 
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483644
         GridColorFixed  =   -2147483644
         HighLight       =   2
         GridLinesFixed  =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
      End
      Begin VB.TextBox txtNotes 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   5400
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label9 
         Caption         =   "Package:"
         Height          =   255
         Left            =   5400
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Sub-Category:"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblNumberItems 
         Alignment       =   2  'Center
         Caption         =   "/ 000"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   5860
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Datasheet URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
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
         ItemData        =   "frmMain.frx":030A
         Left            =   4320
         List            =   "frmMain.frx":0311
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

' Browse for PartsCatalog database.
Private Sub OpenDatabaseFile()
    ' Setup open dialog.
    dlgCommon.DialogTitle = "Open Database"
    dlgCommon.DefaultExt = "mdb"
    dlgCommon.Filter = "Microsoft Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*"
    dlgCommon.FileName = ""
    dlgCommon.ShowOpen
    
    ' TODO: Set the new database.
    
    ' Set the path.
    If dlgCommon.FileName <> "" Then
        txtWorkspace.Text = GetComponentsDir(dlgCommon.FileName)
    End If
End Sub

' Imports an order into the system.
Public Sub ImportOrder()
    ' Parse order and populate the components array.
    ParseFarnellOrder txtOrderLocation.Text
    
    ' Update the components record counter and show the first component.
    lblNumberItems.Caption = "of " & LastComponentIndex
    ShowComponent 0
End Sub

' Imports the current component into the database.
Private Sub ImportCurrentComponent()
    Dim component As component
    
    ' Check if there's a component selected.
    If Not fraComponent.Enabled Then
        MsgBox "There isn't a component selected. We can't import this.", _
            vbOKOnly + vbCritical, "No Component Selected"
        Exit Sub
    End If
    
    ' Don't forget to save any changes and get the current component as well.
    SaveCurrentComponent
    Set component = GetCurrentComponent
    
    ' Check if the current component has already been exported.
    If component.Exported Then
        MsgBox "Currently we can't modify a component that has already been imported.", _
            vbOKOnly + vbCritical, "Operation Not Permitted"
        Exit Sub
    End If
    
    ' Set the component as exported.
    component.Export
    ShowComponent CLng(txtItemNumber.Text)
    
    ' Give the user some feedback.
    If component.Exported Then
        MsgBox component.Name & " imported successfully.", vbOKOnly + vbInformation, _
            "Component Exported"
    Else
        MsgBox component.Name & " import failed.", vbOKOnly + vbCritical, _
            "Failed to Export Component"
    End If
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

' Opens the component distributor website with a search in place.
Private Sub LoadCurrentComponentWebsite()
    Dim component As component
    Set component = GetCurrentComponent
    
    OpenURL "https://pt.farnell.com/search?st=" & component.SearchCode
End Sub

' Browse for the order file to load.
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

' Import current component into the database.
Private Sub cmdExport_Click()
    ImportCurrentComponent
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
    LoadCurrentComponentWebsite
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
    ' If no database is associated, browse for one first.
    If Not IsDatabaseAssociated Then
        OpenDatabaseFile
    End If
    
    ' Import the current component if a database is associated.
    If IsDatabaseAssociated Then
        ImportCurrentComponent
    End If
End Sub

' Component > Load Website clicked.
Private Sub mniComponentLoadWebsite_Click()
    LoadCurrentComponentWebsite
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
