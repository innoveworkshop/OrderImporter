VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "PartCat Order Importer"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
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
      Height          =   5175
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   6855
      Begin VB.CheckBox chkExported 
         Caption         =   "Exported"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   29
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   315
         Left            =   6360
         TabIndex        =   28
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   315
         Left            =   5880
         TabIndex        =   27
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox txtItemNumber 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4560
         TabIndex        =   25
         Text            =   "0"
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   315
         Left            =   4080
         TabIndex        =   24
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   315
         Left            =   3600
         TabIndex        =   23
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdLoadWebsite 
         Caption         =   "Load Website"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox txtDatasheetURL 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid grdProperties 
         Height          =   2055
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   6615
         _ExtentX        =   11668
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
         TabIndex        =   18
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   5400
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblNumberItems 
         Alignment       =   2  'Center
         Caption         =   "/ 000"
         Height          =   255
         Left            =   5280
         TabIndex        =   26
         Top             =   4750
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Datasheet URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
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
      Width           =   6855
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export Component"
         Height          =   615
         Left            =   5640
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
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
         Caption         =   "Workspace Export Location:"
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
         ItemData        =   "frmMain.frx":0000
         Left            =   4320
         List            =   "frmMain.frx":0007
         TabIndex        =   5
         Text            =   "Farnell"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
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
      Begin VB.Menu mniFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
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

' Browse for order file.
Private Sub cmdBrowseOrder_Click()
    ' Setup open dialog.
    dlgCommon.DialogTitle = "Import Distributor Order File"
    dlgCommon.DefaultExt = "csv"
    dlgCommon.Filter = "Comma Separated Files (*.csv)|*.csv|All Files (*.*)|*.*"
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
    dlgCommon.ShowOpen
    
    ' Set the path.
    txtWorkspace.Text = dlgCommon.FileName
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Setup the Flex Grid.
    grdProperties.TextMatrix(0, 0) = "Property"
    grdProperties.TextMatrix(0, 1) = "Value"
    grdProperties.ColWidth(0) = (grdProperties.Width / 2) - 45
    grdProperties.ColWidth(1) = (grdProperties.Width / 2) - 45
End Sub

' Exits the application.
Private Sub mniFileExit_Click()
    Unload Me
End Sub
