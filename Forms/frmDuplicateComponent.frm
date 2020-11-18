VERSION 5.00
Begin VB.Form frmDuplicateComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existing Component"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6855
   Icon            =   "frmDuplicateComponent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuantity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "000000"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Text            =   "Component Name"
      Top             =   360
      Width           =   3855
   End
   Begin VB.ComboBox cmbCategory 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Text            =   "Category"
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cmbSubCategory 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Text            =   "Sub-Category"
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cmbPackage 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Text            =   "Package"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtNotes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "Notes"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Qnt:"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Sub-Category:"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Package:"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmDuplicateComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmDuplicateComponent
''' Duplicate component dialog box.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_lngID As Long

' Positions this dialog by the side of an anchor frame in the parent window.
Public Sub PositionBySide(frmParent As Form, fraAnchor As Frame)
    Top = frmParent.Top + fraAnchor.Top + 800
    Left = frmParent.Left + frmParent.Width + 150
    
    ' Check if we should move to the right side of the parent form.
    If Screen.Width < Left Then
        Left = frmParent.Left - Width - 150
    End If
End Sub

' Populate Form from Recordset.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    Dim intIndex As Integer
    
    ' Store the component ID.
    m_lngID = rs.Fields("ID")
    
    ' Set text fields.
    txtName.Text = rs.Fields("Name")
    txtQuantity.Text = rs.Fields("Quantity")
    txtNotes.Text = rs.Fields("Notes")
    
    ' Set the categories.
    cmbSubCategory.Clear
    LoadCategories cmbCategory, False
    If rs.Fields("CategoryID") >= 0 Then
        For intIndex = 0 To cmbCategory.ListCount
            If cmbCategory.ItemData(intIndex) = rs.Fields("CategoryID") Then
                cmbCategory.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Load the sub-categories.
    LoadSubCategories rs.Fields("CategoryID"), cmbSubCategory, False
    If rs.Fields("SubCategoryID") >= 0 Then
        For intIndex = 0 To cmbSubCategory.ListCount
            If cmbSubCategory.ItemData(intIndex) = rs.Fields("SubCategoryID") Then
                cmbSubCategory.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Set the packages.
    LoadPackages cmbPackage, False
    If rs.Fields("PackageID") >= 0 Then
        For intIndex = 0 To cmbPackage.ListCount
            If cmbPackage.ItemData(intIndex) = rs.Fields("PackageID") Then
                cmbPackage.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    
    ' Show component image.
    ShowImage rs.Fields("Name")
End Sub

' Shows the component image.
Private Sub ShowImage(strName As String)
    Dim strImage As String
    
    ' Get component image.
    strImage = GetComponentImagePath(strName, cmbPackage.Text)
    
    ' Set the component image.
    If strImage <> vbNullString Then
        Dim picBitmap As Picture
        On Error GoTo PictureError
        Set picBitmap = LoadPicture(strImage)
        
        picImage.AutoRedraw = True
        picImage.PaintPicture picBitmap, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight
        Set picImage.Picture = picImage.Image
    Else
        Set picImage.Picture = Nothing
    End If
    
    ' Handle image setting errors.
    Exit Sub
PictureError:
    Set picImage.Picture = Nothing
    MsgBox "An error occured while trying to load the image for this component.", _
        vbOKOnly + vbCritical, "Image Loading Error"
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ' Reset variables.
    m_lngID = -1
End Sub
