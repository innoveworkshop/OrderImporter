VERSION 5.00
Begin VB.Form dlgProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Component Property"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbName 
      Height          =   315
      ItemData        =   "dlgProperty.frx":0000
      Left            =   120
      List            =   "dlgProperty.frx":000D
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "dlgProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' dlgProperty
''' A property editor dialog.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' User wants to cancel it.
Private Sub CancelButton_Click()
    Unload Me
End Sub
