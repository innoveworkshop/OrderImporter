VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "PartCat Order Importer"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
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

' Exits the application.
Private Sub mniFileExit_Click()
    Unload Me
End Sub
