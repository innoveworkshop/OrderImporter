Attribute VB_Name = "modPartCat"
''' modPartCat
''' A PartCat helper module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Gets the component directory from a PartCat workspace file path.
Public Function GetComponentsDir(strWorkspaceFilePath As String)
    Dim strPath As String
    
    ' Go to the parent path and then append the components directory.
    strPath = Left(strWorkspaceFilePath, InStrRev(strWorkspaceFilePath, "\"))
    strPath = strPath + COMPONENTS_PATH + "\"
    
    GetComponentsDir = strPath
End Function
