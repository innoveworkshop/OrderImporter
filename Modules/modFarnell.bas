Attribute VB_Name = "modFarnell"
''' modFarnell
''' Farnell Portugal order parser module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

Private CSVParser As ICSVParser

' Parse the Farnell order CSV file.
Public Sub ParseFarnellOrder(strPath As String)
    Dim strCSV() As String
    Dim lngRows As Long
    Dim lngCols As Long
    Dim hndFile As Integer
    Dim strContents As String

    ' Read the entire file into a string.
    hndFile = FreeFile()
    Open strPath For Input As #hndFile
        strContents = Input(LOF(1), #hndFile)
    Close #hndFile

    ' Actually parse the CSV file.
    ParseCSV strContents, strCSV, lngCols, lngRows

    ' Debug output.
    Dim idx As Long
    For idx = 0 To UBound(strCSV)
        Debug.Print "Row " & (idx \ lngCols), "Column " & (idx Mod lngCols), _
            "Data: " & strCSV(idx)
    Next idx
End Sub
