Attribute VB_Name = "modFarnell"
''' modFarnell
''' Farnell Portugal order parser module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Order file columns.
Private Enum Columns
    colOrderNumber
    colOrderConfirmationNumber
    colDeliveryETA
    colOrderStatus
    colTrackingCode
    colOrderDate
    colCurrency
    colTotal
    colShippingCost
    colImportTax
    colTaxes
    colOrderTotal
    colVouchers
    colOrigin
    colOrderCode
    colCustomPartNumber
    colLineNote
    colDescription
    colManufacturer
    colMfgPartNumber
    colQuantity
    colUnitPrice
    colItemTotalPrice
End Enum

' Parse the Farnell order CSV file.
Public Sub ParseFarnellOrder(strPath As String)
    Dim astrOrder() As String
    Dim lngRows As Long
    Dim lngCols As Long
    Dim hndFile As Integer
    Dim strContents As String

    ' Read the entire file into a string.
    hndFile = FreeFile()
    Open strPath For Input As #hndFile
        strContents = Input(LOF(1), #hndFile)
    Close #hndFile

    ' Parse the CSV file and initialize the components array.
    ParseCSV strContents, astrOrder, lngCols, lngRows
    InitializeComponentsArray (lngRows - 1)
    
    ' Populate components array.
    Dim idxRow As Long
    For idxRow = 1 To (lngRows - 1)
        ' Skip empty rows.
        If astrOrder(idxRow * lngCols + colQuantity) <> "" Then
            ' Add component to the array.
            AddComponent astrOrder(idxRow * lngCols + colMfgPartNumber), _
                         astrOrder(idxRow * lngCols + colDescription), _
                         astrOrder(idxRow * lngCols + colDescription), _
                         CLng(astrOrder(idxRow * lngCols + colQuantity))
        End If
    Next idxRow
    
    ' Check how all the components were added.
    Dim cmpComponent As Component
    For idxRow = 0 To LastComponentIndex
        Set cmpComponent = GetComponent(idxRow)
        Debug.Print idxRow, cmpComponent.Name
    Next idxRow
End Sub
