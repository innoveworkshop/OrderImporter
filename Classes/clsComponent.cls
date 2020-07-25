VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Component"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' clsComponent
''' Abstraction of a PartCat component.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Properties.
Private m_strName As String
Private m_lngQuantity As Long
Private m_strNotes As String
Private m_strDatasheet As String
Private m_astrProperties() As String
Private m_strSearchCode As String
Private m_blnExported As Boolean

' Initializes the properties array.
Public Sub InitializeProperties(intSize As Integer)
    ReDim m_astrProperties(intSize)
End Sub

' Export the component to a root path where the folder will be placed in.
Public Sub Export(strPath As String)
    Dim strFolder As String
    Dim hndFile As Integer
    
    ' Build folder path and check if it already exists.
    strFolder = strPath & Name & "\"
    If Dir(strFolder, vbDirectory) <> "" Then
        MsgBox "Component folder already exists.", vbOKOnly + vbCritical, _
            "Component Already Exists"
        Exit Sub
    End If
    
    ' Create the component folder.
    MkDir strFolder
    
    ' Write manifest.
    hndFile = FreeFile
    Open strFolder & MANIFEST_FILE For Output As #hndFile
        Dim intIndex As Integer
        For intIndex = 0 To UBound(m_astrProperties)
            Print #hndFile, ManifestLine(m_astrProperties(intIndex))
        Next intIndex
    Close #hndFile
    
    ' Write quantity.
    hndFile = FreeFile
    Open strFolder & QUANTITY_FILE For Output As #hndFile
        Print #hndFile, CStr(Quantity)
    Close #hndFile
    
    ' Write notes.
    hndFile = FreeFile
    Open strFolder & NOTES_FILE For Output As #hndFile
        Print #hndFile, Notes
    Close #hndFile
    
    ' Download datasheet.
    If m_strDatasheet <> "" Then
        If Not DownloadFile(m_strDatasheet, strFolder & DATASHEET_FILE) Then
            MsgBox "Failed to download the component datasheet.", _
                vbOKOnly + vbExclamation, "Download Failed"
            
            Exported = False
            Exit Sub
        End If
    End If
    
    Exported = True
End Sub

' Adds a property to the array.
Public Sub AddProperty(strKey As String, strValue As String)
    If UBound(m_astrProperties) > 0 Then
        ReDim Preserve m_astrProperties(UBound(m_astrProperties) + 1)
    End If
    
    m_astrProperties(UBound(m_astrProperties)) = strKey & ": " & strValue
End Sub

' Edit a property from the array.
Public Sub EditProperty(strKey As String, strValue As String)
    Dim intIndex As Integer
    Dim strProperty As String
    
    ' Go through the properties looking for the specified key.
    For intIndex = 0 To UBound(m_astrProperties)
        strProperty = m_astrProperties(intIndex)
        
        If InStr(strProperty, strKey & ":") Then
            m_astrProperties(intIndex) = strKey & ": " & strValue
            Exit Sub
        End If
    Next intIndex
    
    ' Key wasn't found.
    MsgBox "The property key you want to edit wasn't found.", _
        vbOKOnly + vbCritical, "Property Key Not Found"
End Sub

' Deletes a property from the array.
Public Sub DeleteProperty(strKey As String)
    Dim astrBuffer() As String
    Dim intIndex As Integer
    Dim intOffset As Integer
    
    ' Check if we are deleting the last available property.
    If UBound(m_astrProperties) = 0 Then
        ReDim m_astrProperties(0)
        Exit Sub
    End If
    
    ' Populate the buffer array.
    intOffset = 0
    ReDim astrBuffer(UBound(m_astrProperties) - 1)
    For intIndex = 0 To UBound(m_astrProperties)
        ' Check if we should delete the property.
        If InStr(m_astrProperties(intIndex), strKey & ":") Then
            ' Add offset to the index.
            intOffset = -1
        Else
            If intIndex = 0 Then
                intOffset = 0
            End If
            
            ' Copy a non-deleted property
            astrBuffer(intIndex + intOffset) = m_astrProperties(intIndex)
        End If
    Next intIndex
    
    ' Copy the buffer to the actual properties array.
    ReDim m_astrProperties(UBound(astrBuffer))
    For intIndex = 0 To UBound(m_astrProperties)
        m_astrProperties(intIndex) = astrBuffer(intIndex)
    Next intIndex
End Sub

' Removes invalid properties from the properties array.
Public Sub RemoveInvalidProperties()
    Dim astrTemp() As String
    Dim intCountValid As Integer
    Dim intIndex As Integer
    
    ' Count valid properties.
    intCountValid = 0
    For intIndex = 0 To UBound(m_astrProperties)
        If IsPropertyLineValid(m_astrProperties(intIndex)) Then
            intCountValid = intCountValid + 1
        End If
    Next intIndex
    
    ' Go through adding valid items to the array.
    astrTemp = m_astrProperties
    ReDim m_astrProperties(intCountValid - 1)
    Dim intNewIndex As Integer
    intNewIndex = 0
    For intIndex = 0 To UBound(astrTemp)
        If IsPropertyLineValid(astrTemp(intIndex)) Then
            m_astrProperties(intNewIndex) = astrTemp(intIndex)
            intNewIndex = intNewIndex + 1
        End If
    Next intIndex
End Sub

' Makes a manifest file line from a property in the array.
Private Function ManifestLine(strProperty As String) As String
    Dim astrProperty() As String
    
    ' Split the key and value.
    astrProperty = Split(strProperty, ": ")
    
    ' Append the "Value-" to needed properties.
    If (astrProperty(0) <> "Category") And (astrProperty(0) <> "Sub-Category") And _
            (astrProperty(0) <> "Package") Then
        astrProperty(0) = VALUE_PROPERTY & astrProperty(0)
    End If
    
    ' Substitute spaces with dashes.
    astrProperty(0) = Replace(astrProperty(0), " ", "-")
    
    ManifestLine = astrProperty(0) & ": " & astrProperty(1)
End Function

' Checks if a property line is valid.
Private Function IsPropertyLineValid(strLine As String) As Boolean
    If InStr(strLine, ":") <> 0 Then
        If Right(strLine, 1) <> "-" Then
            IsPropertyLineValid = True
            Exit Function
        End If
    End If
    
    IsPropertyLineValid = False
End Function

' Removes invalid file characters from a string, making it filename-safe.
Public Static Function RemoveInvalidFileCharacters(strName As String) As String
    Dim strBuffer As String
    strBuffer = strName
    
    ' Remove invalid characters.
    strBuffer = Replace(strBuffer, ".", "")
    strBuffer = Replace(strBuffer, "<", "")
    strBuffer = Replace(strBuffer, ">", "")
    strBuffer = Replace(strBuffer, ":", "")
    strBuffer = Replace(strBuffer, """", "")
    strBuffer = Replace(strBuffer, "/", "")
    strBuffer = Replace(strBuffer, "\", "")
    strBuffer = Replace(strBuffer, "|", "")
    strBuffer = Replace(strBuffer, "?", "")
    strBuffer = Replace(strBuffer, "*", "")
    
    ' Trim whitespace and return.
    RemoveInvalidFileCharacters = Trim(strBuffer)
End Function

' Name getter.
Public Property Get Name() As String
    Name = m_strName
End Property

' Name setter.
Public Property Let Name(strName As String)
    m_strName = RemoveInvalidFileCharacters(strName)
End Property

' Quantity getter.
Public Property Get Quantity() As Long
    Quantity = m_lngQuantity
End Property

' Quantity setter.
Public Property Let Quantity(lngQuantity As Long)
    m_lngQuantity = lngQuantity
End Property

' Notes getter.
Public Property Get Notes() As String
    Notes = m_strNotes
End Property

' Notes setter.
Public Property Let Notes(strNotes As String)
    m_strNotes = strNotes
End Property

' Datasheet getter.
Public Property Get Datasheet() As String
    Datasheet = m_strDatasheet
End Property

' Datasheet setter.
Public Property Let Datasheet(strDatasheet As String)
    m_strDatasheet = strDatasheet
End Property

' Search code getter.
Public Property Get SearchCode() As String
    SearchCode = m_strSearchCode
End Property

' Search code setter.
Public Property Let SearchCode(strSearchCode As String)
    m_strSearchCode = strSearchCode
End Property

' Exported getter.
Public Property Get Exported() As Boolean
    Exported = m_blnExported
End Property

' Exported setter.
Public Property Let Exported(blnExported As Boolean)
    m_blnExported = blnExported
End Property

' Property getter.
Public Property Get Property(intIndex As Integer) As String
    Property = m_astrProperties(intIndex)
End Property

' Property setter.
Public Property Let Property(intIndex As Integer, strProperty As String)
    Dim astrProperty() As String
    astrProperty = Split(strProperty, ": ")
    
    ' Looks like this will just be another invalid property...
    If UBound(astrProperty) < 1 Then
        m_astrProperties(intIndex) = strProperty
        Exit Property
    End If
    
    ' Remove invalid characters from property key.
    astrProperty(0) = Replace(astrProperty(0), "  ", " ")
    astrProperty(0) = Replace(astrProperty(0), "'", "")
    astrProperty(0) = Replace(astrProperty(0), "(", "")
    astrProperty(0) = Replace(astrProperty(0), ")", "")
    astrProperty(0) = Replace(astrProperty(0), "@", "")
    astrProperty(0) = Replace(astrProperty(0), "%", "")
    astrProperty(0) = Replace(astrProperty(0), ".", "")
    
    m_astrProperties(intIndex) = astrProperty(0) & ": " & astrProperty(1)
End Property

' Properties getter.
Public Property Get Properties() As String()
    Properties = m_astrProperties
End Property
