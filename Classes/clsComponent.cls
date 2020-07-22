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

' Initializes the properties array.
Public Sub InitializeProperties(intSize As Integer)
    ReDim m_astrProperties(intSize)
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

' Name getter.
Public Property Get Name() As String
    Name = m_strName
End Property

' Name setter.
Public Property Let Name(strName As String)
    m_strName = strName
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

' Property getter.
Public Property Get Property(intIndex As Integer) As String
    Property = m_astrProperties(intIndex)
End Property

' Property setter.
Public Property Let Property(intIndex As Integer, strProperty As String)
    m_astrProperties(intIndex) = strProperty
End Property

' Properties getter.
Public Property Get Properties() As String()
    Properties = m_astrProperties
End Property
