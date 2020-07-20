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
