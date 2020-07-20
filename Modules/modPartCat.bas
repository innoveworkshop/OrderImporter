Attribute VB_Name = "modPartCat"
''' modPartCat
''' A PartCat helper module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Components array.
Private m_arrComponents() As Component
Private m_idxLastComponent As Long

' Initializes the components array.
Public Sub InitializeComponentsArray(lngSize As Long)
    ReDim m_arrComponents(lngSize)
    m_idxLastComponent = -1
End Sub

' Adds a component to the array.
Public Sub AddComponent(strName As String, strNotes As String, strProperties As String, _
                        lngQuantity As Long)
    ' Increment the last component index and instantiate a new component.
    m_idxLastComponent = m_idxLastComponent + 1
    Set m_arrComponents(m_idxLastComponent) = New Component
    
    ' Set the component attributes.
    With m_arrComponents(m_idxLastComponent)
        .Name = strName
        .Notes = strNotes
        .Properties = strProperties
        .Quantity = lngQuantity
    End With
End Sub

' Gets a component from the components array.
Public Function GetComponent(lngIndex As Long) As Component
    Set GetComponent = m_arrComponents(lngIndex)
End Function

' Gets the number of components in the array.
Public Function LastComponentIndex() As Long
    LastComponentIndex = m_idxLastComponent
End Function

' Gets the component directory from a PartCat workspace file path.
Public Function GetComponentsDir(strWorkspaceFilePath As String)
    Dim strPath As String
    
    ' Go to the parent path and then append the components directory.
    strPath = Left(strWorkspaceFilePath, InStrRev(strWorkspaceFilePath, "\"))
    strPath = strPath + COMPONENTS_PATH + "\"
    
    GetComponentsDir = strPath
End Function
