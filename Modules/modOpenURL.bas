Attribute VB_Name = "modOpenURL"
''' modOpenURL
''' Opens a URL with the default browser.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public Function OpenURL(strURL As String) As Long
    OpenURL = ShellExecute(0, "open", strURL, 0, 0, 1)
End Function
