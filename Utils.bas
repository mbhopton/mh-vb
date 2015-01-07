Attribute VB_Name = "Utils"
Option Explicit

' Function to change date to the first day in the month
Public Function SOM(d As Date) As Date
    SOM = DateSerial(Year(d), Month(d), 1)
End Function

' Function to change date to the last day in the month
Public Function EOM(d As Date) As Date
    EOM = EoMonth(d, 0)
End Function

' Function to change to end of month n months in the future
Public Function EoMonth(d As Date, n As Integer) As Date
    EoMonth = DateSerial(Year(d), Month(d) + n + 1, 1) - 1
End Function

' Function returns true if the named workbook is open
Function WorkbookOpen(sName As String) As Boolean
    WorkbookOpen = False
    On Error GoTo WorkbookNotOpen
    If Len(Workbooks(sName).name) > 0 Then
        WorkbookOpen = True
        Exit Function
    End If
WorkbookNotOpen:
End Function

' Function to open a workbook unless it is already open
Function OpenWorkbook(sPath As String, sName As String) As Workbook
    On Error GoTo NoFile
    If WorkbookOpen(sName) Then
        Set OpenWorkbook = Workbooks(sName)
    Else
        Set OpenWorkbook = Workbooks.Open(Filename:=sPath & sName, UpdateLinks:=0)
    End If
    Exit Function
NoFile:
    MsgBox "File " & sPath & sName & " not found"
    Set OpenWorkbook = Nothing
    Exit Function
End Function

Function CharCount(text As String, ch As String) As Integer
    Dim i As Integer
    
    CharCount = 0
    For i = 1 To Len(text)
        If Mid(text, i, Len(ch)) = ch Then CharCount = CharCount + 1
    Next
End Function
