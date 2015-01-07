Attribute VB_Name = "URL"
Option Explicit

'==============================================================
' Written: 	M Hopton
' Purpose:	Function to decode URL
' Saved: 	8 Jan 2015
'==============================================================

' Function to decode URL
'	it replaces:
'		+ with space
'		%uXXXX with equivalent unicode character
'		%XX with equivalent hex character
'	uses StringBuilder class
Function urlDecode(sEncodedURL As String) As String

On Error GoTo Catch

Dim iLoop   As Integer
Dim iLen    As Integer
Dim sBuild  As StringBuilder

' Loop through each char
iLoop = 1
Set sBuild = New StringBuilder
Do Until iLoop > Len(sEncodedURL)
    Select Case Mid(sEncodedURL, iLoop, 1)
        Case "+"
            sBuild.Append " "
        Case "%"
            If Mid(sEncodedURL, iLoop + 1, 1) = "u" Then
                ' convert 4 chars HEX to decimal
                sBuild.Append Chr(CDec("&H" & Mid(sEncodedURL, iLoop + 2, 4)))
                iLoop = iLoop + 5
            Else
                ' convert 2 chars HEX to decimal
                sBuild.Append Chr(CDec("&H" & Mid(sEncodedURL, iLoop + 1, 2)))
                iLoop = iLoop + 2
            End If
        Case Else
            sBuild.Append Mid(sEncodedURL, iLoop, 1)
    End Select
    iLoop = iLoop + 1
Loop
urlDecode = sBuild.text

Finally:
    Exit Function
Catch:
    urlDecode = ""
    Resume Finally
End Function

