VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "String builder class"
Option Explicit

'==============================================================
' Written: 	M Hopton
' Purpose:	StringBuilder class
' Saved: 	8 Jan 2015
'==============================================================
'	The string buffer size is handled automatically in all cases
'
'	set sb = new StringBuilder
'	sb.Length 		r/w		get or set string length
'	sb.Capacity		r/w		get or set buffer length, will not truncate the string
'	sb.Text			r/w		get or set text
'	sb.SubString	r		get a substring
'	sb.Chars		r/w		get or set a single character
'	sb.LineDelim	r/w		get or set the line delimiter (default vbNewLine)
'	sb.Clear		a		clear text and reset buffer
'	sb.AppendLine	a		add a newline, with optional text
'	sb.Append		a		add text to end of buffer
'	sb.Insert		a		insert text
'	sb.Remove		a		remove text
'	sb.Replace		a		replace text
'	sb.Find			a		find text
'	sb.FindRev		a		find text backwards
'	sb.DebugFlag	r/w		enable debugging to VBA immediate window

Private Const initialLength As Long = 64

Private totalLength As Long     ' Length of the buffer
Private curLength As Long       ' Length of the string value within the buffer
Private buffer As String        ' The buffer
Private delim As String         ' Line delimiter
Private dbg As Boolean          ' Debug flag

Private Sub Class_Initialize()
    Me.Clear
    dbg = False
    delim = vbNewLine
End Sub

Public Property Get DebugFlag() As Boolean
    DebugFlag = dbg
End Property

Public Property Let DebugFlag(ByVal vNewValue As Boolean)
    dbg = vNewValue
End Property

'Get length
Public Property Get Length() As Long
    Length = curLength
End Property

'Set length which will truncate string or extend buffer
Public Property Let Length(newLength As Long)
    If newLength < 0 Then
        Err.Raise vbObjectError + 513, "StringBuilder", "Index out of range"
        Exit Property
    ElseIf newLength < curLength Then
        Dim clrLen As Long
        clrLen = curLength - newLength
        Mid(buffer, newLength + 1, clrLen) = Space(clrLen)
        curLength = newLength
        If dbg Then debugP
    ElseIf newLength > totalLength Then
        totalLength = newLength
        Resize
    End If
End Property

'Get capacity
Public Property Get Capacity() As Long
    Capacity = totalLength
End Property

'Set capacity which will not change the current string
Public Property Let Capacity(newLength As Long)
    If newLength > curLength Then
        totalLength = newLength
        Resize
    End If
End Property

'Get the contents as a string
Public Property Get Text() As String
    Text = Left(buffer, curLength)
End Property

'Set the string
Public Property Let Text(newStr As String)
    Me.Clear
    Me.Append newStr
End Property

'Get a substring
Public Property Get SubString(start As Long, strLength As Long) As String
    SubString = Mid(buffer, start, strLength)
End Property

'Get a single character
Public Property Get Chars(ByVal index As Long) As String
    Chars = Mid(buffer, index, 1)
End Property

'Set a single character
Public Property Let Chars(ByVal index As Long, theChar As String)
    Mid(buffer, index, 1) = Left(theChar, 1)
End Property

'Get the line delimiter
Public Property Get LineDelim() As String
    LineDelim = delim
End Property

'Set the line delimiter
Public Property Let LineDelim(ByVal vNewValue As String)
    delim = vNewValue
End Property

'Clear the object
Public Sub Clear()
    totalLength = initialLength
    buffer = Space(totalLength)
    curLength = 0
End Sub

'Append a new line with optional text
Public Sub AppendLine(Optional ByVal text As String)
    If Not IsMissing(text) Then
        Me.Append text & Me.LineDelim
    Else
        Me.Append Me.LineDelim
    End If
End Sub

'Append text to the end of the buffer
Public Sub Append(text As String)
    Dim incLen As Long ' the length that the value will be increased by
    Dim newLen As Long ' the length of the value after being appended
    
    incLen = Len(text)
    newLen = curLength + incLen
    
    If newLen > totalLength Then
        ' buffer does not have enough room so double buffer size until its big enough
        While totalLength < newLen
            totalLength = totalLength + totalLength
        Wend
        Resize
    End If
    Mid(buffer, curLength + 1, incLen) = text
    curLength = newLen
    If dbg Then debugP
End Sub

'Insert text into the buffer which moves characters to make room
Public Sub Insert(index As Long, text As String)
    Dim moveStr As String   ' the string moved

    If index > curLength Or index <= 0 Then
        Err.Raise vbObjectError + 513, "StringBuilder", "Index out of range"
        Exit Sub
    End If
    moveStr = Me.SubString(index, curLength - index + 1)
    curLength = index - 1
    Me.Append text
    Me.Append moveStr
End Sub

'Remove text from buffer
Public Sub Remove(index As Long, removeLen As Long)
    Dim trail As Long       ' start index after the removed string
    Dim moveStr As String   ' the string moved

    trail = index + removeLen
    If trail > curLength + 1 Or index <= 0 Or removeLen <= 0 Then
        Err.Raise vbObjectError + 513, "StringBuilder", "Index out of range"
        Exit Sub
    End If
    moveStr = Me.SubString(trail, curLength - trail + 1)
    curLength = index - 1
    Me.Append moveStr
    Mid(buffer, curLength + 1, removeLen) = Space(removeLen)
End Sub

'Replace text in the buffer
Public Sub Replace(findStr As String, newStr As String)
    Me.text = VBA.Replace(Me.Text, findStr, newStr)
End Sub

'Find text in the buffer
Public Function Find(findStr As String) As Long
    Find = VBA.InStr(buffer, findStr)
End Function

'Reverse find text in the buffer
Public Function FindRev(findStr As String) As Long
    FindRev = VBA.InStrRev(Me.text, findStr)
End Function

Private Sub Resize()
    ' change the buffer size to totalLength which is assumed greater than curLength
    buffer = Left(buffer, curLength) & Space(totalLength - curLength)
End Sub

Private Sub debugP(Optional s As String)
    If IsMissing(s) Then
        Debug.Print curLength
        Debug.Print ":" & Me.Text & ":"
    Else
        Debug.Print curLength & s
        Debug.Print ":" & Me.Text & ":"
    End If
End Sub
