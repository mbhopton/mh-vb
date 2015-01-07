Attribute VB_Name = "FaceID"
Sub ShowFaceIDs2()
    Dim NewToolbar As CommandBar
    Dim TopPos As Long, LeftPos As Long
    Dim i As Long, NumPics As Long

'- - - - - Change These - - - - -
    Const ID_START As Long = 1
    Const ID_END As Long = 2000
'- - - - - - - - - - - - - - - - - - - -

'   Delete existing TempFaceIds toolbar if it exists
    On Error Resume Next
    Application.CommandBars("TempFaceIds").Delete
    On Error GoTo 0

'   Clear the sheet
    ActiveSheet.Pictures.Delete
    Application.ScreenUpdating = False
    
'   Add an empty toolbar
    Set NewToolbar = Application.CommandBars.Add(name:="TempFaceIds")

'   Starting positions
    TopPos = 5
    LeftPos = 5
    NumPics = 0
    
    For i = ID_START To ID_END
        Application.StatusBar = "Doing icon " & i
        On Error Resume Next
        NewToolbar.Controls(1).Delete
        With NewToolbar.Controls.Add(Type:=msoControlButton)
            .FaceID = i
            .CopyFace
        End With
        On Error GoTo 0
        
        NumPics = NumPics + 1
        ActiveSheet.Paste
        With ActiveSheet.Shapes(NumPics)
            .Top = TopPos
            .Left = LeftPos
            .name = "FaceID " & i
            .PictureFormat.TransparentBackground = True
            .PictureFormat.TransparencyColor = RGB(224, 223, 227)
        End With
        
'       Update top and left positions for the next one
        LeftPos = LeftPos + 16
        If NumPics Mod 50 = 0 Then
            TopPos = TopPos + 16
            LeftPos = 5
        End If
    Next i
    ActiveWindow.RangeSelection.Select
    Application.CommandBars("TempFaceIds").Delete
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
