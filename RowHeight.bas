Attribute VB_Name = "RowHeight"
Option Explicit

'adjust the row height by iDir lines (up or down)
Sub AdjustHeight(iDir As Integer)
    Dim fH As Double, iFont As Integer, fStd As Double
    
    fStd = ActiveSheet.StandardHeight
    If fStd > 12.75 Then
        iFont = ActiveCell.Font.Size
        fStd = 1.275 * iFont
    End If
    fH = ActiveCell.RowHeight
    fH = fH + iDir * fStd
    fH = fStd * Round(fH / fStd)
    ActiveCell.RowHeight = fH
End Sub

'increase the row height by 1 line
Sub IncHeight()
    AdjustHeight 1
End Sub

'decrease the row height by 1 line
Sub DecHeight()
    AdjustHeight -1
End Sub

'adjust row height in selection to multiple of standard height
'   with some additional space
Sub SpaceH()
    Dim fH As Double, iFont As Integer, fStd As Double
    Dim r As Range
    
    fStd = ActiveSheet.StandardHeight
    If fStd > 12.75 Then
        iFont = ActiveCell.Font.Size
        fStd = 1.275 * iFont
    End If
    For Each r In Selection.Rows
        fH = r.RowHeight
        fH = fStd * Round(fH / fStd) + 5 * 0.75
        r.RowHeight = fH
    Next
End Sub

