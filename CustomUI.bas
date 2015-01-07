Attribute VB_Name = "CustomUI"
Option Explicit

'==============================================================
' Written: 	M Hopton
' Purpose:	Routines for Custom Excel Ribbon
' Saved: 	8 Jan 2015
'==============================================================

Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Public MyRibbon As IRibbonUI
Public showJobs As Boolean

Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set MyRibbon = ribbon
    ThisWorkbook.Sheets(1).Range("A1").Value = ObjPtr(ribbon)
End Sub

Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
    Dim objRibbon As Object
    CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing
End Function

Sub RefreshRibbon()
    If MyRibbon Is Nothing Then
        Set MyRibbon = GetRibbon(ThisWorkbook.Sheets(1).Range("A1").Value)
    End If
    MyRibbon.Invalidate
End Sub

Public Sub mhCtrl_getDescription(control As IRibbonControl, ByRef returnedValue)
    mhCtrl_getLabel control, returnedValue
End Sub

Public Sub mhCtrl_getSize(control As IRibbonControl, ByRef returnedValue)
    Select Case control.ID
        Case "mhButton4", "mhButton5"
            returnedValue = 0
        Case Else
            returnedValue = 1
    End Select
End Sub

Public Sub mhCtrl_getImage(control As IRibbonControl, ByRef returnedValue)
    Select Case control.ID
        Case "mhButton1"
            returnedValue = "HappyFace"
        Case "mhButton2"
            returnedValue = "HappyFace"
        Case "mhButton3"
            returnedValue = "StartAfterPrevious"
        Case Else
            returnedValue = "ListMacros"
    End Select
End Sub
Public Sub mhCtrl_getLabel(control As IRibbonControl, ByRef returnedValue)
    Select Case control.ID
        Case "mhButton1"
            returnedValue = "Combine Cells"
        Case "mhButton2"
            returnedValue = "Apply Structure"
        Case "mhButton3"
            returnedValue = "Allegiance"
        Case Else
            returnedValue = control.ID
    End Select
End Sub

Public Sub mhCtrl_getScreentip(control As IRibbonControl, ByRef returnedValue)
    Select Case control.ID
        Case "mhButton1"
            returnedValue = "Combine Cells"
        Case "mhButton2"
            returnedValue = "Apply Structure"
        Case "mhButton3"
            returnedValue = "Allegiance"
        Case Else
            returnedValue = ThisWorkbook.name
    End Select
End Sub

Sub mhCtrl_getEnabled(control As IRibbonControl, ByRef returnedValue)
    Select Case control.ID
        Case "mhJobs"
            returnedValue = showJobs
        Case "mhButton4", "mhButton5"
            returnedValue = False
        Case Else
            returnedValue = True
    End Select
End Sub

Sub mhCtrl_getVisible(control As IRibbonControl, ByRef returnedValue)
    Select Case control.ID
        Case "mhButton4" To "mhButton5"
            returnedValue = False
        Case Else
            returnedValue = True
    End Select
End Sub

Sub mhCtrl_onAction(control As IRibbonControl)
Dim wb As Workbook
    Select Case control.ID
        Case "mhTimeIt"
            ' open TimeIt and run SetTask macro
            Set wb = OpenWorkbook("M:\Personal Documents\", "TimeIt.xlsm")
            If Not wb Is Nothing Then
                Application.Run "'" & wb.name & "'!SetTask"
            End If
        Case "mhJobs"
            Jobs
        Case "mhRefresh"
            RefreshRibbon
        Case "mhFmtRAS"
            RASFmt
        Case "mhSATCOM"
            SATCOM
        Case "mhButton1"
            CombineCells
        Case "mhButton2"
            FindWBS
        Case "mhButton3"
            Set wb = OpenWorkbook("", "sw.xlsm")
            Application.Run "'" & wb.name & "'!ShowForm"
        Case "mhButton4"
            FilterGateway
        Case "mhColIndex"
            ColourIndex
        Case "mhFaceID"
            ShowFaceIDs2
        Case "mhNames", "mhEmployees", "mhGraph", "mhByName", "mhClean", "mhDefSort", "mhNameSort"
            ' RAS routines
            Set wb = OpenWorkbook("M:\RAS\", "VBA SATCOM.xlsm")
            Application.Run "'" & wb.name & "'!mhRAS_onAction", control
    End Select
End Sub

Public Sub mhDyn_getContent(control As IRibbonControl, ByRef returnedVal)
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.Append "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >"
    Select Case control.ID
        Case "mhRASDyn"
            sb.Append "<button id=""mhNames"" label=""Check Names"" onAction=""mhCtrl_onAction"" imageMso=""CheckNames"" />"
            sb.Append "<button id=""mhEmployees"" label=""Employees"" onAction=""mhCtrl_onAction"" imageMso=""DistributionListSelectMembers"" />"
            sb.Append "<button id=""mhGraph"" label=""Graph RAS"" onAction=""mhCtrl_onAction"" imageMso=""ChartAreaChart"" />"
            sb.Append "<button id=""mhByName"" label=""Sum by person"" onAction=""mhCtrl_onAction"" imageMso=""BusinessCardInsertMenu"" />"
            sb.Append "<button id=""mhClean"" label=""Clean RAS"" onAction=""mhCtrl_onAction"" imageMso=""OmsDelete"" />"
            sb.Append "<button id=""mhDefSort"" label=""Default sort"" onAction=""mhCtrl_onAction"" imageMso=""PivotChartSortByTotalMenu"" />"
            sb.Append "<button id=""mhNameSort"" label=""Sort by person"" onAction=""mhCtrl_onAction"" imageMso=""SortUp"" />"
            sb.Append "<menuSeparator id=""menusep1"" />"
            sb.Append "<button id=""mhFmtRAS"" label=""Format RAS"" onAction=""mhCtrl_onAction"" imageMso=""AccessFormDatasheet"" />"
            sb.Append "<button id=""mhSATCOM"" label=""Filter SATCOM"" onAction=""mhCtrl_onAction"" imageMso=""FilterBySelection"" />"
        Case "mhGenDyn"
            sb.Append "<button id=""mhColIndex"" label=""Colour Index"" onAction=""mhCtrl_onAction"" imageMso=""ThemeColorsGallery"" />"
            sb.Append "<button id=""mhFaceID"" label=""Button faces"" supertip=""Show command button faces"" onAction=""mhCtrl_onAction"" imageMso=""SadFace"" />"
            sb.Append "<menuSeparator id=""menusep4"" />"
            sb.Append "<button id=""mhRefresh"" label=""Refresh Ribbon"" onAction=""mhCtrl_onAction"" imageMso=""SignatureLineInsert"" />"
        Case Else
    End Select
    sb.Append "</menu>"
    returnedVal = sb.text
End Sub

