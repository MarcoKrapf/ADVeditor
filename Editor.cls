VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Advertisement editor V1.1 (15 Oct 2017)
'Written by Marco Krapf for GaloppSim ad design
'excel@marco-krapf.de
'https://marco-krapf.de/galoppsim

Dim CLICK As Boolean
Dim currentColor As Long
Dim rng As Range
Dim rngM As Range
Dim meter As Integer
Dim fee As Integer
Dim remember() As Variant

Private Sub START()
    meter = Range("GU5").Value
    Select Case meter
        Case 50
            Set rngM = Range("B2:AY12")
        Case 100
            Set rngM = Range("B2:CW12")
        Case 150
            Set rngM = Range("B2:EU12")
        Case 200
            Set rngM = Range("B2:GS12")
    End Select
    currentColor = PAINT.BackColor
End Sub

Private Sub expADV_Click()
    Dim EXP As Variant
    Dim i As Integer, j As Integer
    EXP = Application.GetSaveAsFilename(FileFilter:="GaloppSim Advertising File (*.gsadv), *.gsadv")
    If EXP <> False Then
        Open EXP For Output As #99
            Print #99, meter
            For i = 2 To meter + 1
                For j = 2 To 12
                    Print #99, Cells(j, i).Interior.Color
                Next j
            Next i
        Close #99
    End If
End Sub

Private Sub impADV_Click()
    Dim IMPO As Variant
    Dim advertising() As String
    Dim m As String
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    
    If MsgBox("Import data and overwrite cells?", vbExclamation + vbYesNo) = vbYes Then
        IMPO = Application.GetOpenFilename(FileFilter:="GaloppSim Adv File (*.gsadv), *.gsadv")
        If IMPO <> False Then
            ActiveSheet.Unprotect
            With Range("B2:GS12")
                .Interior.Color = xlNone
                .ClearContents
            End With
            ActiveSheet.Protect

            Open IMPO For Input As #99
                Line Input #99, m
            Close #99
            
            meter = CInt(m)
            ReDim advertising(0 To meter * 11) As String
            
            Open IMPO For Input As #99
                For i = 0 To meter * 11
                    Line Input #99, advertising(i)
                Next i
            Close #99
            
            Select Case meter
                Case 50
                    Call setm50
                Case 100
                    Call setm100
                Case 150
                    Call setm150
                Case 200
                    Call setm200
            End Select
        
            ActiveSheet.Unprotect
            Dim line As Integer
            line = 1
            For i = 2 To meter + 1
                For j = 2 To 12
                    Cells(j, i).Interior.Color = CLng(advertising(line))
                    line = line + 1
                Next j
            Next i
            ActiveSheet.Protect
            Call Sync
            Call FeeCalc
            Range("B2").Select
        End If
    End If
End Sub

Private Sub gotoRIGHT_Click()
    ActiveWindow.ScrollColumn = 202
End Sub

Private Sub gotoLEFT_Click()
    ActiveWindow.ScrollColumn = 1
End Sub

Private Sub RememberSelection()
    ReDim remember(Selection.Count - 1, 1)
    Dim i As Integer
    For Each rng In Selection
        remember(i, 0) = rng.AddressLocal
        remember(i, 1) = rng.Interior.Color
        i = i + 1
    Next rng
End Sub

Private Sub RememberAll()
    If rngM Is Nothing Then Call START
    
    ReDim remember(rngM.Count - 1, 1)
    Dim i As Integer
    For Each rng In rngM
        remember(i, 0) = rng.AddressLocal
        remember(i, 1) = rng.Interior.Color
        i = i + 1
    Next rng
End Sub

Private Sub BACKGROUND_ALL_Click()
    ActiveSheet.Unprotect

    Call RememberAll
    
    For Each rng In rngM
        If rng.Interior.Color = BACKGROUND_ALL.BackColor Then
            rng.Interior.Color = currentColor
        End If
    Next
    
    UNDO.Caption = "Undo (" & rngM.AddressLocal(False, False) & ")"
    
    Call Sync
    ActiveSheet.Protect
End Sub

Private Sub BACKGROUND_SEL_Click()
    ActiveSheet.Unprotect
    Call RememberSelection
    Call checkCLICK
    For Each rng In Selection
        If rng.Interior.Color = BACKGROUND_ALL.BackColor Then
            rng.Interior.Color = currentColor
        End If
    Next
    Call Sync
    ActiveSheet.Protect
End Sub

Private Sub CLEAR_Click()
    ActiveSheet.Unprotect
    If MsgBox("Clear all selected cells?", vbExclamation + vbYesNo) = vbYes Then
        
        Call RememberAll
        
        For Each rng In Selection
            rng.Interior.Color = xlNone
            rng.ClearContents
        Next
        UNDO.Caption = "Undo (" & Selection.AddressLocal(False, False) & ")"
        Call Sync
    End If
    ActiveSheet.Protect
End Sub

Private Sub CLEARALL_Click()
    ActiveSheet.Unprotect
    If MsgBox("Clear all visible cells?", vbExclamation + vbYesNo) = vbYes Then
        If rngM Is Nothing Then
            Call START
        End If
        For Each rng In rngM
            rng.Interior.Color = xlNone
            rng.ClearContents
        Next
        UNDO.Caption = "Undo (" & rngM.AddressLocal(False, False) & ")"
        Call Sync
    End If
    ActiveSheet.Protect
End Sub

Private Sub GRIDOFF_Click()
    ActiveSheet.Unprotect
    If rngM Is Nothing Then
        Call START
    End If
    With rngM
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    ActiveSheet.Protect
End Sub

Private Sub GRIDON_Click()
    ActiveSheet.Unprotect
    If rngM Is Nothing Then
        Call START
    End If
    With rngM
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    ActiveSheet.Protect
End Sub

Private Sub LOAD_Click()
    Dim LOAD As Variant
    Dim col(1 To 15) As String
    Dim i As Integer
    
    On Error Resume Next
    
    If MsgBox("Import colour scheme and overwrite current colours?", vbExclamation + vbYesNo) = vbYes Then
        LOAD = Application.GetOpenFilename(FileFilter:="GaloppSim Adv Colours (*.gsadvcol), *.gsadvcol")
        If LOAD <> False Then
            Open LOAD For Input As #99
                For i = 1 To 15
                    Line Input #99, col(i)
                Next i
            Close #99
            
            Col1.BackColor = col(1)
            Col2.BackColor = col(2)
            Col3.BackColor = col(3)
            Col4.BackColor = col(4)
            Col5.BackColor = col(5)
            Col6.BackColor = col(6)
            Col7.BackColor = col(7)
            Col8.BackColor = col(8)
            Col9.BackColor = col(9)
            Col10.BackColor = col(10)
            Col11.BackColor = col(11)
            Col12.BackColor = col(12)
            Col13.BackColor = col(13)
            Col14.BackColor = col(14)
            Col15.BackColor = col(15)
        End If
    End If
End Sub

Private Sub m50_Click()
    meter = 50
    Call setm50
End Sub

Private Sub setm50()
    ActiveSheet.Unprotect
    Set rngM = Range("B2:AY12")
    Call meterReset
    m50.BackColor = 8454143
    Range("AZ:GS").Columns.Hidden = True
    Range("IS14:OL24").Interior.Color = 12566463
    Call FeeCalc
    ActiveSheet.Protect
End Sub

Private Sub m100_Click()
    meter = 100
    Call setm100
    Call FeeCalc
    Call Sync
End Sub

Private Sub setm100()
    ActiveSheet.Unprotect
    Set rngM = Range("B2:CW12")
    Call meterReset
    m100.BackColor = 8454143
    Range("AZ:CW").Columns.Hidden = False
    Range("CX:GS").Columns.Hidden = True
    Range("KQ14:OL24").Interior.Color = 12566463
    ActiveSheet.Protect
End Sub

Private Sub m150_Click()
    meter = 150
    Call setm150
    Call FeeCalc
    Call Sync
End Sub

Private Sub setm150()
    ActiveSheet.Unprotect
    Set rngM = Range("B2:EU12")
    Call meterReset
    m150.BackColor = 8454143
    Range("AZ:EU").Columns.Hidden = False
    Range("EV:GS").Columns.Hidden = True
    Range("MO14:OL24").Interior.Color = 12566463
    ActiveSheet.Protect
End Sub

Private Sub m200_Click()
    meter = 200
    Call setm200
    Call FeeCalc
    Call Sync
End Sub

Private Sub setm200()
    ActiveSheet.Unprotect
    Set rngM = Range("B2:GS12")
    Call meterReset
    m200.BackColor = 8454143
    Range("AZ:GS").Columns.Hidden = False
    ActiveSheet.Protect
End Sub

Private Sub meterReset()
    m50.BackColor = -2147483633
    m100.BackColor = -2147483633
    m150.BackColor = -2147483633
    m200.BackColor = -2147483633
End Sub

Private Sub checkCLICK()
    If CLICK = False Then
        Call RememberSelection
        currentColor = PAINT.BackColor
        CLICK = True
    End If
End Sub

Private Sub PAINT_Click()
    ActiveSheet.Unprotect
    Call checkCLICK
    For Each rng In Selection
        rng.Interior.Color = currentColor
    Next
    Call Sync
    ActiveSheet.Protect
End Sub

Private Sub PICK_Click()
    currentColor = ActiveCell.Interior.Color
    Call colorPrev
End Sub

Private Sub colorPrev()
    currCol.BackColor = currentColor
    PAINT.BackColor = currentColor
    PAINT.Caption = "Fill the selected cells" & vbNewLine & _
        "with this colour" & vbNewLine & _
        "(Interior.Color " & currentColor & ")"
End Sub

Private Sub SpinButton1_Change()
    meter = Range("GU5").Value
    fee1.Caption = SpinButton1.Value
    Call FeeCalc
End Sub

Private Sub SpinButton2_Change()
    meter = Range("GU5").Value
    fee2.Caption = SpinButton2.Value
    Call FeeCalc
End Sub

Private Sub SpinButton3_Change()
    meter = Range("GU5").Value
    fee3.Caption = SpinButton3.Value
    Call FeeCalc
End Sub

Private Sub FeeCalc()
    ActiveSheet.Unprotect
    Range("GU5").Value = meter
    feeS.Caption = SpinButton1.Value * 20 * meter / 100 & " EUR"
    feeM.Caption = SpinButton2.Value * 10 * meter / 100 & " EUR"
    feeF.Caption = SpinButton3.Value * 15 * meter / 100 & " EUR"
    fee = SpinButton1.Value * 20 * meter / 100 + SpinButton2.Value * 10 * meter / 100 + SpinButton3.Value * 15 * meter / 100
    feeTotal.Caption = fee & " EUR"
    lengthTotal.Caption = (SpinButton1.Value + SpinButton2.Value + SpinButton3.Value) * meter & "m"
    ActiveSheet.Protect
End Sub

Private Sub SAVE_Click()
    Dim SAVE As Variant
    Dim i As Integer, j As Integer
    SAVE = Application.GetSaveAsFilename(FileFilter:="GaloppSim Advertising Colours (*.gsadvcol), *.gsadvcol")
    If SAVE <> False Then
        Open SAVE For Output As #99
            Print #99, Col1.BackColor
            Print #99, Col2.BackColor
            Print #99, Col3.BackColor
            Print #99, Col4.BackColor
            Print #99, Col5.BackColor
            Print #99, Col6.BackColor
            Print #99, Col7.BackColor
            Print #99, Col8.BackColor
            Print #99, Col9.BackColor
            Print #99, Col10.BackColor
            Print #99, Col11.BackColor
            Print #99, Col12.BackColor
            Print #99, Col13.BackColor
            Print #99, Col14.BackColor
            Print #99, Col15.BackColor
        Close #99
    End If
End Sub

Private Sub UNDO_Click()
    ActiveSheet.Unprotect
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To UBound(remember)
        Range(remember(i, 0)).Interior.Color = remember(i, 1)
    Next i
    
    On Error GoTo 0
    ActiveSheet.Protect
    
    Call Sync
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ActiveSheet.Unprotect
    Call RememberSelection

    If boxAutopaint.Value = True Then
        For Each rng In Selection
            rng.Interior.Color = PAINT.BackColor
        Next
    End If

    Call Sync

    PICK.Caption = "Pick colour from the selected cell (" & ActiveCell.AddressLocal(False, False) & ")"
    CLEAR.Caption = "Clear selected cells (" & Selection.AddressLocal(False, False) & ")"
    UNDO.Caption = "Undo (" & Selection.AddressLocal(False, False) & ")"
    BACKGROUND_ALL.BackColor = ActiveCell.Interior.Color
    BACKGROUND_SEL.BackColor = ActiveCell.Interior.Color
    colBefore.BackColor = ActiveCell.Interior.Color

    ActiveSheet.Protect
End Sub

Private Sub Sync()
    ActiveSheet.Unprotect
    If rngM Is Nothing Then
        Call START
    End If
    For Each rng In rngM
        rng.Offset(12, 201).Interior.Color = rng.Interior.Color
    Next
    ActiveSheet.Protect
End Sub

Function ColPick(curr As Long) As Long
    Dim returncode As Integer
    returncode = Application.Dialogs(xlDialogEditColor).Show(10)
    If returncode <> 0 Then
        ColPick = ThisWorkbook.Colors(10)
    Else
        ColPick = curr
    End If
End Function

Private Sub Col1_Click()
    currentColor = Col1.BackColor
    Call colorPrev
End Sub

Private Sub Col1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col1.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col2_Click()
    currentColor = Col2.BackColor
    Call colorPrev
End Sub

Private Sub Col2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col2.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col3_Click()
    currentColor = Col3.BackColor
    Call colorPrev
End Sub

Private Sub Col3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col3.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col4_Click()
    currentColor = Col4.BackColor
    Call colorPrev
End Sub

Private Sub Col4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col4.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col5_Click()
    currentColor = Col5.BackColor
    Call colorPrev
End Sub

Private Sub Col5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col5.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col6_Click()
    currentColor = Col6.BackColor
    Call colorPrev
End Sub

Private Sub Col6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col6.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col7_Click()
    currentColor = Col7.BackColor
    Call colorPrev
End Sub

Private Sub Col7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col7.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col8_Click()
    currentColor = Col8.BackColor
    Call colorPrev
End Sub

Private Sub Col8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col8.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col9_Click()
    currentColor = Col9.BackColor
    Call colorPrev
End Sub

Private Sub Col9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col9.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col10_Click()
    currentColor = Col10.BackColor
    Call colorPrev
End Sub

Private Sub Col10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col10.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col11_Click()
    currentColor = Col11.BackColor
    Call colorPrev
End Sub

Private Sub Col11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col11.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col12_Click()
    currentColor = Col12.BackColor
    Call colorPrev
End Sub

Private Sub Col12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col12.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col13_Click()
    currentColor = Col13.BackColor
    Call colorPrev
End Sub

Private Sub Col13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col13.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col14_Click()
    currentColor = Col14.BackColor
    Call colorPrev
End Sub

Private Sub Col14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col14.BackColor = currentColor
    Call colorPrev
End Sub

Private Sub Col15_Click()
    currentColor = Col15.BackColor
    Call colorPrev
End Sub

Private Sub Col15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    currentColor = ColPick(currentColor)
    Col15.BackColor = currentColor
    Call colorPrev
End Sub
