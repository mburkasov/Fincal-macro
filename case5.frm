VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sexy 
   Caption         =   "Sexy macro"
   ClientHeight    =   8610.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3450
   OleObjectBlob   =   "sexy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sexy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

 Dim s As Range
 Dim d As Range
 Dim p As Range
 Dim strike As Range
 Dim rf As Range
 Dim dt As Range
 Dim start As Range
 Dim rng As Range
 Dim n As Integer
 Dim t As Integer
 Dim style As Integer
 Dim form As String
 Dim up As String
 Dim down As String
 Dim confirm As Integer

 If Not IsEmpty(RefEdit1.Text) And Not IsEmpty(RefEdit2.Text) And Not IsEmpty(RefEdit3.Text) And Not IsEmpty(RefEdit4.Text) And Not IsEmpty(RefEdit5.Text) And Not IsEmpty(RefEdit6.Text) And Not IsEmpty(RefEdit7.Text) And Not IsEmpty(TextBox1.Value) Then

 Set s = Range(RefEdit1.Text)
 Set d = Range(RefEdit2.Text)
 Set p = Range(RefEdit3.Text)
 Set strike = Range(RefEdit4.Text)
 Set rf = Range(RefEdit5.Text)
 n = TextBox1.Value
 Set start = Range(RefEdit6.Text)
 Set dt = Range(RefEdit7.Text)

 If OptionButton1.Value = True Then style = 2
 If OptionButton2.Value = True Then style = 1

 If OptionButton3.Value = True Then t = 0
 If OptionButton4.Value = True Then t = 1

 If s.Count = 1 And d.Count = 1 And p.Count = 1 And strike.Count = 1 And rf.Count = 1 And n > 49 And start.Count = 2 And IsNumeric(n) And IsNumeric(t) And IsNumeric(s) Then
    MsgBox "OK"

    If start.Cells(1, 1).Row < n * style Then
        MsgBox "Not enough space"
    Else

    Set rng = Range(start.Address)
    For i = 1 To n
        Set rng = Range(rng.Resize(rng.Rows.Count + style * 2, 1).Offset(-style, 1).Address)
        If (Not IsEmpty(rng)) And confirm <> vbYes Then confirm = MsgBox("Clear the range at " & start.Address & " length " & n & " ?", vbYesNo, "Are you sure?")
        If confirm = vbYes Then rng.Clear
        If confirm = vbNo Then Exit Sub
    Next i

        If t = 0 Then
            form = "=MAX((R[-" & style & "]C[1]*" & p.Address(True, True, xlR1C1) & "+R[" _
                & style & "]C[1]*(1-" & p.Address(True, True, xlR1C1) & "))*EXP(-" & rf.Address(True, True, xlR1C1) & "*" _
                & dt.Address(True, True, xlR1C1) & ")," & strike.Address(True, True, xlR1C1) & "-R[-1]C)"
        Else
            form = "=(R[-" & style & "]C[1]*" & p.Address(True, True, xlR1C1) & "+R[" _
                & style & "]C[1]*(1-" & p.Address(True, True, xlR1C1) & "))*EXP(-" & rf.Address(True, True, xlR1C1) & "*" _
                & dt.Address(True, True, xlR1C1) & ")"
        End If

        up = "=R[" & style & "]C[-1]/" & d.Address(True, True, xlR1C1)
        down = "=R[" & -style & "]C[-1]*" & d.Address(True, True, xlR1C1)

        Set rng = Range(start.Address)
        rng.Cells(1, 1).Formula2R1C1 = down
        rng.Cells(2, 1).Formula2R1C1 = form

        For i = 1 To n

            rng.Copy
            rng.Offset(style, 1).PasteSpecial (xlPasteAll)
            Set rng = Range(rng.Resize(rng.Rows.Count + style * 2, 1).Offset(-style, 1).Address)
            start.Copy
            rng.Cells(1, 1).PasteSpecial (xlPasteFormats)
            rng.Cells(1, 1).Formula2R1C1 = up
            rng.Cells(2, 1).Formula2R1C1 = form

        Next i

        For i = 2 To rng.Rows.Count Step 2 * style
            rng.Cells(i, 1).Formula2R1C1 = "=MAX(0," & strike.Address(True, True, xlR1C1) & "-R[-1]C)"
        Next i

        start.Cells(1, 1).Formula = "=" & s.Address
        start.Cells(2, 1).Formula2R1C1 = form
        MsgBox "Tree planted"
    End If

 ElseIf n < 50 Then
     MsgBox "n too small"
 Else
     MsgBox "Something is wrong with input"
 End If
 End If
 If MsgBox("Finish execution?", vbYesNo, "Finished?") = vbYes Then Unload Me

End Sub

Private Sub CommandButton2_Click()

 Dim n As Integer
 Dim style As Integer
 Dim rng As Range

 n = TextBox1.Value
 If OptionButton1.Value = True Then style = 2
 If OptionButton2.Value = True Then style = 1
 Set rng = Range(RefEdit6.Text)

 If MsgBox("Delete the range at " & rng.Address & " length " & n & " ?", vbYesNo, "Are you sure?") = vbYes Then
    For i = 1 To n
        rng.Clear
        Set rng = Range(rng.Resize(rng.Rows.Count + style * 2, 1).Offset(-style, 1).Address)
    Next i
 Else: Exit Sub
 If MsgBox("Finish execution?", vbYesNo, "Finished?") = vbYes Then Unload Me
 End If
End Sub

Private Sub CommandButton3_Click()

    Dim i As Integer
    Dim off As Integer
    Dim rng As Range
    Dim s As String
    Dim t As Integer
    Dim sol As Integer

    If OptionButton5.Value = True Then t = 2
    If OptionButton6.Value = True Then t = 1
    If OptionButton7.Value = True Then sol = 2
    If OptionButton8.Value = True Then sol = 1


    s = ChrW(10) & ChrW(10) & ChrW(66) & ChrW(121) & ChrW(32) & ChrW(77) & ChrW(105) & ChrW(107) & ChrW(104) & ChrW(97) & ChrW(105)

    Range(RefEdit8.Text).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng = Selection

    If IsEmpty(rng.Cells(1, 1)) Then
        MsgBox "Chosen range is empty."
    ElseIf IsEmpty(rng.Offset(0, -1).Cells(1, 1)) Or IsEmpty(rng.Offset(0, -2).Cells(1, 1)) Or IsEmpty(rng.Offset(0, -3).Cells(1, 1)) Or IsEmpty(rng.Offset(0, -4).Cells(1, 1)) Then
        MsgBox "Something wrong with coefficients"
    ElseIf Not IsEmpty(rng.Offset(0, 1).Cells(1, 1)) Then
        MsgBox "Clear working column 2"
    ElseIf Not IsEmpty(rng.Offset(-1, 0).Cells(1, 1)) Then
            MsgBox "Not the first cell selected"
    Else

    rng.FormulaR1C1 = "=RC[-3]*R[1]C[1]+RC[-2]*RC[1]+RC[-1]*R[-1]C[1]-RC[-5]"
    SolverReset
    SolverAdd CellRef:=Range(rng.Address), Relation:=2, FormulaText:="=0"

    For i = rng.Count To 0 Step -1

        SolverOk SetCell:=rng.Cells(i, 1), MaxMinVal:=3, ValueOf:=0, ByChange:=rng.Offset(0, 1), Engine:=sol
        If i = rng.Count Then
            SolverSolve
        Else
            SolverSolve UserFinish:=True
        End If

        rng.Offset(0, t).Copy

        n = -rng.Count + i - 6
        rng.Offset(rowOffset:=0, columnOffset:=n).Select

        If Selection.HasFormula Then
             s = s & ChrW(108) & ChrW(32) & ChrW(66) & ChrW(117) & ChrW(114) & ChrW(107) & ChrW(97) & ChrW(115) & ChrW(111) & ChrW(118) & ChrW(32) & ChrW(169)
             Exit For
        End If

        Selection.PasteSpecial Paste:=xlPasteValues
        rng.Offset(0, 1).Clear

        rng.Cells(1, 1).Select
        rng.FormulaR1C1 = "=RC[-3]*R[1]C[1]+RC[-2]*RC[1]+RC[-1]*R[-1]C[1]-RC[" & CStr(n) & "]"

    Next i

    rng.FormulaR1C1 = "=RC[-3]*R[1]C[1]+RC[-2]*RC[1]+RC[-1]*R[-1]C[1]-RC[-5]"

    s = "All done!" & s
    MsgBox s

    End If
    If MsgBox("Finish execution?", vbYesNo, "Finished?") = vbYes Then Unload Me
End Sub

Private Sub CommandButton4_Click()

    Range(RefEdit8.Text).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng = Selection

    If MsgBox("Are you sure you want to reset the field", vbYesNo, "Reset?") = vbYes Then
        Range(rng.Offset(0, 1).Address).ClearContents
        For i = 1 To rng.Count
            rng.Offset(0, -i - 5).Select
            If Selection.HasFormula Then Exit For
            Selection.Clear

        Next i
    End If
    rng.Select
    rng.FormulaR1C1 = "=RC[-3]*R[1]C[1]+RC[-2]*RC[1]+RC[-1]*R[-1]C[1]-RC[-5]"
    If MsgBox("Finish execution?", vbYesNo, "Finished?") = vbYes Then Unload Me
End Sub

Private Sub OptionButton7_Click()

End Sub

Private Sub UserForm_Click()

End Sub
