Attribute VB_Name = "Module11"
Sub c()
Attribute c.VB_Description = "Unmerge and Auto-fill"
Attribute c.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' checkcurrent Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'
    Dim Cindex As Integer
    Cindex = 1
    ActiveSheet.Cells(1, Cindex).Select
    Do While Not (IsEmpty(ActiveCell))
        Do While Not (IsEmpty(ActiveCell))
            Dim cellTitle As String
            cellTitle = CStr(ActiveCell.Value)
            Dim rCurr As Range
            Set rCurr = Selection

            Selection.Offset(1, Ê0).Select
            rCurr.UnMerge
            rCurr.FormulaR1C1 = cellTitle
        Loop
        Cindex = Cindex + 1
        ActiveSheet.Cells(1, Cindex).Select
    Loop
End Sub
