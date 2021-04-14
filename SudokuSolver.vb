Dim free() As Integer
Sub Solve()

Dim oldfree() As Integer
Dim fr, fc, lr, lc, rsubbox, csubbox, rr, cc, ans, count, mark, mark2, mark3 As Integer
Dim revert() As Integer

fr = 2
fc = 2

lr = fr + 8
lc = fc + 8
mark2 = 0
mark3 = 0
initial:
ReDim free(fr To lr, fc To lc, 0 To 9) As Integer

For c = fc To lc
    For r = fr To lr
        If Cells(r, c) <> "" Then
            free(r, c, 0) = Cells(r, c)
            For p = 1 To 9
                free(r, c, p) = 0
            Next p
        Else
            For p = 1 To 9
                free(r, c, p) = p
            Next p
        End If
    Next r
Next c


wave1:
Dim mto, only As Integer
ReDim oldfree(fr To lr, fc To lc, 0 To 9) As Integer
mark = 0
'eliminate possibilities
For c = fc To lc
    For r = fr To lr
        For i = 1 To 9
            oldfree(r, c, i) = free(r, c, i)
        Next i
        If free(r, c, 0) = 0 Then
            
            For rr = fr To lr
                If free(rr, c, 0) <> 0 Then free(r, c, free(rr, c, 0)) = 0
            Next rr
            
            For i = 1 To 9
                mto = 0
                If free(r, c, i) <> 0 Then
                    For rr = fr To lr
                        If rr <> r Then
                            If free(r, c, i) = free(rr, c, i) Then
                                mto = mto + 1
                            Else
                                only = free(r, c, i)
                            End If
                        End If
                    Next rr
                    If mto = 0 Then
                        For j = 1 To 9
                            If j <> only Then free(r, c, j) = 0
                        Next j
                    End If
                End If
            Next i
            
            For cc = fc To lc
                If free(r, cc, 0) <> 0 Then free(r, c, free(r, cc, 0)) = 0
            Next cc
            
            For i = 1 To 9
                mto = 0
                If free(r, c, i) <> 0 Then
                    For cc = fc To lc
                        If cc <> c Then
                            If free(r, c, i) = free(r, cc, i) Then
                                mto = mto + 1
                            Else
                                only = free(r, c, i)
                            End If
                        End If
                    Next cc
                    If mto = 0 Then
                        For j = 1 To 9
                            If j <> only Then free(r, c, j) = 0
                        Next j
                    End If
                End If
            Next i
            
            rsubbox = Application.WorksheetFunction.RoundDown((r - fr) / 3, 0) * 3 + fr
            csubbox = Application.WorksheetFunction.RoundDown((c - fc) / 3, 0) * 3 + fc
            For rr = rsubbox To rsubbox + 2
                For cc = csubbox To csubbox + 2
                    If free(rr, cc, 0) <> 0 Then free(r, c, free(rr, cc, 0)) = 0
                Next
            Next
            
            For i = 1 To 9
                mto = 0
                If free(r, c, i) <> 0 Then
                    For rr = rsubbox To rsubbox + 2
                        For cc = csubbox To csubbox + 2
                            If cc <> c Or rr <> r Then
                                If free(r, c, i) = free(rr, cc, i) Then
                                    mto = mto + 1
                                Else
                                    only = free(r, c, i)
                                End If
                            End If
                        Next cc
                    Next rr
                    If mto = 0 Then
                        For j = 1 To 9
                            If j <> only Then free(r, c, j) = 0
                        Next j
                    End If
                End If
            Next i
            
        End If
    Next r
Next c

For c = fc To lc
    For r = fr To lr
        If Cells(r, c) <> "" Then Cells(r, c).Font.ColorIndex = xlAutomatic
    Next
Next

GoTo fillin

wave2:
ReDim uncert(fc To lc) As Integer
Dim match() As Integer
mark = 1
For r = fr To lr
    For c = fc To lc
        count = 0
        For j = 1 To 9
            If free(r, c, j) <> 0 Then count = count + 1
        Next j
        uncert(c) = count
    Next c
    For c = fc To lc
        count = 0
        For cc = fc To lc
            If cc <> c Then
                
                If uncert(cc) <> 0 And uncert(cc) <= uncert(c) Then
                    count = count + 1
                    ReDim Preserve match(count - 1)
                    match(count - 1) = cc
                End If
            End If
        Next cc
        For i = 0 To count - 1
            For j = 1 To 9
                If match(i) <> 0 Then
                    If free(r, match(i), j) <> 0 And free(r, c, j) = 0 Then match(i) = 0
                End If
            Next j
        Next i
        count1 = 0
        For i = 0 To count - 1
            If match(i) <> 0 Then count1 = count1 + 1
        Next i
        If count1 = uncert(c) - 1 Then
            For i = fc To lc
                count2 = 0
                For k = 0 To count - 1
                    If i = match(k) Then count2 = count2 + 1
                Next k
                If count2 = 0 And i <> c Then
                    For j = 1 To 9
                        If free(r, c, j) <> 0 And free(r, i, j) <> 0 Then
                        free(r, i, j) = 0
                        End If
                    Next j
                End If
            Next i
        End If
    Next c
Next r


ReDim uncert(fr To lr) As Integer
ReDim match(0) As Integer
For c = fc To lc
    For r = fr To lr
        count = 0
        For j = 1 To 9
            If free(r, c, j) <> 0 Then count = count + 1
        Next j
        uncert(r) = count
    Next r
    For r = fr To lr
        count = 0
        For rr = fr To lr
            If rr <> r Then
                
                If uncert(rr) <> 0 And uncert(rr) <= uncert(r) Then
                    count = count + 1
                    ReDim Preserve match(count - 1)
                    match(count - 1) = rr
                End If
            End If
        Next rr
        For i = 0 To count - 1
            For j = 1 To 9
                If match(i) <> 0 Then
                    If free(match(i), c, j) <> 0 And free(r, c, j) = 0 Then match(i) = 0
                End If
            Next j
        Next i
        count1 = 0
        For i = 0 To count - 1
            If match(i) <> 0 Then count1 = count1 + 1
        Next i
        If count1 = uncert(r) - 1 Then
            For i = fr To lr
                count2 = 0
                For k = 0 To count - 1
                    If i = match(k) Then count2 = count2 + 1
                Next k
                If count2 = 0 And i <> r Then
                    For j = 1 To 9
                        If free(r, c, j) <> 0 And free(i, c, j) <> 0 Then
                        free(i, c, j) = 0
                        End If
                    Next j
                End If
            Next i
        End If
    Next r
Next c




Dim match2
For bc = fc To lc Step 3
    For br = fr To lr Step 3
        rsubbox = Application.WorksheetFunction.RoundDown((br - fr) / 3, 0) * 3 + fr
        csubbox = Application.WorksheetFunction.RoundDown((bc - fc) / 3, 0) * 3 + fc
        ReDim uncert2(rsubbox To rsubbox + 2, csubbox To csubbox + 2) As Integer
        ReDim match2(0)
        For r = rsubbox To rsubbox + 2
            For c = csubbox To csubbox + 2
                count = 0
                For j = 1 To 9
                    If free(r, c, j) <> 0 Then count = count + 1
                Next j
                uncert2(r, c) = count
            Next c
        Next r
        For r = rsubbox To rsubbox + 2
            For c = csubbox To csubbox + 2
                count = 0
                For rr = rsubbox To rsubbox + 2
                    For cc = csubbox To csubbox + 2
                        If rr <> r Or cc <> c Then
                            
                            If uncert2(rr, cc) <> 0 And uncert2(rr, cc) <= uncert2(r, c) Then
                                count = count + 1
                                ReDim Preserve match2(count - 1)
                                match2(count - 1) = Array(rr, cc)
                                
                            End If
                        End If
                    Next cc
                Next rr
                For i = 0 To count - 1
                    For j = 1 To 9
                        If match2(i)(0) + match2(i)(1) <> 0 Then
                            If free(match2(i)(0), match2(i)(1), j) <> 0 And free(r, c, j) = 0 Then
                                match2(i) = Array(0, 0)
                            End If
                        End If
                    Next j
                Next i
                count1 = 0
                For i = 0 To count - 1
                    If match2(i)(0) + match2(i)(1) <> 0 Then count1 = count1 + 1
                Next i
                If count1 = uncert2(r, c) - 1 Then
                    For l = rsubbox To rsubbox + 2
                        For m = csubbox To csubbox + 2
                            count2 = 0
                            For k = 0 To count - 1
                                If l = match2(k)(0) And m = match2(k)(1) Then count2 = count2 + 1
                            Next k
                            If count2 = 0 And (l <> r Or m <> c) Then
                                For j = 1 To 9
                                    If free(r, c, j) <> 0 And free(l, m, j) <> 0 Then
                                    free(l, m, j) = 0
                                    End If
                                Next j
                            End If
                        Next m
                    Next l
                End If
            Next c
        Next r
    Next br
Next bc
GoTo fillin

guess:
mark2 = 1
Dim op1, op2 As Integer
For r = fr To lr
    For c = fc To lc
        count = 0
        For i = 1 To 9
            If free(r, c, i) <> 0 Then
                count = count + 1
                If count = 1 Then op1 = i
                If count = 2 Then op2 = i
            End If
        Next i
        If count = 2 Then GoTo fiftyfifty
    Next c
Next r
MsgBox "No 50/50 found"
Exit Sub
fiftyfifty:
ReDim revert(fr To lr, fc To lc)
For rr = fr To lr
    For cc = fc To lc
        revert(rr, cc) = ActiveSheet.Cells(rr, cc)
    Next cc
Next rr

If mark3 = 1 Then
    free(r, c, op1) = 0
Else
    free(r, c, op2) = 0
End If

fillin:
'check remaining possibilities
For c = fc To lc
    For r = fr To lr
        count = 0
        For p = 1 To 9
            If free(r, c, p) <> 0 Then
                count = count + 1
                ans = free(r, c, p)
            End If
        Next
        If count = 1 Then
            free(r, c, 0) = ans
            free(r, c, ans) = 0
        End If
        If free(r, c, 0) <> 0 Then Cells(r, c) = free(r, c, 0)
    Next
Next


For c = fc To lc
    For r = fr To lr
        For i = 1 To 9
            If free(r, c, i) <> oldfree(r, c, i) Then GoTo wave1
        Next i
    Next r
Next c

If mark2 = 1 Then
    With Range(Cells(fr, fc), Cells(lr, lc)).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    For r = fr To lr
        For c = fc To lc
            If Cells(r, c) <> "" Then
                For rr = fr To lr
                    If Cells(r, c) = Cells(rr, c) And r <> rr Then
                        Cells(r, c).Interior.Color = 65535
                        Cells(rr, c).Interior.Color = 65535
                        mark = 3
                    End If
                Next rr
                
                For cc = fc To lc
                    If Cells(r, c) = Cells(r, cc) And c <> cc Then
                        Cells(r, c).Interior.Color = 65535
                        Cells(r, cc).Interior.Color = 65535
                        mark = 3
                    End If
                Next cc
            
            
                rsubbox = Application.WorksheetFunction.RoundDown((r - fr) / 3, 0) * 3 + fr
                csubbox = Application.WorksheetFunction.RoundDown((c - fc) / 3, 0) * 3 + fc
                For rr = rsubbox To rsubbox + 2
                    For cc = csubbox To csubbox + 2
                        If Cells(r, c) = Cells(rr, cc) And (rr <> r Or cc <> c) Then
                            Cells(r, c).Interior.Color = 65535
                            Cells(rr, c).Interior.Color = 65535
                            mark = 3
                        End If
                    Next cc
                Next rr
            End If
        Next c
    Next r
    
    If mark = 3 Then
        For rr = fr To lr
            For cc = fc To lc
                If revert(rr, cc) = 0 Then ActiveSheet.Cells(rr, cc) = ""
            Next cc
        Next rr
        mark3 = 1
        GoTo initial
    End If
End If
    
For c = fc To lc
    For r = fr To lr
        If mark = 1 Then GoTo guess
        If Cells(r, c) = "" Then GoTo wave2
    Next r
Next c

MsgBox "Done!"
here:
End Sub
Private Sub check()

    Dim fr, fc, lr, lc, rsubbox, csubbox, rr, cc, val, count As Integer
    fr = 2
    fc = 2
    
    lr = fr + 8
    lc = fc + 8
    
    With Range(Cells(fr, fc), Cells(lr, lc)).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    For r = fr To lr
        For c = fc To lc
            If Cells(r, c) <> "" Then
                For rr = fr To lr
                    If Cells(r, c) = Cells(rr, c) And r <> rr Then
                        Cells(r, c).Interior.Color = 65535
                        Cells(rr, c).Interior.Color = 65535
                    End If
                Next rr
                
                For cc = fc To lc
                    If Cells(r, c) = Cells(r, cc) And c <> cc Then
                        Cells(r, c).Interior.Color = 65535
                        Cells(r, cc).Interior.Color = 65535
                    End If
                Next cc
            
            
                rsubbox = Application.WorksheetFunction.RoundDown((r - fr) / 3, 0) * 3 + fr
                csubbox = Application.WorksheetFunction.RoundDown((c - fc) / 3, 0) * 3 + fc
                For rr = rsubbox To rsubbox + 2
                    For cc = csubbox To csubbox + 2
                        If Cells(r, c) = Cells(rr, cc) And (rr <> r Or cc <> c) Then
                            Cells(r, c).Interior.Color = 65535
                            Cells(rr, c).Interior.Color = 65535
                        End If
                    Next cc
                Next rr
            End If
        Next c
    Next r
            
                    
End Sub
