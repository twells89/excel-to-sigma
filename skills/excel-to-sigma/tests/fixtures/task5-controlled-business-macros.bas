Attribute VB_Name = "Task5ControlledBusiness"
Option Explicit
' NEW GUARDED FIXTURE. Not extracted legacy behavior. Never executed by migration.
' Raw module accompanies an XLSX; it is not a compiled VBA/XLSM container.

Public Function AllocateOverheadByHeadcount() As Boolean
    On Error GoTo Failed
    Dim hc As Variant, result(1 To 4, 1 To 1) As Variant
    Dim i As Long, v As Double, total As Double, overhead As Double
    hc = ThisWorkbook.Worksheets("Headcount").Range("B2:B5").Value2
    For i = 1 To 4
        If Not TryFiniteDouble(hc(i, 1), v) Then Exit Function
        If v < 0 Then Exit Function
        hc(i, 1) = v
        total = total + v
    Next i
    If Not IsFiniteDouble(total) Then Exit Function
    If total <= 0 Then Exit Function
    If Not TryFiniteDouble(ThisWorkbook.Worksheets("Assumptions").Range("B12").Value2, overhead) Then Exit Function
    If overhead < 0 Then Exit Function
    For i = 1 To 4
        result(i, 1) = (hc(i, 1) / total) * overhead * (1# + 0.5 * Sin(CDbl(i)))
        If Not IsFiniteDouble(CDbl(result(i, 1))) Then Exit Function
    Next i
    ThisWorkbook.Worksheets("Alloc").Range("B2:B5").Value2 = result
    AllocateOverheadByHeadcount = True
Failed:
End Function

Public Function SolveForTarget() As Boolean
    On Error GoTo Failed
    Dim solvedGrowth As Double
    If Not TrySolveBounded(2000000#, 0#, 0.25, 0.01, 64, solvedGrowth) Then Exit Function
    ThisWorkbook.Worksheets("Assumptions").Range("B5").Value2 = solvedGrowth
    SolveForTarget = True
Failed:
End Function

Private Function TryFiniteDouble(ByVal raw As Variant, ByRef value As Double) As Boolean
    On Error GoTo Failed
    If IsError(raw) Or IsEmpty(raw) Or IsNull(raw) Then Exit Function
    If VarType(raw) = vbString Or VarType(raw) = vbBoolean Or VarType(raw) = vbDate Then Exit Function
    If Not IsNumeric(raw) Then Exit Function
    value = CDbl(raw)
    TryFiniteDouble = IsFiniteDouble(value)
Failed:
End Function

Private Function IsFiniteDouble(ByVal value As Double) As Boolean
    On Error GoTo Failed
    If value <> value Then Exit Function
    If Abs(value) > 1.7976931348623157E+308 Then Exit Function
    IsFiniteDouble = True
Failed:
End Function

Private Function TrySectionSign(ByVal section As String, ByRef sign As Double) As Boolean
    Select Case section
        Case "Revenue": sign = 1#
        Case "Cost of Revenue", "Operating Expenses", "Below the Line": sign = -1#
        Case Else: Exit Function
    End Select
    TrySectionSign = True
End Function

Private Function TryLoadActualRows(ByRef actualRows As Variant) As Boolean
    On Error GoTo Failed
    Dim r As Long, j As Long, n As Long, found As Long
    Dim value As Double, sign As Double, line As String, month As String, months As Variant
    Dim definitions As Variant
    definitions = ThisWorkbook.Worksheets("Line Definitions").Range("A2:B13").Value2
    actualRows = ThisWorkbook.Worksheets("Actuals").Range("A2:D163").Value2
    months = Array("Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26")
    For r = 1 To 162
        If VarType(actualRows(r, 1)) <> vbString Or VarType(actualRows(r, 2)) <> vbString Or VarType(actualRows(r, 3)) <> vbString Then Exit Function
        If Not TryFiniteDouble(actualRows(r, 4), value) Then Exit Function
        If Not TrySectionSign(CStr(actualRows(r, 3)), sign) Then Exit Function
        found = 0
        For j = 1 To 12
            If actualRows(r, 1) = definitions(j, 1) And actualRows(r, 3) = definitions(j, 2) Then found = found + 1
        Next j
        If found <> 1 Then Exit Function
        found = 0
        For n = 0 To 5
            If actualRows(r, 2) = months(n) Then found = found + 1
        Next n
        If found <> 1 Then Exit Function
    Next r
    For j = 1 To 12
        For n = 0 To 5
            found = 0
            For r = 1 To 162
                If actualRows(r, 1) = definitions(j, 1) And actualRows(r, 2) = months(n) Then found = found + 1
            Next r
            If found = 0 Then Exit Function
        Next n
    Next j
    TryLoadActualRows = True
Failed:
End Function

Private Function TryLoadForecastInputs(ByRef inputRows As Variant) As Boolean
    On Error GoTo Failed
    Dim r As Long, j As Long, k As Long, found As Long, value As Double
    Dim definitions As Variant
    definitions = ThisWorkbook.Worksheets("Line Definitions").Range("A2:B13").Value2
    inputRows = ThisWorkbook.Worksheets("Forecast Inputs").Range("A6:I17").Value2
    For r = 1 To 12
        If VarType(inputRows(r, 1)) <> vbString Or VarType(inputRows(r, 2)) <> vbString Then Exit Function
        found = 0
        For j = 1 To 12
            If definitions(j, 1) = inputRows(r, 1) Then found = found + 1
            If j <> r And inputRows(j, 1) = inputRows(r, 1) Then Exit Function
        Next j
        If found <> 1 Then Exit Function
        Select Case inputRows(r, 2)
            Case "growth"
                If Not TryFiniteDouble(inputRows(r, 3), value) Then Exit Function
                If value <= -1# Then Exit Function
            Case "manual", "flat"
                If Not IsEmpty(inputRows(r, 3)) Then Exit Function
            Case Else: Exit Function
        End Select
        For k = 4 To 9
            If inputRows(r, 2) = "manual" Then
                If Not TryFiniteDouble(inputRows(r, k), value) Then Exit Function
            Else
                If Not IsEmpty(inputRows(r, k)) Then Exit Function
            End If
        Next k
        If inputRows(r, 1) = "Subscription Revenue" And inputRows(r, 2) <> "growth" Then Exit Function
    Next r
    TryLoadForecastInputs = True
Failed:
End Function

Private Function TryJuneBase(ByVal lineItem As String, ByRef actualRows As Variant, ByRef juneBase As Double) As Boolean
    On Error GoTo Failed
    Dim r As Long, value As Double, found As Boolean
    juneBase = 0#
    For r = 1 To 162
        If actualRows(r, 1) = lineItem And actualRows(r, 2) = "Jun-26" Then
            If Not TryFiniteDouble(actualRows(r, 4), value) Then Exit Function
            juneBase = juneBase + value
            found = True
        End If
    Next r
    If Not found Or Not IsFiniteDouble(juneBase) Then Exit Function
    If lineItem = "Subscription Revenue" And juneBase <= 0 Then Exit Function
    TryJuneBase = True
Failed:
End Function

Private Function TryForecastAmount(ByVal lineItem As String, ByVal method As String, ByVal juneBase As Double, ByVal candidateGrowth As Double, ByVal monthIndex As Long, ByRef inputRows As Variant, ByVal rowIndex As Long, ByRef amount As Double) As Boolean
    On Error GoTo Failed
    Dim growth As Double
    If monthIndex < 1 Or monthIndex > 6 Then Exit Function
    Select Case method
        Case "manual"
            If Not TryFiniteDouble(inputRows(rowIndex, monthIndex + 3), amount) Then Exit Function
        Case "flat": amount = juneBase
        Case "growth"
            If Not TryFiniteDouble(inputRows(rowIndex, 3), growth) Then Exit Function
            If lineItem = "Subscription Revenue" Then growth = candidateGrowth
            amount = juneBase * (1# + growth) ^ monthIndex
        Case Else: Exit Function
    End Select
    TryForecastAmount = IsFiniteDouble(amount)
Failed:
End Function

Private Function TryAnnualNetIncome(ByVal candidateGrowth As Double, ByRef netIncome As Double) As Boolean
    On Error GoTo Failed
    Dim actualRows As Variant, inputRows As Variant, definitions As Variant
    Dim r As Long, j As Long, n As Long, section As String, lineItem As String
    Dim sign As Double, value As Double, juneBase As Double, amount As Double
    If Not IsFiniteDouble(candidateGrowth) Then Exit Function
    If Not TryLoadActualRows(actualRows) Then Exit Function
    If Not TryLoadForecastInputs(inputRows) Then Exit Function
    definitions = ThisWorkbook.Worksheets("Line Definitions").Range("A2:B13").Value2
    netIncome = 0#
    For r = 1 To 162
        If Not TrySectionSign(CStr(actualRows(r, 3)), sign) Then Exit Function
        If Not TryFiniteDouble(actualRows(r, 4), value) Then Exit Function
        netIncome = netIncome + sign * value
    Next r
    For r = 1 To 12
        lineItem = CStr(inputRows(r, 1))
        For j = 1 To 12
            If definitions(j, 1) = lineItem Then section = CStr(definitions(j, 2))
        Next j
        If Not TrySectionSign(section, sign) Then Exit Function
        If Not TryJuneBase(lineItem, actualRows, juneBase) Then Exit Function
        For n = 1 To 6
            If Not TryForecastAmount(lineItem, CStr(inputRows(r, 2)), juneBase, candidateGrowth, n, inputRows, r, amount) Then Exit Function
            netIncome = netIncome + sign * amount
        Next n
    Next r
    TryAnnualNetIncome = IsFiniteDouble(netIncome)
Failed:
End Function

Private Function TrySolveBounded(ByVal target As Double, ByVal lower As Double, ByVal upper As Double, ByVal tolerance As Double, ByVal maxIterations As Long, ByRef solvedGrowth As Double) As Boolean
    On Error GoTo Failed
    Dim lo As Double, hi As Double, mid As Double, value As Double, iteration As Long, converged As Boolean
    If Not IsFiniteDouble(target) Or Not IsFiniteDouble(lower) Or Not IsFiniteDouble(upper) Then Exit Function
    If lower < 0# Or upper > 0.25 Or lower >= upper Or tolerance <> 0.01 Or maxIterations < 1 Or maxIterations > 64 Then Exit Function
    If Not TryAnnualNetIncome(lower, lo) Then Exit Function
    If Not TryAnnualNetIncome(upper, hi) Then Exit Function
    If lo = hi Then Exit Function
    If Abs(lo - target) <= tolerance Then
        mid = lower: converged = True
    ElseIf Abs(hi - target) <= tolerance Then
        mid = upper: converged = True
    Else
        If (lo - target) * (hi - target) >= 0 Then Exit Function
        For iteration = 1 To maxIterations
            mid = (lower + upper) / 2#
            If Not TryAnnualNetIncome(mid, value) Then Exit Function
            If Abs(value - target) <= tolerance Then
                converged = True
                Exit For
            End If
            If (value - target) * (lo - target) > 0 Then
                lower = mid: lo = value
            Else
                upper = mid
            End If
        Next iteration
    End If
    If Not converged Then Exit Function
    mid = Round(mid, 12)
    If Not TryAnnualNetIncome(mid, value) Then Exit Function
    If Abs(value - target) > tolerance Then Exit Function
    solvedGrowth = mid
    TrySolveBounded = True
Failed:
End Function

Public Function Task5TypedEcho(ByVal operationId As Variant, ByVal amount As Variant, ByVal count As Variant, ByVal enabled As Variant, ByVal occurredAt As Variant) As String
    ' Controlled scalar contract, no database call and no persistence.
    ' SQL counterpart is CALLER rights, null -20051, invalid operation -20052, negative count -20053.
    Dim a As Double, c As Double, timestamp As String
    If IsNull(operationId) Or IsNull(amount) Or IsNull(count) Or IsNull(enabled) Or IsNull(occurredAt) Then Err.Raise 20051, , "TASK5_REQUIRED_INPUT"
    If VarType(operationId) <> vbString Then Err.Raise 20052, , "TASK5_INVALID_OPERATION_ID"
    If Len(Trim$(operationId)) = 0 Or InStr(operationId, "|") > 0 Then Err.Raise 20052, , "TASK5_INVALID_OPERATION_ID"
    If operationId <> Trim$(operationId) Then Err.Raise 20052, , "TASK5_INVALID_OPERATION_ID"
    If InStr(operationId, vbTab) > 0 Or InStr(operationId, vbCr) > 0 Or InStr(operationId, vbLf) > 0 Then Err.Raise 20052, , "TASK5_INVALID_OPERATION_ID"
    If Not TryFiniteDouble(amount, a) Or Not TryFiniteDouble(count, c) Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    If a <> Round(a, 2) Or Abs(a) >= 1E+15 Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    If c < 0 Then Err.Raise 20053, , "TASK5_NEGATIVE_COUNT"
    If c <> Fix(c) Or c >= 1E+15 Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    If VarType(enabled) <> vbBoolean Or VarType(occurredAt) <> vbString Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    timestamp = CStr(occurredAt)
    If Len(timestamp) <> 23 Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    If Not timestamp Like "####-##-##T##:##:##.###" Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    If Not IsDate(Left$(timestamp, 10) & " " & Mid$(timestamp, 12, 8)) Then Err.Raise 20054, , "TASK5_INVALID_TYPED_INPUT"
    Task5TypedEcho = "operation_id=" & operationId & "|amount=" & Replace(Format$(a, "0.00"), ",", ".") & "|count=" & Format$(c, "0") & "|enabled=" & LCase$(CStr(enabled)) & "|occurred_at=" & timestamp
End Function
