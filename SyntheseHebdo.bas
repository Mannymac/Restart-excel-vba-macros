Attribute VB_Name = "SyntheseHebdo"
Option Explicit

Public Sub GenererSyntheseHebdo()
    Const HEADER_EMP As String = "nom_emp"
    Const HEADER_UNIT As String = "num_csst"
    Const HEADER_DATE As String = "date_debut"
    Const HEADER_SALARY As String = "sal_csst"
    Const HEADER_BONUS As String = "montant_prime"

    Dim wsSource As Worksheet
    Set wsSource = ActiveSheet

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(wsSource)

    If Not headerMap.Exists(HEADER_EMP) Or _
       Not headerMap.Exists(HEADER_UNIT) Or _
       Not headerMap.Exists(HEADER_DATE) Or _
       Not headerMap.Exists(HEADER_SALARY) Or _
       Not headerMap.Exists(HEADER_BONUS) Then
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    Dim lastRow As Long
    lastRow = LastDataRow(wsSource)
    If lastRow < 2 Then
        WriteSyntheseHeaders GetOrCreateSyntheseSheet(wsSource.Parent)
        Exit Sub
    End If

    Dim data As Variant
    data = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, lastCol)).Value2

    Dim endP01 As Date
    endP01 = DateSerial(2024, 12, 28)

    Dim endP53 As Date
    endP53 = DateSerial(2025, 12, 27)

    Dim idxEmp As Long, idxUnit As Long, idxDate As Long, idxSalary As Long, idxBonus As Long
    idxEmp = CLng(headerMap(HEADER_EMP))
    idxUnit = CLng(headerMap(HEADER_UNIT))
    idxDate = CLng(headerMap(HEADER_DATE))
    idxSalary = CLng(headerMap(HEADER_SALARY))
    idxBonus = CLng(headerMap(HEADER_BONUS))

    Dim sumsByGroupPeriod As Object
    Set sumsByGroupPeriod = CreateObject("Scripting.Dictionary")

    Dim boundsByGroup As Object
    Set boundsByGroup = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim periodIndex As Long
        periodIndex = ResolvePeriodIndex(data(r, idxDate), endP01, endP53)

        If periodIndex > 0 Then
            Dim emp As String
            emp = Trim$(CStr(data(r, idxEmp)))

            Dim unitValue As String
            unitValue = Trim$(CStr(data(r, idxUnit)))

            Dim groupKey As String
            groupKey = BuildGroupKey(emp, unitValue)

            Dim periodKey As String
            periodKey = BuildPeriodKey(groupKey, periodIndex)

            Dim salary As Double
            salary = ParseAmount(data(r, idxSalary))

            Dim bonus As Double
            bonus = ParseAmount(data(r, idxBonus))

            AddPeriodSums sumsByGroupPeriod, periodKey, salary, bonus
            UpdateGroupBounds boundsByGroup, groupKey, periodIndex
        End If
    Next r

    WriteSyntheseOutput wsSource.Parent, sumsByGroupPeriod, boundsByGroup, endP01
End Sub

Private Function BuildHeaderMap(ByVal ws As Worksheet) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        Dim rawHeader As String
        rawHeader = CStr(ws.Cells(1, c).Value2)

        Dim normalized As String
        normalized = NormalizeHeader(rawHeader)

        If Len(normalized) > 0 Then
            map(normalized) = c
        End If
    Next c

    Set BuildHeaderMap = map
End Function

Private Function NormalizeHeader(ByVal header As String) As String
    NormalizeHeader = LCase$(Trim$(header))
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        LastDataRow = 1
    Else
        LastDataRow = lastCell.Row
    End If
End Function

Private Function ResolvePeriodIndex(ByVal value As Variant, ByVal endP01 As Date, ByVal endP53 As Date) As Long
    If Not IsDate(value) Then
        ResolvePeriodIndex = 0
        Exit Function
    End If

    Dim d As Date
    d = CDate(value)

    If d < (endP01 - 6) Or d > endP53 Then
        ResolvePeriodIndex = 0
        Exit Function
    End If

    Dim diff As Long
    diff = CLng(d - endP01)

    Dim idx As Long
    idx = Int((diff + 6) / 7) + 1

    If idx < 1 Or idx > 53 Then
        ResolvePeriodIndex = 0
        Exit Function
    End If

    ResolvePeriodIndex = idx
End Function

Private Function BuildGroupKey(ByVal emp As String, ByVal unitValue As String) As String
    BuildGroupKey = emp & vbNullChar & unitValue
End Function

Private Function BuildPeriodKey(ByVal groupKey As String, ByVal periodIndex As Long) As String
    BuildPeriodKey = groupKey & vbNullChar & CStr(periodIndex)
End Function

Private Sub AddPeriodSums(ByVal sumsByGroupPeriod As Object, ByVal periodKey As String, ByVal salary As Double, ByVal bonus As Double)
    Dim values As Variant

    If sumsByGroupPeriod.Exists(periodKey) Then
        values = sumsByGroupPeriod(periodKey)
    Else
        ReDim values(1 To 2)
        values(1) = 0#
        values(2) = 0#
    End If

    values(1) = CDbl(values(1)) + salary
    values(2) = CDbl(values(2)) + bonus

    sumsByGroupPeriod(periodKey) = values
End Sub

Private Sub UpdateGroupBounds(ByVal boundsByGroup As Object, ByVal groupKey As String, ByVal periodIndex As Long)
    Dim bounds As Variant

    If boundsByGroup.Exists(groupKey) Then
        bounds = boundsByGroup(groupKey)
        If periodIndex < CLng(bounds(1)) Then bounds(1) = periodIndex
        If periodIndex > CLng(bounds(2)) Then bounds(2) = periodIndex
        boundsByGroup(groupKey) = bounds
    Else
        ReDim bounds(1 To 2)
        bounds(1) = periodIndex
        bounds(2) = periodIndex
        boundsByGroup.Add groupKey, bounds
    End If
End Sub

Private Function ParseAmount(ByVal value As Variant) As Double
    If IsError(value) Or IsEmpty(value) Then
        ParseAmount = 0#
        Exit Function
    End If

    If IsNumeric(value) Then
        ParseAmount = CDbl(value)
        Exit Function
    End If

    Dim s As String
    s = CStr(value)
    s = Replace$(s, "$", "")
    s = Replace$(s, " ", "")
    s = Replace$(s, ChrW$(160), "")

    If Len(s) = 0 Then
        ParseAmount = 0#
        Exit Function
    End If

    Dim dotPos As Long
    dotPos = InStrRev(s, ".")

    Dim commaPos As Long
    commaPos = InStrRev(s, ",")

    If dotPos > 0 And commaPos > 0 Then
        If dotPos > commaPos Then
            s = Replace$(s, ",", "")
        Else
            s = Replace$(s, ".", "")
            s = Replace$(s, ",", ".")
        End If
    ElseIf commaPos > 0 Then
        s = Replace$(s, ",", ".")
    End If

    If IsNumeric(s) Then
        ParseAmount = CDbl(s)
    Else
        ParseAmount = 0#
    End If
End Function

Private Function GetOrCreateSyntheseSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets("Synthese")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Synthese"
    End If

    Set GetOrCreateSyntheseSheet = ws
End Function

Private Sub WriteSyntheseHeaders(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Employé", "Unité", "Période", "Fin période", "Salaire", "Prime", "Total")
End Sub

Private Sub WriteSyntheseOutput(ByVal wb As Workbook, ByVal sumsByGroupPeriod As Object, ByVal boundsByGroup As Object, ByVal endP01 As Date)
    Dim ws As Worksheet
    Set ws = GetOrCreateSyntheseSheet(wb)

    WriteSyntheseHeaders ws

    If boundsByGroup.Count = 0 Then Exit Sub

    Dim groups() As String
    groups = DictionaryKeysToStringArray(boundsByGroup)
    QuickSortStrings groups, LBound(groups), UBound(groups)

    Dim totalRows As Long
    totalRows = CountOutputRows(boundsByGroup)

    Dim outData() As Variant
    ReDim outData(1 To totalRows, 1 To 7)

    Dim outRow As Long
    outRow = 1

    Dim g As Long
    For g = LBound(groups) To UBound(groups)
        Dim groupKey As String
        groupKey = groups(g)

        Dim parts As Variant
        parts = Split(groupKey, vbNullChar)

        Dim emp As String
        emp = parts(0)

        Dim unitValue As String
        unitValue = vbNullString
        If UBound(parts) >= 1 Then unitValue = parts(1)

        Dim bounds As Variant
        bounds = boundsByGroup(groupKey)

        Dim p As Long
        For p = CLng(bounds(1)) To CLng(bounds(2))
            Dim periodKey As String
            periodKey = BuildPeriodKey(groupKey, p)

            Dim salary As Double, bonus As Double
            salary = 0#
            bonus = 0#

            If sumsByGroupPeriod.Exists(periodKey) Then
                Dim sums As Variant
                sums = sumsByGroupPeriod(periodKey)
                salary = CDbl(sums(1))
                bonus = CDbl(sums(2))
            End If

            outData(outRow, 1) = emp
            outData(outRow, 2) = unitValue
            outData(outRow, 3) = "P" & Format$(p, "00")
            outData(outRow, 4) = endP01 + (p - 1) * 7
            outData(outRow, 5) = salary
            outData(outRow, 6) = bonus
            outData(outRow, 7) = salary + bonus

            outRow = outRow + 1
        Next p
    Next g

    ws.Range("A2").Resize(totalRows, 7).Value = outData
    ws.Columns("A:G").AutoFit
    ws.Columns("D").NumberFormat = "dd/mm/yyyy"
    ws.Columns("E:G").NumberFormat = "#,##0.00"
End Sub

Private Function CountOutputRows(ByVal boundsByGroup As Object) As Long
    Dim key As Variant
    Dim n As Long
    n = 0

    For Each key In boundsByGroup.Keys
        Dim bounds As Variant
        bounds = boundsByGroup(key)
        n = n + (CLng(bounds(2)) - CLng(bounds(1)) + 1)
    Next key

    CountOutputRows = n
End Function

Private Function DictionaryKeysToStringArray(ByVal dict As Object) As String()
    Dim keys As Variant
    keys = dict.Keys

    Dim arr() As String
    ReDim arr(LBound(keys) To UBound(keys))

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        arr(i) = CStr(keys(i))
    Next i

    DictionaryKeysToStringArray = arr
End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop

        Do While arr(j) > pivot
            j = j - 1
        Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub
