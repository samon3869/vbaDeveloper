Option Explicit
Option Compare Text
'########################
' 특정 워크시트에서 앞으로 추가해야 할 최대 ID번호 리턴 (시트 DB 우측 첫번째 머릿글)
' i = Get_MaxID(Sheet1)
'########################
Function Get_MaxID(WS As Worksheet) As Long
With WS
    Get_MaxID = .Cells(1, .Columns.Count).End(xlToLeft).Value
    .Cells(1, .Columns.Count).End(xlToLeft).Value = .Cells(1, .Columns.Count).End(xlToLeft).Value + 1
End With
End Function
'########################
' 워크시트에 새로운 데이터를 추가해야 할 열번호 반환
' i = Get_InsertRow(Sheet1)
'########################
Function Get_InsertRow(WS As Worksheet) As Long
With WS:    Get_InsertRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1: End With
End Function
'########################
' 시트의 열 개수 반환 (이번 예제파일에서만 사용)
' i  = Get_ColumnCnt(Sheet1)
'########################
Function Get_ColumnCnt(WS As Worksheet, Optional Offset As Long = -1) As Long
With WS:    Get_ColumnCnt = .Cells(1, .Columns.Count).End(xlToLeft).Column + Offset: End With
End Function
'########################
' 시트에서 특정 ID 의 행 번호 반환 (-> 해당 행 번호 데이터 업데이트)
' i = get_UpdateRow(Sheet1, ID)
'########################
Function get_UpdateRow(WS As Worksheet, ID)
Dim i As Long
Dim cRow As Long
With WS
    cRow = Get_InsertRow(WS) - 1
    For i = 1 To cRow
        If .Cells(i, 1).Value = ID Then get_UpdateRow = i: Exit For
    Next
End With
End Function


'########################
' 특정 시트의 DB 정보를 배열로 반환 (이번 예제파일에서만 사용)
' Array = Get_DB(Sheet1)
'########################
Function Get_DB(WS As Worksheet, Optional NoID As Boolean = False) As Variant

Dim cRow As Long
Dim cCol As Long
Dim offCol As Long

If NoID = False Then offCol = -1

With WS
    cRow = Get_InsertRow(WS) - 1
    cCol = Get_ColumnCnt(WS, offCol)
    Get_DB = .Range(.Cells(2, 1), .Cells(cRow, cCol))
End With
    
End Function
'########################
'특정 시트에서 지정한 ID의 필드 값 반환 (이번 예제파일 전용)
' Value = Get_Records(Sheet1, ID, "필드명")
'########################
Function Get_Records(WS As Worksheet, ID, Fields)

Dim cRow As Long: Dim cCol As Long
Dim vFields As Variant: Dim vField As Variant
Dim vFieldNo As Variant
Dim i As Long: Dim j As Long


cRow = Get_InsertRow(WS) - 1
cCol = Get_ColumnCnt(WS)

If InStr(1, Fields, ",") > 0 Then vFields = Split(Fields, ",") Else vFields = Array(Fields)
ReDim vFieldNo(0 To UBound(vFields))

With WS
    For Each vField In vFields
        For i = 1 To cCol
            If .Cells(1, i).Value = Trim(vField) Then vFieldNo(j) = i: j = j + 1
        Next
    Next

    For i = 2 To cRow
        If .Cells(i, 1).Value = ID Then
            For j = 0 To UBound(vFieldNo)
                vFieldNo(j) = .Cells(i, vFieldNo(j))
            Next
            Exit For
        End If
    Next
    
Get_Records = vFieldNo

End With

End Function

'########################
' 시트에 새로운 레코드 추가 (반드시 첫번째 값은 ID, 나머지 값 순서대로 입력)
' Insert_Record Sheet1, ID, 필드1, 필드2, 필드3, ..
'########################
Sub Insert_Record(WS As Worksheet, ParamArray vaParamArr() As Variant)

Dim cID As Long
Dim cRow As Long
Dim vaArr As Variant: Dim i As Long: i = 2

With WS
    cRow = Get_InsertRow(WS)
    If InStr(1, .Cells(1, 1).Value, "ID") > 0 Then
        cID = Get_MaxID(WS)
        .Cells(cRow, 1).Value = cID
        For Each vaArr In vaParamArr
            .Cells(cRow, i).Value = vaArr
            i = i + 1
        Next
    Else
        For Each vaArr In vaParamArr
            .Cells(cRow, i - 1).Value = vaArr
            i = i + 1
        Next
    End If
    
End With

End Sub
'########################
' 시트에서 ID 를 갖는 레코드의 모든 값 업데이트 (반드시 첫번째 값은 ID여야 하며, 나머지 값을 순서대로 입력)
' Update_Record Sheet1, ID, 필드1, 필드2, 필드3, ...
'########################
Sub Update_Record(WS As Worksheet, ParamArray vaParamArr() As Variant)

Dim cRow As Long
Dim i As Long
Dim ID As Variant

If IsNumeric(vaParamArr(0)) = True Then ID = CLng(vaParamArr(0)) Else ID = vaParamArr(0)

With WS
    cRow = get_UpdateRow(WS, ID)
    
    For i = 1 To UBound(vaParamArr)
        .Cells(cRow, i + 1).Value = vaParamArr(i)
    Next
    
End With

End Sub
'########################
' 시트에서 ID 를 갖는 레코드 삭제
' Delete_Record Sheet1, ID
'########################
Sub Delete_Record(WS As Worksheet, ID)

Dim cRow As Long

If IsNumeric(ID) = True Then ID = CLng(ID)

With WS
    cRow = get_UpdateRow(WS, ID)
    .Cells(cRow, 1).EntireRow.Delete
End With

End Sub

'########################
' 배열의 외부ID키 필드를 본 시트DB와 연결하여 해당 외부ID키의 연관된 값을 배열로 반환
' Array = Connect_DB(Get_DB(Sheet1),2,Sheet2, "필드1, 필드2, 필드3")
'########################
Function Connect_DB(DB As Variant, ForeignID_Fields As Variant, FromWS As Worksheet, Fields As String)

Dim cRow As Long: Dim cCol As Long
Dim vForeignID_Fields As Variant: Dim vForeignID_Field As Variant
Dim ForeignID As Long
Dim vFields As Variant: Dim vField As Variant
Dim vID As Variant: Dim vFieldNo As Variant
Dim dict As Object
Dim i As Long: Dim j As Long
Dim AddCols As Long


cRow = UBound(DB, 1)
cCol = UBound(DB, 2)
If InStr(1, Fields, ",") > 1 Then
    AddCols = Len(Fields) - Len(Replace(Fields, ",", "")) + 1
    vFields = Split(Fields, ",")
Else
    AddCols = 1
    vFields = Array(Fields)
End If

ReDim Preserve DB(1 To cRow, 1 To cCol + AddCols)
        
Set dict = Get_Dict(FromWS)
vID = dict("ID")

ReDim vFieldNo(0 To UBound(vFields))

For Each vField In vFields
    For i = 1 To UBound(vID)
        If vID(i) = Trim(vField) Then vFieldNo(j) = i: j = j + 1
    Next
Next

If InStr(1, ForeignID_Fields, ",") > 0 Then vForeignID_Fields = Split(ForeignID_Fields, ",") Else vForeignID_Fields = Array(ForeignID_Fields)

For Each vForeignID_Field In vForeignID_Fields
    For i = 1 To cRow
        ForeignID = DB(i, Trim(vForeignID_Field))
        If dict.Exists(ForeignID) Then
            For j = 1 To AddCols
                DB(i, cCol + j) = dict(ForeignID)(vFieldNo(j - 1))
            Next
        End If
    Next
Next

Connect_DB = DB
    
End Function
'########################
' 특정 배열에서 Value를 포함하는 레코드만 찾아 다시 배열로 반환
' Array = Filtered_DB(Array, "검색값", False)
'########################
Function Filtered_DB(DB, Value, Optional FilterCol, Optional ExactMatch As Boolean = False) As Variant

Dim cRow As Long
Dim cCol As Long
Dim vArr As Variant: Dim s As String: Dim filterArr As Variant:  Dim Cols As Variant: Dim Col As Variant: Dim Colcnt As Long
Dim isDateVal As Boolean
Dim vReturn As Variant: Dim vResult As Variant
Dim dict As Object: Dim dictKey As Variant
Dim i As Long: Dim j As Long
Dim Operator As String

Set dict = CreateObject("Scripting.Dictionary")

If Value <> "" Then
    cRow = UBound(DB, 1)
    cCol = UBound(DB, 2)
    ReDim vArr(1 To cRow)
    For i = 1 To cRow
        s = ""
        For j = 1 To cCol
            s = s & DB(i, j) & "|^"
        Next
        vArr(i) = s
    Next
    
    If IsMissing(FilterCol) Then
        filterArr = vArr
    Else
        Cols = Split(FilterCol, ",")
        ReDim filterArr(1 To cRow)
        For i = 1 To cRow
            s = ""
            For Each Col In Cols
                s = s & DB(i, Trim(Col)) & "|^"
            Next
            filterArr(i) = s
        Next
    End If
    
    If Left(Value, 2) = ">=" Or Left(Value, 2) = "<=" Or Left(Value, 2) = "=>" Or Left(Value, 2) = "=<" Then
        Operator = Left(Value, 2)
        If IsDate(Right(Value, Len(Value) - 2)) Then isDateVal = True
    ElseIf Left(Value, 1) = ">" Or Left(Value, 1) = "<" Then
        Operator = Left(Value, 1)
        If IsDate(Right(Value, Len(Value) - 1)) Then isDateVal = True
    Else: End If
    
    If Operator <> "" Then
        If isDateVal = False Then
            Select Case Operator
                Case ">"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) > CDbl(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
                Case "<"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) < CDbl(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
                Case ">=", "=>"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) >= CDbl(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
                 Case "<=", "=<"
                    For i = 1 To cRow
                        If CDbl(Left(filterArr(i), Len(filterArr(i)) - 2)) <= CDbl(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
            End Select
        Else
            Select Case Operator
                Case ">"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) > CDate(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
                Case "<"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) < CDate(Right(Value, Len(Value) - 1)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
                Case ">=", "=>"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) >= CDate(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
                 Case "<=", "=<"
                    For i = 1 To cRow
                        If CDate(Left(filterArr(i), Len(filterArr(i)) - 2)) <= CDate(Right(Value, Len(Value) - 2)) Then: vArr(i) = Left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): dict.Add i, vReturn
                    Next
            End Select
        End If
    Else
        If ExactMatch = False Then
            For i = 1 To cRow
                If filterArr(i) Like "*" & Value & "*" Then
                    vArr(i) = Left(vArr(i), Len(vArr(i)) - 2)
                    vReturn = Split(vArr(i), "|^")
                    dict.Add i, vReturn
                End If
            Next
        Else
            For i = 1 To cRow
                If filterArr(i) Like Value & "|^" Then
                    vArr(i) = Left(vArr(i), Len(vArr(i)) - 2)
                    vReturn = Split(vArr(i), "|^")
                    dict.Add i, vReturn
                End If
            Next
        End If
    End If
        
    If dict.Count > 0 Then
        ReDim vResult(1 To dict.Count, 1 To cCol)
        i = 1
        For Each dictKey In dict.Keys
            For j = 1 To cCol
                vResult(i, j) = dict(dictKey)(j - 1)
            Next
            i = i + 1
        Next
    End If
    
    Filtered_DB = vResult
Else
    Filtered_DB = DB
End If

End Function

'########################
' 특정 시트의 DB 정보를 Dictionary로 반환 (이번 예제파일에서만 사용)
' Dict = GetDict(Sheet1)
'########################
Function Get_Dict(WS As Worksheet) As Object

Dim cRow As Long: Dim cCol As Long
Dim dict As Object
Dim vArr As Variant
Dim i As Long: Dim j As Long

Set dict = CreateObject("Scripting.Dictionary")

With WS
    cRow = Get_InsertRow(WS) - 1
    cCol = Get_ColumnCnt(WS)
    
    For i = 1 To cRow
        ReDim vArr(1 To cCol - 1)
        For j = 2 To cCol
            vArr(j - 1) = .Cells(i, j)
        Next
        dict.Add .Cells(i, 1).Value, vArr
    Next
End With

Set Get_Dict = dict

End Function

'########################
' 시트의 특정 필드 내에서 추가되는 값이 고유값인지 확인. 고유값일 경우 TRUE를 반환
' boolean = IsUnique(Sheet1, "사과", 1)
'########################
Function IsUnique(DB As Variant, uniqueVal, Optional ColNo As Long = 1, Optional Exclude) As Boolean

Dim endRow As Long
Dim i As Long

For i = LBound(DB, 1) To UBound(DB, 1)
    If DB(i, ColNo) = uniqueVal Then
        If Not IsMissing(Exclude) Then
            If Exclude <> uniqueVal Then
                IsUnique = False
                Exit Function
            End If
        Else
            IsUnique = False: Exit Function
        End If
    End If
Next

IsUnique = True

End Function
