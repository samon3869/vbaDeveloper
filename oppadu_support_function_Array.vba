Sub ArrayToRng(startRng As Range, Arr As Variant)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ArrayToRng 함수
'▶ 배열을 범위 위로 반환합니다.
'▶ 인수 설명
'_____________startRng      : 배열을 반환할 기준 범위(셀) 입니다.
'_____________Arr               : 반환할 배열입니다.
'▶ 사용 예제
'Dim v As Variant
'ReDim v(0 to 1)
''v(0) = "a" : v(1) = "b"
'ArrayToRng Sheet1.Range("A1"), v
'##############################################################

On Error GoTo SingleDimension:
startRng.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, UBound(Arr, 2) - LBound(Arr, 2) + 1) = Arr

Exit Sub
SingleDimension:
Dim tempArr As Variant: Dim i As Long
ReDim tempArr(LBound(Arr, 1) To UBound(Arr, 1), 1 To 1)
For i = LBound(Arr, 1) To UBound(Arr, 1)
    tempArr(i, 1) = Arr(i)
Next
startRng.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, 1) = tempArr

End Sub

