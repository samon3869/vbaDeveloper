Private Sub btnEdit_MouseDown(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
EditCustomer
End Sub

' 리스트박스가 클릭되었을 때 선택된 항목의 데이터를 각 텍스트박스에 채우기
Private Sub lstMain_Click()

Dim vArr As Variant
'Get_ListItm 보조함수: 리스트박스에서 선택된 항목을 Array로 리턴
vArr = Get_ListItm(Me.lstMain)

'vArr에 할당된 배열을 각 텍스트박스에 입력
Me.txtID.Value = vArr(0)
Me.txtCustomer.Value = vArr(1)
Me.txtContact.Value = vArr(2)
Me.txtPIC.Value = vArr(3)
Me.txtAddress.Value = vArr(4)

End Sub


' 유저폼이 실행되었을 때 리스트박스를 거래처 정보로 채우기
Private Sub UserForm_Initialize()

Dim DB As Variant

'Get_DB 보조함수: sheet의 전체 range값을 받아와서 Array로 리턴
DB = Get_DB(shtCustomer)

'Update_list 보조함수: 리스트박스를 DB(array형태)로 채우기
'Update_list Listbox, DB, 열넓이
'Me는 해당 함수가 실행되는 object를 의미함. 여기서는 frmCustomer
Update_List Me.lstMain, DB, "0pt;120pt;100pt;80pt;150pt;"

End Sub

'-------------------------------------------------------------

Sub EditCustomer()

Dim DB As Variant

Update_Record shtCustomer, Me.txtID.Value, _
Me.txtCustomer.Value, Me.txtContact.Value, _
Me.txtPIC.Value, Me.txtAddress.Value

Filter_ListBox

Select_ListItm Me.lstMain, Me.txtID.Value

MsgBox "Customer Information update completed.", vbInformation

End Sub


Sub Filter_ListBox()

Dim DB As Variant

DB = Get_DB(shtCustomer)
DB = Filterd_DB(DB, Me.txtSearch.Value)

Update_List Me.lstMain, DB, "0pt;120pt;100pt;80pt;150pt;"

End Sub


Sub Select_ListItm(lstBox As Control, ID, Optional ColNo As Long = 1)

Dim i As Long

If IsNumeric(ID) Then ID = CLng(ID)

With lstBox
    For i = 0 To .ListCount - 1
        If .List(i, ColNo - 1) = ID Then .Selected(i) = True: Exit For
    Next
End With

End Sub









