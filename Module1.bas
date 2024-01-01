Attribute VB_Name = "Module1"
Option Explicit

Sub OpenFiles()

Dim SelectionStr As String
Dim Vars As Variant: Dim Var As Variant

SelectionStr = Multiple_FileDialog

Vars = Split(SelectionStr, "|")

For Each Var In Vars
    Application.Workbooks.Open Var
Next

MsgBox "���õ� ���� ������ ��� �����Ͽ����ϴ�."

End Sub

Public Function Multiple_FileDialog(Optional Title As String = "������ �����ϼ���", Optional FilterName As String = "��������", _
                                    Optional FilterExt As String = "*.xls; *.xlsx; *.xlsm", Optional InitialFolder As String = "", _
                                    Optional InitialView As MsoFileDialogView = msoFileDialogViewList, Optional MultiSelection As Boolean = True, _
                                    Optional PathDelimiter As String = "|", Optional withPath As Boolean = True, Optional withExt As Boolean = True) As String

Dim FDG As FileDialog
Dim Selected As Integer: Dim i As Integer
Dim ReturnStr As String: Dim tempStr As Variant

Set FDG = Application.FileDialog(msoFileDialogFilePicker)

With FDG
    .Title = Title
    .Filters.Add FilterName, FilterExt
    .InitialView = InitialView
    .InitialFileName = InitialFolder
    .AllowMultiSelect = MultiSelection
    Selected = .Show

    If Selected = -1 Then
        For i = 1 To FDG.SelectedItems.Count - 1
            If withPath = False Then tempStr = Right(FDG.SelectedItems(i), Len(FDG.SelectedItems(i)) - InStrRev(FDG.SelectedItems(i), "\")) Else tempStr = FDG.SelectedItems(i)
            If withExt = False Then tempStr = Left(tempStr, InStrRev(tempStr, ".") - 1)
            ReturnStr = ReturnStr & tempStr & PathDelimiter
        Next i
        If withPath = False Then tempStr = Right(FDG.SelectedItems(.SelectedItems.Count), Len(FDG.SelectedItems(.SelectedItems.Count)) - InStrRev(FDG.SelectedItems(.SelectedItems.Count), "\")) Else tempStr = FDG.SelectedItems(.SelectedItems.Count)
        If withExt = False Then tempStr = Left(tempStr, InStrRev(tempStr, ".") - 1)
        ReturnStr = ReturnStr & tempStr
        
        Multiple_FileDialog = ReturnStr
    ElseIf Selected = 0 Then
        MsgBox "���õ� ������ �����Ƿ� ���α׷��� �����մϴ�."
        End
    End If
    
End With

End Function

