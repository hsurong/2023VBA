Attribute VB_Name = "Module2"
Sub AdvanceDemo()
Attribute AdvanceDemo.VB_ProcData.VB_Invoke_Func = "k\n14"
'�p�GD3-D24�x�s�檺�Ȭ��s�_�� ��
'today-�令���U��
Dim targetValue As Variant
targetValue = InputBox("�п�J�n�z���")

'end
'----------step4
Dim targetCId As Integer
targetCId = InputBox("�п�J���z�諸�ů���")

'���ӷ|�� typename
If (IsNumeric(targetValue)) Then
targetValue = CDbl(targetValue)
End If

Dim row As Integer
For row = 3 To 24

If Cells(row, targetCId).Value = targetValue Then
'�r�鬰����
Cells(row, targetCId).Font.ColorIndex = 3
'�I��������
Cells(row, targetCId).Interior.ColorIndex = 6

'�_�h
Else
'�r�鬰�¦�
Cells(row, targetCId).Font.ColorIndex = 1
'�I�����z��
Cells(row, targetCId).Interior.ColorIndex = 0

End If
Next
End Sub
