VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' �I��͈͂̃Z���l�����ɉ��s�R�[�h��ǉ�����
' �E�����̓Z���ɂ͒ǉ����Ȃ�
' �E�I��͈͕͂�������
Public Sub InsertLineBreak()

    Dim Worksheet As Worksheet
    Set Worksheet = Application.ActiveSheet
    
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    ' �l�����͂���Ă���Z���̖����ɉ��s�R�[�h��ǉ�
    For Each cell In selectedRange
        If cell.Value <> "" Then
            cell.Value = cell.Value & vbLf
        End If
    Next
    
    Cells.EntireRow.AutoFit
    
End Sub


' �I��͈͂̃Z���l�����̉��s�R�[�h���ꊇ�폜����
' �E�I��͈͕͂�������
Public Sub DeleteLineBreak()

    Dim Worksheet As Worksheet
    Set Worksheet = Application.ActiveSheet
    
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    ' ���������s�R�[�h�̏ꍇ�A���s�R�[�h���폜
    For Each cell In selectedRange
        If cell.Value <> "" And Right(cell.Value, 1) = vbLf Then
            cell.Value = Left(cell.Value, Len(cell.Value) - 1)
        End If
    Next
    
    Cells.EntireRow.AutoFit
    
End Sub