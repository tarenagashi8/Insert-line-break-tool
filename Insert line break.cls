VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 選択範囲のセル値末尾に改行コードを追加する
' ・未入力セルには追加しない
' ・選択範囲は複数も可
Public Sub InsertLineBreak()

    Dim Worksheet As Worksheet
    Set Worksheet = Application.ActiveSheet
    
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    ' 値が入力されているセルの末尾に改行コードを追加
    For Each cell In selectedRange
        If cell.Value <> "" Then
            cell.Value = cell.Value & vbLf
        End If
    Next
    
    Cells.EntireRow.AutoFit
    
End Sub


' 選択範囲のセル値末尾の改行コードを一括削除する
' ・選択範囲は複数も可
Public Sub DeleteLineBreak()

    Dim Worksheet As Worksheet
    Set Worksheet = Application.ActiveSheet
    
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    ' 末尾が改行コードの場合、改行コードを削除
    For Each cell In selectedRange
        If cell.Value <> "" And Right(cell.Value, 1) = vbLf Then
            cell.Value = Left(cell.Value, Len(cell.Value) - 1)
        End If
    Next
    
    Cells.EntireRow.AutoFit
    
End Sub
