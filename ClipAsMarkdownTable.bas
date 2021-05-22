Attribute VB_Name = "ClipAsMarkdownTable"
'<License>------------------------------------------------------------
'
' Copyright (c) 2021 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'選択セルの内容を Markdown テーブルの書式でクリップボードに貼り付ける
'
Sub ClipAsMarkdownTable()
    
    '変数宣言
    Dim startOfRow As Long
    Dim lastOfRow As Long
    Dim startOfCol As Long
    Dim lastOfCol As Long
    Dim isFirstCol As Boolean
    Dim buf As String
    
    'シート範囲全選択されていた場合は、UsedRange内に収まるようにトリミング
    Set range_selection = trimWithUsedRange(Selection)
    
    '初期化
    startOfRow = range_selection.Row
    lastOfRow = startOfRow + range_selection.Rows.Count - 1
    startOfCol = range_selection.Column
    lastOfCol = startOfCol + range_selection.Columns.Count - 1
    buf = ""
    
    '文字列取り込みループ
    rowFocus = startOfRow
    Do '行ループ
    
        If Not (Application.ActiveSheet.Rows(rowFocus).Hidden) Then '対象行が表示状態なら
            
            colFocus = startOfCol
            isFirstCol = True
            
            Do '列ループ
            
                If Not (Application.ActiveSheet.Columns(colFocus).Hidden) Then '対象列が表示状態なら
                
                    If Not (isFirstCol) Then '2列目以降なら
                        buf = buf & " | " '列区切り文字
                        
                    Else
                        buf = buf & "| " '列区切り文字
                        
                    End If
                    
                    '文字列取り込み
                    buf = buf & Application.ActiveSheet.Cells(rowFocus, colFocus).Text
                    
                    isFirstCol = False
                
                End If
                
                colFocus = colFocus + 1
            
            Loop While colFocus <= lastOfCol
            
            buf = buf & " |" & vbCrLf '改行挿入
        
        End If
        
        If (rowFocus = startOfRow) Then
            buf = buf & "|" & WorksheetFunction.Rept(" - |", lastOfCol - startOfCol + 1) & vbCrLf '改行挿入
            
        End If
        
        rowFocus = rowFocus + 1
    
    Loop While rowFocus <= lastOfRow
    
    SetCB buf

End Sub

'
' セル参照範囲が UsedRange 範囲に収まるようにトリミングする
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '指定範囲の右下位置の取得
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRangeの右下位置の取得
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    'トリミング
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '格納して終了
    Set trimWithUsedRange = ret
    
End Function

'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>
 


