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
'�I���Z���̓��e�� Markdown �e�[�u���̏����ŃN���b�v�{�[�h�ɓ\��t����
'
Sub ClipAsMarkdownTable()
    
    '�ϐ��錾
    Dim startOfRow As Long
    Dim lastOfRow As Long
    Dim startOfCol As Long
    Dim lastOfCol As Long
    Dim isFirstCol As Boolean
    Dim buf As String
    
    '�V�[�g�͈͑S�I������Ă����ꍇ�́AUsedRange���Ɏ��܂�悤�Ƀg���~���O
    Set range_selection = trimWithUsedRange(Selection)
    
    '������
    startOfRow = range_selection.Row
    lastOfRow = startOfRow + range_selection.Rows.Count - 1
    startOfCol = range_selection.Column
    lastOfCol = startOfCol + range_selection.Columns.Count - 1
    buf = ""
    
    '�������荞�݃��[�v
    rowFocus = startOfRow
    Do '�s���[�v
    
        If Not (Application.ActiveSheet.Rows(rowFocus).Hidden) Then '�Ώۍs���\����ԂȂ�
            
            colFocus = startOfCol
            isFirstCol = True
            
            Do '�񃋁[�v
            
                If Not (Application.ActiveSheet.Columns(colFocus).Hidden) Then '�Ώۗ񂪕\����ԂȂ�
                
                    If Not (isFirstCol) Then '2��ڈȍ~�Ȃ�
                        buf = buf & " | " '���؂蕶��
                        
                    Else
                        buf = buf & "| " '���؂蕶��
                        
                    End If
                    
                    '�������荞��
                    buf = buf & Application.ActiveSheet.Cells(rowFocus, colFocus).Text
                    
                    isFirstCol = False
                
                End If
                
                colFocus = colFocus + 1
            
            Loop While colFocus <= lastOfCol
            
            buf = buf & " |" & vbCrLf '���s�}��
        
        End If
        
        If (rowFocus = startOfRow) Then
            buf = buf & "|" & WorksheetFunction.Rept(" - |", lastOfCol - startOfCol + 1) & vbCrLf '���s�}��
            
        End If
        
        rowFocus = rowFocus + 1
    
    Loop While rowFocus <= lastOfRow
    
    SetCB buf

End Sub

'
' �Z���Q�Ɣ͈͂� UsedRange �͈͂Ɏ��܂�悤�Ƀg���~���O����
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '�w��͈͂̉E���ʒu�̎擾
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRange�̉E���ʒu�̎擾
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    '�g���~���O
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '�i�[���ďI��
    Set trimWithUsedRange = ret
    
End Function

'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>
 


