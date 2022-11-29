Attribute VB_Name = "ModuleCodeCheck"
'************************************************************************************
'関数名          :チェック結果
'
'
'************************************************************************************
Sub チェック結果()
Attribute チェック結果.VB_ProcData.VB_Invoke_Func = "M\n14"

    Dim strFileName As String
    Dim imax, i, j, jmax As Long
    Dim str As String
    Dim FSO As FileSystemObject
    Dim Txt As TextStream
    
    
    Application.ScreenUpdating = False
    
    For j = 5 To Sheet1.Cells(Rows.Count, 2).End(xlUp).Row
        strFileName = Sheet1.Cells(j, 2).Value
        If Dir(strFileName) <> "" Then
            'オブジェクト作成
            Set FSO = CreateObject("Scripting.FileSystemObject")
            'FSO.Charset = "UTF-8"
            Set Txt = FSO.OpenTextFile(strFileName, ForReading)
            str = Txt.ReadAll
            Txt.Close
            
            For i = 3 To 17
                If (Sheet1.Cells(4, i).Value) <> "" Then
                    If Sheet1.Cells(4, i).Value = "タブ" Then
                        If InStr(str, Chr(9)) > 0 Then
                            Sheet1.Cells(j, i).Value = "●"
                        End If
                    Else
                        If InStr(str, Sheet1.Cells(4, i).Value) > 0 Then
                            Sheet1.Cells(j, i).Value = "●"
                        End If
                    End If
                End If
            Next i
        End If
    Next j
            
    Application.ScreenUpdating = True
    MsgBox "終了しました"
    Exit Sub
            
End Sub

