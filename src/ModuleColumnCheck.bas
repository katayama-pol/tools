Attribute VB_Name = "ModuleColumnCheck"
'************************************************************************************
'関数名          :百カラム
'
'
'************************************************************************************
Sub 百カラム()
Attribute 百カラム.VB_ProcData.VB_Invoke_Func = "M\n14"

    Dim cnt, GYO As Long
    Dim strFileName As String
    Dim imax, i, j, jmax, intFF As Long
    Dim str As String
    Dim st As Object
    
    
    Application.ScreenUpdating = False
    
    Set st = CreateObject("ADODB.Stream")
    
    For j = 4 To Sheet1.Cells(Rows.Count, 2).End(xlUp).Row
        strFileName = Sheet1.Cells(j, 2).Value
        If Dir(strFileName) <> "" Then
            'オブジェクトに保存するデータの種類を文字列型に指定
            st.Type = 2
            st.Charset = "utf-8"
            '改行LF(10)
            st.LineSeparator = 10
            st.Open
            st.LoadFromFile (strFileName)
            cnt = 1
            Do While Not st.EOS
                str = st.ReadText(-2)
                Sheet2.Cells(cnt, 3).Value = str
                cnt = cnt + 1
            Loop
            
            st.Close
        End If
        
        For i = 1 To Sheet2.Cells(Rows.Count, 3).End(xlUp).Row
            If Len(Sheet2.Cells(i, 3).Value) > 101 Then
                Sheet1.Cells(j, 3).Value = "●"
                Sheet1.Cells(j, 4).Value = Sheet1.Cells(j, 4).Value & "," & i
            End If
        Next i
        
        Sheet2.Range("C:C").Clear
        
    Next j
    
    Set st = Nothing
    
            
    Application.ScreenUpdating = True
    MsgBox "終了しました"
    Exit Sub
            
End Sub

