Attribute VB_Name = "ModuleCreateInsert"
'************************************************************************************
'関数名          :Insert文作成
'
'
'************************************************************************************
Sub Insert文作成()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim iOutRow As Long: iOutRow = 6
    Dim iRow As Long: iRow = 0
    Dim iCol As Long: iCol = 0
    Dim sInsertSql As String
    Dim sNull As String: sNull = "NULL"
    Dim isheetNum As Integer: isheetNum = 0
    Dim isheetCnt As Integer: isheetCnt = 0
    Dim sVal As String

    
    Dim wTBLsheet As Worksheet
    Dim iTBLLastRow As Long
    Dim iTBLLastCol As Long
    
    'ワークシート数の取得
    isheetCnt = Worksheets.Count
    
    For isheetNum = Cells(12, 2).Value + 1 To isheetCnt
        Set wTBLsheet = Sheets(isheetNum)
        iOutRow = 6
        
        'シートの最終行、列の取得
        iTBLLastRow = wTBLsheet.Cells(Rows.Count, 3).End(xlUp).Row
        iTBLLastCol = wTBLsheet.Cells(4, Columns.Count).End(xlToLeft).Column
     
        'Insert文列名の取得
        Cells(7, 2).Copy
        wTBLsheet.Cells(4, iTBLLastCol).PasteSpecial (xlPasteFormats)
        wTBLsheet.Cells(4, iTBLLastCol).Value = wTBLsheet.Cells(3, 3).Value & vbLf & "生成Insert文"
               
        '書式のコピー
        Cells(9, 2).Copy
        
        For iRow = 6 To iTBLLastRow
            '書式の貼付け
            wTBLsheet.Cells(iOutRow, iTBLLastCol).PasteSpecial (xlPasteFormats)
            
            sInsertSql = "INSERT INTO " & wTBLsheet.Cells(3, 3).Value & " VALUES ("
            
            For iCol = 3 To iTBLLastCol - 1
                sType = wTBLsheet.Cells(5, iCol).Value
                sVal = wTBLsheet.Cells(iRow, iCol).Value
                
                '空白が設定されている場合
                If sVal = "" Then
                    sInsertSql = sInsertSql & "''" & ","
                    GoTo NextLoop1
                End If
                'NULLが設定されている場合
                If UCase(sVal) = sNull Then
                    sInsertSql = sInsertSql & "NULL" & ","
                    GoTo NextLoop1
                End If
                'defaultが設定されている場合
                If UCase(sVal) = "DEFAULT" Then
                    sInsertSql = sInsertSql & "default" & ","
                    GoTo NextLoop1
                End If

                'INTの場合
                If UCase(sType) = "INT" Then
                    sInsertSql = sInsertSql & sVal & ","
                'BOOLEANの場合
                ElseIf UCase(sType) = "BOOLEAN" Then
                    sInsertSql = sInsertSql & sVal & ","
                'VARCHARの場合
                ElseIf UCase(sType) = "VARCHAR" Then
                    sInsertSql = sInsertSql & "'" & sVal & "',"
                End If
NextLoop1:
            Next
            '末尾の「,」を削除
            sInsertSql = Left(sInsertSql, Len(sInsertSql) - 1)
            wTBLsheet.Cells(iOutRow, iTBLLastCol).Value = sInsertSql & ");"
            iOutRow = iOutRow + 1
        Next
        'ハイパーリンクの設定
        ActiveSheet.Cells(isheetNum + 6, 4).Hyperlinks.Add Anchor:=ActiveSheet.Cells(isheetNum - Cells(12, 2).Value + 7, 4), _
                                                                    Address:="", _
                                                                    SubAddress:=wTBLsheet.Cells(3, 3).Value & "!" & Split(wTBLsheet.Cells(iTBLLastCol).Address, "$")(1) & "6", _
                                                                    TextToDisplay:=wTBLsheet.Cells(3, 3).Value
    Next
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Insert文生成が完了しました。"
    
            
End Sub
