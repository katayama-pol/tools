Attribute VB_Name = "ModuleCreateInsert"
'************************************************************************************
'�֐���          :Insert���쐬
'
'
'************************************************************************************
Sub Insert���쐬()

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
    
    '���[�N�V�[�g���̎擾
    isheetCnt = Worksheets.Count
    
    For isheetNum = Cells(12, 2).Value + 1 To isheetCnt
        Set wTBLsheet = Sheets(isheetNum)
        iOutRow = 6
        
        '�V�[�g�̍ŏI�s�A��̎擾
        iTBLLastRow = wTBLsheet.Cells(Rows.Count, 3).End(xlUp).Row
        iTBLLastCol = wTBLsheet.Cells(4, Columns.Count).End(xlToLeft).Column
     
        'Insert���񖼂̎擾
        Cells(7, 2).Copy
        wTBLsheet.Cells(4, iTBLLastCol).PasteSpecial (xlPasteFormats)
        wTBLsheet.Cells(4, iTBLLastCol).Value = wTBLsheet.Cells(3, 3).Value & vbLf & "����Insert��"
               
        '�����̃R�s�[
        Cells(9, 2).Copy
        
        For iRow = 6 To iTBLLastRow
            '�����̓\�t��
            wTBLsheet.Cells(iOutRow, iTBLLastCol).PasteSpecial (xlPasteFormats)
            
            sInsertSql = "INSERT INTO " & wTBLsheet.Cells(3, 3).Value & " VALUES ("
            
            For iCol = 3 To iTBLLastCol - 1
                sType = wTBLsheet.Cells(5, iCol).Value
                sVal = wTBLsheet.Cells(iRow, iCol).Value
                
                '�󔒂��ݒ肳��Ă���ꍇ
                If sVal = "" Then
                    sInsertSql = sInsertSql & "''" & ","
                    GoTo NextLoop1
                End If
                'NULL���ݒ肳��Ă���ꍇ
                If UCase(sVal) = sNull Then
                    sInsertSql = sInsertSql & "NULL" & ","
                    GoTo NextLoop1
                End If
                'default���ݒ肳��Ă���ꍇ
                If UCase(sVal) = "DEFAULT" Then
                    sInsertSql = sInsertSql & "default" & ","
                    GoTo NextLoop1
                End If

                'INT�̏ꍇ
                If UCase(sType) = "INT" Then
                    sInsertSql = sInsertSql & sVal & ","
                'BOOLEAN�̏ꍇ
                ElseIf UCase(sType) = "BOOLEAN" Then
                    sInsertSql = sInsertSql & sVal & ","
                'VARCHAR�̏ꍇ
                ElseIf UCase(sType) = "VARCHAR" Then
                    sInsertSql = sInsertSql & "'" & sVal & "',"
                End If
NextLoop1:
            Next
            '�����́u,�v���폜
            sInsertSql = Left(sInsertSql, Len(sInsertSql) - 1)
            wTBLsheet.Cells(iOutRow, iTBLLastCol).Value = sInsertSql & ");"
            iOutRow = iOutRow + 1
        Next
        '�n�C�p�[�����N�̐ݒ�
        ActiveSheet.Cells(isheetNum + 6, 4).Hyperlinks.Add Anchor:=ActiveSheet.Cells(isheetNum - Cells(12, 2).Value + 7, 4), _
                                                                    Address:="", _
                                                                    SubAddress:=wTBLsheet.Cells(3, 3).Value & "!" & Split(wTBLsheet.Cells(iTBLLastCol).Address, "$")(1) & "6", _
                                                                    TextToDisplay:=wTBLsheet.Cells(3, 3).Value
    Next
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Insert���������������܂����B"
    
            
End Sub
