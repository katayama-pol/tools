Attribute VB_Name = "ModuleCreateScript"
'************************************************************************************
'�֐���          :�X�N���v�g�t�@�C���쐬
'
'
'************************************************************************************
Sub �X�N���v�g�t�@�C���쐬()
    
    Dim fso As Object
    Dim folderObj As Object
    Dim tso As Object
   
    Dim iCnt As Long: iCnt = 0
    Dim sFileName As String
    Dim sBookPath As String
    Dim iOutRow As Long: iOutRow = 6
    Dim isheetNum As Integer: isheetNum = 0
    Dim isheetCnt As Integer: isheetCnt = 0
    Dim sLastColAddr As String
    Dim rVisibleRange As Range
    Dim rVisibleCell As Range

    Dim wTBLsheet As Worksheet
    Dim iTBLLastRow As Long
    Dim iTBLLastCol As Long

    '���[�N�V�[�g���̎擾
    isheetCnt = Worksheets.Count
    '�p�X�̎擾
    sBookPath = ThisWorkbook.Path
    
    '�t�@�C���V�X�e���I�u�W�F�N�g�̐���
    Set fso = CreateObject("Scripting.FileSystemObject")
    '�t�H���_�I�u�W�F�N�g�̎擾
    Set folderObj = fso.GetFolder(sBookPath & "\")
    
    For isheetNum = Cells(12, 2).Value + 1 To isheetCnt
        iCnt = iCnt + 1
        Set wTBLsheet = Sheets(isheetNum)
        
        With wTBLsheet
            '�V�[�g�̍ŏI�s�A��̎擾
            iTBLLastRow = .Cells(Rows.Count, 3).End(xlUp).Row
            iTBLLastCol = .Cells(4, Columns.Count).End(xlToLeft).Column
        
            '�t�@�C�����̐ݒ�
            sFileName = Cells(15, 2).Value & " " & .Name & ".sql"
            '�e�L�X�g�X�g���[���I�u�W�F�N�g�̎擾
            Set tso = folderObj.CreateTextFile(sFileName)
            '�V�[�g�ŏI��̃A�h���X�擾
            sLastColAddr = Split(.Cells(iTBLLastCol).Address, "$")(1)
            
            '��ʏ�\������Ă���Z���͈͎擾(�i���݌��ʎ擾)
            Set rVisibleRange = .Range(sLastColAddr & iOutRow, .Range(sLastColAddr & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible)
        End With
        
        For Each rVisibleCell In rVisibleRange
            'Insert���o��(���s����)
            tso.WriteLine rVisibleCell.Value
        Next
        
        tso.Close
    Next
    
    Set fso = Nothing
    Set folderObj = Nothing
    Set tso = Nothing

    MsgBox "�X�N���v�g�t�@�C���쐬���������܂����B" & vbLf & "�쐬�t�@�C�����F" & iCnt & "��"
    

End Sub
