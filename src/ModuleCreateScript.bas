Attribute VB_Name = "ModuleCreateScript"
'************************************************************************************
'関数名          :スクリプトファイル作成
'
'
'************************************************************************************
Sub スクリプトファイル作成()
    
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

    'ワークシート数の取得
    isheetCnt = Worksheets.Count
    'パスの取得
    sBookPath = ThisWorkbook.Path
    
    'ファイルシステムオブジェクトの生成
    Set fso = CreateObject("Scripting.FileSystemObject")
    'フォルダオブジェクトの取得
    Set folderObj = fso.GetFolder(sBookPath & "\")
    
    For isheetNum = Cells(12, 2).Value + 1 To isheetCnt
        iCnt = iCnt + 1
        Set wTBLsheet = Sheets(isheetNum)
        
        With wTBLsheet
            'シートの最終行、列の取得
            iTBLLastRow = .Cells(Rows.Count, 3).End(xlUp).Row
            iTBLLastCol = .Cells(4, Columns.Count).End(xlToLeft).Column
        
            'ファイル名の設定
            sFileName = Cells(15, 2).Value & " " & .Name & ".sql"
            'テキストストリームオブジェクトの取得
            Set tso = folderObj.CreateTextFile(sFileName)
            'シート最終列のアドレス取得
            sLastColAddr = Split(.Cells(iTBLLastCol).Address, "$")(1)
            
            '画面上表示されているセル範囲取得(絞込み結果取得)
            Set rVisibleRange = .Range(sLastColAddr & iOutRow, .Range(sLastColAddr & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible)
        End With
        
        For Each rVisibleCell In rVisibleRange
            'Insert文出力(改行あり)
            tso.WriteLine rVisibleCell.Value
        Next
        
        tso.Close
    Next
    
    Set fso = Nothing
    Set folderObj = Nothing
    Set tso = Nothing

    MsgBox "スクリプトファイル作成が完了しました。" & vbLf & "作成ファイル数：" & iCnt & "件"
    

End Sub
