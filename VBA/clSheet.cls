VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'==================================================
'列αに設定されたIDと関連するIDが列βに存在した場合、列βの値を使って列αを検索し、レコードを取得する。
'＊型番管理されている商品の後継型番をたどって、最新の商品型番を見つけるような時に利用。
'＊MSProjectの先行タスクから後続タスクを見つけるような時に利用。
'  [i]wb            対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]predCol     PredesessorのIDが書いてある列（一番左の列)
'  [i]succCol     SuccessorのIDが書いてある列
'  [i]predID      Successorを検索するPredesessorのID
'  [o]succID     検索結果
' ***Predesessor列よりもSuccessor列が右側にあることが前提(predCol < succCol)
'--------------------------------------------------
Public Function getSuccessorID(ByVal wb As Workbook, _
                                                    ByVal shName As String, _
                                                    ByVal datRowS As Long, _
                                                    ByVal predCol As Long, _
                                                    ByVal succCol As Long, _
                                                    ByVal predID As String, _
                                                    ByRef succID As String _
                                                    ) As Boolean
    Dim bRet As Boolean
    Dim dat As Variant
    Dim retID As String
    Dim prevID As String
    Dim retRow As Long
    
    prevID = predID
    retID = predID
    Do While retID <> ""
       '指定した文字が、指定列に存在した場合、その行を取得しArrayに格納。
        bRet = getRowDataVLookUp(wb, shName, datRowS, predCol, succCol, predCol, retID, dat, retRow)
        If bRet = True Then
            '2レコード以上取得できたら、レコードに重複があるためエラー
            If UBound(dat, 1) > 1 Then
                 Debug.Print "err ::: the predesessor ID is duplicated.  (" & retID & ")"
                getSuccessorID = False
                Exit Function
            End If
        
            'predesessorとsuccessorのIDが一致した場合、検索が循環してしまうのでエラー
            retID = dat(1, succCol - predCol + 1)
            If retID = prevID Then
                Debug.Print "err ::: the successor ID is same as the predesessor ID (" & retID & ")"
                getSuccessorID = False
                Exit Function
            Else
                prevID = retID
            End If
        Else
            'successorIDは設定されているが、それに対応するレコードが無い場合エラー
            Debug.Print "err ::: no record correspond with successor with ID"
            getSuccessorID = False
            Exit Function
        End If
    Loop
    
    succID = dat(1, 1)
    getSuccessorID = True
    
End Function

'==================================================
'2つの列番号をRangeを表すアルファベットに変換する
'  [i]colNoFrom     Rangeの始まり
'  [i]colNoTo         Rangeの終わり
'  [o]rngStr           Rangeを表すアルファベット
'--------------------------------------------------
Public Function colNo2Rng(ByVal colNoFrom As Long, _
                                        ByVal colNoTo As Long, _
                                        ByRef rngStr As String _
                                        ) As Boolean
    Dim colNoFromTxt As String
    Dim colNoToTxt As String
    Dim bRet1 As Boolean
    Dim bRet2 As Boolean
    
    '文字列で検索元ファルのRange取得
    bRet1 = colNo2Txt(colNoFrom, colNoFromTxt)
    bRet2 = colNo2Txt(colNoTo, colNoToTxt)
    
    If bRet1 And bRet2 Then
        If colNoFrom < colNoTo Then
            rngStr = colNoFromTxt & ":" & colNoToTxt
        Else
            rngStr = colNoToTxt & ":" & colNoFromTxt
        End If
        colNo2Rng = True
    Else
        rngStr = ""
        colNo2Rng = False
    End If
End Function

'==================================================
'列番号をアルファベットに変換する
'  [i]colNo         列番号
'  [o]colStr       列番号を表すアルファベット
'--------------------------------------------------
Public Function colNo2Txt(ByVal colNo As Long, _
                                        ByRef colStr As String _
                                        ) As Boolean
  On Error GoTo ErrHandler
  Dim strAddr As String

  strAddr = Cells(1, colNo).Address(False, False)
  colStr = left(strAddr, Len(strAddr) - 1)
  colNo2Txt = True
  Exit Function

ErrHandler:

  colStr = ""
  colNo2Txt = False
End Function
'
'==================================================
'VLOOKUP関数を用いて、別シートまたは別ファイルのデータを参照する（サイレントモード）
'  [i]main_wb               取得したデータを書き込むWorkbook
'  [i]main_shName       取得したデータを書き込むWorkbookのシート名
'  [i]main_dataRow      シート内 検索キーが書いてある列のスタート行
'  [i]main_keyCol         シート内 検索キーが書いてある列
'  [i]main_dataCol        シート内 検索結果を書き込む列
'  [i]from_filePath        参照ファイル名
'  [i]from_shName       参照ファイルのシート名
'  [i]from_keyCol         参照シート内 検索キーが書いてある列
'  [i]from_dataCol       参照シート内  値が書いてある列
'--------------------------------------------------
Public Function setDataByVlookupSilently(ByVal main_wb As Workbook, _
                                                                ByVal main_shName As String, _
                                                                ByVal main_dataRow As Long, _
                                                                ByVal main_keyCol As Long, _
                                                                ByVal main_dataCol As Long, _
                                                                ByVal from_filePath As String, _
                                                                ByVal from_shName As String, _
                                                                ByVal from_keyCol As Long, _
                                                                ByVal from_dataCol As Long _
                                                                ) As Boolean
    Dim fl As New clFiles
    Dim bRet As Boolean
    Dim from_rng As Range
    
    '検索元のファイルを開きオブジェクト取得
    Dim from_wb As Workbook
    bRet = fl.getWorkbookObj(from_filePath, from_wb)
    
    '2つの列番号をRangeを表すアルファベットに変換する
    Dim from_rng_txt As String
    bRet = colNo2Rng(from_keyCol, from_dataCol, from_rng_txt)
    
    If bRet = True Then
        'シートを開き、値を検索
        With from_wb.Worksheets(from_shName)
            '検索元シート内のRangeを取得
            Set from_rng = .Range(from_rng_txt)
            '値を設定
            bRet = setDataByVlookup(main_wb, main_shName, _
                                                    main_dataRow, main_dataCol, _
                                                    main_keyCol, from_rng, from_dataCol - from_keyCol + 1)
        End With
        
        Set from_rng = Nothing
        from_wb.Close
        Set from_wb = Nothing
    End If

    setDataByVlookupSilently = bRet
End Function

'==================================================
'VLOOKUP関数を用いて、別シートまたは別ファイルのデータを参照する
'  [i]main_wb               取得したデータを書き込むWorkbook
'  [i]main_shName       取得したデータを書き込むWorkbookのシート名
'  [i]main_dataRow      シート内 検索キーが書いてある列のスタート行
'  [i]main_dataCol        シート内 検索結果を書き込む列
'  [i]main_keyCol         シート内 検索キーが書いてある列
'  [i]from_rng             参照先のRange
'  [i]from_dataIndex       参照先のRange内で値が書いてある列
'--------------------------------------------------
Public Function setDataByVlookup(ByVal main_wb As Workbook, _
                                                    ByVal main_shName As String, _
                                                    ByVal main_dataRow As Long, _
                                                    ByVal main_dataCol As Long, _
                                                    ByVal main_keyCol As Long, _
                                                    ByVal from_rng As Range, _
                                                    ByVal from_dataIndex As Long _
                                                    ) As Boolean
    Dim i As Long
    Dim lastRow As Long
    Dim rngSet As Range
    Dim rngValue As Range
    Dim result As Variant
    
    With main_wb.Worksheets(main_shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        
        For i = main_dataRow To lastRow Step 1
            '関数設定セルを取得
            Set rngSet = .Range(.Cells(i, main_dataCol), .Cells(i, main_dataCol))
            Set rngValue = .Range(.Cells(i, main_keyCol), .Cells(i, main_keyCol))

            If rngValue.Value <> "" Then
                 On Error Resume Next
                 result = WorksheetFunction.VLookup( _
                                        rngValue.Value, _
                                        from_rng, _
                                        from_dataIndex, _
                                        False)
                 If Err <> 0 Then result = ""
                 rngSet = result
            End If
        Next i
    End With
    Set rngSet = Nothing
    Set rngValue = Nothing
    setDataByVlookup = True
End Function

'==================================================
'指定Sheetの指定列にAutoFilterをかける
'  [i]wb        対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [i]tgtFields Dictionaryオブジェクトで作った{field, Criteria}ハッシュテーブル
'  ***フィルターは、datRowS-1の行に設定されるためdatRowS>2でないといけない
'--------------------------------------------------
Public Function setFilter(ByVal wb As Workbook, _
                            ByVal shName As String, _
                            ByVal datRowS As Long, _
                            ByVal datColS As Long, _
                            ByVal datColE As Long, _
                            ByVal tgtFields As Variant _
                            ) As Boolean
    Dim lastRow As Long
    Dim datRng As Range
    Dim fieldKeyArr As Variant
    Dim i As Long
    
    With wb.Worksheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '最終行までのデータRangeを取得
        Set datRng = .Range(.Cells(datRowS - 1, datColS), .Cells(lastRow, datColE))
        'フィルター設定を解除
        datRng.AutoFilter
        'ハッシュテーブルのキーの配列を取得
        fieldKeyArr = tgtFields.keys
        'フィルター設定
        For i = 0 To tgtFields.count - 1
            datRng.AutoFilter field:=fieldKeyArr(i), Criteria1:=tgtFields(fieldKeyArr(i))
        Next i
    End With
    
    setFilter = True
End Function

'==================================================
'指定した文字が、指定列に存在した場合、その行を取得しArrayに格納
'シートの有無チェックは呼び出し側でやっておく
'  [i]wb        対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [i]col       検索対象列(データ領域の最初の列を検索したい場合、datColSを設定する。1ではない)
'  [i]val       検索文字列
'  [o]dat       取得したデータ
'  [o]row       最終行数
'--------------------------------------------------
Public Function getRowDataVLookUp(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal datColS As Long, _
                                    ByVal datColE As Long, _
                                    ByVal col As Long, _
                                    ByVal val As String, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    Dim lastRow As Long
    Dim rngCol As Range
    Dim rngTmp As Range
    Dim firstAddress As String
    Dim retTmpBucket As Variant
    ReDim retTmpBucket(1 To MAX_ROW, 1 To datColE - datColS + 1)
    Dim retTmp As Variant
    Dim tgtRng As Range
    Dim datArr As New clDatArr
    Dim lastIndex As Long
    Dim bRet As Boolean
    
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '最終行までの1列分のデータを取得
        Set rngCol = .Range(.Cells(datRowS, col), .Cells(lastRow, col))
        '検索する
        '***.Findは最初のRangeを基点にして次のRangeから検索してしまう。
        '***.Findの検索基点を最後のRangeにする。
        Set rngTmp = rngCol.Find(What:=val, LookIn:=xlValues, LookAt:=xlPart, _
                                After:=.Range(.Cells(lastRow, col), .Cells(lastRow, col)))
        '検索できた場合
        If Not rngTmp Is Nothing Then
            '最初に見つかったセルのアドレスを取得
            firstAddress = rngTmp.Address
            Dim i As Long
            i = 1
            Do
                '検索で一致した行のデータを取得
                retTmp = .Range(.Cells(rngTmp.row, datColS), _
                                .Cells(rngTmp.row, datColE)).Value
                'retTmpBucketに格納
                bRet = datArr.addArray(retTmp, i, retTmpBucket, lastIndex)

                '次を検索
                Set rngTmp = rngCol.FindNext(rngTmp)
                i = i + 1
            Loop Until rngTmp.Address = firstAddress '最初に戻ったら終了
        Else
            getRowDataVLookUp = False
            Exit Function
        End If
    End With
        
    '返り値配列を作成
    Dim recNum As Long
    Dim ret As Variant
    Dim newCol As Long
    recNum = i - 1
    newCol = datColE - datColS + 1
    bRet = datArr.formatArray(retTmpBucket, recNum, newCol, ret)
    dat = ret
    row = recNum
    getRowDataVLookUp = True
End Function


'==================================================
'指定列の最後の行までのデータを削除
'  [i]wb                対象のWorkbook
'  [i]shName            シート名
'  [i]datRowS           データ領域のスタート行
'  [i]col               ターゲットの列
'--------------------------------------------------
Public Function deleteColData(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long _
                                    ) As Boolean
    '指定列の全データ取得
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        'データ削除
        .Range(.Cells(datRowS, col), .Cells(lastRow, col)).Clear
    End With
    
    deleteColData = True
End Function

'==================================================
'指定列の行数を取得
'※指定列の一番下(MAX_ROW)から検索して、値がある行をの数を返す。スタート行まで値が無い場合、スタート行を返す。
'  [i]wb                対象のWorkbook
'  [i]shName            シート名
'  [i]datRowS           データ領域のスタート行
'  [i]col               ターゲットの列
'  [o]row               最終行
'--------------------------------------------------
Public Function getLastRow(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long, _
                                    ByRef row As Long _
                                    ) As Boolean
    '指定列の全データ取得
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .Range(.Cells(MAX_ROW, col), .Cells(MAX_ROW, col)).End(xlUp).row
    End With
    
    If lastRow > datRowS Then
        row = lastRow
    Else
        row = datRowS
    End If
    getLastRow = True
End Function

'==================================================
'指定列の最後の行までのデータを取得しArrayに格納
'シートの有無チェックは呼び出し側でやっておく
'  [i]wb                対象のWorkbook
'  [i]shName            シート名
'  [i]datRowS           データ領域のスタート行
'  [i]col               ターゲットの列
'  [i]allowDuplicate    重複を許すかどうか(true=許す)
'  [o]dat               取得データ
'  [o]row               取得データの行数
'--------------------------------------------------
Public Function getColDataAsArray(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long, _
                                    ByVal allowDuplicate As Boolean, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    '指定列の全データ取得
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        'データ取得
        ret = .Range(.Cells(datRowS, col), .Cells(lastRow, col)).Value
    End With
    recNum = lastRow - datRowS + 1
    
    
    '重複削除処理
    Dim retTmp(1 To MAX_ROW, 1 To 1) As Variant
    Dim datArr As New clDatArr
    Dim bRet As Boolean
    If allowDuplicate = False Then
        Dim dic As Variant
        Dim i As Long
        Dim j As Long

        j = 1
        Set dic = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(ret) Step 1
            'dic配列の中に未登録の場合、追加する。
            If Not dic.Exists(ret(i, 1)) Then
                dic.Add ret(i, 1), ret(i, 1)
                retTmp(j, 1) = ret(i, 1)
                j = j + 1
            End If
        Next i
        
        '返り値配列を作成
        recNum = j - 1
        bRet = datArr.formatArray(retTmp, recNum, 1, ret)
    End If
    
    '返り値設定
    If lastRow - datRowS > 0 Then
        dat = ret
        row = recNum
        getColDataAsArray = True
    Else
        getColDataAsArray = False
    End If
End Function

'==================================================
'Sheet内のデータ領域をArrayに格納する
'シートの有無チェックは呼び出し側でやっておく
'  [i]wb        対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]datRowE   データ領域の最終行「0」の場合シートの最後の行まで検索する。
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [o]dat       取得したデータ
'  [o]row       最終行数
'  [o]col       列数
'--------------------------------------------------
Public Function getDataAsArray(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal datRowE As Long, _
                                    ByVal datColS As Long, _
                                    ByVal datColE As Long, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long, _
                                    ByRef col As Long) As Boolean
    Dim lastRow As Long
    Dim ret As Variant
    With wb.Sheets(shName)
        '最終行取得
        If datRowE = 0 Then
            lastRow = .UsedRange.Cells(.UsedRange.count).row
        Else
            lastRow = datRowE
        End If
        'データ取得
        ret = .Range(.Cells(datRowS, datColS), .Cells(lastRow, datColE)).Value
    End With
    
    '返り値設定
    If lastRow - datRowS >= 0 Then
        dat = ret
        col = datColE - datColS + 1
        row = lastRow - datRowS + 1
        getDataAsArray = True
    Else
        getDataAsArray = False
    End If
End Function

'==================================================
'名前を指定してSheetを作成。
'同名のシートが存在した場合その中身を削除する。
'  [i]wb        対象のWorkbook
'  [i]shName    作成するシート名
'--------------------------------------------------
Public Function initSheet(ByVal wb As Workbook, _
                            ByVal shName As String) As Boolean
                            
    If Not existSheet(wb, shName) Then
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.count)
        ActiveSheet.Name = shName
    Else
        With wb.Sheets(shName)
            Dim verLastRow As Long
            Dim delRow As String
            verLastRow = .UsedRange.Cells(.UsedRange.count).row + 1
            delRow = "1:" & verLastRow
            '行削除
            .Range(delRow).Delete
            'すべてのコントロールを削除
            .DrawingObjects.Delete
        End With
    End If

End Function

'==================================================
'名前を指定して既存Sheetをコピー。
'同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。
'  [i]wb        対象のWorkbook
'  [i]baseSh    コピー元のシート
'  [i]shName    作成するシート名
'  [o]newSheet  作成された新しいシートの名前
'--------------------------------------------------
Public Function newCopyedSheet(ByVal wb As Workbook, _
                            ByVal baseSh As Worksheet, _
                            ByVal shName As String, _
                            ByRef newSheetName As String _
                            ) As Boolean

    Dim xlsheet As Worksheet
    Dim num As Integer
    Dim flg As Boolean
    Dim newName As String
    
    newName = shName
    num = 1
    flg = True
    Do While flg
        If Not existSheet(wb, newName) Then
            flg = False
        Else
            newName = shName & " " & Format(num, "(#)")
            num = num + 1
        End If
    Loop
    
    baseSh.Copy After:=wb.Worksheets(Worksheets.count)
    ActiveSheet.Name = newName
    newSheetName = newName
    newCopyedSheet = True
    
End Function



'==================================================
'名前を指定してSheetを作成。
'同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。
'  [i]wb        対象のWorkbook
'  [i]shName    作成するシート名
'  [o]newSheet  作成された新しいシートの名前
'--------------------------------------------------
Public Function newSheet(ByVal wb As Workbook, _
                            ByVal shName As String, _
                            ByRef newSheetName As String _
                            ) As Boolean

    Dim xlsheet As Worksheet
    Dim num As Integer
    Dim flg As Boolean
    Dim newName As String
    
    newName = shName
    num = 1
    flg = True
    Do While flg
        If Not existSheet(wb, newName) Then
            flg = False
        Else
            newName = shName & " " & Format(num, "(#)")
            num = num + 1
        End If
    Loop
    
    wb.Worksheets.Add After:=wb.Worksheets(Worksheets.count)
    ActiveSheet.Name = newName
    newSheetName = newName
    newSheet = True
    
End Function

'==================================================
'名前を指定してSheetの有無をチェック
'  [i]wb        対象のWorkbook
'  [i]shName    確認するシート名
'--------------------------------------------------
Public Function existSheet(ByVal wb As Workbook, _
                            ByVal shName As String) As Boolean
    Dim xlsheet As Worksheet
    On Error Resume Next
    Set xlsheet = wb.Worksheets(shName)
    On Error GoTo 0
    If xlsheet Is Nothing Then
        existSheet = False
    Else
        existSheet = True
    End If
    Set xlsheet = Nothing
End Function

'==================================================
'指定したRange範囲内にある図形,ActiveXコントロールを削除する
'  [i]wb            対象のWorkbook
'  [i]shName    対象のシート名
'  [i]rowUL       削除対象範囲 左上行番号
'  [i]colUL        削除対象範囲 左上列番号
'  [i]rowLR       削除対象範囲 右下行番号
'  [i]colLR        削除対象範囲 右下列番号
'--------------------------------------------------
Public Function deleteObjectInRange(ByVal wb As Workbook, _
                                                        ByVal shName As String, _
                                                        ByVal rowUL As Long, _
                                                        ByVal colUL As Long, _
                                                        ByVal rowLR As Long, _
                                                        ByVal colLR As Long _
                                                        ) As Boolean
    Dim shp As Shape
    Dim rng_shp As Range
    Dim rng As Range
 
    With wb.Worksheets(shName)
        .Select
        'Rangeをセレクト
        Set rng = .Range(.Cells(rowUL, colUL), .Cells(rowLR, colLR))
    End With
 
    'すべての図形にループ処理
    For Each shp In wb.Worksheets(shName).Shapes
        '図形の配置されているセル範囲をオブジェクト変数にセット
        Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)
 
        '図形の配置されているセル範囲と選択されているセル範囲が重なっていれば図形を削除
        If Not Intersect(rng_shp, rng) Is Nothing Then
            shp.Delete
        End If
    Next
    Set rng = Nothing
    deleteObjectInRange = True
End Function
