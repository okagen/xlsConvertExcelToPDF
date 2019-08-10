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
'�񃿂ɐݒ肳�ꂽID�Ɗ֘A����ID������ɑ��݂����ꍇ�A����̒l���g���ė񃿂��������A���R�[�h���擾����B
'���^�ԊǗ�����Ă��鏤�i�̌�p�^�Ԃ����ǂ��āA�ŐV�̏��i�^�Ԃ�������悤�Ȏ��ɗ��p�B
'��MSProject�̐�s�^�X�N����㑱�^�X�N��������悤�Ȏ��ɗ��p�B
'  [i]wb            �Ώۂ�Workbook
'  [i]shName    �V�[�g��
'  [i]datRowS   �f�[�^�̈�̃X�^�[�g�s
'  [i]predCol     Predesessor��ID�������Ă����i��ԍ��̗�)
'  [i]succCol     Successor��ID�������Ă����
'  [i]predID      Successor����������Predesessor��ID
'  [o]succID     ��������
' ***Predesessor�����Successor�񂪉E���ɂ��邱�Ƃ��O��(predCol < succCol)
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
       '�w�肵���������A�w���ɑ��݂����ꍇ�A���̍s���擾��Array�Ɋi�[�B
        bRet = getRowDataVLookUp(wb, shName, datRowS, predCol, succCol, predCol, retID, dat, retRow)
        If bRet = True Then
            '2���R�[�h�ȏ�擾�ł�����A���R�[�h�ɏd�������邽�߃G���[
            If UBound(dat, 1) > 1 Then
                 Debug.Print "err ::: the predesessor ID is duplicated.  (" & retID & ")"
                getSuccessorID = False
                Exit Function
            End If
        
            'predesessor��successor��ID����v�����ꍇ�A�������z���Ă��܂��̂ŃG���[
            retID = dat(1, succCol - predCol + 1)
            If retID = prevID Then
                Debug.Print "err ::: the successor ID is same as the predesessor ID (" & retID & ")"
                getSuccessorID = False
                Exit Function
            Else
                prevID = retID
            End If
        Else
            'successorID�͐ݒ肳��Ă��邪�A����ɑΉ����郌�R�[�h�������ꍇ�G���[
            Debug.Print "err ::: no record correspond with successor with ID"
            getSuccessorID = False
            Exit Function
        End If
    Loop
    
    succID = dat(1, 1)
    getSuccessorID = True
    
End Function

'==================================================
'2�̗�ԍ���Range��\���A���t�@�x�b�g�ɕϊ�����
'  [i]colNoFrom     Range�̎n�܂�
'  [i]colNoTo         Range�̏I���
'  [o]rngStr           Range��\���A���t�@�x�b�g
'--------------------------------------------------
Public Function colNo2Rng(ByVal colNoFrom As Long, _
                                        ByVal colNoTo As Long, _
                                        ByRef rngStr As String _
                                        ) As Boolean
    Dim colNoFromTxt As String
    Dim colNoToTxt As String
    Dim bRet1 As Boolean
    Dim bRet2 As Boolean
    
    '������Ō������t�@����Range�擾
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
'��ԍ����A���t�@�x�b�g�ɕϊ�����
'  [i]colNo         ��ԍ�
'  [o]colStr       ��ԍ���\���A���t�@�x�b�g
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
'VLOOKUP�֐���p���āA�ʃV�[�g�܂��͕ʃt�@�C���̃f�[�^���Q�Ƃ���i�T�C�����g���[�h�j
'  [i]main_wb               �擾�����f�[�^����������Workbook
'  [i]main_shName       �擾�����f�[�^����������Workbook�̃V�[�g��
'  [i]main_dataRow      �V�[�g�� �����L�[�������Ă����̃X�^�[�g�s
'  [i]main_keyCol         �V�[�g�� �����L�[�������Ă����
'  [i]main_dataCol        �V�[�g�� �������ʂ��������ޗ�
'  [i]from_filePath        �Q�ƃt�@�C����
'  [i]from_shName       �Q�ƃt�@�C���̃V�[�g��
'  [i]from_keyCol         �Q�ƃV�[�g�� �����L�[�������Ă����
'  [i]from_dataCol       �Q�ƃV�[�g��  �l�������Ă����
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
    
    '�������̃t�@�C�����J���I�u�W�F�N�g�擾
    Dim from_wb As Workbook
    bRet = fl.getWorkbookObj(from_filePath, from_wb)
    
    '2�̗�ԍ���Range��\���A���t�@�x�b�g�ɕϊ�����
    Dim from_rng_txt As String
    bRet = colNo2Rng(from_keyCol, from_dataCol, from_rng_txt)
    
    If bRet = True Then
        '�V�[�g���J���A�l������
        With from_wb.Worksheets(from_shName)
            '�������V�[�g����Range���擾
            Set from_rng = .Range(from_rng_txt)
            '�l��ݒ�
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
'VLOOKUP�֐���p���āA�ʃV�[�g�܂��͕ʃt�@�C���̃f�[�^���Q�Ƃ���
'  [i]main_wb               �擾�����f�[�^����������Workbook
'  [i]main_shName       �擾�����f�[�^����������Workbook�̃V�[�g��
'  [i]main_dataRow      �V�[�g�� �����L�[�������Ă����̃X�^�[�g�s
'  [i]main_dataCol        �V�[�g�� �������ʂ��������ޗ�
'  [i]main_keyCol         �V�[�g�� �����L�[�������Ă����
'  [i]from_rng             �Q�Ɛ��Range
'  [i]from_dataIndex       �Q�Ɛ��Range���Œl�������Ă����
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
        '�ŏI�s�擾
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        
        For i = main_dataRow To lastRow Step 1
            '�֐��ݒ�Z�����擾
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
'�w��Sheet�̎w����AutoFilter��������
'  [i]wb        �Ώۂ�Workbook
'  [i]shName    �V�[�g��
'  [i]datRowS   �f�[�^�̈�̃X�^�[�g�s
'  [i]datColS   �f�[�^�̈�̍ŏ��̗�i��ԍ��̗�)
'  [i]datColE   �f�[�^�̈�̍Ō�̗�i��ԉE�̗�)
'  [i]tgtFields Dictionary�I�u�W�F�N�g�ō����{field, Criteria}�n�b�V���e�[�u��
'  ***�t�B���^�[�́AdatRowS-1�̍s�ɐݒ肳��邽��datRowS>2�łȂ��Ƃ����Ȃ�
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
        '�ŏI�s�擾
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '�ŏI�s�܂ł̃f�[�^Range���擾
        Set datRng = .Range(.Cells(datRowS - 1, datColS), .Cells(lastRow, datColE))
        '�t�B���^�[�ݒ������
        datRng.AutoFilter
        '�n�b�V���e�[�u���̃L�[�̔z����擾
        fieldKeyArr = tgtFields.keys
        '�t�B���^�[�ݒ�
        For i = 0 To tgtFields.count - 1
            datRng.AutoFilter field:=fieldKeyArr(i), Criteria1:=tgtFields(fieldKeyArr(i))
        Next i
    End With
    
    setFilter = True
End Function

'==================================================
'�w�肵���������A�w���ɑ��݂����ꍇ�A���̍s���擾��Array�Ɋi�[
'�V�[�g�̗L���`�F�b�N�͌Ăяo�����ł���Ă���
'  [i]wb        �Ώۂ�Workbook
'  [i]shName    �V�[�g��
'  [i]datRowS   �f�[�^�̈�̃X�^�[�g�s
'  [i]datColS   �f�[�^�̈�̍ŏ��̗�i��ԍ��̗�)
'  [i]datColE   �f�[�^�̈�̍Ō�̗�i��ԉE�̗�)
'  [i]col       �����Ώۗ�(�f�[�^�̈�̍ŏ��̗�������������ꍇ�AdatColS��ݒ肷��B1�ł͂Ȃ�)
'  [i]val       ����������
'  [o]dat       �擾�����f�[�^
'  [o]row       �ŏI�s��
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
        '�ŏI�s�擾
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '�ŏI�s�܂ł�1�񕪂̃f�[�^���擾
        Set rngCol = .Range(.Cells(datRowS, col), .Cells(lastRow, col))
        '��������
        '***.Find�͍ŏ���Range����_�ɂ��Ď���Range���猟�����Ă��܂��B
        '***.Find�̌�����_���Ō��Range�ɂ���B
        Set rngTmp = rngCol.Find(What:=val, LookIn:=xlValues, LookAt:=xlPart, _
                                After:=.Range(.Cells(lastRow, col), .Cells(lastRow, col)))
        '�����ł����ꍇ
        If Not rngTmp Is Nothing Then
            '�ŏ��Ɍ��������Z���̃A�h���X���擾
            firstAddress = rngTmp.Address
            Dim i As Long
            i = 1
            Do
                '�����ň�v�����s�̃f�[�^���擾
                retTmp = .Range(.Cells(rngTmp.row, datColS), _
                                .Cells(rngTmp.row, datColE)).Value
                'retTmpBucket�Ɋi�[
                bRet = datArr.addArray(retTmp, i, retTmpBucket, lastIndex)

                '��������
                Set rngTmp = rngCol.FindNext(rngTmp)
                i = i + 1
            Loop Until rngTmp.Address = firstAddress '�ŏ��ɖ߂�����I��
        Else
            getRowDataVLookUp = False
            Exit Function
        End If
    End With
        
    '�Ԃ�l�z����쐬
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
'�w���̍Ō�̍s�܂ł̃f�[�^���폜
'  [i]wb                �Ώۂ�Workbook
'  [i]shName            �V�[�g��
'  [i]datRowS           �f�[�^�̈�̃X�^�[�g�s
'  [i]col               �^�[�Q�b�g�̗�
'--------------------------------------------------
Public Function deleteColData(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long _
                                    ) As Boolean
    '�w���̑S�f�[�^�擾
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '�ŏI�s�擾
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '�f�[�^�폜
        .Range(.Cells(datRowS, col), .Cells(lastRow, col)).Clear
    End With
    
    deleteColData = True
End Function

'==================================================
'�w���̍s�����擾
'���w���̈�ԉ�(MAX_ROW)���猟�����āA�l������s���̐���Ԃ��B�X�^�[�g�s�܂Œl�������ꍇ�A�X�^�[�g�s��Ԃ��B
'  [i]wb                �Ώۂ�Workbook
'  [i]shName            �V�[�g��
'  [i]datRowS           �f�[�^�̈�̃X�^�[�g�s
'  [i]col               �^�[�Q�b�g�̗�
'  [o]row               �ŏI�s
'--------------------------------------------------
Public Function getLastRow(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long, _
                                    ByRef row As Long _
                                    ) As Boolean
    '�w���̑S�f�[�^�擾
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '�ŏI�s�擾
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
'�w���̍Ō�̍s�܂ł̃f�[�^���擾��Array�Ɋi�[
'�V�[�g�̗L���`�F�b�N�͌Ăяo�����ł���Ă���
'  [i]wb                �Ώۂ�Workbook
'  [i]shName            �V�[�g��
'  [i]datRowS           �f�[�^�̈�̃X�^�[�g�s
'  [i]col               �^�[�Q�b�g�̗�
'  [i]allowDuplicate    �d�����������ǂ���(true=����)
'  [o]dat               �擾�f�[�^
'  [o]row               �擾�f�[�^�̍s��
'--------------------------------------------------
Public Function getColDataAsArray(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long, _
                                    ByVal allowDuplicate As Boolean, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    '�w���̑S�f�[�^�擾
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '�ŏI�s�擾
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '�f�[�^�擾
        ret = .Range(.Cells(datRowS, col), .Cells(lastRow, col)).Value
    End With
    recNum = lastRow - datRowS + 1
    
    
    '�d���폜����
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
            'dic�z��̒��ɖ��o�^�̏ꍇ�A�ǉ�����B
            If Not dic.Exists(ret(i, 1)) Then
                dic.Add ret(i, 1), ret(i, 1)
                retTmp(j, 1) = ret(i, 1)
                j = j + 1
            End If
        Next i
        
        '�Ԃ�l�z����쐬
        recNum = j - 1
        bRet = datArr.formatArray(retTmp, recNum, 1, ret)
    End If
    
    '�Ԃ�l�ݒ�
    If lastRow - datRowS > 0 Then
        dat = ret
        row = recNum
        getColDataAsArray = True
    Else
        getColDataAsArray = False
    End If
End Function

'==================================================
'Sheet���̃f�[�^�̈��Array�Ɋi�[����
'�V�[�g�̗L���`�F�b�N�͌Ăяo�����ł���Ă���
'  [i]wb        �Ώۂ�Workbook
'  [i]shName    �V�[�g��
'  [i]datRowS   �f�[�^�̈�̃X�^�[�g�s
'  [i]datRowE   �f�[�^�̈�̍ŏI�s�u0�v�̏ꍇ�V�[�g�̍Ō�̍s�܂Ō�������B
'  [i]datColS   �f�[�^�̈�̍ŏ��̗�i��ԍ��̗�)
'  [i]datColE   �f�[�^�̈�̍Ō�̗�i��ԉE�̗�)
'  [o]dat       �擾�����f�[�^
'  [o]row       �ŏI�s��
'  [o]col       ��
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
        '�ŏI�s�擾
        If datRowE = 0 Then
            lastRow = .UsedRange.Cells(.UsedRange.count).row
        Else
            lastRow = datRowE
        End If
        '�f�[�^�擾
        ret = .Range(.Cells(datRowS, datColS), .Cells(lastRow, datColE)).Value
    End With
    
    '�Ԃ�l�ݒ�
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
'���O���w�肵��Sheet���쐬�B
'�����̃V�[�g�����݂����ꍇ���̒��g���폜����B
'  [i]wb        �Ώۂ�Workbook
'  [i]shName    �쐬����V�[�g��
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
            '�s�폜
            .Range(delRow).Delete
            '���ׂẴR���g���[�����폜
            .DrawingObjects.Delete
        End With
    End If

End Function

'==================================================
'���O���w�肵�Ċ���Sheet���R�s�[�B
'�����̃V�[�g�����݂����ꍇ�A�V�[�g��������(#)��t���ăJ�E���g�A�b�v�B
'  [i]wb        �Ώۂ�Workbook
'  [i]baseSh    �R�s�[���̃V�[�g
'  [i]shName    �쐬����V�[�g��
'  [o]newSheet  �쐬���ꂽ�V�����V�[�g�̖��O
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
'���O���w�肵��Sheet���쐬�B
'�����̃V�[�g�����݂����ꍇ�A�V�[�g��������(#)��t���ăJ�E���g�A�b�v�B
'  [i]wb        �Ώۂ�Workbook
'  [i]shName    �쐬����V�[�g��
'  [o]newSheet  �쐬���ꂽ�V�����V�[�g�̖��O
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
'���O���w�肵��Sheet�̗L�����`�F�b�N
'  [i]wb        �Ώۂ�Workbook
'  [i]shName    �m�F����V�[�g��
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
'�w�肵��Range�͈͓��ɂ���}�`,ActiveX�R���g���[�����폜����
'  [i]wb            �Ώۂ�Workbook
'  [i]shName    �Ώۂ̃V�[�g��
'  [i]rowUL       �폜�Ώ۔͈� ����s�ԍ�
'  [i]colUL        �폜�Ώ۔͈� �����ԍ�
'  [i]rowLR       �폜�Ώ۔͈� �E���s�ԍ�
'  [i]colLR        �폜�Ώ۔͈� �E����ԍ�
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
        'Range���Z���N�g
        Set rng = .Range(.Cells(rowUL, colUL), .Cells(rowLR, colLR))
    End With
 
    '���ׂĂ̐}�`�Ƀ��[�v����
    For Each shp In wb.Worksheets(shName).Shapes
        '�}�`�̔z�u����Ă���Z���͈͂��I�u�W�F�N�g�ϐ��ɃZ�b�g
        Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)
 
        '�}�`�̔z�u����Ă���Z���͈͂ƑI������Ă���Z���͈͂��d�Ȃ��Ă���ΐ}�`���폜
        If Not Intersect(rng_shp, rng) Is Nothing Then
            shp.Delete
        End If
    Next
    Set rng = Nothing
    deleteObjectInRange = True
End Function