'************************************************************************************
'*  ProgramID  ：KHPrice26
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*
'*  概要       ：ブロックマニホールド用電磁弁単体
'*             ：３ＧＡ１／３ＧＡ２／３ＧＡ３
'*             ：４ＧＡ１／４ＧＡ２／４ＧＡ３／４ＧＡ４
'*             ：３ＧＢ１／３ＧＢ２
'*             ：４ＧＢ１／４ＧＢ２／４ＧＢ３／４ＧＢ４
'*
'*【修正履歴】
'*                                      更新日：2007/05/09   更新者：NII A.Takahashi
'*  ・オプション「K」において、4GB4の場合のみの価格積上げロジック削除
'*                                      更新日：2007/09/26   更新者：NII A.Takahashi
'*  ・継手オプション追加により、継手加算ロジック追加
'*                                      更新日：2008/04/15   更新者：T.Sato
'*  ・受付No：RM0803048対応　3GA1/3GA2/4GA1/4GA2/3GB1/3GB2/4GB1/4GB2にオプションボックス追加
'*  ・受付No：RM0904031  4GD2/4GE2機種追加
'*                                      更新日：2009/06/23   更新者：Y.Miura
'*  二次電池対応                         更新日：2010/05/25   更新者：Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice60

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolC5Flag As Boolean

        Dim strKiriIchikbn As String = ""   '切換位置区分
        Dim strSosakbn As String = ""       '操作区分
        Dim strKokei As String = ""         '接続口径
        Dim strSyudoSochi As String = ""    '手動装置
        Dim strDensen As String = ""        '電線接続
        Dim strTanshi As String = ""        '端子･ｺﾈｸﾀﾋﾟﾝ配列
        Dim strOption As String = ""        'オプション
        Dim strTaiki As String = ""         '大気開放タイプ
        Dim strDenatsu As String = ""       '電圧
        Dim strCleanShiyo As String = ""    'クリーン仕様
        Dim strHosyo As String = ""         '保証
        Dim strLion As String = ""          '二次電池

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData, False)
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = Divisions.AccumulatePriceDiv.C5
            End If

            '機種によりボックス数が変わる為、当ロジック先頭で分岐させる
            Select Case selectedData.Series.series_kataban.Trim
                'RM0904031 2009/06/23 Y.Miura
                'Case "3GA1", "3GA2", _
                '     "4GA1", "4GA2", _
                '     "3GB1", "3GB2", _
                '     "4GB1", "4GB2"
                Case "3GA1", "3GA2", _
                     "4GA1", "4GA2", _
                     "3GB1", "3GB2", _
                     "4GB1", "4GB2"
                    If selectedData.Series.key_kataban.Trim = "R" Or _
                        selectedData.Series.key_kataban.Trim = "S" Then
                        strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(4).Trim               '接続口径
                        strDensen = selectedData.Symbols(5).Trim              '電線接続
                        strTanshi = selectedData.Symbols(6).Trim              '端子・ｺﾈｸﾀﾋﾟﾝ配列
                        strSyudoSochi = selectedData.Symbols(7).Trim          '手動装置
                        strOption = selectedData.Symbols(8).Trim              'オプション
                        strDenatsu = selectedData.Symbols(9).Trim             '電圧
                        strCleanShiyo = selectedData.Symbols(10).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(11).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 12 Then
                            strLion = selectedData.Symbols(12).Trim           '二次電池
                        End If
                    Else
                        strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(3).Trim               '接続口径
                        strDensen = selectedData.Symbols(4).Trim              '電線接続
                        strSyudoSochi = selectedData.Symbols(5).Trim          '手動装置
                        strOption = selectedData.Symbols(6).Trim              'オプション
                        strDenatsu = selectedData.Symbols(7).Trim             '電圧
                        strCleanShiyo = selectedData.Symbols(8).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(9).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 10 Then
                            strLion = selectedData.Symbols(10).Trim           '二次電池
                        End If
                    End If
                Case "3GD1", "3GD2", _
                     "4GD1", "4GD2", _
                     "3GE1", "3GE2", _
                     "4GE1"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(4).Trim               '接続口径
                        strDensen = selectedData.Symbols(5).Trim              '電線接続
                        strSyudoSochi = selectedData.Symbols(6).Trim          '手動装置
                        strOption = selectedData.Symbols(7).Trim              'オプション
                        strDenatsu = selectedData.Symbols(9).Trim             '電圧
                        strTaiki = selectedData.Symbols(8).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(10).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 11 Then
                            strLion = selectedData.Symbols(11).Trim           '二次電池
                        End If
                    Else
                        strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(3).Trim               '接続口径
                        strDensen = selectedData.Symbols(4).Trim              '電線接続
                        strSyudoSochi = selectedData.Symbols(5).Trim          '手動装置
                        strOption = selectedData.Symbols(6).Trim              'オプション
                        strDenatsu = selectedData.Symbols(7).Trim             '電圧
                        strCleanShiyo = selectedData.Symbols(8).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(9).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 10 Then
                            strLion = selectedData.Symbols(10).Trim           '二次電池
                        End If
                    End If
                    '↓RM1310067 2013/10/23 追加
                Case "4GE2"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        If selectedData.Series.key_kataban.Trim <> "1" Then
                            strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                            strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                            strKokei = selectedData.Symbols(4).Trim               '接続口径
                            strDensen = selectedData.Symbols(5).Trim              '電線接続
                            strSyudoSochi = selectedData.Symbols(6).Trim          '手動装置
                            strOption = selectedData.Symbols(7).Trim              'オプション
                            strDenatsu = selectedData.Symbols(9).Trim             '電圧
                            strTaiki = selectedData.Symbols(8).Trim          'クリーン仕様
                            strHosyo = selectedData.Symbols(10).Trim               '保証
                            If UBound(selectedData.Symbols.ToArray()) >= 11 Then
                                strLion = selectedData.Symbols(11).Trim           '二次電池
                            End If
                        Else
                            strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                            strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                            strKokei = selectedData.Symbols(4).Trim               '接続口径
                            strDensen = selectedData.Symbols(5).Trim              '電線接続
                            strOption = selectedData.Symbols(6).Trim              'オプション
                            strDenatsu = selectedData.Symbols(9).Trim             '電圧
                        End If
                    Else
                        'キー型番の変更、およびオプション数の変更に伴い、以下の内容を合わせて修正  2016/11/22 修正 松原
                        If selectedData.Series.key_kataban.Trim <> "T" Then
                            'If selectedData.Series.key_kataban.Trim <> "1" Then
                            strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                            strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                            strKokei = selectedData.Symbols(3).Trim               '接続口径
                            strDensen = selectedData.Symbols(4).Trim              '電線接続
                            strSyudoSochi = selectedData.Symbols(5).Trim          '手動装置
                            strOption = selectedData.Symbols(6).Trim              'オプション
                            strDenatsu = selectedData.Symbols(7).Trim             '電圧
                            strCleanShiyo = selectedData.Symbols(8).Trim          'クリーン仕様
                            strHosyo = selectedData.Symbols(9).Trim               '保証
                            If UBound(selectedData.Symbols.ToArray()) >= 10 Then
                                strLion = selectedData.Symbols(10).Trim           '二次電池
                            End If
                        Else
                            strKiriIchikbn = selectedData.Symbols(1).Trim        '切換位置区分
                            strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                            '以下一項目分ずらす  2016/11/22 修正 松原
                            strKokei = selectedData.Symbols(4).Trim               '接続口径
                            strDensen = selectedData.Symbols(5).Trim              '電線接続
                            strOption = selectedData.Symbols(6).Trim              'オプション
                            strDenatsu = selectedData.Symbols(8).Trim             '電圧
                        End If
                    End If
                Case "3GA3", "4GA3", "4GA4", _
                     "4GB3", "4GB4"
                    If selectedData.Series.key_kataban.Trim = "R" Or _
                        selectedData.Series.key_kataban.Trim = "S" Then
                        strKiriIchikbn = selectedData.Symbols(1).Trim         '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(4).Trim               '接続口径
                        strDensen = selectedData.Symbols(5).Trim              '電線接続
                        strTanshi = selectedData.Symbols(6).Trim              '端子・ｺﾈｸﾀﾋﾟﾝ配列
                        strOption = selectedData.Symbols(7).Trim              'オプション
                        strDenatsu = selectedData.Symbols(8).Trim             '電圧
                        strCleanShiyo = selectedData.Symbols(9).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(10).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 11 Then
                            strLion = selectedData.Symbols(11).Trim            '二次電池
                        End If
                    Else
                        strKiriIchikbn = selectedData.Symbols(1).Trim         '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(3).Trim               '接続口径
                        strDensen = selectedData.Symbols(4).Trim              '電線接続
                        strOption = selectedData.Symbols(5).Trim              'オプション
                        strDenatsu = selectedData.Symbols(6).Trim             '電圧
                        strCleanShiyo = selectedData.Symbols(7).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(8).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 9 Then
                            strLion = selectedData.Symbols(9).Trim            '二次電池
                        End If
                    End If
                Case "3GD3", "4GD3","4GE3"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        strKiriIchikbn = selectedData.Symbols(1).Trim         '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(4).Trim               '接続口径
                        strDensen = selectedData.Symbols(5).Trim              '電線接続
                        strOption = selectedData.Symbols(6).Trim              'オプション
                        strDenatsu = selectedData.Symbols(8).Trim             '電圧
                        strTaiki = selectedData.Symbols(7).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(9).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 10 Then
                            strLion = selectedData.Symbols(10).Trim            '二次電池
                        End If
                    Else
                        strKiriIchikbn = selectedData.Symbols(1).Trim         '切換位置区分
                        strSosakbn = selectedData.Symbols(2).Trim             '操作区分
                        strKokei = selectedData.Symbols(3).Trim               '接続口径
                        strDensen = selectedData.Symbols(4).Trim              '電線接続
                        strOption = selectedData.Symbols(5).Trim              'オプション
                        strDenatsu = selectedData.Symbols(6).Trim             '電圧
                        strCleanShiyo = selectedData.Symbols(7).Trim          'クリーン仕様
                        strHosyo = selectedData.Symbols(8).Trim               '保証
                        If UBound(selectedData.Symbols.ToArray()) >= 9 Then
                            strLion = selectedData.Symbols(9).Trim            '二次電池
                        End If
                    End If
            End Select

            '基本価格キー
            '↓RM1310067 2013/10/23 追加
            Select Case selectedData.Series.key_kataban.Trim
                'キー型番の変更に伴い修正  2016/11/22 修正 松原
                Case "R", "S", "T"
                    'Case "R", "S"
                    Select Case selectedData.Series.series_kataban.Trim
                        Case "4GE2"
                            'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                            If selectedData.Series.key_kataban.Trim <> "T" Then
                                'If selectedData.Series.key_kataban.Trim <> "1" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strKiriIchikbn & strSosakbn & "R"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & strKiriIchikbn & strSosakbn & "R-" & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strKiriIchikbn & strSosakbn & "R"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case Else
                    Select Case selectedData.Series.series_kataban.Trim
                        Case "4GE2"
                            'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                            If selectedData.Series.key_kataban.Trim <> "T" Then
                                'If selectedData.Series.key_kataban.Trim <> "1" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strKiriIchikbn & strSosakbn
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & strKiriIchikbn & strSosakbn & "-" & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strKiriIchikbn & strSosakbn
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            '配管ねじ加算価格キー
            Select Case Right(strKokei, 1)
                Case "G", "N"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        'If strKiriIchikbn = "66" Or strKiriIchikbn = "76" Or _
                        '    strKiriIchikbn = "77" Or strKiriIchikbn = "67" Then

                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "66R" & _
                        '                                               MyControlChars.Hyphen & _
                        '                                               Right(strKokei, 1)
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                        'Else
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & _
                        '                                               MyControlChars.Hyphen & _
                        '                                               Right(strKokei, 1)
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                        'End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   Right(strKokei, 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.Screw
                    End If
            End Select

            '接続口径
            Select Case strKokei
                Case "C18", "CL18", "CD18", "CD4", "CD6", "CD8", "CD10", "CF"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If InStr(selectedData.Series.series_kataban.Trim, "3G") <> 0 And _
                       (InStr(strKiriIchikbn, "1") <> 0 Or _
                       InStr(strKiriIchikbn, "11") <> 0) Then
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strKokei & MyControlChars.Hyphen & "S"
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strKokei
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '接続口径によってチェック区分を変える
            Select Case selectedData.Series.series_kataban.Trim
                Case "4GA1", "3GA1"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        Select Case strKokei
                            Case "C3N", "C4N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GA2", "3GA2"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        Select Case strKokei
                            Case "C8N", "C6N", "06N", "C4G", "C6G", "C8G", "06G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GA3", "3GA3"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        Select Case strKokei
                            Case "C8N", "C10N", "08N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GB1", "3GB1"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        Select Case strKokei
                            Case "06N", "06G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GB2", "3GB2"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        Select Case strKokei
                            Case "08N", "08G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case "4GB3", "3GB3"
                    If selectedData.Series.key_kataban.Trim = "R" Then
                        Select Case strKokei
                            Case "10N", "08N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R" & MyControlChars.Hyphen & strKokei
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
            End Select

            '大気開放加算価格キー
            If strTaiki <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & strTaiki
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'クリーン仕様加算価格キー
            If strCleanShiyo <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strKiriIchikbn & _
                                                           strSosakbn & MyControlChars.Hyphen & strCleanShiyo
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '電線接続・省配線接続加算価格キー
            If strDensen <> "" Then
                '↓RM1310067 2013/10/23 追加
                Select Case selectedData.Series.series_kataban.Trim
                    Case "4GE2"
                        'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                        If selectedData.Series.key_kataban.Trim <> "T" Then
                            'If selectedData.Series.key_kataban.Trim <> "1" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & strDensen
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & strDensen
                        End If
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & strDensen
                End Select
                Select Case strKiriIchikbn
                    Case "1", "11"
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "66", "67", "76", "77", "2", "3", "4", "5"
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 2
                End Select
            End If

            '端子・ｺﾈｸﾀﾋﾟﾝ配列
            If strTanshi <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           strTanshi
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            strOpArray = Split(strOption, MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "F"
                        Select Case strKiriIchikbn
                            Case "66", "67", "76", "77"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & "DUAL"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Case "K"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "S", "E", "Q"     'オプション「Q」を同処理に追加 2017/01/17 追加

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        'ダブルソレノイドは２倍加算
                        If strKiriIchikbn <> "1" And strKiriIchikbn <> "11" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "H"
                        If selectedData.Series.key_kataban.Trim = "R" Then
                            If strSosakbn = "9" Then
                            Else
                                '↓RM1310067 2013/10/23 追加
                                Select Case selectedData.Series.series_kataban.Trim
                                    Case "4GE2"
                                        'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                                        If selectedData.Series.key_kataban.Trim <> "T" Then
                                            'If selectedData.Series.key_kataban.Trim <> "1" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                        Else
                            '↓RM1310067 2013/10/23 追加
                            Select Case selectedData.Series.series_kataban.Trim
                                Case "4GE2"
                                    'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                                    If selectedData.Series.key_kataban.Trim <> "T" Then
                                        'If selectedData.Series.key_kataban.Trim <> "1" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        End If
                    Case Else
                        '↓RM1310067 2013/10/23 追加
                        Select Case selectedData.Series.series_kataban.Trim
                            Case "4GE2"
                                'キー型番の変更、およびオプション数の変更に伴い修正  2016/11/22 修正 松原
                                If selectedData.Series.key_kataban.Trim <> "T" Then
                                    'If selectedData.Series.key_kataban.Trim <> "1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

            '電圧加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "4GA4", "4GB4"
                    If strDenatsu = "5" Then
                        If strKiriIchikbn = "1" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "4G4-AC"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "4G4-AC(2)"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

            '2011/06/16 ADD RM1106028(7月VerUP:M4G-ULシリーズ　価格積上げ) START --->
            'RM1210067 2013/04/04 不具合対応
            'ＵＬ仕様加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "3GA1", "3GA2", "3GA3", "4GA1", "4GA2", "4GA3", _
                     "3GB1", "3GB2", "4GB1", "4GB2", "4GB3"
                    If strHosyo = "UL" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 4) & MyControlChars.Hyphen & strHosyo
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select
            '2011/06/16 ADD RM1106028(7月VerUP:M4G-ULシリーズ　価格積上げ) <--- END

            '二次電池加算    'RM1005030 2010/05/25 Y.Miura 追加
            If strLion <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                Select Case strKiriIchikbn
                    Case "1", "11"
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   "-OP-" & strLion & MyControlChars.Hyphen & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "2", "3", "4", "5"
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                   "-OP-" & strLion & MyControlChars.Hyphen & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "66", "67", "76", "77"
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & strKiriIchikbn & _
                                                                   "-OP-" & strLion & MyControlChars.Hyphen & strKokei
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            '電圧
            Select Case selectedData.Series.key_kataban
                Case "R"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                               MyControlChars.Hyphen & strDenatsu
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
            End Select

            'オプション(H)
            If selectedData.Series.key_kataban.Trim = "R" Or _
               selectedData.Series.key_kataban.Trim = "S" Then
                If strSosakbn = "9" Then
                    If Not strOption.Contains("H") Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "R-H"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module