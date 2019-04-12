'************************************************************************************
'*  ProgramID  ：KHPriceL0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/30   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：レギュレータ
'*             ：ＲＢ５００
'*             ：ＭＮＲＢ５００
'*             ：ＮＲＢ５００
'*
'*  二次電池追加                         RM1004012 2010/04/23 Y.Miura
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceL0

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String = Nothing
        Dim intLoopCnt As Integer
        Dim bolOptionL As Boolean
        Dim bolOptionT As Boolean
        Dim bolFirst As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case selectedData.Series.series_kataban.Trim
                Case "RB500"
                    Select Case selectedData.Series.key_kataban.Trim
                        '2010/08/24 MOD RM1008009(9月VerUP RB500シリーズ機種追加)
                        'Case "1"
                        Case "1", "4"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & _
                                                                       selectedData.Symbols(2).Trim & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                            bolOptionL = False
                            bolOptionT = False
                            bolFirst = True
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "L"
                                        bolOptionL = True
                                    Case "T"
                                        bolOptionT = True
                                End Select
                            Next
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case "L", "T"
                                        If bolOptionL = True And bolOptionT = True Then
                                            If bolFirst = True Then
                                                bolFirst = False
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "LT"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next

                            '二次電池加算価格キー
                            '2010/08/24 MOD RM1008009(9月VerUP RB500シリーズ機種追加) START--->
                            If selectedData.Series.key_kataban.Trim = "4" AndAlso _
                            selectedData.Symbols(6) <> "" Then
                                ''RM1004012 2010/04/23 Y.Miura
                                'If selectedData.Symbols(6) <> "" Then
                                '2010/08/24 MOD RM1008009(9月VerUP RB500シリーズ機種追加) <---END
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                        Case "2"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                            bolOptionL = False
                            bolOptionT = False
                            bolFirst = True
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "L"
                                        bolOptionL = True
                                    Case "T"
                                        bolOptionT = True
                                End Select
                            Next
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case "L", "T"
                                        If bolOptionL = True And bolOptionT = True Then
                                            If bolFirst = True Then
                                                bolFirst = False
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "LT"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                    End Select
                Case "MNRB500"
                    '基本価格キー
                    Select Case selectedData.Symbols(1).Trim
                        Case "A"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & _
                                                                       Left(selectedData.Symbols(4).Trim, 1) & _
                                                                       Right(selectedData.Symbols(4).Trim, 1)
                            decOpAmount(UBound(decOpAmount)) = CInt(selectedData.Symbols(5).Trim)
                        Case "B"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & _
                                                                       selectedData.Symbols(3).Trim & _
                                                                       Left(selectedData.Symbols(4).Trim, 1) & _
                                                                       Right(selectedData.Symbols(4).Trim, 1)
                            decOpAmount(UBound(decOpAmount)) = CInt(selectedData.Symbols(5).Trim)
                    End Select

                    'エンドブロック(右側)加算価格キー
                    Select Case selectedData.Symbols(8).Trim
                        Case "D"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & "NE" & _
                                                                       selectedData.Symbols(8).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & "NE"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'エンドブロック(左側)加算価格キー
                    Select Case selectedData.Symbols(8).Trim
                        Case "D"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & "NE" & _
                                                                       selectedData.Symbols(8).Trim & "L"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & "NEL"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'ＤＩＮレール加算価格キー
                    Select Case selectedData.Symbols(8).Trim
                        Case "D"
                        Case Else
                            Select Case selectedData.Symbols(1).Trim
                                Case "A"
                                    Select Case selectedData.Symbols(5).Trim
                                        Case "1"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA125"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "2"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "3"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA175"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "4"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA212.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "5"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA237.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "6"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA262.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "7"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA287.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "8"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA325"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "9"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA350"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA375"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "B"
                                    Select Case selectedData.Symbols(5).Trim
                                        Case "1"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "2"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA137.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "3"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA162.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "4"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA187.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "5"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA212.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "6"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA250"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "7"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA275"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "8"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "9"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA325"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       MyControlChars.Hyphen & "BAA362.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select

                    '集中給気ブロック加算価格キー
                    Select Case selectedData.Symbols(1).Trim
                        Case "A"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & MyControlChars.Hyphen & "NP" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & _
                                                                       Left(selectedData.Symbols(4).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                    bolOptionL = False
                    bolOptionT = False
                    bolFirst = True
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                bolOptionL = True
                            Case "T"
                                bolOptionT = True
                        End Select
                    Next
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "L", "T"
                                If bolOptionL = True And bolOptionT = True Then
                                    If bolFirst = True Then
                                        bolFirst = False
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 3) & "LT"
                                        decOpAmount(UBound(decOpAmount)) = CInt(selectedData.Symbols(5).Trim)
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 3) & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = CInt(selectedData.Symbols(5).Trim)
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 3) & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = CInt(selectedData.Symbols(5).Trim)
                        End Select
                    Next
                    'オプション加算価格キー
                    strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 3) & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = CInt(selectedData.Symbols(5).Trim)
                        End Select
                    Next
                Case "NRB500"
                    Select Case selectedData.Series.key_kataban.Trim
                        Case "1"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & _
                                                                       selectedData.Symbols(3).Trim & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                            bolOptionL = False
                            bolOptionT = False
                            bolFirst = True
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "L"
                                        bolOptionL = True
                                    Case "T"
                                        bolOptionT = True
                                End Select
                            Next
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case "L", "T"
                                        If bolOptionL = True And bolOptionT = True Then
                                            If bolFirst = True Then
                                                bolFirst = False
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & "LT"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        Case "2"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim & _
                                                                       selectedData.Symbols(4).Trim & _
                                                                       selectedData.Symbols(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                            strOpArray = Split(selectedData.Symbols(7), MyControlChars.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        Case "4"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'リード線長さ加算価格キー
                            If selectedData.Symbols(3).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(4).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
