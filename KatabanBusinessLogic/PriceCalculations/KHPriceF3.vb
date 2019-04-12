'************************************************************************************
'*  ProgramID  ：KHPriceF3
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ニューハンドリングシステム　ＮＨＳ－Ｓ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceF3

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer

        Dim intZStroke As Integer
        Dim strXSymbol As String
        Dim strKSymbol As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '可搬質量を変換する(1)
            Select Case selectedData.Symbols(2).Trim
                Case "1003", "1503"
                    strXSymbol = "10"
                Case "1507", "1512"
                    strXSymbol = "15"
                Case "3007", "3012"
                    strXSymbol = "30"
                Case "5007", "5012", "5033"
                    strXSymbol = "50"
                Case Else
                    strXSymbol = ""
            End Select

            '可搬質量を変換する(2)
            Select Case Mid(selectedData.Symbols(2).Trim, 3, 2)
                Case "03"
                    strKSymbol = "16"
                Case "07"
                    strKSymbol = "25"
                Case "12"
                    strKSymbol = "32"
                Case "33"
                    strKSymbol = "50"
                Case Else
                    strKSymbol = ""
            End Select

            'Z軸ストロークの値を求める(中間STまるめ処理)
            Select Case True
                Case CInt(selectedData.Symbols(5).Trim) <= 50
                    intZStroke = 50
                Case 51 <= CInt(selectedData.Symbols(5).Trim) And CInt(selectedData.Symbols(5).Trim) <= 75
                    intZStroke = 75
                Case 76 <= CInt(selectedData.Symbols(5).Trim) And CInt(selectedData.Symbols(5).Trim) <= 100
                    intZStroke = 100
                Case 101 <= CInt(selectedData.Symbols(5).Trim) And CInt(selectedData.Symbols(5).Trim) <= 125
                    intZStroke = 125
                Case 126 <= CInt(selectedData.Symbols(5).Trim) And CInt(selectedData.Symbols(5).Trim) <= 150
                    intZStroke = 150
                Case 151 <= CInt(selectedData.Symbols(5).Trim) And CInt(selectedData.Symbols(5).Trim) <= 175
                    intZStroke = 175
                Case 176 <= CInt(selectedData.Symbols(5).Trim)
                    intZStroke = 200
            End Select

            'X軸処理
            Select Case selectedData.Symbols(1).Trim
                Case "1"
                    'ストローク設定
                    intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                          CInt(selectedData.Symbols(1).Trim), _
                                                          CInt(selectedData.Symbols(3).Trim))

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & MyControlChars.Hyphen & _
                                                               strXSymbol & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    'ストローク設定
                    intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                          CInt(selectedData.Symbols(1).Trim), _
                                                          CInt(selectedData.Symbols(3).Trim) + CInt(selectedData.Symbols(6).Trim))

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & MyControlChars.Hyphen & _
                                                               strXSymbol & MyControlChars.Hyphen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(2ヘッド)
            Select Case selectedData.Symbols(1).Trim
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-2HEAD-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(ストローク調整ブロック)
            Select Case selectedData.Symbols(7).Trim
                Case "L", "R", "D"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-STAB-" & selectedData.Symbols(7).Trim & MyControlChars.Hyphen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(スピードコントローラ)
            Select Case selectedData.Symbols(8).Trim
                Case "3", "4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-SCLB-" & selectedData.Symbols(8).Trim & MyControlChars.Hyphen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(ケーブルベア)
            Select Case selectedData.Symbols(9).Trim
                Case "B", "W"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-CABLE-" & selectedData.Symbols(9).Trim & MyControlChars.Hyphen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(スイッチ)
            Select Case selectedData.Symbols(10).Trim
                Case "A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & MyControlChars.Hyphen & selectedData.Symbols(10).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'Z軸処理
            '(Z軸)基本価格キー
            Select Case selectedData.Symbols(1).Trim
                Case "1"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-" & strKSymbol & MyControlChars.Hyphen & intZStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-" & strKSymbol & MyControlChars.Hyphen & intZStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 2
            End Select

            '(Z軸)オプション加算価格キー(スピードコントローラ)
            Select Case selectedData.Symbols(8).Trim
                Case "3", "4"
                    Select Case selectedData.Symbols(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-SC-" & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-SC-" & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(スイッチ)
            Select Case selectedData.Symbols(10).Trim
                Case "A"
                    Select Case selectedData.Symbols(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-" & selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-" & selectedData.Symbols(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(落下防止機構)
            Select Case selectedData.Symbols(11).Trim
                Case "Q"
                    Select Case selectedData.Symbols(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-" & selectedData.Symbols(11).Trim & MyControlChars.Hyphen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-" & selectedData.Symbols(11).Trim & MyControlChars.Hyphen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(アタッチメント)
            Select Case selectedData.Symbols(1).Trim
                Case "1"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-ATATCHMENT-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-S-ATATCHMENT-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 2
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
