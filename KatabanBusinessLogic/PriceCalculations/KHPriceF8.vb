'************************************************************************************
'*  ProgramID  ：KHPriceF8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ハイブリロボ　２アクション空圧ロボット　ＨＲＬ－２Ｇ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceF8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intXZStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '中間STまるめ処理
            Select Case True
                Case CInt(selectedData.Symbols(4).Trim) + CInt(selectedData.Symbols(5).Trim) <= 149
                    intXZStroke = 75
                Case 150 <= CInt(selectedData.Symbols(4).Trim) + CInt(selectedData.Symbols(5).Trim) And _
                            CInt(selectedData.Symbols(4).Trim) + CInt(selectedData.Symbols(5).Trim) <= 349
                    intXZStroke = 150
                Case 350 <= CInt(selectedData.Symbols(4).Trim) + CInt(selectedData.Symbols(5).Trim)
                    intXZStroke = 350
            End Select

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-" & selectedData.Symbols(3).Trim & MyControlChars.Hyphen & intXZStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー(スイッチ)
            If selectedData.Symbols(6).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-" & selectedData.Symbols(6).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'オプション加算価格キー(レール)
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-RAIL"
                decOpAmount(UBound(decOpAmount)) = 1

                'オプション加算価格キー(リード線長さ)
                If selectedData.Symbols(7).Trim <> "" Then
                    '2010/10/19 RM1010017(11月VerUP:HRLシリーズ) START--->
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case selectedData.Symbols(6).Trim
                        Case "T2YD"
                            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-T2YD-" & _
                                                                    selectedData.Symbols(7).Trim
                        Case "T2YDT"
                            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-T2YDT-" & _
                                                                    selectedData.Symbols(7).Trim
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-" & _
                                                                    selectedData.Symbols(7).Trim
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1

                    'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    'strOpRefKataban(UBound(strOpRefKataban)) = "HRL-2G-" & selectedData.Symbols(7).Trim
                    'decOpAmount(UBound(decOpAmount)) = 1
                    '2010/10/19 RM1010017(11月VerUP:HRLシリーズ) <---END
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
