'************************************************************************************
'*  ProgramID  ：KHPriceC1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/06   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：精密レギュレータ　ＲＰ１０００／ＲＰ２０００
'*
'*  ・受付No：RM1003086　二次電池対応   2010/03/26 Y.Miura        
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceC3

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'RM1003086 2010/03/26 Y.Miura 追加
            'RP1000は3番目の要素『二次電池仕様』が存在するのでstrOpSymbol(3)以降はプラス1する
            Dim intOpt As Integer
            Select Case selectedData.Series.series_kataban.Trim
                Case "RP1000"
                    intOpt = 1
                Case "RP2000"
                    intOpt = 0
            End Select

            '基本価格キー            'RM1003086 2010/03/26 Y.Miura 変更（intOptを追加）
            If selectedData.Symbols(3 + intOpt).Trim = "P70" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3 + intOpt).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'RM1001045 2010/02/24 Y.Miura 二次電池機器追加　RP1000のみ
            '二次電池加算価格キー
            Select Case selectedData.Symbols(3).Trim
                Case "P4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OP-" & _
                                                               selectedData.Symbols(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(4 + intOpt), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        If selectedData.Symbols(3 + intOpt).Trim = "P70" Then
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "B", "B3", "GX49", "GY49", "E1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(3 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
