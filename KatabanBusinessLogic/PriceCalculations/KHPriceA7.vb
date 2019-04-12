'************************************************************************************
'*  ProgramID  ：KHPriceA7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＮＷ３ＧＡ２
'*             ：ＮＷ４ＧＡ２
'*             ：ＮＷ４ＧＢ２
'*             ：ＮＷ４ＧＺ２
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceA7

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If selectedData.Symbols(4).Trim = "R1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                           selectedData.Symbols(1).Trim & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'RM1805036_二次電池価格加算対応
                If selectedData.Series.key_kataban = "P" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(4).Trim & "-P40"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
                If selectedData.Series.key_kataban = "F" Or selectedData.Series.key_kataban = "H" Or selectedData.Series.key_kataban = "O" Or selectedData.Series.key_kataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'バルブブロックキーの作成を追加  2017/03/07 追加
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & "-V-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                           selectedData.Symbols(1).Trim & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'RM1805036_二次電池価格加算対応
                If selectedData.Series.key_kataban = "P" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & "-P40"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
                If selectedData.Series.key_kataban = "F" Or selectedData.Series.key_kataban = "H" Or selectedData.Series.key_kataban = "O" Or selectedData.Series.key_kataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'バルブブロックキーの作成を追加  2017/03/07 追加
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 6) & _
                                                               selectedData.Symbols(1).Trim & _
                                                               selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(3).Trim & "-V-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                End If
            End If

            'オプション価格加算キー
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A", "F"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2, 5) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "M7"
                        Select Case selectedData.Symbols(1).Trim
                            Case "1", "11"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "2", "3", "4", "5"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & "D"
                                decOpAmount(UBound(decOpAmount)) = 1

                        End Select

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
                    Case "M"
                        If selectedData.Series.key_kataban = "F" Or selectedData.Series.key_kataban = "H" Or selectedData.Series.key_kataban = "O" Or selectedData.Series.key_kataban = "L" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                End Select
            Next

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If selectedData.Symbols(7).Trim = "4" Then
                If selectedData.Series.key_kataban = "F" Or selectedData.Series.key_kataban = "H" Or selectedData.Series.key_kataban = "O" Or selectedData.Series.key_kataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-DC-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If selectedData.Symbols(7).Trim = "1" Then
                If selectedData.Series.key_kataban = "F" Or selectedData.Series.key_kataban = "H" Or selectedData.Series.key_kataban = "O" Or selectedData.Series.key_kataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-AC-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'Hが含まれない場合は排気誤作動防止弁価格キー設定
            If selectedData.Symbols(6).IndexOf("H") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Mid(selectedData.Series.series_kataban.Trim, 2, 5) & MyControlChars.Hyphen & "H"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            ''食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            'If selectedData.Series.key_kataban = "F" Or selectedData.Series.key_kataban = "H" Or selectedData.Series.key_kataban = "O" Then

            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-FP1"
            '    decOpAmount(UBound(decOpAmount)) = 1

            'End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
