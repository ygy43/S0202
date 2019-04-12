'************************************************************************************
'*  ProgramID  ：KHPriceE7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：レギュレータ・リバースレギュレータ
'*             ：Ｌ１０００／Ｌ３０００／Ｌ４０００／Ｌ８０００／Ｒ１０００
'*             ：Ｒ１１００／Ｒ２０００／Ｒ２１００／Ｒ３０００／Ｒ３１００
'*             ：Ｒ４０００／Ｒ４１００／Ｒ６０００／Ｒ６１００／Ｒ８０００／Ｒ８１００
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceE7

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If selectedData.Symbols(4).Trim = "P74" Then
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'アタッチメント加算価格キー
                strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            Else
                If selectedData.Symbols(2).IndexOf("T") >= 0 And _
                   selectedData.Symbols(2).IndexOf("T8") < 0 Then

                    'T6の場合の条件を追加  2017/03/02 更新 RM1702049   ------------------------------------------------------------------------------------------->

                    If selectedData.Symbols(2).IndexOf("T") >= 0 And _
                       selectedData.Symbols(2).IndexOf("T6") < 0 Then

                        If selectedData.Symbols(4).Trim = "P70" Then
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "T" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "T"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else

                        If selectedData.Symbols(4).Trim = "P70" Then
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    End If

                    'T6の場合の条件を追加  2017/03/02 更新 RM1702049   <-------------------------------------------------------------------------------------------

                    'If selectedData.Symbols(4).Trim = "P70" Then
                    '    '基本価格キー
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "T" & MyControlChars.Hyphen & _
                    '                                               selectedData.Symbols(4).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'Else
                    '    '基本価格キー
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "T"
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If

                Else
                    If selectedData.Symbols(4).Trim = "P70" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If

                'オプション加算価格キー
                strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            '2010/07/27 MOD RM1007012(8月VerUp：FRLクリーン仕様シリーズ) START --->
                            'If selectedData.Symbols(4).Trim = "P70" Then
                            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                            '                                               strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                            '                                               selectedData.Symbols(4).Trim
                            '    decOpAmount(UBound(decOpAmount)) = 1
                            'Else
                            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                            '                                               strOpArray(intLoopCnt).Trim
                            '    decOpAmount(UBound(decOpAmount)) = 1
                            'End If
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'T6の場合の条件を追加  2017/03/02 更新 RM1702049 
                            If (selectedData.Symbols(4).Trim = "P70" _
                                AndAlso strOpArray(intLoopCnt).Trim = "T8") Or
                               (selectedData.Symbols(4).Trim = "P70" _
                                AndAlso strOpArray(intLoopCnt).Trim = "T6") Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) _
                                                    & MyControlChars.Hyphen & selectedData.Symbols(4).Trim
                            End If
                            '2010/07/27 MOD RM1007012(8月VerUp：FRLクリーン仕様シリーズ) <--- END
                    End Select
                Next

                'アタッチメント加算価格キー
                strOpArray = Split(selectedData.Symbols(5), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            Select Case True
                                Case Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "G"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 2)
                                Case Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "G"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 3)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                               strOpArray(intLoopCnt).Trim
                            End Select

                            If selectedData.Symbols(4).Trim = "P70" Then
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "A8", "A10", "A15", "A20", "A25", _
                                         "A32", "B", "B3", "B4", "E1", _
                                         "GX59", "GY59"
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                                                   selectedData.Symbols(4).Trim
                                End Select
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
