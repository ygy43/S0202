'************************************************************************************
'*  ProgramID  ：KHPriceL4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ショービックシリンダ　ＳＨＣ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceL4

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(4).Trim))

            '基本価格キー
            If Mid(selectedData.Series.series_kataban.Trim, 5, 1) = "K" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '支持形式加算価格キー
            Select Case selectedData.Symbols(1).Trim
                Case "FA", "FB", "CA", "CB", "TA", "TB", "TC"
                    If Mid(selectedData.Series.series_kataban.Trim, 5, 1) = "K" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            'スイッチ加算価格キー
            If selectedData.Symbols(6).Trim <> "" Then
                Select Case selectedData.Symbols(7).Trim
                    Case "A", "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim & _
                                                                   selectedData.Symbols(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)

                        'リード線長さ加算価格キー
                        If selectedData.Symbols(7).Trim.Length <> 0 Then
                            '耐強磁界スイッチの時
                            If InStr(selectedData.Series.series_kataban, "L2") > 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(6).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(7).Trim
                                decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(8).Trim)
                            End If
                        End If
                End Select
            End If

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(9), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "J"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                   intStroke.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "F", "G1", "P6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "A"
                        If Mid(selectedData.Series.series_kataban.Trim, 5, 1) = "K" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 5) & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                       strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(selectedData.Symbols(10), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'L2(耐強磁界スイッチ)加算価格キー
            If InStr(selectedData.Series.series_kataban, "L2") > 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 3) & "-L2-" & _
                                                           selectedData.Symbols(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
