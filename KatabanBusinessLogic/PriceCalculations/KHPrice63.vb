'************************************************************************************
'*  ProgramID  ：KHPrice63
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：エアオペレイト式２ポート弁
'*             ：ＳＡＢ／ＳＶＢ
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice63

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If selectedData.Symbols(2).Trim = "S" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'コイルオプション加算価格キー
            If selectedData.Symbols(5).Trim <> "" Then
                If selectedData.Symbols(5).Trim <> "2C" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                               selectedData.Symbols(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'その他オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "L"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "S", "B", "B2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算価格キー
            If selectedData.Series.series_kataban.Trim = "SVB" Then
                strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                               selectedData.Symbols(8).Trim)
                If strStdVoltageFlag <> Divisions.VoltageDiv.Standard Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OTHER-VOL"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '食品製造工程向け商品
            If selectedData.Series.series_kataban.Trim = "SAB" And selectedData.Series.key_kataban.Trim = "F" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(8).Trim
                decOpAmount(UBound(decOpAmount)) = 1

            End If

            'RM1806035_二次電池機種追加対応
            If (selectedData.Series.series_kataban.Trim = "SAB" And selectedData.Series.key_kataban.Trim = "4") _
            Or (selectedData.Series.series_kataban.Trim = "SVB" And selectedData.Series.key_kataban.Trim = "4") Then

                Select Case selectedData.Series.series_kataban
                    Case "SAB"

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(8).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "SVB"

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(9).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
