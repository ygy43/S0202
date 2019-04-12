'************************************************************************************
'*  ProgramID  ：KHPrice04
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*
'*  概要       ：低圧損形エアオペレイト式２ポート弁
'*             ：ＣＶＥ２／ＣＶＳＥ２シリーズ
'*
'*  更新履歴   ：                       更新日：2008/11/06   更新者：T.Sato
'*  ・受付No.RM0810052 電磁弁上面搭載オプション（Ｔ）追加
'* 
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice09

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

            '価格キー設定
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(4).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'コイルオプション加算価格キー
            Select Case selectedData.Series.series_kataban.Trim
                Case "CVE2", "CVSE2"
                    If selectedData.Series.key_kataban.Trim = "" Then
                        Select Case selectedData.Symbols(5).Trim
                            Case "", "2C"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        Select Case selectedData.Symbols(5).Trim
                            Case "", "2G"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                           selectedData.Symbols(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case Else
                    Select Case selectedData.Symbols(5).Trim
                        Case "", "2G"
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            'その他オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "B", "B2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "S", "ST"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-S"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算価格キー
            Select Case Left(selectedData.Series.series_kataban.Trim, 4)
                Case "CVSE"
                    If selectedData.Symbols(8).Trim <> "" Then
                        '電圧区分取得
                        strStdVoltageFlag = KatabanUtility.GetVoltageInfo(selectedData, _
                                                                       selectedData.Symbols(8).Trim)
                        '標準電圧以外の場合は電圧加算
                        If strStdVoltageFlag <> Divisions.VoltageDiv.Standard Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & "-OTHER-VOL"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
