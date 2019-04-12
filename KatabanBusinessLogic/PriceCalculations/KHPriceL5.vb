'************************************************************************************
'*  ProgramID  ：KHPriceL5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：マニホールド電磁弁単体　３ＳＡ１／４ＳＡ１／４ＳＢ１
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceL5

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intQuantity As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '数量設定
            Select Case selectedData.Symbols(1).Trim
                Case "1"
                    intQuantity = 1
                Case "11"
                    intQuantity = 1
                Case "2"
                    intQuantity = 2
                Case "3"
                    intQuantity = 2
                Case "4"
                    intQuantity = 2
                Case "5"
                    intQuantity = 2
                Case Else
                    intQuantity = 1
            End Select


            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                       selectedData.Symbols(1).Trim & _
                                                       selectedData.Symbols(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'A・Bポートフィルタ加算価格キー
            If Mid(selectedData.Symbols(3).Trim, 3, 1) = "F" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           Mid(selectedData.Symbols(3).Trim, 3, 1)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '手動装置加算価格キー
            If selectedData.Symbols(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            '配線方式加算価格キー
            If selectedData.Symbols(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "4S1" & MyControlChars.Hyphen & selectedData.Symbols(5).Trim
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            'その他オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "4S1" & MyControlChars.Hyphen & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
