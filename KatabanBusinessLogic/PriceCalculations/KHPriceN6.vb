'************************************************************************************
'*  ProgramID  ：KHPriceN6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ベース搭載用　電磁弁単品　Ｗ４ＧＢ４
'*
'*  更新履歴   ：                       更新日：2008/04/09   更新者：T.Sato
'*   ・受付No.RM0803048対応  オプションに『無記号』を追加したので価格キー作成ロジックを追加
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceN6

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
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
                                                       selectedData.Symbols(1).Trim & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & "00"
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー(M/M7)
            Select Case Trim(selectedData.Symbols(5).Trim)
                Case ""
                    Select Case selectedData.Symbols(1).Trim
                        Case "1"
                            '2位置シングル
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G4" & MyControlChars.Hyphen & "BLANK" & MyControlChars.Hyphen & "S"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2", "3", "4", "5"
                            '2位置ダブル,3位置
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G4" & MyControlChars.Hyphen & "BLANK" & MyControlChars.Hyphen & "D"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "M7"
                    Select Case selectedData.Symbols(1).Trim
                        Case "1"
                            '2位置シングル
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G4" & MyControlChars.Hyphen & selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "S"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2", "3", "4", "5"
                            '2位置ダブル,3位置
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G4" & MyControlChars.Hyphen & selectedData.Symbols(5).Trim & MyControlChars.Hyphen & "D"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            'オプション加算価格キー(A/K)
            strOpArray = Split(selectedData.Symbols(6), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算(AC110Vは加算する)
            If selectedData.Symbols(7).Trim = "5" Then
                If selectedData.Symbols(1).Trim = "1" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-AC"
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-AC(2)"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
