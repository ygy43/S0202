'************************************************************************************
'*  ProgramID  ：KHPrice86
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：真空パッド
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice86

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            ' 基本価格キー
            Select Case True
                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "C"
                    ' ロングストロークホルダ付
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut("VSP" & MyControlChars.Hyphen & _
                                                                                      selectedData.Symbols(1).Trim & _
                                                                                      selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                                                      selectedData.Symbols(3).Trim & _
                                                                                      selectedData.Symbols(4).Trim & _
                                                                                      selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                                      selectedData.Symbols(6).Trim & _
                                                                                      selectedData.Symbols(7).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "M"
                    ' 小形ホルダタイプ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut("VSP" & MyControlChars.Hyphen & _
                                                                                     selectedData.Symbols(1).Trim & _
                                                                                     selectedData.Symbols(2).Trim & _
                                                                                     selectedData.Symbols(3).Trim & _
                                                                                     selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                     selectedData.Symbols(5).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "Q"
                    ' 吸着痕防止タイプ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut("VSP" & MyControlChars.Hyphen & _
                                                                                     selectedData.Symbols(1).Trim & _
                                                                                     selectedData.Symbols(2).Trim & _
                                                                                     selectedData.Symbols(3).Trim & _
                                                                                     selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                                                                                     selectedData.Symbols(6).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = KatabanUtility.HyphenCut("VSP" & MyControlChars.Hyphen & _
                                                                                      selectedData.Symbols(1).Trim & _
                                                                                      selectedData.Symbols(2).Trim & _
                                                                                      selectedData.Symbols(3).Trim & _
                                                                                      selectedData.Symbols(4).Trim & _
                                                                                      selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                                                                                      selectedData.Symbols(6).Trim)
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            ' フリーホルダ加算価格キー
            Select Case True
                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "M"
                    '小形ホルダタイプ
                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "C"
                    ' ロングストロークホルダ付
                    Select Case selectedData.Symbols(8).Trim
                        Case "F1", "F2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "VSP" & MyControlChars.Hyphen & "P" & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                       selectedData.Symbols(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case Else
                    Select Case selectedData.Symbols(7).Trim
                        Case "F1", "F2"
                            '機種毎に価格キーを設定
                            Select Case True
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "R"
                                    'スタンダードタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-R/A-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "S"
                                    'スポンジタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-S-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "B"
                                    'ベローズタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-B-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "E"
                                    '長円タイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-E-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "L"
                                    ' ソフトタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-L-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "1"
                                    ' ソフトベローズタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "K"
                                    ' 滑り止めタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-K-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "F"
                                    ' フラットタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-F-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    'RM1610027 Start
                                Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "A"
                                    ' ソフトベローズタイプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                               selectedData.Symbols(7).Trim & MyControlChars.Hyphen & _
                                                                               selectedData.Symbols(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    'RM1610027 End
                            End Select
                    End Select
            End Select

            ' フリーホルダ加算価格キー
            Dim fullKataban = PriceManager.GetFullKataban(selectedData)
            If Right(fullKataban.Trim, 2) = "-V" Then
                ' 機種毎に価格キーを設定
                Select Case True
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "R"
                        'スタンダードタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-R/A-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "S"
                        'スポンジタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-S-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "B"
                        'ベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-B-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "W"
                        '多段ベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-W-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "E"
                        '長円タイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-E-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "L"
                        'ソフトタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-L-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "1"
                        'ソフトベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "K"
                        '滑り止めタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-K-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "C"
                        'ロングストロークホルダ付
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-P-" & _
                                                                   selectedData.Symbols(9).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "M"
                        '小型真空パッド　スタンダードタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-M-" & _
                                                                   selectedData.Symbols(6).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "Q"
                        ' 吸着痕防止タイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-Q-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "F"
                        'フラットタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-F-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1610027 Start
                    Case selectedData.Series.series_kataban.Trim = "VSP" And selectedData.Series.key_kataban.Trim = "A"
                        'ソフトベローズタイプ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "VSP-LB-" & _
                                                                   selectedData.Symbols(8).Trim & MyControlChars.Hyphen & _
                                                                   selectedData.Symbols(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1610027 End
                End Select
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
