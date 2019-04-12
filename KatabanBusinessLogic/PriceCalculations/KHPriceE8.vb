'************************************************************************************
'*  ProgramID  ：KHPriceE8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：レギュレータ・リバースレギュレータ
'*             ：Ｆ１０００／８０００
'*             ：Ｗ１０００／１１００／８０００／８１００
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceE8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionF As Boolean = False
        Dim bolOptionF1 As Boolean = False
        Dim bolOptionFF As Boolean = False
        Dim bolOptionFF1 As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionT As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'オプション選択判定
            strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "F"
                        bolOptionF = True
                    Case "F1"
                        bolOptionF1 = True
                    Case "FF"
                        bolOptionFF = True
                    Case "FF1"
                        bolOptionFF1 = True
                    Case "Y"
                        bolOptionY = True
                    Case "T"
                        bolOptionT = True
                End Select
            Next

            Select Case True
                Case bolOptionF = True
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "F"

                    If bolOptionT = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                    End If

                    If bolOptionY = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                    End If
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case bolOptionF1 = True
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "F1"

                    If bolOptionT = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                    End If

                    If bolOptionY = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                    End If
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case bolOptionFF = True
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "FF"

                    If bolOptionT = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                    End If

                    If bolOptionY = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                    End If
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case bolOptionFF1 = True
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & MyControlChars.Hyphen & "FF1"

                    If bolOptionT = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                    End If

                    If bolOptionY = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                    End If
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00"

                    If bolOptionT = True Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "T"
                    End If

                    If bolOptionY = True Then
                        If Mid(strOpRefKataban(UBound(strOpRefKataban)).Trim, 6, 1) = MyControlChars.Hyphen Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "Y"
                        End If
                    End If
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            strOpArray = Split(selectedData.Symbols(2), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'アタッチメント加算価格キー
            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        Select Case True
                            Case Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 3)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 2) & "*00" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
