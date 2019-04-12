﻿'************************************************************************************
'*  ProgramID  ：KHPrice07
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：エアフィルタ
'*             ：ＦＭ３／４／６／８０００
'*             ：ＭＭ３／４／６／８０００
'*             ：ＦＭ３／４／６／８０００ーＷ
'*             ：ＭＭ３／４／６／８０００ーＷ
'*
'*  更新履歴   ：                       更新日：2008/01/22   更新者：NII A.Takahashi
'*               ・FM3/4/6/8000-W,MM3/4/6/8000-Wを追加したため、単価見積りロジック変更
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPrice07

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim bolOptionF As Boolean = False
        Dim bolOptionQ As Boolean = False
        Dim bolOptionS As Boolean = False
        Dim bolOptionX As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim intOptionPos As Integer
        Dim intDispUnitPos As Integer
        Dim intAttachPos As Integer
        Dim bolWFlg As Boolean
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If selectedData.Series.key_kataban.Trim = "W" Then
                bolWFlg = True
                intOptionPos = 3
                intDispUnitPos = 4
                intAttachPos = 5
            Else
                bolWFlg = False
                intOptionPos = 2
                intDispUnitPos = 3
                intAttachPos = 4
            End If

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            If bolWFlg = True Then
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & "W"
            Else
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション判定
            strOpArray = Split(selectedData.Symbols(intOptionPos), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "F", "F1"
                        bolOptionF = True
                    Case "Q"
                        bolOptionQ = True
                    Case "S"
                        bolOptionS = True
                    Case "X"
                        bolOptionX = True
                    Case "Y"
                        bolOptionY = True
                End Select
            Next

            'オプションＦ付加
            If bolOptionF = True Then
                If Left(selectedData.Series.series_kataban.Trim, 2) = "FM" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "F"
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "F1"
                End If
            End If
            'オプションＹ付加
            If bolOptionY = True Then
                If bolOptionF = True Then
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "Y"
                End If
            End If
            'オプションＳ付加
            If bolOptionS = True Then
                If bolOptionF = True Then
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "S"
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "S"
                End If
            End If
            'オプションＸ付加
            If bolOptionX = True Then
                If bolOptionF = True Then
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "X"
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & "X"
                End If
            End If

            '表示単位付加
            If selectedData.Symbols(intDispUnitPos).Trim <> "" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(intDispUnitPos).Trim
            End If

            'オプション「Ｑ」加算価格キー
            If bolOptionQ = True Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & "Q"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '配管アダプタセット・ブラケット加算価格キー
            If selectedData.Symbols(intAttachPos).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(intAttachPos).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
