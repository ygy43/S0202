Imports System.Text
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Namespace Managers
    Public Class PriceManager
        ''' <summary>
        '''     フル形番を取得
        ''' </summary>
        ''' <param name="selectedData">選択した情報</param>
        ''' <returns></returns>
        Public Shared Function GetFullKataban(selectedData As SelectedInfo) As String

            Dim result As New StringBuilder

            Select Case selectedData.Series.division
                Case Divisions.DataTypeDiv.Series


                    With selectedData
                        Dim series = .Series.series_kataban
                        Dim keyKataban = .Series.key_kataban

                        '機種
                        result.Append(series)

                        '機種ハイフン
                        If .Series.hyphen_div = HyphenDiv.Necessary Then result.Append(MyControlChars.Hyphen)

                        '????????
                        'If Left(Me.strcSelection.strSeriesKataban.Trim, 3) = "NAB" Then
                        '    If Me.strcSelection.strOpSymbol(4).Trim = "" And Me.strcSelection.strOpSymbol(5) = "B" Then
                        '        Me.strcSelection.strOpSymbol(4) = "0"
                        '    End If
                        'End If

                        '正常の場合
                        For i = 0 To .KatabanStructures.Count - 1

                            '選択した構成情報を追加
                            result.Append(.Symbols(i).Replace(MyControlChars.Comma, String.Empty))

                            If .KatabanStructures(i).hyphen_div = HyphenDiv.Necessary Then

                                If _
                                    ((series.StartsWith("AMD3") AndAlso (keyKataban = "1" OrElse keyKataban = "2")) OrElse
                                     (series.StartsWith("AMD4") AndAlso keyKataban = "0") OrElse
                                     (series.StartsWith("AMD5") AndAlso keyKataban = "0") OrElse
                                     (series.StartsWith("AMD0") AndAlso keyKataban = "1")) AndAlso
                                    Not String.IsNullOrEmpty(.Symbols(8)) AndAlso
                                    i = 4 Then

                                Else
                                    If _
                                        Not result.ToString.EndsWith(MyControlChars.Hyphen) Then
                                        result.Append(MyControlChars.Hyphen)
                                    End If
                                End If

                            End If
                        Next

                        'ロッド先端

                        'オプション外

                    End With

                    Return result.ToString.TrimEnd(MyControlChars.Hyphen)

                Case Else
                    Return selectedData.Series.series_kataban
            End Select
        End Function

#Region "価格情報を取得"

        ''' <summary>
        '''     価格情報を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetPricesInfo(kataban As String,
                                             currency As String,
                                             selectedData As SelectedInfo,
                                             userData As UserInfo) As PriceInfo
            Dim result As New PriceInfo

            Select Case selectedData.Series.division
                Case DataTypeDiv.Shiire
                    '仕入れ品の場合はCommonDbServiceから価格情報を取得
                Case Else

                    Using client As New DbAccessServiceClient
                        'フル形番として単価情報を取得
                        Dim fullKatabanPriceInfo = client.SelectFullKatabanPriceInfo(kataban, currency)

                        If fullKatabanPriceInfo IsNot Nothing Then
                            result = fullKatabanPriceInfo
                        Else
                            '積上げ価格情報を取得
                            Dim accumulatePriceInfo = client.SelectAccumulatePriceInfo(kataban, currency)

                            If accumulatePriceInfo IsNot Nothing Then
                                result = accumulatePriceInfo
                            Else
                                '形番から価格キーを分解する
                                Dim priceKeyInfos As List(Of PriceKeyInfo) = GetPriceKeys(selectedData,
                                                                                          userData.country_cd,
                                                                                          userData.office_cd)

                                '価格キーにより価格計算
                                For Each keyInfo As PriceKeyInfo In priceKeyInfos
                                    Dim priceKeyPriceInfo As New PriceInfo

                                    '積上げ価格を取得
                                    Dim priceKeyAccumulatePriceInfo = client.SelectAccumulatePriceInfo(keyInfo.PriceKey,
                                                                                                       currency)
                                    If priceKeyAccumulatePriceInfo IsNot Nothing Then
                                        priceKeyPriceInfo = priceKeyAccumulatePriceInfo
                                    Else
                                        'フル形番価格を取得
                                        Dim priceKeyFullKatabanPriceInfo =
                                                client.SelectFullKatabanPriceInfo(keyInfo.PriceKey, currency)
                                        If priceKeyFullKatabanPriceInfo IsNot Nothing Then
                                            priceKeyPriceInfo = priceKeyFullKatabanPriceInfo
                                        End If
                                    End If

                                    '価格計算
                                    result = client.AddPriceInfo(result, priceKeyPriceInfo)
                                Next
                            End If
                        End If
                    End Using
            End Select

            Return result
        End Function

        ''' <summary>
        '''     価格キーを取得
        ''' </summary>
        ''' <param name="selectedData">選択した情報</param>
        ''' <param name="countryCd">国コード</param>
        ''' <param name="officeCd">営業所コード</param>
        ''' <returns></returns>
        Private Shared Function GetPriceKeys(selectedData As SelectedInfo,
                                             ByRef countryCd As String,
                                             ByRef officeCd As String) As List(Of PriceKeyInfo)
            Dim result As New List(Of PriceKeyInfo)

            Dim strOpRefKataban() As String = Nothing
            Dim decOpAmount() As Decimal = Nothing
            Dim strPriceDiv() As String = Nothing

            Select Case selectedData.Series.price_no.Trim()
                'Case "01"
                '    Call KHPrice01.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv, countryCd, officeCd)
                Case "02"
                    Call KHPrice02.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv, countryCd, officeCd)
                Case "03"
                    Call KHPrice03.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv, countryCd, officeCd)
                Case "04"
                    Call KHPrice04.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "05"
                    Call KHPrice05.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "06"
                '    Call KHPrice06.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "07"
                    Call KHPrice07.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "08"
                    Call KHPrice08.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "09"
                    Call KHPrice09.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "10"
                    Call KHPrice10.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "11"
                    Call KHPrice11.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "12"
                    Call KHPrice12.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "13"
                    Call KHPrice13.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "14"
                    Call KHPrice14.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "15"
                    Call KHPrice15.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "16"
                '    Call KHPrice16.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "17"
                    Call KHPrice17.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv, countryCd, officeCd)
                Case "18"
                    Call KHPrice18.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv, countryCd, officeCd)
                Case "19"
                    Call KHPrice19.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "20"
                    Call KHPrice20.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "21"
                    Call KHPrice21.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "22"
                    Call KHPrice22.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "23"
                    Call KHPrice23.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "24"
                    Call KHPrice24.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "25"
                    Call KHPrice25.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "26"
                    Call KHPrice26.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)  'RM1306001 2013/06/06
                Case "27"
                    Call KHPrice27.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "28"
                    Call KHPrice28.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "29"
                    Call KHPrice29.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "30"
                    Call KHPrice30.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "31"
                    Call KHPrice31.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "32"
                    Call KHPrice32.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "33"
                    Call KHPrice33.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "34"
                '    Call KHPrice34.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "35"
                    Call KHPrice35.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "36"
                    Call KHPrice36.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "37"
                    Call KHPrice37.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "38"
                    Call KHPrice38.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "39"
                    Call KHPrice39.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "40"
                    Call KHPrice40.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "41"
                    Call KHPrice41.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "42"
                    Call KHPrice42.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "43"
                    Call KHPrice43.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "44"
                '    Call KHPrice44.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "45"
                '    Call KHPrice45.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "46"
                    Call KHPrice46.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "47"
                    Call KHPrice47.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "48"
                    Call KHPrice48.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "49"
                    Call KHPrice49.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "50"
                    Call KHPrice50.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "51"
                    Call KHPrice51.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "52"
                '    Call KHPrice52.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "53"
                    Call KHPrice53.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "54"
                    Call KHPrice54.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "55"
                    Call KHPrice55.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "56"
                    Call KHPrice56.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "57"
                '    Call KHPrice57.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "58"
                    Call KHPrice58.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "59"
                '    Call KHPrice59.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "60"
                    Call KHPrice60.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "61"
                '    Call KHPrice61.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "62"
                    Call KHPrice62.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "63"
                    Call KHPrice63.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "64"
                    Call KHPrice64.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "65"
                    Call KHPrice65.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "66"
                    Call KHPrice66.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "67"
                    Call KHPrice67.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "68"
                    Call KHPrice68.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "69"
                    Call KHPrice69.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "70"
                    Call KHPrice70.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "71"
                '    Call KHPrice71.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "72"
                    Call KHPrice72.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "73"
                    Call KHPrice73.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "74"
                    Call KHPrice74.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "75"
                    Call KHPrice75.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "76"
                    Call KHPrice76.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "77"
                    Call KHPrice77.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "78"
                    Call KHPrice78.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "79"
                    Call KHPrice79.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "80"
                    Call KHPrice80.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "81"
                    Call KHPrice81.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "82"
                    Call KHPrice82.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "83"
                    Call KHPrice83.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "84"
                    Call KHPrice84.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "85"
                    Call KHPrice85.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "86"
                    Call KHPrice86.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "87"
                    Call KHPrice87.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "88"
                    Call KHPrice88.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "89"
                '    Call KHPrice89.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "90"
                    Call KHPrice90.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "91"
                'Call KHPrice91.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "92"
                'Call KHPrice92.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "93"
                'Call KHPrice93.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "94"
                'Call KHPrice94.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "95"
                'Call KHPrice95.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "96"
                '    Call KHPrice96.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "97"
                    Call KHPrice97.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "98"
                    Call KHPrice98.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "99"
                    Call KHPrice99.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A0"
                    Call KHPriceA0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A1"
                    Call KHPriceA1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A2"
                    Call KHPriceA2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A3"
                    Call KHPriceA3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A4"
                    Call KHPriceA4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A5"
                    Call KHPriceA5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A6"
                    Call KHPriceA6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A7"
                    Call KHPriceA7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A8"
                    Call KHPriceA8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "A9"
                    Call KHPriceA9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "B0"
                    Call KHPriceB0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "B1"
                    Call KHPriceB1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "B2"
                    Call KHPriceB2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "B3"
                    Call KHPriceB3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "B4"
                    Call KHPriceB4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "B5"
                '    Call KHPriceB5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "B6"
                '    Call KHPriceB6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "B7"
                    Call KHPriceB7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "B8"
                '    Call KHPriceB8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "B9"
                    Call KHPriceB9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C0"
                    Call KHPriceC0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C1"
                    Call KHPriceC1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C2"
                    Call KHPriceC2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C3"
                    Call KHPriceC3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C4"
                    Call KHPriceC4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C5"
                    Call KHPriceC5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C6"
                    Call KHPriceC6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C7"
                    Call KHPriceC7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "C8"
                    Call KHPriceC8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "C9"
                    Call KHPriceC9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D0"
                    Call KHPriceD0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D1"
                    Call KHPriceD1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D2"
                    Call KHPriceD2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "D3"
                '    Call KHPriceD3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D4"
                    Call KHPriceD4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D5"
                    Call KHPriceD5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D6"
                    Call KHPriceD6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D7"
                    Call KHPriceD7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D8"
                    Call KHPriceD8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "D9"
                    Call KHPriceD9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E0"
                    Call KHPriceE0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E1"
                    Call KHPriceE1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E2"
                    Call KHPriceE2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E3"
                    Call KHPriceE3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "E4"
                '    Call KHPriceE4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "E5"
                    Call KHPriceE5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E6"
                    Call KHPriceE6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E7"
                    Call KHPriceE7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E8"
                    Call KHPriceE8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "E9"
                    Call KHPriceE9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F0"
                    Call KHPriceF0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F1"
                    Call KHPriceF1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F2"
                    Call KHPriceF2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F3"
                    Call KHPriceF3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F4"
                    Call KHPriceF4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F5"
                '    Call KHPriceF5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F6"
                '    Call KHPriceF6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F7"
                    Call KHPriceF7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F8"
                    Call KHPriceF8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "F9"
                    Call KHPriceF9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "G0"
                    Call KHPriceG0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G1"
                    Call KHPriceG1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv) 'RM1610011 strPriceDiv追加
                Case "G2"
                'Call KHPriceG2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G3"
                'Call KHPriceG3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G4"
                'Call KHPriceG4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G5"
                'Call KHPriceG5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G6"
                    Call KHPriceG6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G7"
                    Call KHPriceG7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G8"
                    Call KHPriceG8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "G9"
                    Call KHPriceG9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "H0"
                    Call KHPriceH0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "H1"
                    Call KHPriceH1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H2"
                    Call KHPriceH2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H3"
                    Call KHPriceH3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H4"
                    Call KHPriceH4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H5"
                    Call KHPriceH5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H6"
                    Call KHPriceH6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H7"
                    Call KHPriceH7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H8"
                    Call KHPriceH8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "H9"
                    Call KHPriceH9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I0"
                    Call KHPriceI0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I1"
                    Call KHPriceI1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I2"
                    Call KHPriceI2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I3"
                    Call KHPriceI3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I4"
                    Call KHPriceI4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I5"
                    Call KHPriceI5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I6"
                    Call KHPriceI6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I7"
                    Call KHPriceI7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I8"
                    Call KHPriceI8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "I9"
                    Call KHPriceI9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J0"
                    Call KHPriceJ0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J1"
                    Call KHPriceJ1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J2"
                    Call KHPriceJ2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J3"
                    Call KHPriceJ3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J4"
                    Call KHPriceJ4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J5"
                    Call KHPriceJ5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J6"
                    Call KHPriceJ6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J7"
                    Call KHPriceJ7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J8"
                    Call KHPriceJ8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "J9"
                    Call KHPriceJ9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K0"
                    Call KHPriceK0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K1"
                    Call KHPriceK1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K2"
                    Call KHPriceK2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K3"
                    Call KHPriceK3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K4"
                    Call KHPriceK4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K5"
                    Call KHPriceK5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K6"
                    Call KHPriceK6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K7"
                    Call KHPriceK7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "K8"
                    Call KHPriceK8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "K9"
                    Call KHPriceK9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "L0"
                    Call KHPriceL0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "L1"
                    Call KHPriceL1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "L2"
                '    Call KHPriceL2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "L3"
                    Call KHPriceL3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "L4"
                    Call KHPriceL4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "L5"
                    Call KHPriceL5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "L6"
                '    Call KHPriceL6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "L7"
                    Call KHPriceL7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "L8"
                    Call KHPriceL8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "L9"
                    Call KHPriceL9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "M0"
                    Call KHPriceM0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "M1"
                    Call KHPriceM1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "M2"
                    Call KHPriceM2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "M3"
                    Call KHPriceM3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, countryCd, officeCd)
                Case "M4"
                    Call KHPriceM4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "M5"
                    Call KHPriceM5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)  'RM1306001 2013/06/06 追加
                Case "M6"
                    Call KHPriceM6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "M7"
                    Call KHPriceM7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "M8"
                    Call KHPriceM8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "M9"
                    Call KHPriceM9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "N0"
                    Call KHPriceN0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "N1"
                    Call KHPriceN1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "N2"
                    Call KHPriceN2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "N3"
                    Call KHPriceN3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "N4"
                    Call KHPriceN4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "N5"
                '    Call KHPriceN5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "N6"
                    Call KHPriceN6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "N7"
                    Call KHPriceN7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "N8"
                    Call KHPriceN8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "N9"
                    Call KHPriceN9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "O0"
                    Call KHPriceO0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                'Case "O1"
                '    Call KHPriceO1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "O2"
                    Call KHPriceO2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "O3"
                    Call KHPriceO3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "O5"
                    Call KHPriceO5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "O6"
                    Call KHPriceO6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "O7"
                    Call KHPriceO7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "O8"
                    Call KHPriceO8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "O9"
                    Call KHPriceO9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "OA"
                    Call KHPriceOA.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P0"
                    Call KHPriceP0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "P1"
                    Call KHPriceP1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P2"
                    Call KHPriceP2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P3"
                    Call KHPriceP3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "P4"
                    Call KHPriceP4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "P5"
                    Call KHPriceP5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P6"
                    Call KHPriceP6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P7"
                    Call KHPriceP7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P8"
                    Call KHPriceP8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "P9"
                    Call KHPriceP9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "Q0"
                    Call KHPriceQ0.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "Q1"
                    Call KHPriceQ1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q2"
                    Call KHPriceQ2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q3"
                    Call KHPriceQ3.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q4"
                    Call KHPriceQ4.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q5"
                    Call KHPriceQ5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q6"
                    Call KHPriceQ6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "Q7"
                    Call KHPriceQ7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q8"
                    Call KHPriceQ8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "Q9"
                    Call KHPriceQ9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "R1"
                '    Call KHPriceR1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "R2"
                    Call KHPriceR2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                'Case "R5"
                '    Call KHPriceR5.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "R6"
                    Call KHPriceR6.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "R7"
                    Call KHPriceR7.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "R8"
                    Call KHPriceR8.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
                Case "R9"
                    Call KHPriceR9.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "S1"
                    Call KHPriceS1.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "S2"
                    'RM1708016 2017/8/22
                    Call KHPriceS2.subPriceCalculation(selectedData, strOpRefKataban, decOpAmount)
            End Select

            'The first item is nothing, so begin from the second
            For i = 1 To strOpRefKataban.Count - 1

                If strPriceDiv Is Nothing Then
                    result.Add(New PriceKeyInfo _
                                  With {.PriceKey = strOpRefKataban(i), .Amount = decOpAmount(i)})
                Else
                    result.Add(New PriceKeyInfo _
                                  With {.PriceKey = strOpRefKataban(i), .Amount = decOpAmount(i), .PriceKeyDiv = strPriceDiv(i)})
                End If
            Next

            Return result
        End Function

#End Region

        ''' <summary>
        '''     チェック区分を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetCheckDiv() As String
        End Function

        ''' <summary>
        '''     標準納期を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetStandardNoukiAndKosuu() As String
        End Function

        ''' <summary>
        '''     販売数量単位を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetQuantityUnitInfo(kataban As String, language As String) As QuantityUnitInfo
            Dim result As New QuantityUnitInfo

            Using client As New DbAccessServiceClient
                result = client.SelectQuantityUnit(kataban, language)
            End Using

            Return IIf(result Is Nothing, New QuantityUnitInfo, result)
        End Function

        ''' <summary>
        '''     在庫情報を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetStokeInfo() As String
        End Function

        ''' <summary>
        '''     出荷場所を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetShipPlace() As List(Of String)
            Dim result As New List(Of String)

            Return result
        End Function

        ''' <summary>
        '''     現地定価+通貨を取得
        ''' </summary>
        ''' <param name="katabanCurrency">形番通貨</param>
        ''' <param name="userCurrency">ユーザー通貨</param>
        ''' <param name="userCountry">ユーザー国コード</param>
        ''' <param name="katabanFirstHyphen">形番第一ハイフン</param>
        ''' <param name="checkDiv">チェック区分</param>
        ''' <param name="listPrice">登録店価格</param>
        ''' <param name="gsPrice">GS店価格</param>
        ''' <returns></returns>
        Public Shared Function GetLocalPrice(katabanCurrency As String,
                                             userCurrency As String,
                                             userCountry As String,
                                             katabanFirstHyphen As String,
                                             checkDiv As String,
                                             listPrice As Decimal,
                                             gsPrice As Decimal) As String

            Using client As New DbAccessServiceClient
                '為替レートを取得
                Dim exchangeRate = client.SelectExchangeRate(katabanCurrency, userCurrency)
                '端数処理方法を取得
                Dim mathInfo = client.SelectMathTypeLocalPrice(userCountry, katabanFirstHyphen)

                If mathInfo Is Nothing Then Return "0"

                Select Case userCountry
                    Case "USA", "MEX", "E09" '欧州代理店明治対応 RM1705008  2017/05/11 更新
                        'Case "USA", "MEX"                           'メキシコ対応  RM1509001 
                        Return FractionProcess(listPrice * mathInfo.list_price_rate1 * exchangeRate,
                                               mathInfo.mathType, mathInfo.mathPosition)
                    Case "PRC"
                        Select Case checkDiv
                            Case KatabanCheckDiv.Parts
                                'フル形番検索時、現地定価表示制御変更
                                If mathInfo.list_price_rate2 = 0 Then
                                    Return FractionProcess(gsPrice * mathInfo.list_price_rate1 * exchangeRate,
                                                           mathInfo.mathType, mathInfo.mathPosition)
                                Else
                                    Return FractionProcess(gsPrice * mathInfo.list_price_rate2 * exchangeRate,
                                                           mathInfo.mathType, mathInfo.mathPosition)
                                End If
                            Case Else
                                Return FractionProcess(gsPrice * mathInfo.list_price_rate1 * exchangeRate,
                                                       mathInfo.mathType, mathInfo.mathPosition)
                        End Select
                    Case Else
                        Return FractionProcess(gsPrice * mathInfo.list_price_rate1 * exchangeRate,
                                               mathInfo.mathType, mathInfo.mathPosition)
                End Select

            End Using
        End Function

        ''' <summary>
        '''     購入価格+通貨を取得
        ''' </summary>
        ''' <param name="dataType">データ種類</param>
        ''' <param name="katabanCurrency">形番通貨</param>
        ''' <param name="userCountry">ユーザー国コード</param>
        ''' <param name="katabanFirstHyphen">形番第一ハイフン</param>
        ''' <param name="shipPlace">選択した出荷場所</param>
        ''' <param name="gsPrice">GS店価格</param>
        ''' <returns></returns>
        Public Shared Function GetFobPrice(dataType As String,
                                           katabanCurrency As String,
                                           userCountry As String,
                                           katabanFirstHyphen As String,
                                           shipPlace As String,
                                           gsPrice As Decimal) As String
            If dataType = DataTypeDiv.Shiire Then
                Return "0"
            Else
                Using client As New DbAccessServiceClient
                    Dim mathTypeInfo = client.SelectMathTypeFobPrice(userCountry, katabanFirstHyphen, shipPlace)
                    If mathTypeInfo Is Nothing Then Return "0"

                    Dim exchangeRate = client.SelectExchangeRate(katabanCurrency, mathTypeInfo.currency_cd)

                    Return _
                        FractionProcess(gsPrice * mathTypeInfo.fob_rate * exchangeRate, mathTypeInfo.mathType,
                                        mathTypeInfo.mathPosition)

                End Using

            End If
        End Function

        ''' <summary>
        '''     ELチェック
        ''' </summary>
        ''' <param name="kataban">形番</param>
        ''' <param name="elFlag">EL区分</param>
        ''' <returns></returns>
        Public Shared Function GetElFlag(kataban As String, elFlag As String) As String
            Using client As New DbAccessServiceClient
                If client.CheckEl(kataban, elFlag) Then
                    Return MyControlChars.Maru
                Else
                    Return String.Empty
                End If
            End Using
        End Function

        ''' <summary>
        '''     在庫情報の取得
        ''' </summary>
        ''' <param name="fullKataban">形番</param>
        ''' <param name="language">言語</param>
        ''' <param name="shipPlace">出荷場所</param>
        ''' <returns></returns>
        Public Shared Function GetStock(fullKataban As String, language As String, shipPlace As String) As String
            Return String.Empty
        End Function

#Region "Private Functions"

        ''' <summary>
        '''     端数処理区分取得
        ''' </summary>
        ''' <param name="decPrice">端数処理対象数値</param>
        ''' <param name="strMathType">計算タイプ（0：なし、1：四捨五入、2：切上げ、3：切捨て）</param>
        ''' <param name="intMathPos">計算位置（例)1：なし、10：整数一位、0.1：少数一位</param>
        ''' <returns></returns>
        ''' <remarks>拠点コードを元に端数処理をする</remarks>
        Private Shared Function FractionProcess(decPrice As Decimal,
                                                strMathType As String,
                                                intMathPos As Decimal) As String
            FractionProcess = "0"
            Select Case strMathType
                Case "1"
                    '四捨五入(丸め)
                    FractionProcess = (Math.Round(decPrice * intMathPos) / intMathPos).ToString
                Case "2"
                    '切上げ
                    FractionProcess = (Math.Ceiling(decPrice * intMathPos) / intMathPos).ToString
                Case "3"
                    '切捨て
                    FractionProcess = (Math.Truncate(decPrice * intMathPos) / intMathPos).ToString
                Case "4"
                    '四捨五入
                    If intMathPos < 1 Then
                        FractionProcess =
                            (Math.Round(decPrice * intMathPos, 0, MidpointRounding.AwayFromZero) / intMathPos).ToString
                    Else
                        FractionProcess =
                            (Math.Round(decPrice, intMathPos.ToString.Length - 1, MidpointRounding.AwayFromZero)).
                                ToString
                    End If
            End Select

            If intMathPos > 1 Then
                Dim str() As String = FractionProcess.Split(".")

                If str.Length = 2 Then
                    FractionProcess = str(0) & "." & str(1).PadRight(intMathPos.ToString.Length - 1, "0")
                ElseIf str.Length = 1 Then
                    FractionProcess = str(0) & "." & "".PadRight(intMathPos.ToString.Length - 1, "0")
                End If
            End If
        End Function

#End Region
    End Class
End Namespace