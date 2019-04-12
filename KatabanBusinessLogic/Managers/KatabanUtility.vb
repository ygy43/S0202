Imports System.Text
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Namespace Managers
    Public Class KatabanUtility

#Region "スイッチ"

        ''' <summary>
        '''     スイッチ数取得
        ''' </summary>
        ''' <param name="strSwitchQty">スイッチ数</param>
        ''' <returns></returns>
        ''' <remarks>スイッチ数を判定し返却する</remarks>
        Public Shared Function SwitchQtyGet(strSwitchQty As String) As Integer
            Try
                'スイッチ数判定
                Select Case Left(strSwitchQty, 1)
                    Case "R", "L", "H", "X"
                        SwitchQtyGet = 1
                    Case "D"
                        SwitchQtyGet = 2
                    Case "T"
                        SwitchQtyGet = 3
                    Case Else
                        SwitchQtyGet = CInt(strSwitchQty)
                End Select
            Catch ex As Exception
                SwitchQtyGet = 1
            End Try
        End Function

#End Region

#Region "ストローク"

        ''' <summary>
        '''     ストローク調整
        ''' </summary>
        ''' <param name="selectedData"></param>
        ''' <param name="intBoreSize">口径</param>
        ''' <param name="intStroke">ストローク</param>
        ''' <returns></returns>
        ''' <remarks>ストロークのサイズを調整する</remarks>
        Public Shared Function GetStrokeSize(selectedData As SelectedInfo,
                                             intBoreSize As Integer,
                                             intStroke As Integer) As Integer
            Dim result = intStroke

            Using client As New DbAccessServiceClient

                With selectedData.Series
                    Dim standardStrokes = client.SelectStroke(.series_kataban, .key_kataban, intBoreSize, .country_cd)

                    For Each standardStroke In standardStrokes

                        If intStroke <= standardStroke.std_stroke Then
                            result = standardStroke.std_stroke
                        Else
                            Exit For
                        End If
                    Next
                End With
                Return result

            End Using
        End Function

        ''' <summary>
        '''     入力電圧のチェック
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function CheckStroke(selectedData As SelectedInfo,
                                           boreSize As Integer,
                                           stroke As Integer,
                                           madeCountry As String) As Boolean
            Dim result = False

            '電圧の取得
            Using client As New DbAccessServiceClient
                Dim strokeInfos = client.SelectStroke(selectedData.Series.series_kataban,
                                                      selectedData.Series.key_kataban,
                                                      boreSize,
                                                      madeCountry
                                                      )

                If strokeInfos.Count > 0 Then

                    If _
                        stroke >= strokeInfos.First.min_stroke AndAlso
                        stroke <= strokeInfos.First.max_stroke Then
                        result = True
                    End If

                End If
            End Using

            Return result
        End Function

#End Region

#Region "電圧"

        ''' <summary>
        '''     電圧情報の取得
        ''' </summary>
        ''' <param name="selectedData">引当情報</param>
        ''' <param name="strVoltage">電圧</param>
        ''' <param name="strCountryCd">国コード</param>
        ''' <param name="strOfficeCd">営業所コード</param>
        ''' <returns></returns>
        Public Shared Function GetVoltageInfo(selectedData As SelectedInfo,
                                              strVoltage As String,
                                              Optional strCountryCd As String = Nothing,
                                              Optional strOfficeCd As String = Nothing) As String
            Dim intVoltage As Integer
            Dim strVoltageDiv As String = Nothing
            Dim strSeriesKataban As String = Nothing
            Dim strKeyKataban As String = Nothing
            Dim strPortSize As String = Nothing
            Dim strCoil As String = Nothing

            '標準電圧にデフォルト設定
            GetVoltageInfo = VoltageDiv.Other

            '電圧検索情報取得
            Call GetVoltageSearchInfo(selectedData, strVoltage, intVoltage, strVoltageDiv,
                                      strSeriesKataban, strKeyKataban, strPortSize, strCoil)

            '電圧の取得
            Using client As New DbAccessServiceClient
                Dim voltageInfos = client.SelectVoltageInfo(strSeriesKataban, strKeyKataban, strPortSize, strCoil,
                                                            strVoltageDiv, intVoltage)

                If voltageInfos.Count > 0 Then

                    GetVoltageInfo = voltageInfos.Item(0).std_voltage_flag
                End If
            End Using

            If GetVoltageInfo <> VoltageDiv.Standard AndAlso strCountryCd IsNot Nothing Then
                If Not GetVoltageIsStandard(strVoltage, strCountryCd, strOfficeCd) Then
                    '標準電圧
                    GetVoltageInfo = VoltageDiv.Standard
                End If
            End If
        End Function

        ''' <summary>
        '''     電圧検索情報取得
        ''' </summary>
        ''' <param name="selectedData"></param>
        ''' <param name="strVoltage"></param>
        ''' <param name="intVoltage"></param>
        ''' <param name="strVoltageDiv"></param>
        ''' <param name="strSeriesKataban"></param>
        ''' <param name="strKeyKataban"></param>
        ''' <param name="strPortSize"></param>
        ''' <param name="strCoil"></param>
        ''' <returns></returns>
        Private Shared Function GetVoltageSearchInfo(selectedData As SelectedInfo,
                                                     strVoltage As String,
                                                     ByRef intVoltage As Integer,
                                                     ByRef strVoltageDiv As String,
                                                     ByRef strSeriesKataban As String,
                                                     ByRef strKeyKataban As String,
                                                     ByRef strPortSize As String,
                                                     ByRef strCoil As String)
            Dim intLoopCnt As Integer
            Try
                intVoltage = 0
                strVoltageDiv = Nothing
                strSeriesKataban = Nothing
                strKeyKataban = Nothing
                strPortSize = Nothing
                strCoil = Nothing
                '電圧設定
                Select Case strVoltage
                    Case PowerSupply.AC100V
                        intVoltage = CInt(Mid(PowerSupply.Const1, 3, PowerSupply.Const1.IndexOf("V") - 2))
                        strVoltageDiv = PowerSupply.Div1
                    Case PowerSupply.AC200V
                        intVoltage = CInt(Mid(PowerSupply.Const2, 3, PowerSupply.Const2.IndexOf("V") - 2))
                        strVoltageDiv = PowerSupply.Div1
                    Case PowerSupply.DC24V
                        intVoltage = CInt(Mid(PowerSupply.Const3, 3, PowerSupply.Const3.IndexOf("V") - 2))
                        strVoltageDiv = PowerSupply.Div2
                    Case PowerSupply.DC12V
                        intVoltage = CInt(Mid(PowerSupply.Const4, 3, PowerSupply.Const4.IndexOf("V") - 2))
                        strVoltageDiv = PowerSupply.Div2
                    Case PowerSupply.AC110V
                        intVoltage = CInt(Mid(PowerSupply.Const5, 3, PowerSupply.Const5.IndexOf("V") - 2))
                        strVoltageDiv = PowerSupply.Div1
                    Case PowerSupply.AC220V
                        intVoltage = CInt(Mid(PowerSupply.Const6, 3, PowerSupply.Const6.IndexOf("V") - 2))
                        strVoltageDiv = PowerSupply.Div1
                    Case Else
                        intVoltage = CInt(Mid(strVoltage, 3, strVoltage.IndexOf("V") - 2))
                        strVoltageDiv = Left(strVoltage, 2)
                End Select

                '接続口径・コイル設定
                For intLoopCnt = 1 To UBound(selectedData.KatabanStructures.ToArray())
                    Select Case selectedData.KatabanStructures(intLoopCnt).element_div
                        Case ElementDiv.Coil
                            strCoil = selectedData.Symbols(intLoopCnt).Trim
                        Case ElementDiv.VolPort
                            strPortSize = selectedData.Symbols(intLoopCnt).Trim
                    End Select
                Next
                If Not strCoil Is Nothing Then
                    Select Case strCoil.Trim
                        Case "0", "00"
                            strCoil = ""
                    End Select
                End If

                'シリーズ形番・キー形番
                Select Case selectedData.Series.price_no.Trim
                    Case "02", "03"
                        Select Case Left(selectedData.Series.series_kataban.Trim, 1)
                            Case "A"
                                If Left(selectedData.Series.series_kataban.Trim, 3) = "AB4" Then
                                    If strCoil Is Nothing Then
                                        strSeriesKataban = Left(selectedData.Series.series_kataban, 3)
                                        strKeyKataban = ""
                                    Else
                                        If _
                                            (strCoil.Trim = "3A" Or strCoil.Trim = "3K") And Left(strVoltage, 2) = "DC" Or
                                            (strCoil.Trim = "5A" Or strCoil.Trim = "5K") And Left(strVoltage, 2) = "AC" _
                                            Then
                                            strSeriesKataban = Left(selectedData.Series.series_kataban, 4)
                                            strKeyKataban = ""
                                        Else
                                            strSeriesKataban = Left(selectedData.Series.series_kataban, 3)
                                            strKeyKataban = ""
                                        End If
                                    End If
                                Else
                                    strSeriesKataban = Left(selectedData.Series.series_kataban, 3)
                                    strKeyKataban = ""
                                End If
                            Case "G"
                                If Mid(selectedData.Series.series_kataban.Trim, 2, 3) = "AB4" Then
                                    If strCoil Is Nothing Then
                                        strSeriesKataban = Mid(selectedData.Series.series_kataban, 2, 3)
                                        strKeyKataban = ""
                                    Else
                                        If _
                                            (strCoil.Trim = "3A" Or strCoil.Trim = "3K") And Left(strVoltage, 2) = "DC" Or
                                            (strCoil.Trim = "5A" Or strCoil.Trim = "5K") And Left(strVoltage, 2) = "AC" _
                                            Then
                                            strSeriesKataban = Mid(selectedData.Series.series_kataban, 2, 4)
                                            strKeyKataban = ""
                                        Else
                                            strSeriesKataban = Mid(selectedData.Series.series_kataban, 2, 3)
                                            strKeyKataban = ""
                                        End If
                                    End If
                                Else
                                    strSeriesKataban = Mid(selectedData.Series.series_kataban, 2, 3)
                                    strKeyKataban = ""
                                End If
                        End Select
                    Case Else
                        strSeriesKataban = selectedData.Series.series_kataban
                        strKeyKataban = selectedData.Series.key_kataban
                End Select
            Catch ex As Exception
                intVoltage = 0
                strVoltageDiv = ""
                strSeriesKataban = selectedData.Series.series_kataban
                strKeyKataban = selectedData.Series.key_kataban
                strPortSize = ""
                strCoil = ""
            End Try
        End Function

        ''' <summary>
        '''     異電圧判定
        ''' </summary>
        ''' <param name="strVoltage">電圧</param>
        ''' <param name="strCountryCd">国コード</param>
        ''' <param name="strOfficeCd">営業所コード</param>
        ''' <returns>True:異電圧、Flase:標準電圧</returns>
        ''' <remarks>
        '''     海外店対応
        '''     海外ユーザーの場合、異電圧加算を行わない
        ''' </remarks>
        Public Shared Function GetVoltageIsStandard(strVoltage As String, strCountryCd As String,
                                                    strOfficeCd As String) As Boolean
            '初期値
            GetVoltageIsStandard = True
            '海外ユーザは異電圧加算しない
            If (strCountryCd <> CountryDiv.DefaultCountry) Or
               (strCountryCd = CountryDiv.DefaultCountry And
                strOfficeCd = OfficeDiv.Overseas) Then

                Select Case strVoltage
                    'AC110V/AC220V/AC120V/AC240V　海外では標準電圧扱い
                    Case PowerSupply.Const5, PowerSupply.Const6,
                        PowerSupply.Const7, PowerSupply.Const8
                        '標準電圧
                        GetVoltageIsStandard = False
                End Select
            End If
        End Function

        ''' <summary>
        '''     入力電圧のチェック
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function CheckVoltage(selectedData As SelectedInfo,
                                            strVoltage As String) As Boolean
            Dim result = False
            Dim intVoltage As Integer
            Dim strVoltageDiv As String = Nothing
            Dim strSeriesKataban As String = Nothing
            Dim strKeyKataban As String = Nothing
            Dim strPortSize As String = Nothing
            Dim strCoil As String = Nothing


            '電圧検索情報取得
            Call GetVoltageSearchInfo(selectedData,
                                      strVoltage,
                                      intVoltage,
                                      strVoltageDiv,
                                      strSeriesKataban,
                                      strKeyKataban,
                                      strPortSize,
                                      strCoil)

            '電圧の取得
            Using client As New DbAccessServiceClient
                Dim voltageInfos = client.SelectVoltageInfo(strSeriesKataban,
                                                            strKeyKataban,
                                                            strPortSize,
                                                            strCoil,
                                                            strVoltageDiv,
                                                            intVoltage)

                If voltageInfos.Count > 0 Then

                    If voltageInfos.First.min_voltage = 0 AndAlso voltageInfos.First.max_voltage = 0 Then

                        If intVoltage = voltageInfos.First.std_voltage_flag Then
                            result = True
                        End If
                    Else
                        If _
                            intVoltage >= voltageInfos.First.min_voltage AndAlso
                            intVoltage <= voltageInfos.First.max_voltage Then
                            result = True
                        End If
                    End If

                End If
            End Using

            Return result
        End Function

#End Region

        ''' <summary>
        '''     形番から重複するハイフンを除去する
        ''' </summary>
        ''' <param name="strKataban">形番</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HyphenCut(strKataban As String) As String
            Dim sbKataban As New StringBuilder(60)
            Dim hyphenFlg = False
            Dim intLoopCnt As Integer

            For intLoopCnt = 1 To strKataban.Length
                If Mid(strKataban, intLoopCnt, 1) = MyControlChars.Hyphen Then
                    If hyphenFlg = True Then
                        '1桁前がハイフンの場合は次へ
                    Else
                        '形番生成
                        sbKataban.Append(Mid(strKataban, intLoopCnt, 1))
                    End If

                    'ハイフンフラグＯＮ
                    hyphenFlg = True
                Else
                    '形番生成
                    sbKataban.Append(Mid(strKataban, intLoopCnt, 1))

                    'ハイフンフラグＯＦＦ
                    hyphenFlg = False
                End If
            Next

            HyphenCut = sbKataban.ToString

            '形番の右側がハイフンの場合は除去する
            If Left(HyphenCut, 1) = MyControlChars.Hyphen Then
                HyphenCut = Mid(HyphenCut, 2, HyphenCut.Length)
            End If
            '形番の左側がハイフンの場合は除去する
            If Right(HyphenCut, 1) = MyControlChars.Hyphen Then
                HyphenCut = Left(HyphenCut, HyphenCut.Length - 1)
            End If
        End Function
    End Class
End Namespace