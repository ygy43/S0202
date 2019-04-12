'************************************************************************************
'*  ProgramID  �FKHPriceS1
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2012/09/27   �쐬�ҁFY.Tachi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�y���V���V�����_�����`�@�r�b�o�c�R�^�r�b�o�c�R�|�k
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceS1

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intIndex As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionP4 As Boolean = False



        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)


            '�X�g���[�N�擾
            'intStroke = objPrice.fncGetStrokeSize(selectedData.Series.series_kataban, _
            '                                      selectedData.Series.key_kataban, _
            '                                      CInt(selectedData.Symbols(3).Trim), _
            '                                      CInt(selectedData.Symbols(4).Trim))

            Select Case Left(selectedData.Series.series_kataban.Trim, 5)
                Case "SCPS3" 'SCPS3
                    Select Case Mid(selectedData.Series.series_kataban.Trim, 7, 1)
                        Case "M" 'SCPS3-M
                            Select Case True
                                Case CInt(selectedData.Symbols(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 100
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                                Case CInt(selectedData.Symbols(3).Trim) >= 101
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            Select Case True
                                Case CInt(selectedData.Symbols(3).Trim) <= 10
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-10"
                                Case CInt(selectedData.Symbols(3).Trim) >= 11 And CInt(selectedData.Symbols(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 20
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-20"
                                Case CInt(selectedData.Symbols(3).Trim) >= 21 And CInt(selectedData.Symbols(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 105
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-105"
                                Case CInt(selectedData.Symbols(3).Trim) >= 106 And CInt(selectedData.Symbols(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "SCPD3" 'SCPD3-D
                    Select Case Mid(selectedData.Series.series_kataban.Trim, 7, 1)
                        Case "D"
                            Select Case True
                                Case CInt(selectedData.Symbols(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 100
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                                Case CInt(selectedData.Symbols(3).Trim) >= 101
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "T" 'SCPD3-T
                            Select Case True
                                Case CInt(selectedData.Symbols(3).Trim) <= 10
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-10"
                                Case CInt(selectedData.Symbols(3).Trim) >= 11 And CInt(selectedData.Symbols(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 20
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-20"
                                Case CInt(selectedData.Symbols(3).Trim) >= 21 And CInt(selectedData.Symbols(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 100
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                                Case CInt(selectedData.Symbols(3).Trim) >= 101 And CInt(selectedData.Symbols(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                                Case CInt(selectedData.Symbols(3).Trim) >= 121 And CInt(selectedData.Symbols(3).Trim) <= 135
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-135"
                                Case CInt(selectedData.Symbols(3).Trim) >= 136 And CInt(selectedData.Symbols(3).Trim) <= 150
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-150"
                                Case CInt(selectedData.Symbols(3).Trim) >= 151 And CInt(selectedData.Symbols(3).Trim) <= 165
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-165"
                                Case CInt(selectedData.Symbols(3).Trim) >= 166 And CInt(selectedData.Symbols(3).Trim) <= 180
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-180"
                                Case CInt(selectedData.Symbols(3).Trim) >= 181 And CInt(selectedData.Symbols(3).Trim) <= 195
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-195"
                                Case CInt(selectedData.Symbols(3).Trim) >= 196 And CInt(selectedData.Symbols(3).Trim) <= 210
                                    Select Case selectedData.Symbols(2).Trim
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-200"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-210"
                                    End Select
                                Case CInt(selectedData.Symbols(3).Trim) >= 211 And CInt(selectedData.Symbols(3).Trim) <= 225
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-225"
                                Case CInt(selectedData.Symbols(3).Trim) >= 226 And CInt(selectedData.Symbols(3).Trim) <= 240
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-240"
                                Case CInt(selectedData.Symbols(3).Trim) >= 241 And CInt(selectedData.Symbols(3).Trim) <= 255
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-255"
                                Case CInt(selectedData.Symbols(3).Trim) >= 256
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-260"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "O" 'SCPD3-O
                            Select Case True
                                Case CInt(selectedData.Symbols(3).Trim) <= 10
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-10"
                                Case CInt(selectedData.Symbols(3).Trim) >= 11 And CInt(selectedData.Symbols(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 20
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-20"
                                Case CInt(selectedData.Symbols(3).Trim) >= 21 And CInt(selectedData.Symbols(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 105
                                    Select Case selectedData.Symbols(2).Trim
                                        Case "6"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-105"
                                    End Select
                                Case CInt(selectedData.Symbols(3).Trim) >= 106 And CInt(selectedData.Symbols(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                                Case CInt(selectedData.Symbols(3).Trim) >= 121 And CInt(selectedData.Symbols(3).Trim) <= 135
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-135"
                                Case CInt(selectedData.Symbols(3).Trim) >= 136 And CInt(selectedData.Symbols(3).Trim) <= 150
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-150"
                                Case CInt(selectedData.Symbols(3).Trim) >= 151 And CInt(selectedData.Symbols(3).Trim) <= 165
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-165"
                                Case CInt(selectedData.Symbols(3).Trim) >= 166 And CInt(selectedData.Symbols(3).Trim) <= 180
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-180"
                                Case CInt(selectedData.Symbols(3).Trim) >= 181 And CInt(selectedData.Symbols(3).Trim) <= 195
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-195"
                                Case CInt(selectedData.Symbols(3).Trim) >= 196 And CInt(selectedData.Symbols(3).Trim) <= 210
                                    Select Case selectedData.Symbols(2).Trim
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-200"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-210"
                                    End Select
                                Case CInt(selectedData.Symbols(3).Trim) >= 211 And CInt(selectedData.Symbols(3).Trim) <= 225
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-225"
                                Case CInt(selectedData.Symbols(3).Trim) >= 226 And CInt(selectedData.Symbols(3).Trim) <= 240
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-240"
                                Case CInt(selectedData.Symbols(3).Trim) >= 241 And CInt(selectedData.Symbols(3).Trim) <= 255
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-255"
                                Case CInt(selectedData.Symbols(3).Trim) >= 256
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-260"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "M", "Z", "K" 'SCPD3-M
                            Select Case True
                                Case CInt(selectedData.Symbols(3).Trim) <= 15
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 30
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 105
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                                Case CInt(selectedData.Symbols(3).Trim) >= 106 And CInt(selectedData.Symbols(3).Trim) <= 120
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                                Case CInt(selectedData.Symbols(3).Trim) >= 121 And CInt(selectedData.Symbols(3).Trim) <= 135
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-135"
                                Case CInt(selectedData.Symbols(3).Trim) >= 136 And CInt(selectedData.Symbols(3).Trim) <= 150
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-150"
                                Case CInt(selectedData.Symbols(3).Trim) >= 151 And CInt(selectedData.Symbols(3).Trim) <= 165
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-165"
                                Case CInt(selectedData.Symbols(3).Trim) >= 166 And CInt(selectedData.Symbols(3).Trim) <= 180
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-180"
                                Case CInt(selectedData.Symbols(3).Trim) >= 181 And CInt(selectedData.Symbols(3).Trim) <= 195
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-195"
                                Case CInt(selectedData.Symbols(3).Trim) >= 196 And CInt(selectedData.Symbols(3).Trim) <= 210
                                    Select Case selectedData.Symbols(2).Trim
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-200"
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-210"
                                    End Select
                                Case CInt(selectedData.Symbols(3).Trim) >= 211 And CInt(selectedData.Symbols(3).Trim) <= 225
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-225"
                                Case CInt(selectedData.Symbols(3).Trim) >= 226 And CInt(selectedData.Symbols(3).Trim) <= 240
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-240"
                                Case CInt(selectedData.Symbols(3).Trim) >= 241 And CInt(selectedData.Symbols(3).Trim) <= 255
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-255"
                                Case CInt(selectedData.Symbols(3).Trim) >= 256
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 7) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-260"
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "F", "L"
                            If selectedData.Symbols(3).Trim = "C" Then
                                Select Case selectedData.Symbols(2).Trim
                                    Case "6"
                                        Select Case True
                                            Case CInt(selectedData.Symbols(4).Trim) <= 15
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-15"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 16 And CInt(selectedData.Symbols(4).Trim) <= 30
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-30"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 31 And CInt(selectedData.Symbols(4).Trim) <= 45
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-45"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 46 And CInt(selectedData.Symbols(4).Trim) <= 60
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-60"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 61 And CInt(selectedData.Symbols(4).Trim) <= 70
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-70"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 71 And CInt(selectedData.Symbols(4).Trim) <= 80
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-80"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 81 And CInt(selectedData.Symbols(4).Trim) <= 90
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-90"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 91
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-100"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        '2013/11/06 ���i�L�[�C��
                                        Select Case True
                                            Case CInt(selectedData.Symbols(4).Trim) <= 15
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-15"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 16 And CInt(selectedData.Symbols(4).Trim) <= 30
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-30"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 31 And CInt(selectedData.Symbols(4).Trim) <= 45
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-45"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 46 And CInt(selectedData.Symbols(4).Trim) <= 60
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-60"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 61 And CInt(selectedData.Symbols(4).Trim) <= 75
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-75"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 76 And CInt(selectedData.Symbols(4).Trim) <= 90
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-90"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 91 And CInt(selectedData.Symbols(4).Trim) <= 100
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-100"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 101 And CInt(selectedData.Symbols(4).Trim) <= 110
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-110"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 111 And CInt(selectedData.Symbols(4).Trim) <= 120
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-120"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 121 And CInt(selectedData.Symbols(4).Trim) <= 130
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-130"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 131 And CInt(selectedData.Symbols(4).Trim) <= 140
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-140"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 141 And CInt(selectedData.Symbols(4).Trim) <= 150
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-150"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 151 And CInt(selectedData.Symbols(4).Trim) <= 160
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-160"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 161 And CInt(selectedData.Symbols(4).Trim) <= 170
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-170"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 171 And CInt(selectedData.Symbols(4).Trim) <= 180
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-180"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 181 And CInt(selectedData.Symbols(4).Trim) <= 190
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-190"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 191 And CInt(selectedData.Symbols(4).Trim) <= 200
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-200"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 201 And CInt(selectedData.Symbols(4).Trim) <= 210
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-210"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 211 And CInt(selectedData.Symbols(4).Trim) <= 220
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-220"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 221 And CInt(selectedData.Symbols(4).Trim) <= 230
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-230"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 231 And CInt(selectedData.Symbols(4).Trim) <= 240
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-240"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 241 And CInt(selectedData.Symbols(4).Trim) <= 250
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-250"
                                            Case CInt(selectedData.Symbols(4).Trim) >= 251
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-260"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                Select Case True
                                    Case CInt(selectedData.Symbols(3).Trim) <= 10
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-10"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 11 And CInt(selectedData.Symbols(3).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 20
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-20"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 21 And CInt(selectedData.Symbols(3).Trim) <= 30
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 105
                                        Select Case selectedData.Symbols(2).Trim
                                            Case "6"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-100"
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-105"
                                        End Select
                                    Case CInt(selectedData.Symbols(3).Trim) >= 106 And CInt(selectedData.Symbols(3).Trim) <= 120
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 121 And CInt(selectedData.Symbols(3).Trim) <= 135
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-135"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 136 And CInt(selectedData.Symbols(3).Trim) <= 150
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-150"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 151 And CInt(selectedData.Symbols(3).Trim) <= 165
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-165"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 166 And CInt(selectedData.Symbols(3).Trim) <= 180
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-180"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 181 And CInt(selectedData.Symbols(3).Trim) <= 195
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-195"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 196 And CInt(selectedData.Symbols(3).Trim) <= 210
                                        Select Case selectedData.Symbols(2).Trim
                                            Case "10"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-200"
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-210"
                                        End Select
                                    Case CInt(selectedData.Symbols(3).Trim) >= 211 And CInt(selectedData.Symbols(3).Trim) <= 225
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-225"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 226 And CInt(selectedData.Symbols(3).Trim) <= 240
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-240"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 241 And CInt(selectedData.Symbols(3).Trim) <= 255
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-255"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 256 And CInt(selectedData.Symbols(3).Trim) <= 260
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-260"
                                    Case CInt(selectedData.Symbols(3).Trim) >= 261
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-270"
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1

                                Select Case selectedData.Symbols(2).Trim '�`���[�u��a����
                                    Case "6"
                                        Select Case True
                                            Case CInt(selectedData.Symbols(3).Trim) <= 30
                                                '�X�g���[�N�P�O�`�R�O
                                                '��"SCPD3-F-6-STR10-30"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR10-30"
                                            Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 60
                                                '�X�g���[�N�R�P�`�U�O
                                                '��"SCPD3-F-6-STR31-60"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR31-60"
                                            Case CInt(selectedData.Symbols(3).Trim) >= 61
                                                '�X�g���[�N�U�P�`�P�O�O
                                                '��"SCPD3-F-6-STR61-100"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR61-100"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "10"
                                        Select Case True
                                            Case CInt(selectedData.Symbols(3).Trim) <= 45
                                                '�X�g���[�N�P�O�`�S�T
                                                '��"SCPD3-F-10-STR10-45"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR10-45"
                                            Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 100
                                                '�X�g���[�N�S�U�`�P�O�O
                                                '��"SCPD3-F-10-STR46-100"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR46-100"
                                            Case CInt(selectedData.Symbols(3).Trim) >= 101
                                                '�X�g���[�N�P�O�P�`�Q�O�O
                                                '��"SCPD3-F-10-STR101-200"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR101-200"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "16"
                                        Select Case True
                                            Case CInt(selectedData.Symbols(3).Trim) <= 45
                                                '�X�g���[�N�P�O�`�S�T
                                                '��"SCPD3-F-16-STR10-45"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR10-45"
                                            Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 100
                                                '�X�g���[�N�S�U�`�P�O�O
                                                '��"SCPD3-F-16-STR46-100"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR46-100"
                                            Case CInt(selectedData.Symbols(3).Trim) >= 101
                                                '�X�g���[�N�P�O�P�`�Q�U�O
                                                '��"SCPD3-F-16-STR101-260"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & "-F-" & selectedData.Symbols(2).Trim & "-STR101-260"
                                        End Select
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                        Case Else
                            'Select Case Mid(selectedData.Series.series_kataban.Trim, 10, 1)
                            Select Case selectedData.Series.key_kataban.Trim
                                Case "C"
                                    Select Case selectedData.Symbols(2).Trim
                                        Case "6"
                                            Select Case True
                                                Case CInt(selectedData.Symbols(4).Trim) <= 15
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-15"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 16 And CInt(selectedData.Symbols(4).Trim) <= 30
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-30"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 31 And CInt(selectedData.Symbols(4).Trim) <= 45
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-45"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 46 And CInt(selectedData.Symbols(4).Trim) <= 60
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-60"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 61 And CInt(selectedData.Symbols(4).Trim) <= 70
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-70"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 71 And CInt(selectedData.Symbols(4).Trim) <= 80
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-80"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 81 And CInt(selectedData.Symbols(4).Trim) <= 90
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-90"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 91
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-100"
                                            End Select
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case Else
                                            Select Case True
                                                Case CInt(selectedData.Symbols(4).Trim) <= 15
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-15"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 16 And CInt(selectedData.Symbols(4).Trim) <= 30
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-30"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 31 And CInt(selectedData.Symbols(4).Trim) <= 45
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-45"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 46 And CInt(selectedData.Symbols(4).Trim) <= 60
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-60"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 61 And CInt(selectedData.Symbols(4).Trim) <= 75
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-75"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 76 And CInt(selectedData.Symbols(4).Trim) <= 90
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-90"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 91 And CInt(selectedData.Symbols(4).Trim) <= 100
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-100"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 101 And CInt(selectedData.Symbols(4).Trim) <= 110
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-110"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 111 And CInt(selectedData.Symbols(4).Trim) <= 120
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-120"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 121 And CInt(selectedData.Symbols(4).Trim) <= 130
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-130"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 131 And CInt(selectedData.Symbols(4).Trim) <= 140
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-140"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 141 And CInt(selectedData.Symbols(4).Trim) <= 150
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-150"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 151 And CInt(selectedData.Symbols(4).Trim) <= 160
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-160"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 161 And CInt(selectedData.Symbols(4).Trim) <= 170
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-170"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 171 And CInt(selectedData.Symbols(4).Trim) <= 180
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-180"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 181 And CInt(selectedData.Symbols(4).Trim) <= 190
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-190"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 191 And CInt(selectedData.Symbols(4).Trim) <= 200
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-200"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 201 And CInt(selectedData.Symbols(4).Trim) <= 210
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-210"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 211 And CInt(selectedData.Symbols(4).Trim) <= 220
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-220"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 221 And CInt(selectedData.Symbols(4).Trim) <= 230
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-230"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 231 And CInt(selectedData.Symbols(4).Trim) <= 240
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-240"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 241 And CInt(selectedData.Symbols(4).Trim) <= 250
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-250"
                                                Case CInt(selectedData.Symbols(4).Trim) >= 251
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "C-260"
                                            End Select
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select
                Case "SCPH3"    'SCPH3
                    Select Case True
                        Case CInt(selectedData.Symbols(3).Trim) <= 10
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-10"
                        Case CInt(selectedData.Symbols(3).Trim) >= 11 And CInt(selectedData.Symbols(3).Trim) <= 15
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-15"
                        Case CInt(selectedData.Symbols(3).Trim) >= 16 And CInt(selectedData.Symbols(3).Trim) <= 20
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-20"
                        Case CInt(selectedData.Symbols(3).Trim) >= 21 And CInt(selectedData.Symbols(3).Trim) <= 30
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-30"
                        Case CInt(selectedData.Symbols(3).Trim) >= 31 And CInt(selectedData.Symbols(3).Trim) <= 45
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-45"
                        Case CInt(selectedData.Symbols(3).Trim) >= 46 And CInt(selectedData.Symbols(3).Trim) <= 60
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-60"
                        Case CInt(selectedData.Symbols(3).Trim) >= 61 And CInt(selectedData.Symbols(3).Trim) <= 75
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-75"
                        Case CInt(selectedData.Symbols(3).Trim) >= 76 And CInt(selectedData.Symbols(3).Trim) <= 90
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-90"
                        Case CInt(selectedData.Symbols(3).Trim) >= 91 And CInt(selectedData.Symbols(3).Trim) <= 105
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-105"
                        Case CInt(selectedData.Symbols(3).Trim) >= 106
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & selectedData.Symbols(2).Trim & "-120"
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select


            ''�o���G�[�V����(����)���Z���i�L�[
            'Select Case selectedData.Symbols(1).Trim
            '    Case "F"
            '        Select Case selectedData.Symbols(3).Trim
            '            Case "6"
            '                Select Case True
            '                    Case CInt(selectedData.Symbols(4).Trim) <= 30
            '                        '�X�g���[�N10�`30
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR10" & MyControlChars.Hyphen & "30"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(selectedData.Symbols(4).Trim) >= 31 And _
            '                         CInt(selectedData.Symbols(4).Trim) <= 60
            '                        '�X�g���[�N31�`60
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR31" & MyControlChars.Hyphen & "60"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(selectedData.Symbols(4).Trim) >= 61
            '                        '�X�g���[�N61�`100
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR61" & MyControlChars.Hyphen & "100"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                End Select
            '                decOpAmount(UBound(decOpAmount)) = 1
            '            Case "10"
            '                Select Case True
            '                    Case CInt(selectedData.Symbols(4).Trim) <= 45
            '                        '�X�g���[�N10�`45
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR10" & MyControlChars.Hyphen & "45"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(selectedData.Symbols(4).Trim) >= 46 And _
            '                         CInt(selectedData.Symbols(4).Trim) <= 100
            '                        '�X�g���[�N46�`100
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR46" & MyControlChars.Hyphen & "100"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(selectedData.Symbols(4).Trim) >= 101
            '                        '�X�g���[�N101�`200
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR101" & MyControlChars.Hyphen & "200"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                End Select
            '                decOpAmount(UBound(decOpAmount)) = 1
            '            Case "16"
            '                Select Case True
            '                    Case CInt(selectedData.Symbols(4).Trim) <= 45
            '                        '�X�g���[�N10�`45
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR10" & MyControlChars.Hyphen & "45"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(selectedData.Symbols(4).Trim) >= 46 And _
            '                         CInt(selectedData.Symbols(4).Trim) <= 100
            '                        '�X�g���[�N46�`100
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR46" & MyControlChars.Hyphen & "100"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                    Case CInt(selectedData.Symbols(4).Trim) >= 101
            '                        '�X�g���[�N101�`260
            '                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 5) & MyControlChars.Hyphen & "F" & MyControlChars.Hyphen & _
            '                                                                   selectedData.Symbols(3).Trim & MyControlChars.Hyphen & "STR101" & MyControlChars.Hyphen & "260"
            '                        decOpAmount(UBound(decOpAmount)) = 1
            '                End Select
            '                decOpAmount(UBound(decOpAmount)) = 1
            '        End Select
            'End Select

            'ϸ�ȯĉ��Z���i�L�[
            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or _
               Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3-L"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '�x���`�����Z���i�L�[
            If Left(selectedData.Symbols(1).Trim, 2) = "CB" Or _
               Left(selectedData.Symbols(1).Trim, 2) = "FA" Or _
               Left(selectedData.Symbols(1).Trim, 2) = "LB" Or _
               Left(selectedData.Symbols(1).Trim, 2) = "LS" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & selectedData.Symbols(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If


            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                Select Case True
                    Case Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "D"
                        intIndex = 4
                    Case Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "F"
                        intIndex = 4
                        'Case Mid(selectedData.Series.series_kataban.Trim, 10, 1) = "C"
                    Case selectedData.Series.key_kataban.Trim = "C"
                        intIndex = 6
                    Case Else
                        intIndex = 5
                End Select
                If selectedData.Symbols(intIndex).Trim <> "" Then
                    '�X�C�b�`���Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & selectedData.Symbols(intIndex).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)

                    If selectedData.Symbols(intIndex + 1).Trim <> "" Then
                        '���[�h���������Z���i�L�[
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & selectedData.Symbols(intIndex + 1).Trim
                        decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(intIndex + 2).Trim)
                    End If
                End If
                'intIndex = intIndex + 2
            End If

            '�I�v�V�����E�t���i���Z���i�L�[
            intIndex = 0
            Select Case Left(selectedData.Series.series_kataban.Trim, 5)
                Case "SCPD3"
                    Select Case selectedData.Series.key_kataban
                        Case "F"
                            '�H�i�����H���������i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                            strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                            Select Case Mid(selectedData.Series.series_kataban.Trim, 7, 1)
                                Case "D"
                                    If Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                                        If Len(selectedData.Symbols(8)) <> 0 Then
                                            intIndex = 8
                                        End If
                                    Else
                                        If Len(selectedData.Symbols(5)) <> 0 Then
                                            intIndex = 5
                                        End If
                                    End If
                                Case "Z", "K", "M"
                                    If Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                                        If Len(selectedData.Symbols(9)) <> 0 Then
                                            intIndex = 9
                                        End If
                                    Else
                                        If Len(selectedData.Symbols(6)) <> 0 Then
                                            intIndex = 6
                                        End If
                                    End If
                                Case Else
                            End Select

                        Case Else
                            Select Case Mid(selectedData.Series.series_kataban.Trim, 7, 1)
                                Case "D"
                                    If Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                                        If Len(selectedData.Symbols(7)) <> 0 Then
                                            intIndex = 7
                                        End If
                                    Else
                                        If Len(selectedData.Symbols(4)) <> 0 Then
                                            intIndex = 4
                                        End If
                                    End If
                                Case "F", "L"
                                    If selectedData.Symbols(3).Trim = "C" Then
                                        If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                                            If Len(selectedData.Symbols(9)) <> 0 Then
                                                intIndex = 9
                                            End If
                                        Else
                                            If Len(selectedData.Symbols(6)) <> 0 Then
                                                intIndex = 6
                                            End If
                                        End If
                                    Else
                                        If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                                            If Len(selectedData.Symbols(7)) <> 0 Then
                                                intIndex = 7
                                            End If
                                        Else
                                            If Len(selectedData.Symbols(4)) <> 0 Then
                                                intIndex = 4
                                            End If
                                        End If
                                    End If
                                Case Else
                                    'Select Case Mid(selectedData.Series.series_kataban.Trim, 10, 1)
                                    Select Case selectedData.Series.key_kataban.Trim
                                        Case "C"
                                            If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Then
                                                If Len(selectedData.Symbols(9)) <> 0 Then
                                                    intIndex = 9
                                                End If
                                            Else
                                                If Len(selectedData.Symbols(6)) <> 0 Then
                                                    intIndex = 6
                                                End If
                                            End If
                                        Case Else
                                            If Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                                                If Len(selectedData.Symbols(8)) <> 0 Then
                                                    intIndex = 8
                                                End If
                                            Else
                                                If Len(selectedData.Symbols(5)) <> 0 Then
                                                    intIndex = 5
                                                End If
                                            End If
                                    End Select
                            End Select
                    End Select

                Case Else
                    If Mid(selectedData.Series.series_kataban.Trim, 7, 1) = "L" Or _
                       Mid(selectedData.Series.series_kataban.Trim, 8, 1) = "L" Then
                        If Len(selectedData.Symbols(8)) <> 0 Then
                            intIndex = 8
                        End If
                    Else
                        If Len(selectedData.Symbols(5)) <> 0 Then
                            intIndex = 5
                        End If
                    End If
            End Select

            If intIndex <> 0 Then
                strOpArray = Split(selectedData.Symbols(intIndex), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    If Left(selectedData.Series.series_kataban.Trim, 7) = "SCPD3-D" Then
                        decOpAmount(UBound(decOpAmount)) = 2
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Next
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
