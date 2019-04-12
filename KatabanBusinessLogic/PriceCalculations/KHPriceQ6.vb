'************************************************************************************
'*  ProgramID  �FKHPriceQ6
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/03/05   �쐬�ҁFT.Yagyu
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FSFR�ASFRT�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceQ6

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strSw As String = ""

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            Dim bolC5Flag As Boolean

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
            selectedData.Symbols(1) & MyControlChars.Hyphen & _
            selectedData.Symbols(2)
            decOpAmount(UBound(decOpAmount)) = 1

            '2011/10/24 ADD RM1110032(11��VerUP:�񎟓d�r) START--->
            Select Case selectedData.Series.key_kataban
                Case ""
                    '2011/10/24 ADD RM1110032(11��VerUP:�񎟓d�r) <---END
                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(3).Trim <> "" Then
                        If (selectedData.Symbols(5) = "R") Or (selectedData.Symbols(5) = "L") Then
                            strSw = "S"
                        Else
                            strSw = "D"
                        End If
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & MyControlChars.Hyphen & _
                        selectedData.Symbols(3) & selectedData.Symbols(4).Trim & MyControlChars.Hyphen & _
                        strSw & MyControlChars.Hyphen & selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "2"
                    Dim intSu As Integer
                    ReDim strPriceDiv(0)
                    '��{���i�L�[�p
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    'C5�`�F�b�N
                    bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(selectedData, False)

                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5
                    End If

                    '�I�v�V�������Z���i�L�[
                    If selectedData.Symbols(3).Trim <> "" Then
                        If (selectedData.Symbols(4) = "R") Or (selectedData.Symbols(4) = "L") Then
                            strSw = "S"
                            intSu = 1
                        Else
                            strSw = "D"
                            intSu = 2
                        End If
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & MyControlChars.Hyphen & _
                                                            selectedData.Symbols(3) & MyControlChars.Hyphen & _
                                                            strSw & MyControlChars.Hyphen & selectedData.Symbols(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & MyControlChars.Hyphen & _
                                                                "SW" & MyControlChars.Hyphen & selectedData.Symbols(5)
                        decOpAmount(UBound(decOpAmount)) = intSu
                    End If

                    '�񎟓d�r���Z
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & MyControlChars.Hyphen & _
                                                            selectedData.Symbols(5)
                    decOpAmount(UBound(decOpAmount)) = 1
                    'strPriceDiv(UBound(strPriceDiv)) = AccumulatePriceDiv.C5

            End Select
            '2011/10/24 ADD RM1110032(11��VerUP:�񎟓d�r) <---END

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

