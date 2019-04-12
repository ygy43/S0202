'************************************************************************************
'*  ProgramID  �FKHPriceP6
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2008/06/09   �쐬�ҁFM.Kojima
'*
'*  �T�v       �F�����h�~�t�G���V�����_�@UFCD�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceP6
    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim strBoreSize As String           '���a
        Dim strStroke As String             '�X�g���[�N

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            strBoreSize = selectedData.Symbols(1).Trim
            strStroke = selectedData.Symbols(3).Trim

            '�σX�g���[�N�ݒ�@
            intStroke = _
                KatabanUtility.GetStrokeSize(selectedData, _
                    CInt(strBoreSize), CInt(strStroke))

            '��{���i�L�[�̐ݒ�
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            '�}�O�l�b�g���(L)���Z
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                "L"
            decOpAmount(UBound(decOpAmount)) = 1

            '�X�C�b�`���Z���i�L�[

            If selectedData.Symbols(5).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                    selectedData.Symbols(5).Trim & MyControlChars.Hyphen & _
                    selectedData.Symbols(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '���[�h���������Z���i�L�[
                If selectedData.Symbols(6).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        selectedData.Series.series_kataban.Trim & _
                        MyControlChars.Hyphen & _
                        selectedData.Symbols(6).Trim
                    decOpAmount(UBound(decOpAmount)) = KatabanUtility.SwitchQtyGet(selectedData.Symbols(7).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
