'************************************************************************************
'*  ProgramID  �FKHPriceO2
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/04/18   �쐬�ҁFNII A.Tatakashi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�G�A�n�C�h���u�[�X�^ �`�g�a�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceO2

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)
        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(4).Trim
            decOpAmount(UBound(decOpAmount)) = 1

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
