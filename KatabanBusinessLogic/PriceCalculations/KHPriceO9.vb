'************************************************************************************
'*  ProgramID  �FKHPriceO9
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2007/04/26   �쐬�ҁFNII A.Takahashi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�J�M���t�c���r�o��   �u�U�O�P�O�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceO9

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '�I�v�V�������Z���i�L�[
            strOpArray = Split(selectedData.Symbols(4), MyControlChars.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           strOpArray(intLoopCnt).Trim()
                decOpAmount(UBound(decOpAmount)) = 1

            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
