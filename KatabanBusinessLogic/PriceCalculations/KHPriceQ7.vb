'************************************************************************************
'*  ProgramID  �FKHPriceQ7
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/04/06   �쐬�ҁFT.Yagyu
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FFSL�V���[�Y
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceQ7

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
            strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & _
            selectedData.Symbols(1) & MyControlChars.Hyphen & _
            selectedData.Symbols(2)
            decOpAmount(UBound(decOpAmount)) = 1

            '�I�v�V�������Z���i�L�[
            If selectedData.Symbols(3).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban, 3) & MyControlChars.Hyphen & _
                selectedData.Symbols(3)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

