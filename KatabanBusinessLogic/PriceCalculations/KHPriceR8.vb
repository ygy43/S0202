'************************************************************************************
'*  ProgramID  �FKHPriceR8
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2012/04/25   �쐬�ҁFY.Tachi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�������M�����[�^             �q�o�d�P�O�O�O
'*
'************************************************************************************
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceR8

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer = 0
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                       selectedData.Symbols(2)
            decOpAmount(UBound(decOpAmount)) = 1

            '�I�v�V�������Z���i�L�[
            '2016/2/18 �ē��C��
            'If Left(selectedData.Symbols(3).Trim, 1) <> " " Then
            If Left(selectedData.Symbols(3).Trim, 1) <> "" Then
                strOpArray = Split(selectedData.Symbols(3), MyControlChars.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & "-" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

            '�񎟓d�r
            If selectedData.Series.key_kataban.ToString = "4" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 7) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
