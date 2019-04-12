'************************************************************************************
'*  ProgramID  �FKHPriceR2
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2010/03/26   �쐬�ҁFY.Miura
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FSCPS�V���[�Y  (�y���V���V�����_)
'*             �FZSF�V���[�Y   (PP�^�C�v �j���[�W���C���g)
'*             �FSC3F�V���[�Y  (PP�^�C�v �X�s�[�h�R���g���[��)
'*
'************************************************************************************
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Module KHPriceR2

    Public Sub subPriceCalculation(selectedData As SelectedInfo, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer = 0

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '�X�g���[�N�ݒ�
            intStroke = KatabanUtility.GetStrokeSize(selectedData, _
                                                  CInt(selectedData.Symbols(2).Trim), _
                                                  CInt(selectedData.Symbols(3).Trim))
            If selectedData.Series.series_kataban.Trim <> "ZSF" Or _
               selectedData.Series.series_kataban.Trim <> "SC3F" Then
                '��{���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = selectedData.Series.series_kataban.Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2) & MyControlChars.Hyphen & _
                                                           intStroke
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '��RM1312XXX 2013/11/28 �C��
            If selectedData.Series.series_kataban.Trim = "ZSF" Then

                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            If selectedData.Series.series_kataban.Trim = "SC3F" Then

                strOpRefKataban(UBound(strOpRefKataban)) = Left(selectedData.Series.series_kataban.Trim, 9) & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(1).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(2).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(3).Trim & MyControlChars.Hyphen & _
                                                           selectedData.Symbols(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

