Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants

Namespace MyHelpers
    Public Class SessionHelper

        ''' <summary>
        '''     セッションから選択情報を取得
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetSelectedData() As SelectedInfo
            Dim selectedData As New SelectedInfo

            '機種情報
            selectedData.Series = HttpContext.Current.Session(SessionKeys.SelectedSeriesData)

            If selectedData.Series.division = DataTypeDiv.Series Then

                '形番構成情報
                selectedData.KatabanStructures = HttpContext.Current.Session(SessionKeys.KatabanStructureData)
                '構成選択情報
                selectedData.Symbols = HttpContext.Current.Session(SessionKeys.SelectedStructureData)
                'ロッド先端情報
                selectedData.RodEnd.RodEndOption = IIf(HttpContext.Current.Session(SessionKeys.SelectedRodEndOption) Is Nothing, String.Empty,
                                                       HttpContext.Current.Session(SessionKeys.SelectedRodEndOption))
                'オプション外情報
                selectedData.OtherOption = IIf(HttpContext.Current.Session(SessionKeys.SelectedOtherOption) Is Nothing, String.Empty,
                                               HttpContext.Current.Session(SessionKeys.SelectedOtherOption))

            ElseIf selectedData.Series.division = DataTypeDiv.FullKataban Then
                'フル形番の場合

            End If

            Return selectedData
        End Function

    End Class
End Namespace