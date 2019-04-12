Imports KatabanCommon.Constants

Namespace ViewModels.Account
    Public Class LayoutViewModel
        Public Sub New()
            Me.SelectedLanguage = Divisions.LanguageDiv.DefaultLang
            Languages = New List(Of SelectListItem) From {
                new SelectListItem With {.Value = LanguageDiv.DefaultLang,.Text = LanguageDiv.DefaultLang}, 
                new SelectListItem With {.Value = LanguageDiv.Japanese,.Text = LanguageDiv.Japanese}, 
                new SelectListItem With {.Value = LanguageDiv.Korean,.Text = LanguageDiv.Korean}, 
                new SelectListItem With {.Value = LanguageDiv.SimplifiedChinese,.Text = LanguageDiv.SimplifiedChinese}, 
                new SelectListItem With {.Value = LanguageDiv.TraditionalChinese,.Text = LanguageDiv.TraditionalChinese}
                }
        End Sub

        ''' <summary>
        '''     選択した言語
        ''' </summary>
        Public Property SelectedLanguage As String

        ''' <summary>
        '''     言語リスト
        ''' </summary>
        Public Property Languages As List(Of SelectListItem)
    End Class
End Namespace