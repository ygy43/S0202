Imports KatabanCommon.Constants
Imports S0202.My.Resources

Namespace MyHelpers
    Public Class LanguageHelper

        Public Shared LanguageList As IEnumerable(Of SelectListItem) = New List(Of SelectListItem) From {
            New SelectListItem With {.Value = LanguageDiv.DefaultLang, .Text = RLayout.English},
            New SelectListItem With {.Value = LanguageDiv.SimplifiedChinese, .Text = RLayout.SimplifiedChinese},
            New SelectListItem With {.Value = LanguageDiv.TraditionalChinese, .Text = RLayout.TraditionalChinese},
            New SelectListItem With {.Value = LanguageDiv.Japanese, .Text = RLayout.Japanese},
            New SelectListItem With {.Value = LanguageDiv.Korean, .Text = RLayout.Korean}
            }
    End Class
End Namespace