Imports KatabanBusinessLogic.KatabanWcfService
Imports PagedList

Namespace ViewModels.Series
    Public Class SeriesSearchViewModel

        ''' <summary>
        '''     入力した機種
        ''' </summary>
        Public Property Series As String

        ''' <summary>
        '''     検索タイプ
        ''' </summary>
        Public Property SearchType As String

        ''' <summary>
        '''     検索結果
        ''' </summary>
        Public Property SearchResults As StaticPagedList(of SeriesInfo)

        '''' <summary>
        ''''     選択された機種
        '''' </summary>
        'Public Property SelectedSeries As String

        '''' <summary>
        ''''     選択されたキー形番
        '''' </summary>
        'Public Property SelectedKeyKataban As String
    End Class
End Namespace