Imports KatabanCommon.Constants

Namespace ViewModels.Options
    ''' <summary>
    '''     ロッド先端表示情報
    ''' </summary>
    Public Class RodEndViewModel
        Public Sub New()
            Me.SelectedPatternSymbol = String.Empty
            Me.SeriesName = String.Empty
            Me.RodEndUnitInfos = New List(Of RodEndUnitViewModel)
        End Sub

        ''' <summary>
        '''     選択したロッド先端情報
        ''' </summary>
        Public Property SelectedPatternSymbol As String

        ''' <summary>
        '''     機種名称
        ''' </summary>
        Public Property SeriesName As String

        ''' <summary>
        '''     ロッド先端情報
        ''' </summary>
        Public Property RodEndUnitInfos As List(Of RodEndUnitViewModel)

    End Class

End Namespace