Imports KatabanBusinessLogic.KatabanWcfService

Namespace ViewModels.Options
    Public Class OptionsIndexViewModel
        ''' <summary>
        '''     選択した機種情報
        ''' </summary>
        Public Property SelectedSeriesInfo As SeriesInfo

        ''' <summary>
        '''     選択した構成情報
        ''' </summary>
        Public Property SelectedStructureInfos As List(Of String)

        ''' <summary>
        '''     形番構成情報
        ''' </summary>
        Public Property KatabanStructureInfos As List(Of KatabanStructureInfo)

        ''' <summary>
        '''     フォカスされた構成番号
        ''' </summary>
        Public Property FocusSeqNo As Integer

        ''' <summary>
        '''     警告メッセージ
        ''' </summary>
        Public Property Messages As List(Of String)

        ''' <summary>
        '''     ロッド先端特注表示フラグ
        ''' </summary>
        Public Property IsShowRodEnd As Boolean

        ''' <summary>
        '''     オプション外表示フラグ
        ''' </summary>
        Public Property IsShowOtherOption As Boolean

        ''' <summary>
        '''     ストッパ位置表示フラグ
        ''' </summary>
        Public Property IsShowStopper As Boolean

        ''' <summary>
        '''     取付モータ仕様表示フラグ
        ''' </summary>
        Public Property IsShowMotor1 As Boolean

        ''' <summary>
        '''     取付方向/取付モータ仕様表示フラグ
        ''' </summary>
        Public Property IsShowMotor2 As Boolean

        ''' <summary>
        '''     操作ポート位置表示フラグ
        ''' </summary>
        Public Property IsShowPortPosition As Boolean

        ''' <summary>
        '''     在庫本数一覧表示フラグ
        ''' </summary>
        Public Property IsShowStock As Boolean

        ''' <summary>
        '''     コンストラクタ
        ''' </summary>
        Public Sub New()
            SelectedSeriesInfo = New SeriesInfo()
            SelectedStructureInfos = New List(Of String)
            KatabanStructureInfos = New List(Of KatabanStructureInfo)
            FocusSeqNo = 1
            Messages = New List(Of String)
            IsShowRodEnd = False
            IsShowOtherOption = False
            IsShowStopper = False
            IsShowMotor1 = False
            IsShowMotor2 = False
            IsShowPortPosition = False
            IsShowStock = False
        End Sub
    End Class
End Namespace