Imports KatabanCommon.Constants

Namespace ViewModels.Options
    ''' <summary>
    '''     ロッド先端表示情報
    ''' </summary>
    Public Class RodEndUnitViewModel
        Public Sub New()
            Me.IsEnable = True
            Me.PatternSymbol = String.Empty
            Me.PatternType = Divisions.RodEndUnitDiv.Normal
        End Sub

        ''' <summary>
        '''     編集可否フラグ
        ''' </summary>
        Public Property IsEnable As Boolean


        ''' <summary>
        '''     パタン名称
        ''' </summary>
        Public Property PatternSymbol As String

        ''' <summary>
        '''     パタン種類
        ''' </summary>
        Public Property PatternType As String
    End Class

    ''' <summary>
    '''     ロッド先端情報（標準）
    ''' </summary>
    Public Class RodEndUnitNormalViewModel
        Inherits RodEndUnitViewModel

        Public Sub New()
            MyBase.New
            Me.Rows = New List(Of RodEndRow)
            me._image = string.Empty
        End Sub

        ''' <summary>
        '''     画像名称
        ''' </summary>
        Private _image As String

        Public Property Image As String
            Get
                Return _image.Replace("../KHImage/", "~/Content/Images/")
            End Get
            Set
                _image = Value
            End Set
        End Property

        ''' <summary>
        '''     標準寸法タイトル
        ''' </summary>
        Public Property TitleStandard As String

        ''' <summary>
        '''     特注寸法タイトル
        ''' </summary>
        Public Property TitleCustom As String

        ''' <summary>
        '''     ロッド先端情報
        ''' </summary>
        Public Property Rows As List(Of RodEndRow)
    End Class

    ''' <summary>
    '''     ロッド先端情報（Other）
    ''' </summary>
    Public Class RodEndUnitOtherViewModel
        Inherits RodEndUnitViewModel

        Public Sub New()
            MyBase.New
            Me.PatternType = RodEndUnitDiv.Other
            Me.RealPattern = String.Empty
        End Sub

        ''' <summary>
        '''     Other
        ''' </summary>
        Public Property TextTitle As String

        ''' <summary>
        '''     Other
        ''' </summary>
        Public Property CustomValue As String

        ''' <summary>
        ''' 入力値から分解したPatternSymbol
        ''' </summary>
        Public Property RealPattern As String
    End Class

    ''' <summary>
    '''     ロッド先端情報（画像のみ）
    ''' </summary>
    Public Class RodEndUnitOnlyImageViewModel
        Inherits RodEndUnitViewModel

        Public Sub New()
            MyBase.New
            Me.IsShowMessage = False
            me._image = string.Empty
            Me.PatternType = RodEndUnitDiv.ImageOnly
        End Sub

        ''' <summary>
        '''     メッセージ
        ''' </summary>
        Public Property Message As String

        ''' <summary>
        '''     メッセージ表示フラグ
        ''' </summary>
        Public Property IsShowMessage As Boolean

        ''' <summary>
        '''     画像名称
        ''' </summary>
        Private _image As String

        Public Property Image As String
            Get
                Return _image.Replace("../KHImage/", "~/Content/Images/")
            End Get
            Set
                _image = Value
            End Set
        End Property
    End Class

    ''' <summary>
    '''     普通行
    ''' </summary>
    Public Class RodEndRow
        Public Sub New()
            Me.IsEnable = True
            Me.IsCalculateC = False
            Me.DisplayExternalForm = String.Empty
            Me.StandardValue = String.Empty
            Me.CustomValue = String.Empty
            Me.ActStandardValue = String.Empty
            Me.CustomValueOptions = New List(Of String)
        End Sub

        ''' <summary>
        '''     編集可否フラグ
        ''' </summary>
        Public Property IsEnable As Boolean

        ''' <summary>
        '''     KK,A,C同時に存在する時にKKとAの値によりCを計算
        ''' </summary>
        Public Property IsCalculateC As Boolean

        ''' <summary>
        '''     外形寸法
        ''' </summary>
        Public Property ExternalForm As String

        ''' <summary>
        '''     表示外形寸法
        ''' </summary>
        Public Property DisplayExternalForm As String

        ''' <summary>
        '''     標準寸法
        ''' </summary>
        Public Property StandardValue As String

        ''' <summary>
        '''     標準寸法差分（KKの場合のみ利用）
        ''' </summary>
        ''' <returns></returns>
        Public Property ActStandardValue As String

        ''' <summary>
        '''     特注寸法
        ''' </summary>
        Public Property CustomValue As String

        ''' <summary>
        '''     KK特注寸法と差分（パイプ区切り）
        ''' </summary>
        Public Property CustomValueOptions As List(Of String)

        Public ReadOnly Property CustomValueOptionSelectList As SelectList
            Get
                Dim items As New List(Of SelectListItem)

                For Each s As String In CustomValueOptions

                    If s = MyControlChars.Pipe Then
                        items.Add(New SelectListItem With {
                                     .Text = String.Empty,
                                     .Value = MyControlChars.Pipe & ActStandardValue}
                                  )
                    Else

                        items.Add(New SelectListItem With {
                                     .Text = s.Split(MyControlChars.Pipe).FirstOrDefault(),
                                     .Value = s}
                                  )
                    End If
                Next

                Return New SelectList(items, "Value", "Text")
            End Get
        End Property

    End Class
End Namespace