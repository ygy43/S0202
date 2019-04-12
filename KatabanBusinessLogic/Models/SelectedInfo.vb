Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanCommon.Constants

Namespace Models
    ''' <summary>
    '''     選択した情報
    ''' </summary>
    Public Class SelectedInfo
        Public Sub New()
            Me.Series = New SeriesInfo()
            Me.KatabanStructures = New List(Of KatabanStructureInfo)
            Me.Symbols = New List(Of String)
            Me.RodEnd = New RodEndInfoSelected()
            Me.OtherOption = String.Empty
        End Sub

        ''' <summary>
        '''     機種情報
        ''' </summary>
        Public Property Series As SeriesInfo

        ''' <summary>
        '''     形番構成情報
        ''' </summary>
        Public Property KatabanStructures As List(Of KatabanStructureInfo)

        ''' <summary>
        '''     引当情報
        ''' </summary>
        Public Property Symbols As List(Of String)

        ''' <summary>
        '''     ロッド先端情報
        ''' </summary>
        Public Property RodEnd As RodEndInfoSelected

        ''' <summary>
        '''     オプション外情報
        ''' </summary>
        Public Property OtherOption As String

        ''' <summary>
        '''     口径
        ''' </summary>
        Public ReadOnly Property BoreSize As Integer
            Get
                For i = 0 To KatabanStructures.Count - 1
                    Dim structureInfo = KatabanStructures(i)

                    If structureInfo.element_div = Divisions.ElementDiv.Port AndAlso
                       Not String.IsNullOrEmpty(Symbols(i)) Then
                        '口径の場合は
                        Return CType(Symbols(i), Integer)
                    End If
                Next

                Return 0
            End Get
        End Property
    End Class
End Namespace