Imports KatabanBusinessLogic.KatabanWcfService

Namespace ViewModels.Options
    Public Class OptionsUpdateOptionsViewModel
        ''' <summary>
        '''     構成名称
        ''' </summary>
        Public Property StructureName As String

        ''' <summary>
        '''     フォカスされた構成の候補情報
        ''' </summary>
        Public Property CurrentOptions As List(Of KatabanStructureOptionInfo)
    End Class
End NameSpace