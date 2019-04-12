Namespace Constants
    Public Module Utility
        ''' <summary>
        '''     機種検索結果ページサイズ
        ''' </summary>
        Public Const PageSize = 10

        ''' <summary>
        '''     オプション選択画面テキストボックスの幅単位
        ''' </summary>
        Public Const TextBoxWidthUnit = 20

        ''' <summary>
        '''     その他電圧言語
        ''' </summary>
        Public Structure OtherVoltageText
            Public Const Japanese = "その他電圧"
            Public Const English = "OTHER"
        End Structure

        ''' <summary>
        '''     電圧の種類
        ''' </summary>
        Public Structure VoltageRegex
            Public Const Ac = "AC%V"
            Public Const Dc = "DC%V"
        End Structure
    End Module
End Namespace