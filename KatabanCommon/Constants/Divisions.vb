Namespace Constants
    ''' <summary>
    '''     各区分
    ''' </summary>
    Public Module Divisions
        ''' <summary>
        '''     言語区分
        ''' </summary>
        Public Structure LanguageDiv
            ''' <summary> デフォルト言語(英語) </summary>
            Public Const DefaultLang = "en"

            ''' <summary> 簡体字 </summary>
            Public Const SimplifiedChinese = "zh"

            ''' <summary> 繁体字 </summary>
            Public Const TraditionalChinese = "tw"

            ''' <summary> 日本語 </summary>
            Public Const Japanese = "ja"

            ''' <summary> 韓国語 </summary>
            Public Const Korean = "ko"
        End Structure

        ''' <summary>
        '''     多言語区分
        ''' </summary>
        Public Structure LocalizationDiv
            ''' <summary> デフォルト言語(英語) </summary>
            Public Const DefaultLang = "en-US"

            ''' <summary> 簡体字 </summary>
            Public Const SimplifiedChinese = "zh-CN"

            ''' <summary> 繁体字 </summary>
            Public Const TraditionalChinese = "zh-TW"

            ''' <summary> 日本語 </summary>
            Public Const Japanese = "ja-JP"

            ''' <summary> 韓国語 </summary>
            Public Const Korean = "ko-KR"
        End Structure

        ''' <summary>
        '''     言語リスト
        ''' </summary>
        Public LanguageList As List(Of String) = New List(Of String) From {
            LanguageDiv.DefaultLang,
            LanguageDiv.SimplifiedChinese,
            LanguageDiv.TraditionalChinese,
            LanguageDiv.Japanese,
            LanguageDiv.Korean
            }

        ''' <summary>
        '''     国区分
        ''' </summary>
        Public Structure CountryDiv
            ''' <summary> デフォルト国(日本) </summary>
            Public Const DefaultCountry = "JPN"
        End Structure

        ''' <summary>
        '''     営業所
        ''' </summary>
        Public Structure OfficeDiv
            Public Const Overseas = "II2"
        End Structure

        ''' <summary>
        '''     稼動状況区分
        ''' </summary>
        Public Structure OperationStateDiv
            ''' <summary> 停止中 </summary>
            Public Const Stopping = "0"

            ''' <summary> 稼動中 </summary>
            Public Const Operating = "1"

            ''' <summary> トラブル </summary>
            Public Const Trouble = "E"
        End Structure

        ''' <summary>
        '''     オプション種類区分
        ''' </summary>
        Public Structure ElementDiv
            ''' <summary> 電圧 </summary>
            Public Const Voltage = "1"

            ''' <summary> ストローク </summary>
            Public Const Stroke = "3"

            ''' <summary> 口径 </summary>
            Public Const Port = "5"

            ''' <summary> コイル </summary>
            Public Const Coil = "6"

            ''' <summary> 口径(電圧用) </summary>
            Public Const VolPort = "7"
        End Structure

        ''' <summary>
        '''     オプション候補判定区分
        ''' </summary>
        Public Structure ElementJudgeDiv
            ''' <summary> IN </summary>
            Public Const InSign = "I"

            ''' <summary> OUT </summary>
            Public Const OutSign = "O"

            ''' <summary> OR </summary>
            Public Const CondOr = "+"

            ''' <summary> AND </summary>
            Public Const CondAnd = "*"

            ''' <summary> イコール </summary>
            Public Const Equal = "EQ"

            ''' <summary> ノットイコール </summary>
            Public Const NotEqual = "NE"
        End Structure

        ''' <summary>
        '''     積上価格区分
        ''' </summary>
        Public Structure AccumulatePriceDiv
            ''' <summary> 国内用(標準) </summary>
            Public Const Domestic = "0"

            ''' <summary> 海外用(価格加算無) </summary>
            Public Const Overseas = "1"

            ''' <summary> C5 </summary>
            Public Const C5 = "C5"

            ''' <summary> DIN Rail </summary>
            Public Const DinRail = "DINRail"

            ''' <summary> 継手 </summary>
            Public Const Joint = "Joint"

            ''' <summary> ねじ </summary>
            Public Const Screw = "Screw"

            ''' <summary> Open Price </summary>
            Public Const Open = "Open"
        End Structure

        ''' <summary>
        '''     電圧区分
        ''' </summary>
        Public Structure VoltageDiv
            ''' <summary> 標準電圧 </summary>
            Public Const Standard = "1"

            ''' <summary> オプション </summary>
            Public Const Options = "2"

            ''' <summary> その他電圧 </summary>
            Public Const Other = "3"
        End Structure

        ''' <summary>
        '''     形番チェック区分
        ''' </summary>
        Public Structure KatabanCheckDiv
            ''' <summary> 在庫品 </summary>
            Public Const Stock = "1"

            ''' <summary> 標準品 </summary>
            Public Const Standard = "2"

            ''' <summary> 特注品 </summary>
            Public Const Special = "3"

            ''' <summary> 部品 </summary>
            Public Const Parts = "4"
        End Structure

        '検索区分
        Public Structure DataTypeDiv
            ''' <summary> 機種検索 </summary>
            Public Const Series = "1"

            ''' <summary> フル形番検索 </summary>
            Public Const FullKataban = "2"

            ''' <summary> 仕入品検索 </summary>
            Public Const Shiire = "3"

            ''' <summary> 全て検索 </summary>
            Public Const All = "4"
        End Structure

        'ハイフン区分
        Public Structure HyphenDiv
            ''' <summary> あり </summary>
            Public Const Necessary = "1"

            ''' <summary> なし </summary>
            Public Const Unnecessary = "0"
        End Structure

        'EL区分
        Public Structure ElDiv
            ''' <summary> ELではない </summary>
            Public Const IsNotEl = "0"

            ''' <summary> EL品 </summary>
            Public Const IsEl = "1"
        End Structure

        '構成オプション検証データ区分
        Public Structure ElePatternDiv
            ''' <summary> 全て </summary>
            Public Const All = "*"

            ''' <summary> 複数選択 </summary>
            Public Const Plural = "#"
        End Structure

        '電源種類
        Public Structure PowerSupply
            Public Const AC = "1"                                                   'AC電源
            Public Const DC = "2"                                                   'DC電源
            Public Const Div1 = "AC"                                                'AC電源
            Public Const Div2 = "DC"                                                'DC電源
            Public Const AC100V = "1"                                               'AC100V
            Public Const AC200V = "2"                                               'AC200V
            Public Const DC24V = "3"                                                'DC24V
            Public Const DC12V = "4"                                                'DC12V
            Public Const AC110V = "5"                                               'AC110V
            Public Const AC220V = "6"                                               'AC220V
            Public Const Const1 = "AC100V"                                          'AC100V
            Public Const Const2 = "AC200V"                                          'AC200V
            Public Const Const3 = "DC24V"                                           'DC24V
            Public Const Const4 = "DC12V"                                           'DC12V
            Public Const Const5 = "AC110V"                                          'AC110V
            Public Const Const6 = "AC220V"                                          'AC220V
            Public Const Const7 = "AC120V"                                          'AC120V
            Public Const Const8 = "AC240V"                                          'AC240V
        End Structure

        ''' <summary>
        '''     ロッド先端ユニット種類
        ''' </summary>
        Public Structure RodEndUnitDiv
            Public Const Normal = "1"
            Public Const Other = "2"
            Public Const ImageOnly = "3"
        End Structure

        ''' <summary>
        '''     ロッド先端パタン記号
        ''' </summary>
        Public Structure RodEndPatternDiv
            Public Const N13 = "N13"
            Public Const N15 = "N15"
            Public Const N11 = "N11"
            Public Const N1 = "N1"
            Public Const N12 = "N12"
            Public Const N14 = "N14"
            Public Const N3 = "N3"
            Public Const N31 = "N31"
            Public Const N2 = "N2"
            Public Const N21 = "N21"
            Public Const N13N11 = "N13-N11"
            Public Const N11N13 = "N11-N13"
        End Structure

        ''' <summary>
        ''' ロッド先端WF最大値区分
        ''' </summary>
        Public Structure RodEndWfMaxDiv
            ''' <summary>
            ''' 最大WF寸法
            ''' </summary>
            Public Const WfMax = "0"

            ''' <summary>
            ''' 標準寸法+最大WF寸法
            ''' </summary>
            Public Const WfMaxAndStandard = "1"
        End Structure
    End Module
End Namespace