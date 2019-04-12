Namespace Constants
    Public Module Levels
        ''' <summary>
        '''     ユーザー種別
        ''' </summary>
        Public Structure UserClassLevel
            ''' <summary> 国内代理店(登録店) </summary>
            Public Const DmAgentRs = "11"

            ''' <summary> 国内代理店(ＳＳ店) </summary>
            Public Const DmAgentSs = "12"

            ''' <summary> 国内代理店(ＢＳ店) </summary>
            Public Const DmAgentBs = "13"

            ''' <summary> 国内代理店(ＧＳ店) </summary>
            Public Const DmAgentGs = "14"

            ''' <summary> 国内代理店(ＰＳ店) </summary>
            Public Const DmAgentPs = "15"

            ''' <summary> 海外代理店(契約店) </summary>
            Public Const OsAgentCs = "16"

            ''' <summary> 海外代理店(契約店、E-con) </summary>
            Public Const OsAgentLs = "17"

            ''' <summary> 国内営業所 </summary>
            Public Const DmSalesOffice = "21"

            ''' <summary> 海外営業部 </summary>
            Public Const OsSalesDep = "41"

            ''' <summary> 海外営業部(管理者) </summary>
            Public Const OsSalesDepMnger = "42"

            ''' <summary> 営業本部 </summary>
            Public Const BizHeadquarters = "45"

            ''' <summary> 営業本部(管理者) </summary>
            Public Const BizHeadquartersMnger = "46"

            ''' <summary> 情報システム部 </summary>
            Public Const InfoSysForce = "91"

            ''' <summary> 情報システム部(管理者) </summary>
            Public Const InfoSysForceMnger = "95"

            ''' <summary> 情報システム部(システム管理者) </summary>
            Public Const InfoSysForceSysAdmin = "99"

            '''' <summary> 海外販社(現地採用者(営業)) </summary>
            'Public Const OsSelComLocEmpBiz = "22"       
            '''' <summary> 海外販社(現地採用者(販管)) </summary>
            'Public Const OsSelComLocEmpRetMnger = "23"  
            '''' <summary> 海外販社(現地採用者) </summary>
            'Public Const OsSelComLocEmp = "24"          
            '''' <summary> 海外販社(日本人駐在員) </summary>
            'Public Const OsSelComJpnRep = "25"          
            '''' <summary> 技術部 </summary>
            'Public Const EngineeringDep = "31"          
        End Structure

        ''' <summary>
        '''     オプション候補判定レベル
        ''' </summary>
        Public structure OptionJudgeLevel
            ''' <summary> 選択条件 </summary>
            Public Const SelectCondition = "1"

            ''' <summary> Skip条件 </summary>
            Public Const SkipCondition = "2"

            ''' <summary> 複数選択 </summary>
            Public Const PluralCondition = "4"
        End Structure
    End Module
End NameSpace