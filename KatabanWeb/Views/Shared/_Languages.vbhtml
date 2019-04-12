@Imports KatabanCommon.Constants

<div class="nav navbar-nav navbar-right">
    <select class="form-control language" id="Language" name="Language" onchange="changeLanguage('@Me.Request.RawUrl')">

        @*英語*@
        @If Session(SessionKeys.Language) IsNot Nothing AndAlso
                            Session(SessionKeys.Language) = LocalizationDiv.DefaultLang Then
            @* 選択された場合 *@
            @<option value="@LocalizationDiv.DefaultLang" selected="selected">@S0202.My.Resources.RLayout.English</option>
        Else
            @<option value="@LocalizationDiv.DefaultLang">@S0202.My.Resources.RLayout.English</option>
        End If

        @*簡体字*@
        @If Session(SessionKeys.Language) IsNot Nothing AndAlso
                                Session(SessionKeys.Language) = LocalizationDiv.SimplifiedChinese Then
            @* 選択された場合 *@
            @<option value="@LocalizationDiv.SimplifiedChinese" selected="selected">@S0202.My.Resources.RLayout.SimplifiedChinese</option>
        Else
            @<option value="@LocalizationDiv.SimplifiedChinese">@S0202.My.Resources.RLayout.SimplifiedChinese</option>
        End If

        @*繁体字*@
        @If Session(SessionKeys.Language) IsNot Nothing AndAlso
                                Session(SessionKeys.Language) = LocalizationDiv.TraditionalChinese Then
            @* 選択された場合 *@
            @<option value="@LocalizationDiv.TraditionalChinese" selected="selected">@S0202.My.Resources.RLayout.TraditionalChinese</option>
        Else
            @<option value="@LocalizationDiv.TraditionalChinese">@S0202.My.Resources.RLayout.TraditionalChinese</option>
        End If

        @*日本語*@
        @If Session(SessionKeys.Language) IsNot Nothing AndAlso
                                Session(SessionKeys.Language) = LocalizationDiv.Japanese Then
            @* 選択された場合 *@
            @<option value="@LocalizationDiv.Japanese" selected="selected">@S0202.My.Resources.RLayout.Japanese</option>
        Else
            @<option value="@LocalizationDiv.Japanese">@S0202.My.Resources.RLayout.Japanese</option>
        End If

        @*韓国語*@
        @If Session(SessionKeys.Language) IsNot Nothing AndAlso
                                Session(SessionKeys.Language) = LocalizationDiv.Korean Then
            @* 選択された場合 *@
            @<option value="@LocalizationDiv.Korean" selected="selected">@S0202.My.Resources.RLayout.Korean</option>
        Else
            @<option value="@LocalizationDiv.Korean">@S0202.My.Resources.RLayout.Korean</option>
        End If

    </select>
</div>