Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanBusinessLogic.Results
Imports KatabanCommon.Constants
Imports S0202.My.Resources
Imports S0202.MyHelpers
Imports S0202.ViewModels.Options

Namespace Controllers
    Public Class RodEndController
        Inherits Controller

        ' GET: RodEnd
        Function Index() As ActionResult

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()

            'ロッド先端画面情報を取得
            Dim rodEndModel As RodEndViewModel = GetRodEndInfo(selectedData)

            Return View(rodEndModel)
        End Function

        ' POST: RodEnd
        <HttpPost>
        Function Index(model As RodEndViewModel) As ActionResult

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()

            '選択したロッド先端情報
            Dim selectUnitInfo =
                    model.RodEndUnitInfos.FirstOrDefault(Function(r) r.PatternSymbol = model.SelectedPatternSymbol)

            If selectUnitInfo Is Nothing Then
                '選択されていません、エラーメッセージを表示
                ModelState.AddModelError("", "選択されていません")
                Return View(model)
            Else
                '入力検証

                'オプションチェック
                Dim checkResult = ValidateInput(selectUnitInfo, selectedData, model.RodEndUnitInfos)

                If checkResult.IsSucceed Then
                    
                    'ロッド先端情報をセッションに保存


                    'オプション選択画面へ
                    'Return View("",)
                Else
                    'エラーメッセージを表示し、フォカスを設定
                    For Each err As String In checkResult.Errors
                        ModelState.AddModelError("", err)
                    Next

                    Return View(model)
                End If
            End If
        End Function

#Region "Private"

        ''' <summary>
        '''     ロッド先端画面情報の取得
        ''' </summary>
        ''' <param name="info">選択した情報</param>
        ''' <returns></returns>
        Private Function GetRodEndInfo(info As SelectedInfo) As RodEndViewModel

            Dim model As New RodEndViewModel

            'ロッド先端マスタ情報
            Dim rodEndInfos = RodEndManager.GetRodEndInfo(info)

            'ロッド先端外形寸法などの情報
            Dim rodEndExternalFormInfos = RodEndManager.GetRodEndExternalFormInfo(info.Series.series_kataban,
                                                                                  info.Series.key_kataban,
                                                                                  info.BoreSize)
            '情報設定
            model.SeriesName = info.Series.disp_name
            model.RodEndUnitInfos = Convert2RodEndUnitInfos(rodEndInfos, rodEndExternalFormInfos, info, info.BoreSize)

            Return model
        End Function

        ''' <summary>
        '''     ロッド先端のModelViewに変換
        ''' </summary>
        ''' <param name="rodEndInfos"></param>
        ''' <param name="rodEndExternalFormInfos"></param>
        ''' <returns></returns>
        Private Function Convert2RodEndUnitInfos(rodEndInfos As List(Of RodEndInfo),
                                                 rodEndExternalFormInfos As List(Of RodEndExternalFormInfo),
                                                 info As SelectedInfo,
                                                 boreSize As Integer) _
            As List(Of RodEndUnitViewModel)

            Dim result As New List(Of RodEndUnitViewModel)

            For Each endInfo As RodEndInfo In rodEndInfos

                If endInfo.rod_pattern_symbol.Contains(MyControlChars.Hyphen) Then

                    '画像のみ
                    Dim unit As New RodEndUnitOnlyImageViewModel

                    unit.IsEnable = RodEndManager.IsRodEndUnitEnable(endInfo.rod_pattern_symbol, info, boreSize)
                    unit.IsShowMessage = RodEndManager.IsRodEndUnitShowMessage(endInfo.rod_pattern_symbol,
                                                                               info.Series.series_kataban)
                    unit.PatternSymbol = endInfo.rod_pattern_symbol
                    unit.PatternType = RodEndUnitDiv.ImageOnly
                    unit.Image = endInfo.url
                    unit.Message = GetRodEndMessage(endInfo.rod_pattern_symbol, info.Series.series_kataban)

                    result.Add(unit)

                ElseIf endInfo.rod_pattern_symbol = "Other" Then

                    'Other
                    Dim unit As New RodEndUnitOtherViewModel

                    unit.IsEnable = RodEndManager.IsRodEndUnitEnable(endInfo.rod_pattern_symbol, info, boreSize)
                    unit.PatternSymbol = endInfo.rod_pattern_symbol
                    unit.PatternType = RodEndUnitDiv.Other
                    unit.TextTitle = RRodEnd.OtherTitle

                    result.Add(unit)
                Else

                    '正常
                    Dim unit As New RodEndUnitNormalViewModel

                    unit.IsEnable = RodEndManager.IsRodEndUnitEnable(endInfo.rod_pattern_symbol, info, boreSize)
                    unit.PatternSymbol = endInfo.rod_pattern_symbol
                    unit.PatternType = RodEndUnitDiv.Normal
                    unit.Image = endInfo.url
                    unit.TitleStandard = "標準寸法"
                    unit.TitleCustom = "特注寸法"

                    '行情報を設定
                    unit.Rows = GetUnitRows(endInfo.rod_pattern_symbol, rodEndExternalFormInfos)

                    result.Add(unit)
                End If

            Next

            Return result
        End Function

        ''' <summary>
        '''     N11-N13/N13-N11の場合、メッセージを取得
        ''' </summary>
        ''' <param name="pattern"></param>
        ''' <param name="seriesKataban"></param>
        ''' <returns></returns>
        Private Function GetRodEndMessage(pattern As String, seriesKataban As String) As String
            Select Case seriesKataban
                Case "SSD", "SCA2", "CMK2"
                    'メッセージ表示
                    If pattern = RodEndPatternDiv.N13N11 Then
                        Return RRodEnd.DqMessage
                    ElseIf pattern = RodEndPatternDiv.N11N13 Then
                        Return RRodEnd.DqMessage
                    End If

                    Return String.Empty
                Case Else
                    Return String.Empty
            End Select
        End Function

        ''' <summary>
        '''     ロッド先端ユニットの行情報を取得
        ''' </summary>
        ''' <param name="rodPatternSymbol"></param>
        ''' <param name="rodEndExternalFormInfos"></param>
        ''' <returns></returns>
        Private Function GetUnitRows(rodPatternSymbol As String,
                                     rodEndExternalFormInfos As List(Of RodEndExternalFormInfo)) As List(Of RodEndRow)
            Dim results As New List(Of RodEndRow)

            '対象パタンを取得
            Dim patterns = rodEndExternalFormInfos.Where(Function(r) r.rod_pattern_symbol = rodPatternSymbol).ToList()
            '行を取得
            Dim forms = patterns.Select(Function(p) p.external_form).Distinct().ToList()

            For Each form As String In forms

                Dim patternsWithSameForm = patterns.Where(Function(p) p.external_form = form).ToList()

                Dim row As New RodEndRow

                row.DisplayExternalForm = patternsWithSameForm.First.disp_external_form
                row.StandardValue = patternsWithSameForm.First.normal_value
                row.ExternalForm = patternsWithSameForm.First.external_form
                'KKの場合のみ利用
                row.ActStandardValue = patternsWithSameForm.First.act_normal_value

                If form = "KK" Then
                    'リストから選択する
                    row.CustomValueOptions =
                        patternsWithSameForm.Select(
                            Function(p) p.selectable_value & MyControlChars.Pipe & p.act_selectable_value).ToList()
                Else
                    'テキストの場合
                    If patternsWithSameForm.First.input_div = "L" Then
                        'ラベルの場合は表示不可
                        row.IsEnable = False
                    ElseIf patternsWithSameForm.First.input_div = "T" Then
                        row.IsEnable = True
                    End If
                End If

                'KK,A,C同時に存在する時にKKとAの値によりCを計算
                If forms.Contains("KK") AndAlso
                   forms.Contains("A") AndAlso
                   forms.Contains("C") AndAlso
                   (form = "KK" OrElse form = "A") Then
                    row.IsCalculateC = True
                Else
                    row.IsCalculateC = False
                End If

                results.Add(row)
            Next

            Return results
        End Function

#Region "検証"

        ''' <summary>
        '''     入力検証
        ''' </summary>
        ''' <param name="info"></param>
        ''' <returns></returns>
        Private Function ValidateInput(info As RodEndUnitViewModel,
                                       selectedData As SelectedInfo,
                                       rodEndUnitInfos As List(Of RodEndUnitViewModel)) As RodEndCheckResult

            Dim result As New RodEndCheckResult
            Dim errorMessage As String = String.Empty

            Dim wfMaxValue As Double = RodEndManager.GetWfMaxValue(selectedData.Series.series_kataban,
                                                                   selectedData.Series.key_kataban,
                                                                   selectedData.BoreSize)

            Select Case selectedData.Series.series_kataban
                Case "SSD"
                    If Not SsdInputCheck(info, wfMaxValue, rodEndUnitInfos, errorMessage) Then
                        result.IsSucceed = False
                        result.Errors = New List(Of String) From {errorMessage}
                    End If
                Case "SCA2"
                    If _
                        Not _
                        Sca2InputCheck(info, wfMaxValue, rodEndUnitInfos, selectedData.Symbols,
                                       selectedData.Series.key_kataban, errorMessage) _
                        Then
                        result.IsSucceed = False
                        result.Errors = New List(Of String) From {errorMessage}
                    End If
                Case "JSC3", "JSC4"
                    If _
                        Not _
                        Jsc3InputCheck(info, wfMaxValue, rodEndUnitInfos, selectedData.Series.key_kataban, errorMessage) _
                        Then
                        result.IsSucceed = False
                        result.Errors = New List(Of String) From {errorMessage}
                    End If
                Case "SCS", "SCS2"
                    If Not ScsInputCheck(info, wfMaxValue, rodEndUnitInfos, errorMessage) Then
                        result.IsSucceed = False
                        result.Errors = New List(Of String) From {errorMessage}
                    End If
                Case "CMK2"
                    If Not Cmk2InputCheck(info, wfMaxValue, rodEndUnitInfos, errorMessage) Then
                        result.IsSucceed = False
                        result.Errors = New List(Of String) From {errorMessage}
                    End If
            End Select

            Return result
        End Function

        ''' <summary>
        '''     SSDチェック
        ''' </summary>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function SsdInputCheck(infoBase As RodEndUnitViewModel,
                                       wfMaxValue As Double,
                                       rodEndUnitInfos As List(Of RodEndUnitViewModel),
                                       ByRef errorMessage As String) As Boolean

            If infoBase.PatternType = RodEndUnitDiv.ImageOnly Then Return True

            Dim info As New RodEndUnitNormalViewModel

            Select Case infoBase.PatternType
                Case RodEndUnitDiv.Other

                    Dim infoOther = CType(infoBase, RodEndUnitOtherViewModel)

                    'ハイフンチェック
                    If Not HyphenCheck(infoOther,
                                       New List(Of String) From {RodEndPatternDiv.N11N13, RodEndPatternDiv.N13N11},
                                       errorMessage) Then Return False

                    '二つ以上のパタン記号を入力するとエラー

                    '入力値からパタン記号と特注寸法を分解
                    info = SeparateRodEndOtherInput(infoOther, rodEndUnitInfos)

                    'WFの後に数値がなかったらエラー
                    If Not NumericCheck(info, "WF", errorMessage) Then Return False

                    'Aの後に数値がなかったらエラー
                    If Not NumericCheck(info, "A", errorMessage) Then Return False

                Case RodEndUnitDiv.Normal

                    info = CType(infoBase, RodEndUnitNormalViewModel)

                    'A/KL寸法チェック
                    If Not AklCheck(info, 0, 0, errorMessage) Then Return False

                    'WF寸法チェック
                    If Not WfCheck(info, errorMessage) Then Return False

            End Select

            'N13/N11チェック
            If info.PatternSymbol = RodEndPatternDiv.N13 Or info.PatternSymbol = RodEndPatternDiv.N11 Then
                If Not NotEqualsStandardCheck(info, errorMessage) Then Return False
            End If

            'WF + A寸法チェック
            If Not WfaCheck(info, wfMaxValue, errorMessage) Then Return False

            '最大WFチェック
            If Not MaxWfCheck(info, wfMaxValue, RodEndWfMaxDiv.WfMaxAndStandard, errorMessage) Then Return False

            Return True
        End Function

        ''' <summary>
        '''     SCA2チェック
        ''' </summary>
        ''' <param name="symbols"></param>
        ''' <param name="keyKataban"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function Sca2InputCheck(infoBase As RodEndUnitViewModel,
                                        wfMaxValue As Double,
                                        rodEndUnitInfos As List(Of RodEndUnitViewModel),
                                        symbols As List(Of String),
                                        keyKataban As String,
                                        ByRef errorMessage As String) As Boolean
            Dim info As New RodEndUnitNormalViewModel

            Select Case infoBase.PatternType
                Case RodEndUnitDiv.Other
                    Dim infoOther = CType(infoBase, RodEndUnitOtherViewModel)
                    'ハイフンチェック
                    If Not HyphenCheck(infoOther,
                                       New List(Of String) From {RodEndPatternDiv.N11N13, RodEndPatternDiv.N13N11},
                                       errorMessage) Then Return False

                    '入力値からパタン記号と特注寸法を分解
                    info = SeparateRodEndOtherInput(infoOther, rodEndUnitInfos)

                    'WFの後に数値がなかったらエラー
                    If Not NumericCheck(info, "WF", errorMessage) Then Return False

                    'Aの後に数値がなかったらエラー
                    If Not NumericCheck(info, "A", errorMessage) Then Return False

                Case Else
                    info = CType(infoBase, RodEndUnitNormalViewModel)
                    'A/KL寸法チェック
                    If Not AklCheck(info, 15, 5, errorMessage) Then Return False

                    'WF寸法チェック
                    If Not WfCheck(info, errorMessage) Then Return False

                    'N13/N11チェック
                    If info.PatternSymbol = RodEndPatternDiv.N13 Or info.PatternSymbol = RodEndPatternDiv.N11 Then
                        If Not NotEqualsStandardCheck(info, errorMessage) Then Return False
                    End If

                    'WF + A寸法チェック
                    If Not WfaCheck(info, wfMaxValue, errorMessage) Then Return False

                    '最大WFチェック
                    If Not MaxWfCheck(info, wfMaxValue, RodEndWfMaxDiv.WfMax, errorMessage) Then Return False

                    'WF寸法SCAチェック
                    If Not WfCheckSca2(info, symbols, keyKataban, errorMessage) Then Return False
            End Select

            Return True
        End Function


        ''' <summary>
        '''     JSC3チェック
        ''' </summary>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function Jsc3InputCheck(infoBase As RodEndUnitViewModel,
                                        wfMaxValue As Double,
                                        rodEndUnitInfos As List(Of RodEndUnitViewModel),
                                        keyKataban As String,
                                        ByRef errorMessage As String) As Boolean
            Dim info As New RodEndUnitNormalViewModel

            Select Case infoBase.PatternType
                Case RodEndUnitDiv.Other
                    Dim infoOther = CType(infoBase, RodEndUnitOtherViewModel)
                    'ハイフンチェック
                    If Not HyphenCheck(infoOther,
                                       New List(Of String) From {RodEndPatternDiv.N11N13, RodEndPatternDiv.N13N11},
                                       errorMessage) Then Return False

                    '入力値からパタン記号と特注寸法を分解
                    info = SeparateRodEndOtherInput(infoOther, rodEndUnitInfos)

                    'WFの後に数値がなかったらエラー
                    If Not NumericCheck(info, "WF", errorMessage) Then Return False

                    'Aの後に数値がなかったらエラー
                    If Not NumericCheck(info, "A", errorMessage) Then Return False

                Case Else
                    info = CType(infoBase, RodEndUnitNormalViewModel)
                    'A/KL寸法チェック
                    Dim minASize = 0
                    If keyKataban = "1" Then
                        minASize = 15
                    ElseIf keyKataban = "2" Then
                        minASize = 20
                    End If

                    If Not AklCheck(info, minASize, 5, errorMessage) Then Return False

                    'WF寸法チェック
                    If Not WfCheck(info, errorMessage) Then Return False

                    'N13チェック
                    If info.PatternSymbol = RodEndPatternDiv.N13 Then
                        If Not NotEqualsStandardCheck(info, errorMessage) Then Return False
                    End If

                    'WF + A寸法チェック
                    If Not WfaCheck(info, wfMaxValue, errorMessage) Then Return False

                    '最大WFチェック
                    If Not MaxWfCheck(info, wfMaxValue, RodEndWfMaxDiv.WfMax, errorMessage) Then Return False
            End Select

            Return True
        End Function

        ''' <summary>
        '''     SCSチェック
        ''' </summary>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function ScsInputCheck(infoBase As RodEndUnitViewModel,
                                       wfMaxValue As Double,
                                       rodEndUnitInfos As List(Of RodEndUnitViewModel),
                                       ByRef errorMessage As String) As Boolean
            Dim info As New RodEndUnitNormalViewModel

            Select Case infoBase.PatternType
                Case RodEndUnitDiv.Other
                    Dim infoOther = CType(infoBase, RodEndUnitOtherViewModel)
                    'ハイフンチェック
                    If Not HyphenCheck(infoOther,
                                       New List(Of String) From {RodEndPatternDiv.N13N11},
                                       errorMessage) Then Return False

                    '入力値からパタン記号と特注寸法を分解
                    info = SeparateRodEndOtherInput(infoOther, rodEndUnitInfos)

                    'WFの後に数値がなかったらエラー
                    If Not NumericCheck(info, "WF", errorMessage) Then Return False

                    'Aの後に数値がなかったらエラー
                    If Not NumericCheck(info, "A", errorMessage) Then Return False

                Case Else
                    info = CType(infoBase, RodEndUnitNormalViewModel)
                    'A/KL寸法チェック
                    If Not AklCheck(info, 20, 5, errorMessage) Then Return False

                    'WF寸法チェック
                    If Not WfCheck(info, errorMessage) Then Return False

                    'N13チェック
                    If info.PatternSymbol = RodEndPatternDiv.N13 Then
                        If Not NotEqualsStandardCheck(info, errorMessage) Then Return False
                    End If

                    'WF + A寸法チェック
                    If Not WfaCheck(info, wfMaxValue, errorMessage) Then Return False

                    '最大WFチェック
                    If Not MaxWfCheck(info, wfMaxValue, RodEndWfMaxDiv.WfMax, errorMessage) Then Return False
            End Select

            Return True
        End Function

        ''' <summary>
        '''     CMK2チェック
        ''' </summary>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function Cmk2InputCheck(infoBase As RodEndUnitViewModel,
                                        wfMaxValue As Double,
                                        rodEndUnitInfos As List(Of RodEndUnitViewModel),
                                        ByRef errorMessage As String) As Boolean
            Dim info As New RodEndUnitNormalViewModel

            Select Case infoBase.PatternType
                Case RodEndUnitDiv.Other
                    Dim infoOther = CType(infoBase, RodEndUnitOtherViewModel)
                    'ハイフンチェック
                    If Not HyphenCheck(infoOther,
                                       New List(Of String) From {RodEndPatternDiv.N13N11},
                                       errorMessage) Then Return False

                    '入力値からパタン記号と特注寸法を分解
                    info = SeparateRodEndOtherInput(infoOther, rodEndUnitInfos)

                    'WFの後に数値がなかったらエラー
                    If Not NumericCheck(info, "WF", errorMessage) Then Return False

                    'Aの後に数値がなかったらエラー
                    If Not NumericCheck(info, "A", errorMessage) Then Return False

                Case Else
                    info = CType(infoBase, RodEndUnitNormalViewModel)
                    'A/KL寸法チェック
                    If Not AklCheck(info, 20, 5, errorMessage) Then Return False

                    'WF寸法チェック
                    If Not WfCheck(info, errorMessage) Then Return False

                    'N13チェック
                    If info.PatternSymbol = RodEndPatternDiv.N13 Then
                        If Not NotEqualsStandardCheck(info, errorMessage) Then Return False
                    End If

                    'WF + A寸法チェック
                    If Not WfaCheck(info, wfMaxValue, errorMessage) Then Return False

                    '最大WFチェック
                    If Not MaxWfCheck(info, wfMaxValue, RodEndWfMaxDiv.WfMaxAndStandard, errorMessage) Then Return False
            End Select

            Return True
        End Function

#Region "検証詳細"

        ''' <summary>
        '''     入力値からパタン記号と特注寸法を分解
        ''' </summary>
        ''' <param name="info"></param>
        Private Function SeparateRodEndOtherInput(ByRef info As RodEndUnitOtherViewModel,
                                                  rodEndUnitInfos As List(Of RodEndUnitViewModel)) _
            As RodEndUnitNormalViewModel
            Dim result As New RodEndUnitNormalViewModel

            Dim patternList =
                    rodEndUnitInfos.Select(Function(r) r.PatternSymbol).OrderByDescending(Function(r) r.Length).ToList()

            For Each selectablePattern As String In patternList
                If info.CustomValue.StartsWith(selectablePattern) Then
                    Dim unitInfo As RodEndUnitNormalViewModel =
                            rodEndUnitInfos.First(Function(r) r.PatternSymbol = selectablePattern)

                    Dim allForms = unitInfo.Rows.Select(Function(r) r.ExternalForm).ToArray()

                    '各formの値を分解
                    Dim otherInputs = info.CustomValue.Split(allForms, StringSplitOptions.None)

                    For Each row As RodEndRow In unitInfo.Rows

                        If Not info.CustomValue.Contains(row.ExternalForm) Then Continue For

                        For Each otherInput As String In otherInputs
                            Dim splitResults = info.CustomValue.Split({row.ExternalForm}, StringSplitOptions.None)
                            If splitResults(1).StartsWith(otherInput) Then
                                '対象formの値をセット
                                result.Rows.Add(New RodEndRow With {.CustomValue = otherInput,
                                                   .ExternalForm = row.ExternalForm,
                                                   .StandardValue = row.StandardValue
                                                   }
                                                )
                            End If
                        Next

                    Next

                    result.PatternSymbol = selectablePattern
                    Exit For
                End If
            Next

            Return result
        End Function

        ''' <summary>
        '''     WFの最大値のチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="maxWfSize"></param>
        ''' <param name="maxDiv"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function MaxWfCheck(info As RodEndUnitNormalViewModel,
                                    maxWfSize As Double,
                                    maxDiv As String,
                                    ByRef errorMessage As String) As Boolean

            Dim wfRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "WF")

            If wfRow IsNot Nothing Then
                Dim maxSize As Double

                If maxDiv = "0" Then
                    maxSize = maxWfSize
                    errorMessage = "W8470"
                Else
                    maxSize = maxWfSize + CType(wfRow.StandardValue, Double)
                    errorMessage = "W8460"
                End If

                If Not String.IsNullOrEmpty(wfRow.CustomValue) Then
                    Dim wfSize = CType(wfRow.CustomValue, Double)
                    If wfSize > maxSize Then
                        Return False
                    End If
                End If
            End If

            Return True
        End Function

        ''' <summary>
        '''     WFとA寸法のチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="maxWfSize"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function WfaCheck(info As RodEndUnitNormalViewModel, maxWfSize As Double, ByRef errorMessage As String) _
            As Boolean

            Select Case info.PatternSymbol
                Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                    Dim aRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "A")
                    Dim wfRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "WF")

                    If aRow IsNot Nothing AndAlso wfRow IsNot Nothing Then
                        Dim aCustom As Double
                        Dim wfCustom As Double
                        Dim aStandard = CType(aRow.StandardValue, Double)
                        Dim wfStandard = CType(wfRow.StandardValue, Double)

                        If Not String.IsNullOrEmpty(aRow.CustomValue) Then
                            aCustom = CType(aRow.CustomValue, Double)
                        Else
                            aCustom = CType(aRow.StandardValue, Double)
                        End If

                        If Not String.IsNullOrEmpty(wfRow.CustomValue) Then
                            wfCustom = CType(wfRow.CustomValue, Double)
                        Else
                            wfCustom = CType(wfRow.StandardValue, Double)
                        End If

                        If wfCustom + aCustom > aStandard + wfStandard + maxWfSize Then
                            If Math.Abs(maxWfSize) < 0.001 Then
                                errorMessage = "W8470"
                            Else
                                errorMessage = "W8460"
                            End If
                            Return False
                        End If

                    End If
                Case Else
                    Dim wfRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "WF")

                    If wfRow IsNot Nothing Then

                        Dim wfCustom As Double
                        Dim wfStandard = CType(wfRow.StandardValue, Double)

                        If Not String.IsNullOrEmpty(wfRow.CustomValue) Then
                            wfCustom = CType(wfRow.CustomValue, Double)
                        Else
                            wfCustom = CType(wfRow.StandardValue, Double)
                        End If

                        If wfCustom > wfStandard + maxWfSize Then
                            If Math.Abs(maxWfSize) < 0.001 Then
                                errorMessage = "W8470"
                            Else
                                errorMessage = "W8460"
                            End If
                            Return False
                        End If
                    End If
            End Select

            Return True
        End Function

        ''' <summary>
        '''     標準寸法と不一致のチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function NotEqualsStandardCheck(info As RodEndUnitNormalViewModel, ByRef errorMessage As String) _
            As Boolean

            If info.Rows.All(Function(r) r.CustomValue = r.StandardValue) Then
                errorMessage = "W0110"
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        '''     WF寸法のチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function WfCheck(info As RodEndUnitNormalViewModel, ByRef errorMessage As String) As Boolean
            Dim wfRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "WF")

            If wfRow IsNot Nothing Then

                Dim wfCustomValue = wfRow.CustomValue
                Dim wfStandardValue = wfRow.StandardValue

                Dim intWfCustom = CType(wfCustomValue, Double)
                Dim intWfStandard = CType(wfStandardValue, Double)

                If intWfCustom < intWfStandard Then
                    errorMessage = "W8460"
                    Return False
                End If
            End If

            Return True
        End Function

        ''' <summary>
        '''     WF寸法のチェック（SCA2）
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function WfCheckSca2(info As RodEndUnitNormalViewModel,
                                     symbols As List(Of String),
                                     keyKataban As String,
                                     ByRef errorMessage As String) As Boolean

            Dim wfRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "WF")

            If wfRow IsNot Nothing Then

                If Not String.IsNullOrEmpty(wfRow.CustomValue) Then

                    Select Case keyKataban
                        Case "", "V", "2"
                            If InStr(1, symbols(13), "J") <> 0 Or InStr(1, symbols(13), "L") <> 0 Then
                                errorMessage = "W8980"
                                Return False
                            End If
                        Case "B", "C"
                            If InStr(1, symbols(17), "J") <> 0 Or InStr(1, symbols(17), "L") <> 0 Then
                                errorMessage = "W8980"
                                Return False
                            End If
                        Case "D", "E"
                            If InStr(1, symbols(12), "J") <> 0 Or InStr(1, symbols(12), "L") <> 0 Then
                                errorMessage = "W8980"
                                Return False
                            End If
                    End Select


                End If

            End If

            Return True
        End Function

        ''' <summary>
        '''     AとKL寸法のチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="minASize"></param>
        ''' <param name="minKlSize"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function AklCheck(info As RodEndUnitNormalViewModel,
                                  minASize As Double,
                                  minKlSize As Double,
                                  ByRef errorMessage As String) As Boolean

            Select Case info.PatternSymbol
                Case RodEndPatternDiv.N13, RodEndPatternDiv.N15
                    Dim aRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "A")

                    If Not String.IsNullOrEmpty(aRow.CustomValue) Then
                        Dim intACustom = CType(aRow.CustomValue, Double)
                        Dim intAStandard = CType(aRow.StandardValue, Double)

                        If intACustom < minASize OrElse
                           intACustom > intAStandard * 2 Then
                            errorMessage = "W8440"
                            Return False
                        End If
                    End If
                Case RodEndPatternDiv.N11, RodEndPatternDiv.N1
                    Dim klRow = info.Rows.FirstOrDefault(Function(r) r.ExternalForm = "KL")

                    If Not String.IsNullOrEmpty(klRow.CustomValue) Then
                        Dim intKlCustom = CType(klRow.CustomValue, Double)
                        Dim intKlStandard = CType(klRow.StandardValue, Double)

                        If intKlCustom < minKlSize OrElse
                           intKlCustom > intKlStandard * 1.5 Then
                            errorMessage = "W8450"
                            Return False
                        End If
                    End If
            End Select

            Return True
        End Function

        ''' <summary>
        '''     数字のチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function NumericCheck(info As RodEndUnitNormalViewModel,
                                      externalForm As String,
                                      ByRef errorMessage As String) As Boolean
            Dim strValue As String = info.Rows.First(Function(r) r.ExternalForm = externalForm).CustomValue
            Dim doubleValue As Double

            If Not Double.TryParse(strValue, doubleValue) Then
                errorMessage = "W0130"
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        '''     ハイフォンのチェック
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="errorMessage"></param>
        ''' <returns></returns>
        Private Function HyphenCheck(info As RodEndUnitOtherViewModel, exceptForm As List(Of String),
                                     ByRef errorMessage As String) As Boolean

            Dim checkInput = info.CustomValue

            For Each s As String In exceptForm
                checkInput = checkInput.Replace(s, String.Empty)
            Next
            If checkInput.Contains(MyControlChars.Hyphen) Then
                errorMessage = "W8570"
                Return False
            End If
            Return True
        End Function

#End Region

#End Region

#End Region
    End Class
End Namespace