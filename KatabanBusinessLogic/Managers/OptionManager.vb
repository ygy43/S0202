Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Models
Imports KatabanBusinessLogic.My.Resources
Imports KatabanBusinessLogic.Results
Imports KatabanCommon.Constants

Namespace Managers
    ''' <summary>
    '''     オプション選択画面ビジネスロジック
    ''' </summary>
    Public Class OptionManager

#Region "共通"

        ''' <summary>
        '''     形番構成情報の取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="language">言語</param>
        ''' <returns></returns>
        Public Shared Function GetKatabanStructureInfo(series As String,
                                                       keyKataban As String,
                                                       language As String) As List(Of KatabanStructureInfo)
            Dim katabanStructureInfos As New List(Of KatabanStructureInfo)

            Using client As New DbAccessServiceClient

                '形番構成の取得
                katabanStructureInfos = client.SelectKatabanStructure(series, keyKataban, language)

                '形番構成オプションの取得
                Dim allOptions = client.SelectKatabanStructureOptions(series, keyKataban)

                For i = 0 To katabanStructureInfos.Count - 1
                    Dim info = katabanStructureInfos(i)

                    '電圧の場合
                    If info.element_div = ElementDiv.Voltage Then
                        info.Width = 100
                        Continue For
                    End If

                    'ストロークの場合
                    If info.element_div = ElementDiv.Stroke Then
                        info.Width = 100
                        Continue For
                    End If

                    If info.structure_div >= OptionJudgeLevel.PluralCondition Then
                        '複数選択可能な場合

                        '幅の設定
                        Dim symbolList =
                                allOptions.Where(Function(o) o.ktbn_strc_seq_no = info.ktbn_strc_seq_no).Select(
                                    Function(o) o.option_symbol)

                        '幅 = 文字数合計 * 20
                        info.Width = symbolList.Sum(Function(s) s.Length) * TextBoxWidthUnit

                        'グループ情報を取得
                        info.PluralGroupData = GetPluralGroupInfo(series, keyKataban, i + 1)

                    Else
                        '正常の場合は幅を設定
                        Dim symbolList =
                                allOptions.Where(Function(o) o.ktbn_strc_seq_no = info.ktbn_strc_seq_no).Select(
                                    Function(o) o.option_symbol)

                        '幅 = 最大文字数 * 20
                        info.Width = symbolList.Max(Function(s) s.Length) * TextBoxWidthUnit

                    End If

                    '最小サイズを40にする
                    info.Width += 20
                    '最大サイズを100にする
                    If info.Width > 100 Then info.Width = 100
                Next
            End Using

            If katabanStructureInfos.Count = 0 Then
                '取得失敗
                Throw New NullReferenceException
            Else
                '取得成功
                'インデックスを統一するためNothingを追加
                katabanStructureInfos.Insert(0, New KatabanStructureInfo())
                Return katabanStructureInfos
            End If
        End Function

        ''' <summary>
        '''     警告メッセージを取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <returns></returns>
        Public Shared Function GetMessages(series As String, keyKataban As String) As List(Of String)

            Dim result As New List(Of String)

            Select Case series
                Case "SCA2", "SCG", "SCG-D", "SCG-G", "SCG-G2", "SCG-G3",
                    "SCG-G4", "SCG-M", "SCG-O", "SCG-Q", "SCG-U"
                    result.Add(Message.M001)
                    result.Add(Message.M002)
                    result.Add(Message.M003)
                Case "SSD", "SSD2", "CMK2", "SCM", "STS-B", "STS-M", "STL-B", "STL-M", "JSC3", "SCA2"
                    result.Add(Message.M001)
                    result.Add(Message.M002)
                Case "JSC4"
                    If keyKataban = "2" Then
                        result.Add(Message.M001)
                        result.Add(Message.M002)
                    End If
                Case "SMD2", "SMD2-L", "SMD2-XL", "SMD2-YL", "SMD2-X", "SMD2-Y", "SMD2-M", "SMD2-ML"
                    result.Add(Message.M004)
                Case "SCS"
                    result.Add(Message.M001)
                    result.Add(Message.M002)
                    result.Add(Message.M005)
                Case "SNP", "V3301"
                    result.Add(Message.M009)
                Case "PPD3"
                    result.Add(Message.M010)
            End Select

            result.Add(Message.M008)

            Return result
        End Function
        
        ''' <summary>
        '''     ロッド先端特注表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowRodEnd(series As String, keyKataban As String) As Boolean

            Select Case series
                Case "SSD"
                    Select Case keyKataban
                        Case "L", "4", "P", "E", "R", "S"
                        Case Else
                            Return True
                    End Select
                Case "SCA2"
                    Return True
                Case "SCS2"
                    If keyKataban <> "4" Then
                        Return True
                    End If
                Case "JSC4"
                    If keyKataban = "2" Then
                        Return True
                    End If
                Case "JSC3"
                    Select Case keyKataban
                        Case "R", "S"
                        Case Else
                            Return True
                    End Select
                Case "CMK2"
                    If keyKataban <> "4" Then
                        Return True
                    End If
                Case "SCS"
                    If keyKataban <> "2" Then
                        Return True
                    End If
            End Select

            Return False
        End Function

        ''' <summary>
        '''     オプション外表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowOtherOption(series As String, keyKataban As String) As Boolean

            Select Case series
                Case "SCA2"
                    Return True
                Case "SCS2"
                    If keyKataban <> "4" Then
                        Return True
                    End If
                Case "JSC4"
                    If keyKataban = "2" Then
                        Return True
                    End If
                Case "JSC3"
                    Select Case keyKataban
                        Case "R", "S"
                        Case Else
                            Return True
                    End Select
                Case "SCS"
                    If keyKataban <> "2" Then
                        Return True
                    End If
            End Select

            Return False
        End Function

        ''' <summary>
        '''     ストッパ位置表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowStopper(series As String, keyKataban As String) As Boolean

            Select Case series
                Case "LCG", "LCG-Q", "LCR", "LCR-Q"
                    Return True
            End Select

            Return False
        End Function

        ''' <summary>
        '''     取付モータ仕様表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowMotor1(series As String, keyKataban As String) As Boolean

            Select Case series
                Case "ETV", "ECS", "ECV", "ESM", "EKS"
                    Return True
                Case "ETS", "EBS", "EBR"
                    Select Case keyKataban
                        Case "A", "B", "C", "D"
                        Case Else
                            Return True
                    End Select
            End Select

            Return False
        End Function

        ''' <summary>
        '''     取付方向/取付モータ仕様表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowMotor2(series As String, keyKataban As String) As Boolean

            Select Case series
                Case "ETS", "EBS", "EBR"
                    Select Case keyKataban
                        Case "A", "B", "C", "D"
                            Return True
                    End Select
            End Select

            Return False
        End Function

        ''' <summary>
        '''     操作ポート位置表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowPortPosition(series As String, keyKataban As String) As Boolean
            Select Case series
                Case "IAVB"
                    Return True
            End Select

            Return False
        End Function

        ''' <summary>
        '''     在庫本数一覧表示判断
        ''' </summary>
        ''' <param name="series"></param>
        ''' <param name="keyKataban"></param>
        ''' <returns></returns>
        Public Shared Function IsShowStock(series As String, keyKataban As String) As Boolean
            Select Case series
                Case "EKS"
                    Return True
            End Select

            Return False
        End Function

#End Region

#Region "ELEパタンチェック"

        ''' <summary>
        '''     形番構成オプション情報の取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="focusSeqNo">構成番号</param>
        ''' <returns></returns>
        Public Shared Function GetKatabanStructureOptionInfo(series As String,
                                                             keyKataban As String,
                                                             focusSeqNo As Integer,
                                                             language As String,
                                                             selectedStructures As List(Of String),
                                                             structureDiv As String) _
            As List(Of KatabanStructureOptionInfo)

            Using client As New DbAccessServiceClient

                '全ての構成オプションを取得
                Dim katabanStructureOptionInfos = client.SelectKatabanStructureOptionsBySeqNo(series,
                                                                                              keyKataban,
                                                                                              focusSeqNo,
                                                                                              language)
                '構成オプションの検証
                Return CheckStructureOptions(series,
                                             keyKataban,
                                             structureDiv,
                                             focusSeqNo,
                                             selectedStructures,
                                             katabanStructureOptionInfos)

            End Using
        End Function

        ''' <summary>
        '''     構成オプションの検証
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="structureDiv">構成区分</param>
        ''' <param name="focusSeqNo">構成番号</param>
        ''' <param name="selectedStructures">選択した構成</param>
        ''' <param name="optionsAll">構成オプション</param>
        ''' <returns>表示オプション</returns>
        Private Shared Function CheckStructureOptions(series As String,
                                                      keyKataban As String,
                                                      structureDiv As String,
                                                      focusSeqNo As Integer,
                                                      selectedStructures As List(Of String),
                                                      optionsAll As List(Of KatabanStructureOptionInfo)) _
            As List(Of KatabanStructureOptionInfo)

            Dim results = optionsAll

            '全ての検証ルール
            Dim elePatternAll = GetElePatternInfoAll(series, keyKataban, focusSeqNo)

            Dim intStructureDiv = CInt(structureDiv)

            '複数選択条件検証
            If intStructureDiv >= CInt(OptionJudgeLevel.PluralCondition) Then
                intStructureDiv -= CInt(OptionJudgeLevel.PluralCondition)

                'javascriptで実装されていたので、
                'Dim elePatternPlural = elePatternAll.Where(Function(elePattern) elePattern.search_seq_no = "1").ToList()

                'results = CheckStructureOptionsPlural(selectedStructures,
                '                                      results,
                '                                      elePatternPlural)
            End If

            'Skip条件検証
            If intStructureDiv >= CInt(OptionJudgeLevel.SkipCondition) Then
                intStructureDiv -= CInt(OptionJudgeLevel.SkipCondition)

                Dim elePatternSkip = elePatternAll.Where(Function(elePattern) elePattern.search_seq_no = "2").ToList()

                results = CheckStructureOptionsSkip(selectedStructures,
                                                    results,
                                                    elePatternSkip)

            End If

            '選択条件検証
            If intStructureDiv >= CInt(OptionJudgeLevel.SelectCondition) Then

                Dim elePatternSelect = elePatternAll.Where(Function(elePattern) elePattern.search_seq_no = "3").ToList()

                results = CheckStructureOptionsSelect(selectedStructures,
                                                      results,
                                                      elePatternSelect)

            End If

            Return results
        End Function

        '''' <summary>
        ''''     複数選択条件検証
        '''' </summary>
        '''' <param name="selectedStructures"></param>
        '''' <param name="optionsAll"></param>
        '''' <param name="elePatternAll"></param>
        '''' <returns></returns>
        'Private Shared Function CheckStructureOptionsPlural(selectedStructures As List(Of String),
        '                                                    optionsAll As List(Of KatabanStructureOptionInfo),
        '                                                    elePatternAll As List(Of ElePatternInfo)) _
        '    As List(Of KatabanStructureOptionInfo)
        'End Function

        ''' <summary>
        '''     選択条件検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="optionsAll"></param>
        ''' <param name="elePatternAll"></param>
        ''' <returns></returns>
        Private Shared Function CheckStructureOptionsSelect(selectedStructures As List(Of String),
                                                            optionsAll As List(Of KatabanStructureOptionInfo),
                                                            elePatternAll As List(Of ElePatternInfo)) _
            As List(Of KatabanStructureOptionInfo)
            Dim result As New List(Of KatabanStructureOptionInfo)

            For Each info As KatabanStructureOptionInfo In optionsAll

                '該当するオプションの全ての検証条件を取得
                Dim elePatterns =
                        elePatternAll.Where(Function(pattern) pattern.option_symbol = info.option_symbol).ToList()

                If CheckByAllElePattern(selectedStructures, elePatterns) Then
                    '検証結果がすべてOKの場合は表示リストに追加
                    result.Add(info)
                End If
            Next

            Return result
        End Function

        ''' <summary>
        '''     Skip条件検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="optionsAll"></param>
        ''' <param name="elePatternAll"></param>
        ''' <returns></returns>
        Private Shared Function CheckStructureOptionsSkip(selectedStructures As List(Of String),
                                                          optionsAll As List(Of KatabanStructureOptionInfo),
                                                          elePatternAll As List(Of ElePatternInfo)) _
            As List(Of KatabanStructureOptionInfo)
            Dim result As New List(Of KatabanStructureOptionInfo)

            '検証結果がすべてOKの場合は表示リストに追加
            If CheckByAllElePattern(selectedStructures, elePatternAll) Then
                result.AddRange(optionsAll)
            End If

            Return result
        End Function

        ''' <summary>
        '''     ElePatternによりの検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="elePatterns"></param>
        ''' <returns></returns>
        Private Shared Function CheckByAllElePattern(selectedStructures As List(Of String),
                                                     elePatterns As List(Of ElePatternInfo)) As Boolean
            Dim checkResults As New List(Of Boolean)

            If elePatterns.Count = 0 Then Return True

            'condition_seq_noにより分類して検証
            Dim seqNos = elePatterns.Select(Function(pattern) pattern.condition_seq_no).Distinct()

            For Each seqNo As String In seqNos

                '各condition_seq_noごとの検証
                Dim elePatternsSameSeqNo =
                        elePatterns.Where(Function(pattern) pattern.condition_seq_no = seqNo).ToList()

                'I/Oから始まるElePattern
                Dim ioMarks =
                        elePatternsSameSeqNo.Where(
                            Function(p) _
                                                      p.condition_cd.StartsWith(ElementJudgeDiv.OutSign) Or
                                                      p.condition_cd.StartsWith(ElementJudgeDiv.InSign)).Select(
                                                          Function(r) r.condition_cd).Distinct().ToList()

                If ioMarks.Count = 1 Then
                    'I/O条件は一つしかいない場合は、分類しない
                    '検証
                    If ioMarks(0).Substring(0, 1) = ElementJudgeDiv.InSign Then

                        'Iの場合は、チェック結果がOKなら表示
                        checkResults.Add(CheckByElePatternSameSeqNoAndIO(selectedStructures, elePatternsSameSeqNo))

                    ElseIf ioMarks(0).Substring(0, 1) = ElementJudgeDiv.OutSign Then

                        'Oの場合は、チェック結果がNGなら表示
                        checkResults.Add(Not CheckByElePatternSameSeqNoAndIO(selectedStructures, elePatternsSameSeqNo))

                    End If
                Else
                    'I/O条件は複数存在する場合は、分類
                    For i = 1 To ioMarks.Count - 1

                        Dim elePatternsSameSeqNoAndIo As New List(Of ElePatternInfo)

                        'グループ分ける
                        If i <> ioMarks.Count - 1 Then
                            For Each pattern As ElePatternInfo In elePatternsSameSeqNo

                                If pattern.condition_cd <> ioMarks(i) Then
                                    elePatternsSameSeqNoAndIo.Add(pattern)
                                Else
                                    Exit For
                                End If
                            Next
                        Else
                            Dim index =
                                    elePatternsSameSeqNo.IndexOf(
                                        elePatternsSameSeqNo.First(Function(e) e.condition_cd = ioMarks(i)))
                            elePatternsSameSeqNoAndIo = elePatternsSameSeqNo.GetRange(index,
                                                                                      elePatternsSameSeqNo.Count - index)
                        End If

                        '検証
                        If ioMarks(i).Substring(0, 1) = ElementJudgeDiv.InSign Then

                            'Iの場合は、チェック結果がOKなら表示
                            checkResults.Add(CheckByElePatternSameSeqNoAndIO(selectedStructures,
                                                                             elePatternsSameSeqNoAndIo))

                        ElseIf ioMarks(i).Substring(0, 1) = ElementJudgeDiv.OutSign Then

                            'Oの場合は、チェック結果がNGなら表示
                            checkResults.Add(
                                Not CheckByElePatternSameSeqNoAndIO(selectedStructures, elePatternsSameSeqNoAndIo))

                        End If
                    Next
                End If

            Next
            Return checkResults.All(Function(r) r = True)
        End Function

        ''' <summary>
        '''     condition_seq_noにより分類して検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="elePatternsSameSeqNoAndIO"></param>
        ''' <returns></returns>
        Private Shared Function CheckByElePatternSameSeqNoAndIO(selectedStructures As List(Of String),
                                                                elePatternsSameSeqNoAndIO As List(Of ElePatternInfo)) _
            As Boolean

            Dim checkResults As New List(Of ElePatternCheckResult)
            Dim elePatternsSameBrackets As New List(Of ElePatternInfo)

            '括弧により分類して検証
            Dim currentMark As String = elePatternsSameSeqNoAndIO.First.condition_cd

            For i = 0 To elePatternsSameSeqNoAndIO.Count - 1

                Dim mark As String = elePatternsSameSeqNoAndIO(i).condition_cd

                If currentMark.EndsWith(MyControlChars.LeftBracket) AndAlso
                   mark.EndsWith(MyControlChars.LeftBracket) AndAlso
                   currentMark <> mark Then
                    '検証
                    checkResults.Add(New ElePatternCheckResult(mark.Substring(0, 1),
                                                               CheckByElePatternSameBrackets(selectedStructures,
                                                                                             elePatternsSameBrackets)
                                                               )
                                     )

                    '検証用条件をリセット
                    elePatternsSameBrackets = New List(Of ElePatternInfo)
                    elePatternsSameBrackets.Add(elePatternsSameSeqNoAndIO(i))
                    currentMark = mark
                Else
                    elePatternsSameBrackets.Add(elePatternsSameSeqNoAndIO(i))
                End If
            Next

            '検証
            checkResults.Add(New ElePatternCheckResult(currentMark.Substring(0, 1),
                                                       CheckByElePatternSameBrackets(selectedStructures,
                                                                                     elePatternsSameBrackets)
                                                       )
                             )

            If checkResults.Count = 1 Then
                Return checkResults.Item(0).Result
            Else
                Dim result As Boolean = checkResults.Item(0).Result

                '*/+によりSeqNoグループのTRUE/FALSEを判断
                For i = 1 To checkResults.Count - 1
                    Select Case checkResults(i).ConditionMark
                        Case ElementJudgeDiv.CondAnd
                            result = result And checkResults(i).Result

                        Case ElementJudgeDiv.CondOr

                            result = result Or checkResults(i).Result
                    End Select
                Next

                Return result

            End If
        End Function

        ''' <summary>
        '''     括弧ごとのグループにより検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="elePatternsSameBrackets"></param>
        ''' <returns></returns>
        Private Shared Function CheckByElePatternSameBrackets(selectedStructures As List(Of String),
                                                              elePatternsSameBrackets As List(Of ElePatternInfo)) _
            As Boolean

            Dim checkResults As New List(Of ElePatternCheckResult)

            'condition_seq_no_brにより分類して検証
            Dim seqNoBrs = elePatternsSameBrackets.Select(Function(pattern) pattern.condition_seq_no_br).Distinct()

            For Each seqNoBr As String In seqNoBrs
                '各condition_seq_no_brごとの検証結果
                Dim elePatternsSameSeqNoBr =
                        elePatternsSameBrackets.Where(Function(pattern) pattern.condition_seq_no_br = seqNoBr).ToList()

                Dim checkResultOfSameSeqNoBr As Boolean

                checkResultOfSameSeqNoBr = CheckByElePatternSameSeqNoBr(selectedStructures, elePatternsSameSeqNoBr)

                '同一BRグループの条件は一致することは前提
                checkResults.Add(New ElePatternCheckResult(elePatternsSameSeqNoBr.Item(0).condition_cd.Substring(0, 1),
                                                           checkResultOfSameSeqNoBr))

            Next

            If checkResults.Count = 1 Then
                Return checkResults.Item(0).Result
            Else
                Dim result As Boolean = checkResults.Item(0).Result

                '*/+によりSeqNoグループのTRUE/FALSEを判断
                For i = 1 To checkResults.Count - 1
                    Select Case checkResults(i).ConditionMark
                        Case ElementJudgeDiv.CondAnd
                            result = result And checkResults(i).Result

                        Case ElementJudgeDiv.CondOr

                            result = result Or checkResults(i).Result
                    End Select
                Next

                Return result

            End If
        End Function

        ''' <summary>
        '''     BRグループごとに検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="elePatternsSameSeqNoBr"></param>
        ''' <returns></returns>
        Private Shared Function CheckByElePatternSameSeqNoBr(selectedStructures As List(Of String),
                                                             elePatternsSameSeqNoBr As List(Of ElePatternInfo)) _
            As Boolean

            Dim checkResults As New List(Of ElePatternCheckResult)

            For Each info As ElePatternInfo In elePatternsSameSeqNoBr

                '単一条件のTRUE・FALSEを判断
                Dim checkResultOfSingleCondition = CheckBySingleCondition(selectedStructures, info)
                checkResults.Add(New ElePatternCheckResult(info.condition_cd.Substring(3, 2),
                                                           checkResultOfSingleCondition))
            Next

            If checkResults.Count = 1 Then
                Return checkResults.Item(0).Result
            Else
                'EQ/NEによりBRグループのTRUE/FALSEを判断
                '同一BRグループの条件は一致することは前提
                Select Case checkResults.Item(0).ConditionMark
                    Case ElementJudgeDiv.Equal
                        'EQの場合は、TRUEが存在すれば、BRグループの結果はTRUE
                        If checkResults.Any(Function(r) r.Result = True) Then
                            Return True
                        Else
                            Return False
                        End If
                    Case ElementJudgeDiv.NotEqual
                        'NEの場合は、FALSEが存在しなければ、BRグループの結果はTRUE
                        If checkResults.All(Function(r) r.Result = True) Then
                            Return True
                        Else
                            Return False
                        End If
                    Case Else
                        Return False
                End Select
            End If
        End Function

        ''' <summary>
        '''     単一条件の検証
        ''' </summary>
        ''' <param name="selectedStructures"></param>
        ''' <param name="info"></param>
        ''' <returns></returns>
        Private Shared Function CheckBySingleCondition(selectedStructures As List(Of String),
                                                       info As ElePatternInfo) As Boolean

            Dim index = CInt(info.condition_cd.Substring(1, 2))
            Dim mark = info.condition_cd.Substring(3, 2)

            Select Case mark
                Case ElementJudgeDiv.Equal
                    Return selectedStructures.Item(index) = info.cond_option_symbol
                Case ElementJudgeDiv.NotEqual
                    Return selectedStructures.Item(index) <> info.cond_option_symbol
            End Select

            Throw New NotImplementedException
        End Function

        ''' <summary>
        '''     複数選択可能なオプションのグループ情報を取得
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function GetPluralGroupInfo(series As String,
                                                   keyKataban As String,
                                                   focusSeqNo As Integer) As String
            Dim values As New List(Of String)
            Dim infos = GetElePatternInfoPlural(series, keyKataban, focusSeqNo)

            Dim seqNoBrs = infos.Select(Function(info) info.condition_seq_no_br).Distinct().ToList()

            For Each seqNobr In seqNoBrs
                Dim sameGroupData As List(Of String) =
                        infos.Where(Function(info) info.condition_seq_no_br = seqNobr).Select(
                            Function(info) info.cond_option_symbol).ToList()

                values.Add(String.Join(MyControlChars.Comma, sameGroupData))
            Next

            Return String.Join(MyControlChars.Pipe, values)
        End Function

        ''' <summary>
        '''     形番構成オプション検証情報の取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="focusSeqNo">構成番号</param>
        ''' <returns></returns>
        Private Shared Function GetElePatternInfoAll(series As String,
                                                     keyKataban As String,
                                                     focusSeqNo As Integer) As List(Of ElePatternInfo)


            Using client As New DbAccessServiceClient

                Return client.SelectElePatternInfoAll(series,
                                                      keyKataban,
                                                      focusSeqNo)
            End Using
        End Function

        ''' <summary>
        '''     複数選択可能なオプションの検証情報を取得
        ''' </summary>
        ''' <param name="series">機種</param>
        ''' <param name="keyKataban">キー形番</param>
        ''' <param name="focusSeqNo">構成番号</param>
        ''' <returns></returns>
        Private Shared Function GetElePatternInfoPlural(series As String,
                                                        keyKataban As String,
                                                        focusSeqNo As Integer) As List(Of ElePatternInfo)

            Using client As New DbAccessServiceClient

                Return client.SelectElePatternInfoPlural(series,
                                                         keyKataban,
                                                         focusSeqNo)
            End Using
        End Function

#End Region

#Region "入力検証"

        ''' <summary>
        '''     入力チェック
        ''' </summary>
        ''' <param name="selectedData">選択した全ての情報</param>
        ''' <returns></returns>
        Public Shared Function ValidateInput(selectedData As SelectedInfo) As OptionCheckResult

            Dim result As New OptionCheckResult

            For i = 1 To selectedData.Symbols.Count - 1

                Dim seqNoResult = ValidateInputBySeqNo(selectedData, i)

                If Not seqNoResult.IsSucceed Then
                    Return seqNoResult
                End If
            Next

            Return result
        End Function

        ''' <summary>
        '''     入力チェック
        ''' </summary>
        ''' <param name="selectedData">選択した全ての情報</param>
        ''' <param name="seqNo">構成番号</param>
        ''' <returns></returns>
        Public Shared Function ValidateInputBySeqNo(selectedData As SelectedInfo, seqNo As Integer) As OptionCheckResult

            Dim result As New OptionCheckResult
            Dim symbol = selectedData.Symbols(seqNo)

            If selectedData.KatabanStructures(seqNo).structure_div >= OptionJudgeLevel.PluralCondition Then
                '複数選択の場合
                If Not ValidatePlural(selectedData, symbol, seqNo, result) Then
                    Return result
                End If
            Else
                '正常の場合

                '入力したオプションが選択候補に存在するかどうかのチェック
                Dim options = GetKatabanStructureOptionInfo(selectedData.Series.series_kataban,
                                                            selectedData.Series.key_kataban,
                                                            seqNo,
                                                            "ja",
                                                            selectedData.Symbols,
                                                            selectedData.KatabanStructures(seqNo).structure_div)

                If String.IsNullOrEmpty(symbol) AndAlso options.Count = 0 Then
                    'OK
                ElseIf options.Count <> 0 Then
                    '存在チェック
                    If options.Exists(Function(o) o.option_symbol = symbol) Then

                        If Not symbol.ToUpper.StartsWith(OtherVoltageText.English) Then
                            'OK
                        Else
                            'その他電圧が選択された場合
                            Return New OptionCheckResult(seqNo, New List(Of String) From {"電圧を直接に入力してください。"})
                        End If
                    Else

                        If selectedData.KatabanStructures(seqNo).element_div = ElementDiv.Voltage AndAlso
                           (symbol Like VoltageRegex.Ac Or symbol Like VoltageRegex.Dc) Then
                            '電圧のチェック
                            If Not ValidateVoltage(selectedData, symbol, seqNo, result) Then
                                Return result
                            End If
                        ElseIf selectedData.KatabanStructures(seqNo).element_div = ElementDiv.Stroke Then
                            'ストロークのチェック
                            If Not ValidateStroke(selectedData, symbol, seqNo, result) Then
                                Return result
                            End If
                        Else
                            Return New OptionCheckResult(seqNo, New List(Of String) From {"下記のリストから選択してください。"})
                        End If
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        '''     複数オプションの分解
        ''' </summary>
        ''' <param name="symbol"></param>
        ''' <param name="options"></param>
        ''' <returns></returns>
        Private Shared Function DecomposeSymbol(symbol As String,
                                                options As List(Of KatabanStructureOptionInfo),
                                                ByRef selectedSymbols As List(Of String)) As Boolean

            Dim symbolOriginal As String = symbol

            'オプションをソート
            Dim symbols = options.Select(Function(o) o.option_symbol).ToList().OrderByDescending(Function(o) o.Length)

            For i = 0 To symbolOriginal.Length - 1
                For j As Integer = symbolOriginal.Length - 1 To i Step -1
                    Dim subs = symbolOriginal.Substring(i, j - i + 1)

                    If symbols.Contains(subs) Then
                        selectedSymbols.Add(subs)
                        symbol = symbol.Replace(subs, String.Empty)
                    End If
                Next
            Next

            If String.IsNullOrEmpty(symbol) Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        '''     複数選択可能項目の検証
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function ValidatePlural(selectedData As SelectedInfo,
                                               symbol As String,
                                               index As Integer,
                                               ByRef result As OptionCheckResult) As Boolean
            '複数選択の場合
            If String.IsNullOrEmpty(symbol) Then
                result = New OptionCheckResult()
                Return True
            End If

            '選択可能な候補を取得
            Dim options = GetKatabanStructureOptionInfo(selectedData.Series.series_kataban,
                                                        selectedData.Series.key_kataban,
                                                        index,
                                                        "ja",
                                                        selectedData.Symbols,
                                                        selectedData.KatabanStructures(index).structure_div)

            'オプションを分解
            Dim selectedSymbols As New List(Of String)

            If Not DecomposeSymbol(symbol, options, selectedSymbols) Then
                '分解失敗
                result = New OptionCheckResult(index, New List(Of String) From {"下記のリストから選択してください。"})
                Return False
            Else

                '存在チェック
                For Each selectedSymbol As String In selectedSymbols

                    If Not options.Exists(Function(o) o.option_symbol = selectedSymbol) Then
                        result = New OptionCheckResult(index, New List(Of String) From {"下記のリストから選択してください。"})
                        Return False
                    End If

                Next

                '順番チェック
                Dim queue As New Queue(Of String)(selectedSymbols)

                For Each info As KatabanStructureOptionInfo In options

                    Dim s = queue.Peek()
                    If info.option_symbol = s Then
                        queue.Dequeue()
                        If queue.Count = 0 Then Exit For
                    End If
                Next

                If queue.Count > 0 Then
                    result = New OptionCheckResult(index, New List(Of String) From {"順番をご確認ください。"})
                    Return False
                Else
                    result = New OptionCheckResult()
                    Return True
                End If
            End If
        End Function

        ''' <summary>
        '''     電圧項目の検証
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function ValidateVoltage(selectedData As SelectedInfo,
                                                symbol As String,
                                                index As Integer,
                                                ByRef result As OptionCheckResult) As Boolean

            '電圧の場合は格式をチェック
            Dim strVoltage = symbol.Remove(symbol.Length - 1).Substring(2)
            Dim intVoltage = 0

            If Integer.TryParse(strVoltage, intVoltage) Then

                '電圧範囲をチェック
                If KatabanUtility.CheckVoltage(selectedData, symbol) Then

                    result = New OptionCheckResult()
                    Return True
                Else
                    result = New OptionCheckResult(index, New List(Of String) From {"電圧範囲外。"})
                    Return False
                End If

            Else
                result = New OptionCheckResult(index, New List(Of String) From {"電圧の格式をご確認ください。"})
                Return False
            End If
        End Function

        ''' <summary>
        '''     ストローク項目の検証
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function ValidateStroke(selectedData As SelectedInfo,
                                               symbol As String,
                                               index As Integer,
                                               ByRef result As OptionCheckResult) As Boolean

            'ストロークチェック
            '口径
            Dim boreSize = 0
            For portIndex = 1 To selectedData.KatabanStructures.Count - 1
                Dim info = selectedData.KatabanStructures(portIndex)

                If info.element_div = ElementDiv.Port Then

                    If Integer.TryParse(selectedData.Symbols(portIndex), boreSize) Then
                    Else
                        If selectedData.Series.series_kataban = "ESM" Then

                            If selectedData.Symbols(portIndex) = "ST" Then
                                boreSize = 2
                            ElseIf selectedData.Symbols(portIndex) = "B" Then
                                boreSize = 1
                            End If
                        End If
                    End If
                End If
            Next

            If boreSize <> 0 Then
                Dim stroke = 0

                If Integer.TryParse(symbol, stroke) Then

                    If KatabanUtility.CheckStroke(selectedData,
                                                  boreSize,
                                                  stroke,
                                                  selectedData.Series.country_cd) Then
                        result = New OptionCheckResult()
                        Return True

                    Else
                        result = New OptionCheckResult(index, New List(Of String) From {"ストロークの値が範囲内ではありません。"})
                        Return False
                    End If
                Else
                    result = New OptionCheckResult(index, New List(Of String) From {"ストロークをご確認してください。。"})
                    Return False

                End If
            Else
                result = New OptionCheckResult(index, New List(Of String) From {"ストロークの値が範囲内ではありません。"})
                Return False

            End If
        End Function

#End Region
    End Class
End Namespace