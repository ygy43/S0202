
Imports KatabanBusinessLogic.KatabanWcfService
Imports KatabanBusinessLogic.Managers
Imports KatabanBusinessLogic.Models
Imports KatabanCommon.Constants
Imports KatabanCommon.Utility
Imports S0202.Filters
Imports S0202.MyHelpers
Imports S0202.ViewModels.Options

Namespace Controllers
    <AuthorizeFilter>
    Public Class OptionsController
        Inherits Controller

#Region "画面ロード"

        ' GET: Options/Index
        Function Index() As ActionResult

            '画面作成
            Dim model As New OptionsIndexViewModel

            '選択した機種情報を取得
            Dim selectedSeriesInfo As SeriesInfo = Session(SessionKeys.SelectedSeriesData)

            model.SelectedSeriesInfo = selectedSeriesInfo
            model.FocusSeqNo = 1

            With selectedSeriesInfo

                '形番構成を取得
                Dim katabanStructureData = OptionManager.GetKatabanStructureInfo(.series_kataban, .key_kataban, "ja")

                '形番構成情報をセッションに保存
                Session(SessionKeys.KatabanStructureData) = katabanStructureData

                '形番構成情報をモデルに保存
                model.KatabanStructureInfos = katabanStructureData

                '選択した構成情報を初期化
                model.SelectedStructureInfos = ListMethods.SetEmptyList(katabanStructureData.Count)

                '警告メッセージの取得
                model.Messages = OptionManager.GetMessages(.series_kataban, .key_kataban)

                'ボタンの表示設定
                model.IsShowRodEnd = OptionManager.IsShowRodEnd(.series_kataban, .key_kataban)
                model.IsShowOtherOption = OptionManager.IsShowOtherOption(.series_kataban, .key_kataban)
                model.IsShowStopper = OptionManager.IsShowStopper(.series_kataban, .key_kataban)
                model.IsShowMotor1 = OptionManager.IsShowMotor1(.series_kataban, .key_kataban)
                model.IsShowMotor2 = OptionManager.IsShowMotor2(.series_kataban, .key_kataban)
                model.IsShowPortPosition = OptionManager.IsShowPortPosition(.series_kataban, .key_kataban)
                model.IsShowStock = OptionManager.IsShowStock(.series_kataban, .key_kataban)
            End With

            Return View(model)
        End Function

        ' POST: Options/Index
        <HttpPost>
        Function Index(model As OptionsIndexViewModel) As ActionResult

            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()

            'オプションチェック
            Dim checkResult = OptionManager.ValidateInput(selectedData)

            If checkResult.IsSucceed Then

                '価格画面へ遷移
                Return RedirectToAction("Index", "Prices")
            Else

                'エラーメッセージを表示し、フォカスを設定
                For Each err As String In checkResult.Errors
                    ModelState.AddModelError("", err)
                Next
                model.SelectedSeriesInfo = selectedData.Series
                model.KatabanStructureInfos = selectedData.KatabanStructures
                model.FocusSeqNo = checkResult.ErrorSeqNo

                Return View(model)
            End If
        End Function

#End Region

#Region "付加情報画面へ遷移"

        ' POST: Options/ShowRodEnd
        <HttpPost>
        Function ShowRodEnd(model As OptionsIndexViewModel) As ActionResult

            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            Return RedirectToAction("Index", "RodEnd")
        End Function

        ' POST: Options/ShowOtherOption
        <HttpPost>
        Function ShowOtherOption(model As OptionsIndexViewModel) As ActionResult

            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()


            Return View("OtherOption")
        End Function

        ' POST: Options/ShowStopper
        <HttpPost>
        Function ShowStopper(model As OptionsIndexViewModel) As ActionResult
            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()


            Return View("Stopper")
        End Function

        ' POST: Options/ShowMotor1
        <HttpPost>
        Function ShowMotor1(model As OptionsIndexViewModel) As ActionResult
            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()


            Return View("Motor1")
        End Function

        ' POST: Options/ShowMotor2
        <HttpPost>
        Function ShowMotor2(model As OptionsIndexViewModel) As ActionResult
            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()


            Return View("Motor2")
        End Function

        ' POST: Options/ShowPortPosition
        <HttpPost>
        Function ShowPortPosition(model As OptionsIndexViewModel) As ActionResult
            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()


            Return View("PortPosition")
        End Function

        ' POST: Options/ShowStock
        <HttpPost>
        Function ShowStock(model As OptionsIndexViewModel) As ActionResult
            Dim selectedStructureInfos = model.SelectedStructureInfos

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructureInfos
            Session(SessionKeys.CurrentSeqNo) = model.FocusSeqNo

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()


            Return View("Stock")
        End Function

#End Region

#Region "AJAX"

        ' Ajax: Options/ValidateInputBySeqNo
        Function ValidateInputBySeqNo(structures As String,
                                      seqNo As Integer) As ActionResult

            If String.IsNullOrEmpty(structures) Then Return Content("")

            Dim selectedStructures = structures.Split("|"c).ToList()

            '選択した構成情報をセッションに保存
            Session(SessionKeys.SelectedStructureData) = selectedStructures

            '形番構成情報を取得
            Dim selectedData = SessionHelper.GetSelectedData()

            '検証
            Dim validateResult = OptionManager.ValidateInputBySeqNo(selectedData, seqNo)

            If validateResult.IsSucceed Then
                Return Content("")
            Else
                Return _
                    Content(
                        validateResult.ErrorSeqNo & MyControlChars.Pipe &
                        String.Join(ControlChars.NewLine, validateResult.Errors))
            End If
        End Function

        ' Ajax: Options/UpdateOptions
        '<OutputCache(NoStore:=True, Duration:=0)>
        Function UpdateOptions(structureName As String,
                               focusSeqNo As Integer,
                               structureNumber As Integer,
                               structures As String,
                               structureDiv As String) As ActionResult

            Dim optionList As New List(Of KatabanStructureOptionInfo)

            Dim seriesInfo As SeriesInfo = Session(SessionKeys.SelectedSeriesData)

            '選択した構成
            Dim selectedStructures = structures.Split("|"c).ToList()

            selectedStructures = ListMethods.SetEmptyList(structureNumber + 1, selectedStructures)

            '構成オプションを取得
            optionList = OptionManager.GetKatabanStructureOptionInfo(seriesInfo.series_kataban,
                                                                     seriesInfo.key_kataban,
                                                                     focusSeqNo,
                                                                     "ja",
                                                                     selectedStructures,
                                                                     structureDiv)

            If optionList.Count = 0 Then
                Return Content("")
            ElseIf optionList.Count = 1 Then
                '唯一のオプションの場合
                Return Content(optionList.First.option_symbol)
            Else
                '複数オプションの場合
                Dim model As New OptionsUpdateOptionsViewModel
                model.StructureName = structureName
                model.CurrentOptions = optionList

                Return PartialView("_CurrentOptions", model)
            End If
        End Function

#End Region
    End Class
End Namespace