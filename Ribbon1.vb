Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Word

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        MsgBox("测试")

    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub SplitButton1_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub DropDown1_SelectionChanged(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        Dim pgCount As Integer
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        For pgCount = 1 To wdApp.Selection.Paragraphs.Count
            With wdApp.Selection.Paragraphs(pgCount)
                .Reset
                .Alignment = WdParagraphAlignment.wdAlignParagraphJustify
                .CharacterUnitFirstLineIndent = 2
                .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
                .LineSpacing = 20
                With .Range.Font
                    .Name = "宋体"
                    .Size = 12
                    .Bold = False
                    .Italic = False
                    .Outline = False
                    .Shadow = False
                    .Underline = WdUnderline.wdUnderlineNone
                    .Scaling = 100
                    .ColorIndex = WdColorIndex.wdBlack

                End With

            End With

        Next pgCount

    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        wdApp.Visible = True
        objSelection = wdApp.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp3 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
        GoTo EndOk

ErrL:
        ZhangTemplate()
        temp3 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
EndOk:
        Dim listLevel = temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplate(ListTemplate:=temp3)

        'Set name_num = objSelection.Range.ListForma.ListString
        objSelection.Font.Name = "黑体"
        objSelection.Font.Size = 12
        If objSelection.Range.ListFormat.ListLevelNumber > 1 Then

            '删除项目符号
            objSelection.TypeBackspace()
            '增加换行
            objSelection.TypeParagraph()
        End If
    End Sub
    Private Sub ZhangTemplate()
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 As Word.ListTemplate = wdApp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)

        With temp3.ListLevels(1)
            .NumberFormat = "%1"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(2)
            .NumberFormat = "%1.%2"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .TabPosition = False
            .ResetOnHigher = 1
            .StartAt = 1
            With .Font

                .Size = 12

                .Name = "Times New Roman"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(3)
            .NumberFormat = "%1.%2.%3"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .TabPosition = False
            .ResetOnHigher = 2
            .StartAt = 1
            With .Font

                .Size = 12

                .Name = "Times New Roman"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(4)
            .NumberFormat = "%1.%2.%3.%4"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .TabPosition = False
            .ResetOnHigher = 3
            .StartAt = 1
            With .Font
                .Size = 12
                .Name = "Times New Roman"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(5)
            .NumberFormat = "%1.%2.%3.%4.%5"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .TabPosition = False
            .ResetOnHigher = 4
            .StartAt = 1
            With .Font

                .Size = 12

                .Name = "Times New Roman"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(6)
            .NumberFormat = "%1.%2.%3.%4.%5.%6"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .TabPosition = False
            .ResetOnHigher = 5
            .StartAt = 1
            With .Font
                .Size = 12
                .Name = "Times New Roman"
            End With
            .LinkedStyle = ""
        End With
        temp3.Name = "初始化章模板3"
    End Sub

    Private Sub Button32_Click(sender As Object, e As RibbonControlEventArgs) Handles Button32.Click, Button20.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp4 As Word.ListTemplate
        Dim objSelection
        wdApp.Visible = True
        objSelection = wdApp.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp4 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
        GoTo EndOk

ErrL:
        ZhangTemplate()

        temp4 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
EndOk:
        'Dim listLevel = temp4.ListLevels.Item(2)
        '删除项目符号
        'Selection.TypeBackspace
        'objSelection.Range.ListFormat.RemoveNumbers
        'Selection.TypeParagraph
        objSelection.HomeKey(Unit:=WdUnits.wdLine)
        objSelection.TypeText(Constants.vbTab)
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp4, ApplyLevel:=2)
        With objSelection.ParagraphFormat
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 22
            .FirstLineIndent = wdApp.CentimetersToPoints(0)
        End With
        objSelection.Font.Name = "宋体"
        objSelection.Font.Size = 12

        'Set name_num = objSelection.Range.ListForma.ListString

        'If objSelection.Range.ListFormat.ListLevelNumber > 1 Then

        '删除项目符号
        ' Selection.TypeBackspace
        '增加换行
        'Selection.TypeParagraph
        'End If
    End Sub

    Private Sub Button33_Click(sender As Object, e As RibbonControlEventArgs) Handles Button33.Click, Button21.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp5 As Word.ListTemplate
        Dim objSelection
        wdApp.Visible = True
        objSelection = wdApp.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp5 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
        GoTo EndOk

ErrL:
        ZhangTemplate()

        temp5 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
EndOk:
        'Dim listLevel = temp5.ListLevels.Item(3)

        '删除项目符号
        'objSelection.TypeBackspace()
        'objSelection.Range.ListFormat.RemoveNumbers
        'Selection.TypeParagraph
        objSelection.HomeKey(Unit:=WdUnits.wdLine)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp5, ApplyLevel:=3)
        With objSelection.ParagraphFormat
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 22
            .FirstLineIndent = wdApp.CentimetersToPoints(0)
        End With
        objSelection.Font.Name = "宋体"
        objSelection.Font.Size = 12

        'Set name_num = objSelection.Range.ListForma.ListString

        'If objSelection.Range.ListFormat.ListLevelNumber > 1 The

        '删除项目符号
        ' Selection.TypeBackspace
        '增加换行
        'Selection.TypeParagraph
        'End If

    End Sub

    Private Sub Button34_Click(sender As Object, e As RibbonControlEventArgs) Handles Button34.Click, Button28.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp6 As Word.ListTemplate
        Dim objSelection
        wdApp.Visible = True
        objSelection = wdApp.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp6 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
        GoTo EndOk

ErrL:
        ZhangTemplate()

        temp6 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
EndOk:
        'Dim listLevel = temp6.ListLevels.Item(4)
        '删除项目符号
        'Selection.TypeBackspace
        'objSelection.Range.ListFormat.RemoveNumbers
        'Selection.TypeParagraph
        objSelection.HomeKey(Unit:=WdUnits.wdLine)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp6, ApplyLevel:=4)
        With objSelection.ParagraphFormat
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 22
            .FirstLineIndent = wdApp.CentimetersToPoints(0)
        End With
        objSelection.Font.Name = "宋体"
        objSelection.Font.Size = 12
        'Set name_num = objSelection.Range.ListForma.ListString

        'If objSelection.Range.ListFormat.ListLevelNumber > 1 Then

        '删除项目符号
        ' Selection.TypeBackspace
        '增加换行
        'Selection.TypeParagraph
        'End If

    End Sub

    Private Sub Button76_Click(sender As Object, e As RibbonControlEventArgs) Handles Button76.Click, Button29.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp7 As Word.ListTemplate
        Dim objSelection
        wdApp.Visible = True
        objSelection = wdApp.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp7 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
        GoTo EndOk

ErrL:
        ZhangTemplate()

        temp7 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
EndOk:
        'Dim listLevel = temp7.ListLevels.Item(5)
        '删除项目符号
        'Selection.TypeBackspace
        'objSelection.Range.ListFormat.RemoveNumbers
        'Selection.TypeParagraph
        objSelection.HomeKey(Unit:=WdUnits.wdLine)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp7, ApplyLevel:=5)
        With objSelection.ParagraphFormat
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 22
            .FirstLineIndent = wdApp.CentimetersToPoints(0)
        End With
        objSelection.Font.Name = "宋体"
        objSelection.Font.Size = 12
        'Set name_num = objSelection.Range.ListForma.ListString

        'If objSelection.Range.ListFormat.ListLevelNumber > 1 Then

        '删除项目符号
        ' Selection.TypeBackspace
        '增加换行
        'Selection.TypeParagraph
        'End If

    End Sub

    Private Sub Button35_Click(sender As Object, e As RibbonControlEventArgs) Handles Button35.Click, Button30.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim temp8 As Word.ListTemplate
        Dim objSelection
        wdApp.Visible = True
        objSelection = wdApp.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp8 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
        GoTo EndOk

ErrL:
        ZhangTemplate()

        temp8 = wdApp.ActiveDocument.ListTemplates.Item(Index:="初始化章模板3")
EndOk:
        'Dim listLevel = temp8.ListLevels.Item(6)
        '删除项目符号
        'Selection.TypeBackspace
        'objSelection.Range.ListFormat.RemoveNumbers
        'Selection.TypeParagraph
        objSelection.HomeKey(Unit:=WdUnits.wdLine)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        objSelection.TypeText(Constants.vbTab)
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp8, ApplyLevel:=6)
        With objSelection.ParagraphFormat
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 22
            .FirstLineIndent = wdApp.CentimetersToPoints(0)
        End With
        objSelection.Font.Name = "宋体"
        objSelection.Font.Size = 12
        'Set name_num = objSelection.Range.ListForma.ListString

        'If objSelection.Range.ListFormat.ListLevelNumber > 1 Then

        '删除项目符号
        ' Selection.TypeBackspace
        '增加换行
        'Selection.TypeParagraph
        'End If

    End Sub

    Private Sub remark()
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        DiyParagraph()
        Dim LT As Word.ListTemplate = wdApp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With LT.ListLevels(1)
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .TextPosition = wdApp.CentimetersToPoints(0.74)
            .NumberFormat = "%" & "注" & "："
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .ResetOnHigher = 0
            .StartAt = 1

            With .Font
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Size = 10.5

                .Name = "仿宋"
            End With
        End With
        wdApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=
            LT, ContinuePreviousList:=
            False, ApplyTo:=WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=
            WdDefaultListBehavior.wdWord10ListBehavior)
        With wdApp.Selection.Paragraphs(1).Range.Font
            .Size = 10.5
            .Name = "仿宋"
        End With
    End Sub
    Private Sub DiyParagraph()
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim pgCount As Integer
        For pgCount = 1 To wdApp.Selection.Paragraphs.Count
            With wdApp.Selection.Paragraphs(pgCount)
                .Reset()
                .Alignment = WdParagraphAlignment.wdAlignParagraphJustify
                .CharacterUnitFirstLineIndent = 2
                .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
                .LineSpacing = 20
                With .Range.Font
                    .Name = "宋体"
                    .Size = 12
                    .Bold = False
                    .Italic = False
                    .Outline = False
                    .Shadow = False
                    .Underline = WdUnderline.wdUnderlineNone
                    .Scaling = 100
                    .ColorIndex = WdColorIndex.wdBlack

                End With

            End With

        Next pgCount

    End Sub


    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        remark()
        '换行
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        wdApp.Selection.TypeText(Text:="" & vbCrLf)
        remark1()
    End Sub
    Sub remark1()

        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim LT As Word.ListTemplate = wdApp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With LT.ListLevels(1)
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .TextPosition = wdApp.CentimetersToPoints(0.74)
            .NumberFormat = "%1"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft

            .ResetOnHigher = 0
            .StartAt = 1

            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False

                .Size = 10.5
                .Name = "仿宋"
            End With
        End With
        wdApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=
            LT, ContinuePreviousList:=
            False, ApplyTo:=WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=
            WdDefaultListBehavior.wdWord10ListBehavior)
        With wdApp.Selection.Paragraphs(1).Range.Font
            .Size = 10.5
            .Name = "仿宋"
        End With


    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        remark()
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        example_no()
    End Sub

    Private Sub example_no()
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例模板")
        GoTo EndOk

ErrL:
        Example_Template()

        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例模板")
EndOk:
        'temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)

    End Sub
    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        DiyParagraph()
        tableTip(1)
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim LT As Word.ListTemplate = wdApp.ListGalleries(WdListGalleryType.wdNumberGallery).ListTemplates(1)
        wdApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=
            LT, ContinuePreviousList:=
            True, ApplyTo:=WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=
            WdDefaultListBehavior.wdWord10ListBehavior)
        With wdApp.Selection.Paragraphs(1)
            .Alignment = WdParagraphAlignment.wdAlignParagraphCenter

            With .Range.Font
                .Size = 10.5
                .Name = "黑体"
            End With
        End With
    End Sub
    Private Sub tableTip(tableTipIndex As Integer)
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim LT As Word.ListTemplate = wdApp.ListGalleries(WdListGalleryType.wdNumberGallery).ListTemplates(1)
        With LT.ListLevels(1)
            .NumberFormat = "表" & "%1" & "  "
            .TrailingCharacter = WdTrailingCharacter.wdTrailingNone
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)

            .ResetOnHigher = 0
            .StartAt = tableTipIndex

            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False

                .Size = 10.5
                .Name = "黑体"
            End With
        End With
        LT.Name = "表题"

    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click
        DiyParagraph()
        pictureTip()
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim LT As Word.ListTemplate = wdApp.ListGalleries(WdListGalleryType.wdNumberGallery).ListTemplates(2)
        wdApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=
            LT, ContinuePreviousList:=
            True, ApplyTo:=WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=
            WdDefaultListBehavior.wdWord10ListBehavior)

        With wdApp.Selection.Paragraphs(1)
            .Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            With .Range.Font
                .Size = 10.5
                .Name = "宋体"
            End With
        End With
    End Sub

    Sub pictureTip()
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim LT As Word.ListTemplate = wdApp.ListGalleries(WdListGalleryType.wdNumberGallery).ListTemplates(2)
        With LT.ListLevels(1)
            .NumberFormat = "图" & "%1" & "  "
            .TrailingCharacter = WdTrailingCharacter.wdTrailingNone
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wdApp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wdApp.CentimetersToPoints(0)
            .ResetOnHigher = 0
            .StartAt = 1

            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                '  .Color = wdUndefined
                .Size = 10.5

                .Name = "宋体"
            End With
        End With
        LT.Name = "图题"

    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click

        Dim doc As Document
        doc = Globals.ThisAddIn.Application.Documents.Add(Template:="C:/StandardDocEditor/template/1.dotx" _
        , NewTemplate:=False, DocumentType:=0)
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        Dim doc As Document
        doc = Globals.ThisAddIn.Application.Documents.Add(Template:="C:/StandardDocEditor/template/2.dotx" _
        , NewTemplate:=False, DocumentType:=0)
    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        Dim doc As Document
        doc = Globals.ThisAddIn.Application.Documents.Add(Template:="C:/StandardDocEditor/template/3.dotx" _
        , NewTemplate:=False, DocumentType:=0)
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        Dim doc As Document
        doc = Globals.ThisAddIn.Application.Documents.Add(Template:="C:/StandardDocEditor/template/4.dotx" _
        , NewTemplate:=False, DocumentType:=0)

    End Sub

    Private Sub SplitButton2_Click(sender As Object, e As RibbonControlEventArgs) Handles SplitButton6.Click, SplitButton14.Click, SplitButton12.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        '删除目录
        Dim muluCount = wdApp.Selection.Application.ActiveDocument.TablesOfContents.Count
        If muluCount > 0 Then
            'Dim muluTable = wdApp.Selection.Application.ActiveDocument.TablesOfContents(1)
            'muluTable.Application.Selection.MoveUp(Count:=1)
            'muluTable.Delete()

            'wdApp.Selection.TypeBackspace()
            MsgBox("目录已经存在，请先手动删除目录")
        Else
            Dim dialo_mulu = New Dialog_mulu
            dialo_mulu.ShowDialog()
            If dialo_mulu.DialogResult = System.Windows.Forms.DialogResult.OK Then

                '插入下一页面
                wdApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
                wdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage)
                wdApp.Selection.TypeParagraph()
                wdApp.Selection.MoveUp(Count:=1)

                'wdApp.Selection.MoveDown(Count:=2)
                wdApp.Selection.Font.Name = "黑体"
                wdApp.Selection.Font.Size = 14
                wdApp.Selection.Font.Bold = WdConstants.wdToggle
                wdApp.Selection.TypeText("目录")
                wdApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter



                Dim level As Integer = 1
                '一级条标题
                Dim checkbox1 = dialo_mulu.CheckBox1
                If checkbox1.Checked Then
                    level = 2
                End If
                '二级条标题
                Dim checkbox2 = dialo_mulu.CheckBox2
                If checkbox2.Checked Then
                    level = 3
                End If
                '三级条标题
                Dim checkbox3 = dialo_mulu.CheckBox3
                If checkbox3.Checked Then
                    level = 4
                End If
                '四级条标题
                Dim checkbox4 = dialo_mulu.CheckBox4
                If checkbox4.Checked Then
                    level = 5
                End If
                '五级条标题
                Dim checkbox5 = dialo_mulu.CheckBox5
                If checkbox5.Checked Then
                    level = 6
                End If
                '图表标题
                Dim checkbox6 = dialo_mulu.CheckBox6
                If checkbox6.Checked Then
                    level = 7
                End If

                '处理二级以上目录左对齐
                Dim pgCount As Integer
                Dim tocLevel As String
                For pgCount = 2 To 7
                    tocLevel = "TOC " & pgCount
                    With wdApp.Selection.Application.ActiveDocument.Styles(tocLevel)
                        .AutomaticallyUpdate = True
                    End With
                    With wdApp.Selection.Application.ActiveDocument.Styles(tocLevel).ParagraphFormat

                        .LeftIndent = wdApp.CentimetersToPoints(0)
                        .RightIndent = wdApp.CentimetersToPoints(0)
                        .CharacterUnitLeftIndent = 0
                    End With
                    wdApp.Selection.Application.ActiveDocument.Styles(tocLevel).NoSpaceBetweenParagraphsOfSameStyle = False
                Next pgCount



                With wdApp.Selection.Application.ActiveDocument
                    .TablesOfContents.Add(wdApp.Selection.Range, RightAlignPageNumbers:=
                    True, UseHeadingStyles:=True, UpperHeadingLevel:=1,
                    LowerHeadingLevel:=level, IncludePageNumbers:=True, AddedStyles:="",
                    UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:=
                    True)
                    .TablesOfContents(1).TabLeader = WdTabLeader.wdTabLeaderDots
                    .TablesOfContents.Format = WdIndexType.wdIndexIndent
                End With
                wdApp.Selection.Find.ClearFormatting()
                wdApp.Selection.Find.Replacement.ClearFormatting()

                'With wdApp.Selection.Find

                With wdApp.Selection.Application.ActiveDocument.TablesOfContents.Application.Selection.Find
                    .Text = "([0-9]{1,})"
                    .Replacement.Text = "（\1）"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchByte = False
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                End With
                ' wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll, Wrap:=WdFindWrap.wdFindStop)
                wdApp.Selection.Find.Execute()
                wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
                wdApp.Selection.Find.ClearFormatting()
                wdApp.Selection.Find.Replacement.ClearFormatting()
                'With wdApp.Selection.Find
                With wdApp.Selection.Application.ActiveDocument.TablesOfContents.Application.Selection.Find
                    .Text = "(?)（"
                    .Replacement.Text = "\1"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchByte = False
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                End With
                'wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll, Wrap:=WdFindWrap.wdFindStop)
                wdApp.Selection.Find.Execute()
                wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
                wdApp.Selection.Find.ClearFormatting()
                wdApp.Selection.Find.Replacement.ClearFormatting()

                'With wdApp.Selection.Find
                With wdApp.Selection.Application.ActiveDocument.TablesOfContents.Application.Selection.Find
                    .Text = "）(?)"
                    .Replacement.Text = "\1"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchByte = False
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                End With
                'wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll, Wrap:=WdFindWrap.wdFindStop)
                wdApp.Selection.Find.Execute()
                wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
                wdApp.Selection.Find.ClearFormatting()
                wdApp.Selection.Find.Replacement.ClearFormatting()

                'With wdApp.Selection.Find
                With wdApp.Selection.Application.ActiveDocument.TablesOfContents.Application.Selection.Find
                    .Text = "（([0-9]{1,}) "
                    .Replacement.Text = "\1 "
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchByte = False
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                End With
                'wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll, Wrap:=WdFindWrap.wdFindStop)
                wdApp.Selection.Find.Execute()
                wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
                ' With wdApp.Selection.Find
                With wdApp.Selection.Application.ActiveDocument.TablesOfContents.Application.Selection.Find
                    .Text = "（([0-9]{1,})."
                    .Replacement.Text = "\1."
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchByte = False
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                End With
                'wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll, Wrap:=WdFindWrap.wdFindStop)
                wdApp.Selection.Find.Execute()
                wdApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
                wdApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
            End If
            dialo_mulu.Dispose()
        End If
    End Sub

    Private Sub connectLine_Click(sender As Object, e As RibbonControlEventArgs) Handles connectLine.Click
        '连接线
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        wapp.Selection.Font.Name = "Times New Roman"
        wapp.Selection.TypeText(" —— ")
        '设置内容格式
        With wapp.Selection
            With .Font
                .Size = 12
            End With

            With .ParagraphFormat
                .LeftIndent = wapp.CentimetersToPoints(0)
                .RightIndent = wapp.CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .CharacterUnitFirstLineIndent = 2
                .LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5


            End With
        End With
    End Sub

    Private Sub itemDescription_Click(sender As Object, e As RibbonControlEventArgs) Handles itemDescription.Click
        '列项说明
        Dim contentArray() As String
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim sTitle As String = InputBox("列项说明开头")

        'sContent = InputBox("列项说明内容，多个用逗号分隔")
        '设置标题格式
        With wapp.Selection
            With .Font
                .Name = "宋体"
                .Size = 12
            End With

            With .ParagraphFormat
                .LeftIndent = wapp.CentimetersToPoints(0)
                .RightIndent = wapp.CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .CharacterUnitFirstLineIndent = 0.2
                .LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5
            End With
        End With
        wapp.Selection.TypeText(sTitle)
        'titleFormat.InsertAfter Text:=sTitle & Chr(10)
        wapp.Selection.TypeParagraph()
        '以下是设置多行内容
        Dim sContent As String = InputBox("列项说明内容，多个之间用中文分号分割")
        contentArray = Split(sContent, "；")
        '设置内容格式
        With wapp.Selection
            With .Font
                .Size = 12
            End With

            With .ParagraphFormat
                .LeftIndent = wapp.CentimetersToPoints(0)
                .RightIndent = wapp.CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .CharacterUnitFirstLineIndent = 2
                .LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5
            End With
        End With
        'Selection.TypeParagraph
        'For Each i In contentArray
        For i = 0 To UBound(contentArray)
            wapp.Selection.Font.Name = "黑体"
            wapp.Selection.TypeText("—— ")
            wapp.Selection.Font.Name = "宋体"
            If UBound(contentArray) = i Then
                wapp.Selection.TypeText(contentArray(i) & "。")
            Else
                wapp.Selection.TypeText(contentArray(i) & "；" & Chr(10))
            End If

        Next i
    End Sub

    Private Sub Button36_Click(sender As Object, e As RibbonControlEventArgs) Handles Button36.Click
        '列项一
        Dim objWord As Word.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord = Globals.ThisAddIn.Application
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        'Set temp3 = objWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化项一模板")
        GoTo EndOk

ErrL:
        First_Mock_Exam_Template()

        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化项一模板")
EndOk:
        'temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)
        '设置段落
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .CharacterUnitFirstLineIndent = 2
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 20
        End With
    End Sub

    Private Sub First_Mock_Exam_Template()
        '项一模板 初始化项一模板
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)

        With temp3.ListLevels(1)
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberFormat = "--"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(0.74)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With

        temp3.Name = "初始化项一模板"
    End Sub

    Private Sub Second_Mock_Exam_Template()
        '项二模板 初始化项二模板
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With temp3.ListLevels(1)
            .NumberFormat = ChrW(61548)
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleBullet
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(0.73)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "Wingdings"
            End With
            .LinkedStyle = ""
        End With

        temp3.Name = "初始化项二模板"
    End Sub

    Private Sub Button37_Click(sender As Object, e As RibbonControlEventArgs) Handles Button37.Click
        '列项二
        Dim objWord As Word.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord = Globals.ThisAddIn.Application
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdBulletGallery).Reset (1)
        'Set temp3 = objWord.ListGalleries(wdBulletGallery).ListTemplates(1)

        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化项二模板")
        GoTo EndOk

ErrL:
        Second_Mock_Exam_Template()


        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化项二模板")
EndOk:
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)
        'Set name_num = objSelection.Range.ListForma.ListString
        '设置段落
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(2.0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .LineSpacing = objWord.LinesToPoints(1)
        End With
    End Sub

    Private Sub Letter_Item_Template()
        '字母项模板
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With temp3.ListLevels(1)
            .NumberFormat = "%1)"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleLowercaseLetter
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(0.35)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "宋体"
            End With
            .LinkedStyle = ""
        End With
        temp3.Name = "初始化字母项模板"
    End Sub

    Private Sub Button38_Click(sender As Object, e As RibbonControlEventArgs) Handles Button38.Click
        '字母项
        Dim objWord As Word.Application
        Dim objDoc
        Dim temp6 As Word.ListTemplate
        Dim objSelection
        objWord = Globals.ThisAddIn.Application
        objWord.Visible = True
        objSelection = objWord.Selection
        objWord.Selection.Range.ListFormat.RemoveNumbers()
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        'Set temp6 = objWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
        On Error GoTo ErrL
        temp6 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化字母项模板")
        GoTo EndOk

ErrL:
        Letter_Item_Template()


        temp6 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化字母项模板")
EndOk:
        'temp6.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp6, ApplyLevel:=1)
        '设置段落
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 2
            .LineUnitBefore = 0
            .LineUnitAfter = 0
        End With
    End Sub
    Private Sub Digital_Item()
        '数字项模板
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With temp3.ListLevels(1)

            .NumberFormat = "%1)"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "宋体"
            End With
            .LinkedStyle = ""
        End With
        temp3.Name = "初始化数字项模板"

    End Sub

    Private Sub Button39_Click(sender As Object, e As RibbonControlEventArgs) Handles Button39.Click
        '数字项
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        'Set temp3 = objWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化数字项模板")
        GoTo EndOk

ErrL:
        Digital_Item()


        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化数字项模板")
EndOk:
        'temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)
        '设置段落
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 3
            .LineUnitBefore = 0
            .LineUnitAfter = 0
        End With
    End Sub

    Private Sub Example_Template()
        '例模板 初始化
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With temp3.ListLevels(1)
            .NumberFormat = "示例："
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With

        temp3.Name = "初始化例模板"
    End Sub
    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs)
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例模板")
        GoTo EndOk

ErrL:
        Example_Template()

        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例模板")
EndOk:
        'temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)
    End Sub
    Private Sub Example_Template_n()
        '例n 模板 初始化
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
        With temp3.ListLevels(1)
            .NumberFormat = "示例%1："
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(0.35)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "宋体"
            End With
            .LinkedStyle = ""
        End With
        temp3.Name = "初始化例n模板1"
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As RibbonControlEventArgs)
        '例n 
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例n模板1")
        GoTo EndOk

ErrL:
        Example_Template_n()

        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例n模板1")
EndOk:
        'temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)
    End Sub

    Private Sub quote_Click(sender As Object, e As RibbonControlEventArgs) Handles quote.Click, Button31.Click, Button41.Click
        '引用
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim sTitle As String = InputBox("引用内容")
        wapp.Selection.Font.Name = "Times New Roman"
        wapp.Selection.TypeText("—— ")
        '设置内容格式
        With wapp.Selection
            With .Font
                .Size = 12
            End With

            With .ParagraphFormat
                .LeftIndent = wapp.CentimetersToPoints(0)
                .RightIndent = wapp.CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .CharacterUnitFirstLineIndent = 2
                .LineSpacingRule = WdLineSpacing.wdLineSpace1pt5
            End With
        End With
        wapp.Selection.TypeText(sTitle)
    End Sub
    Private Sub Appendix_Template()
        '附录模板
        Dim wapp As Word.Application = Globals.ThisAddIn.Application
        Dim temp3 = wapp.ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)
        With temp3.ListLevels(1)
            .NumberFormat = "附录 %A"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 14
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(2)
            .NumberFormat = "%1.%2"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 1
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(3)
            .NumberFormat = "%1.%2.%3"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 2
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(4)
            .NumberFormat = "%1.%2.%3.%4"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 3
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(5)
            .NumberFormat = "%1.%2.%3.%4.%5"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 4
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        With temp3.ListLevels(6)
            .NumberFormat = "%1.%2.%3.%4.%5.%6"
            .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
            .NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
            .NumberPosition = wapp.CentimetersToPoints(0)
            .Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            .TextPosition = wapp.CentimetersToPoints(1.02)
            .TabPosition = False
            .ResetOnHigher = 5
            .StartAt = 1
            With .Font
                .Bold = False
                .Italic = False
                .StrikeThrough = False
                .Subscript = False
                .Superscript = False
                .Shadow = False
                .Outline = False
                .Emboss = False
                .Engrave = False
                .AllCaps = False
                .Hidden = False
                .Underline = False
                .Color = False
                .Size = 12
                .Animation = False
                .DoubleStrikeThrough = False
                .Name = "黑体"
            End With
            .LinkedStyle = ""
        End With
        temp3.Name = "初始化附录模板6"
    End Sub
    Private Sub Button45_Click(sender As Object, e As RibbonControlEventArgs) Handles Button45.Click
        '附录
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
        GoTo EndOk

ErrL:
        Appendix_Template()

        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
EndOk:
        Dim sTitle As String = InputBox("附录标题：")
        '插入下一页面

        objWord.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage)
        objWord.Selection.TypeParagraph()
        objWord.Selection.MoveUp(Count:=1)
        'temp3.ListLevels.Item(1)
        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp3, ApplyLevel:=1)

        'Set name_num = objSelection.Range.ListForma.ListString

        If objSelection.Range.ListFormat.ListLevelNumber > 1 Then

            '删除项目符号
            objWord.Selection.TypeBackspace()
            '增加换行
            objWord.Selection.TypeParagraph()
        End If
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            .Alignment = WdParagraphAlignment.wdAlignParagraphJustify
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = objWord.CentimetersToPoints(0)
            .OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = WdTextboxTightWrap.wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto
        End With
        objWord.Selection.MoveDown(Count:=2)
        objWord.Selection.TypeText(sTitle)
        objWord.Selection.Font.Name = "黑体"
        objWord.Selection.Font.Size = 14
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 6
            .SpaceAfterAuto = False
            .LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 20
            .Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = objWord.CentimetersToPoints(0)
            .OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = WdTextboxTightWrap.wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto
        End With
        '下一行
    End Sub

    Private Sub Button3_Click_2(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp4 As Word.ListTemplate
        Dim objSelection
        objWord.Selection.Range.ListFormat.RemoveNumbers()
        objWord.Visible = True
        objSelection = objWord.Selection

        On Error GoTo ErrL
        temp4 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
        GoTo EndOk

ErrL:
        Appendix_Template()

        temp4 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
EndOk:
        'temp4.ListLevels.Item(2)
        'objWord.Selection.HomeKey Unit:=wdLine
        objSelection.TypeText(Constants.vbTab)
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp4, ApplyLevel:=2)
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .SpaceBefore = 2.5
            .SpaceBeforeAuto = False
            .SpaceAfter = 0.5
            .SpaceAfterAuto = False
            .LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            .Alignment = WdParagraphAlignment.wdAlignParagraphJustify
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = objWord.CentimetersToPoints(0)
            .OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0.5
            .LineUnitAfter = 0.5
            .MirrorIndents = False
            .TextboxTightWrap = WdTextboxTightWrap.wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto
        End With
    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp4 As Word.ListTemplate
        Dim objSelection
        objWord.Selection.Range.ListFormat.RemoveNumbers()
        objWord.Visible = True
        objSelection = objWord.Selection

        On Error GoTo ErrL
        temp4 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
        GoTo EndOk

ErrL:
        Appendix_Template()

        temp4 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
EndOk:
        'temp4.ListLevels.Item(2)
        'objWord.Selection.HomeKey Unit:=wdLine
        objSelection.TypeText(Constants.vbTab)
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp4, ApplyLevel:=3)
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .SpaceBefore = 2.5
            .SpaceBeforeAuto = False
            .SpaceAfter = 0.5
            .SpaceAfterAuto = False
            .LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            .Alignment = WdParagraphAlignment.wdAlignParagraphJustify
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = objWord.CentimetersToPoints(0)
            .OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0.5
            .LineUnitAfter = 0.5
            .MirrorIndents = False
            .TextboxTightWrap = WdTextboxTightWrap.wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto
        End With
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp4 As Word.ListTemplate
        Dim objSelection
        objWord.Selection.Range.ListFormat.RemoveNumbers()
        objWord.Visible = True
        objSelection = objWord.Selection

        On Error GoTo ErrL
        temp4 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
        GoTo EndOk

ErrL:
        Appendix_Template()

        temp4 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化附录模板6")
EndOk:
        'temp4.ListLevels.Item(2)
        'objWord.Selection.HomeKey Unit:=wdLine
        objSelection.TypeText(Constants.vbTab)
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=temp4, ApplyLevel:=4)
        With objWord.Selection.ParagraphFormat
            .LeftIndent = objWord.CentimetersToPoints(0)
            .RightIndent = objWord.CentimetersToPoints(0)
            .SpaceBefore = 2.5
            .SpaceBeforeAuto = False
            .SpaceAfter = 0.5
            .SpaceAfterAuto = False
            .LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            .Alignment = WdParagraphAlignment.wdAlignParagraphJustify
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = objWord.CentimetersToPoints(0)
            .OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0.5
            .LineUnitAfter = 0.5
            .MirrorIndents = False
            .TextboxTightWrap = WdTextboxTightWrap.wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto
        End With
    End Sub

    Private Sub SplitButton13_Click(sender As Object, e As RibbonControlEventArgs) Handles SplitButton13.Click

    End Sub

    Private Sub SplitButton8_Click(sender As Object, e As RibbonControlEventArgs) Handles SplitButton8.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim sTitle As String = InputBox("脚注内容")
        Dim rngIntro As Range = wdApp.Selection.Range
        With rngIntro
            With .FootnoteOptions
                .Location = WdFootnoteLocation.wdBottomOfPage
                .NumberingRule = WdNumberingRule.wdRestartPage
                .StartingNumber = 1
                .NumberStyle = WdNoteNumberStyle.wdNoteNumberStyleNumberInCircle
                .LayoutColumns = 1
            End With

        End With

        wdApp.ActiveDocument.Footnotes.Add(wdApp.Selection.Range, "", sTitle)

        With wdApp.ActiveDocument.Footnotes.Separator
            .ParagraphFormat.RightIndent = wdApp.CentimetersToPoints(12.4)
        End With

        Dim ftCount As Integer
        For ftCount = 1 To wdApp.ActiveDocument.Footnotes.Count

            With wdApp.ActiveDocument.Footnotes.Item(ftCount).Range.Paragraphs(1).Range.Font
                .Size = 9
                .Superscript = False
                .Subscript = False

                .Name = "仿宋"
            End With
        Next ftCount

    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click

        Dim wdApp As Word.Application = Globals.ThisAddIn.Application

    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        example_no()
        '换行
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        wdApp.Selection.TypeText(Text:="" & vbCrLf)
        '例n 
        Dim objWord As Word.Application = Globals.ThisAddIn.Application
        Dim objDoc
        Dim temp3 As Word.ListTemplate
        Dim objSelection
        objWord.Visible = True
        objSelection = objWord.Selection
        objSelection.Range.ListFormat.RemoveNumbers
        'ListGalleries(wdOutlineNumberGallery).Reset (1)
        On Error GoTo ErrL
        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例n模板1")
        GoTo EndOk

ErrL:
        Example_Template_n()

        temp3 = objWord.ActiveDocument.ListTemplates.Item(Index:="初始化例n模板1")
EndOk:
        'temp3.ListLevels.Item(1)

        'Apply formatting to our range
        objSelection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=
            temp3, ContinuePreviousList:=
            False, ApplyTo:=WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=
            WdDefaultListBehavior.wdWord10ListBehavior)
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        Dim dialog_table = New Dialog_table
        dialog_table.ShowDialog()
        If dialog_table.DialogResult = System.Windows.Forms.DialogResult.OK Then
            Dim wdApp As Word.Application = Globals.ThisAddIn.Application
            Dim table As Table = wdApp.ActiveDocument.Tables.Add(wdApp.Selection.Range, dialog_table.NumericUpDown1.Value, dialog_table.NumericUpDown2.Value)
            With table
                .Select()
                wdApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                With wdApp.Selection.Range.Font
                    .Bold = False
                    .Italic = False
                    .StrikeThrough = False
                    .Subscript = False
                    .Superscript = False
                    .Shadow = False
                    .Outline = False
                    .Emboss = False
                    .Engrave = False
                    .AllCaps = False
                    .Hidden = False
                    .Underline = False
                    .Color = False
                    .Size = 10.5
                    .Animation = False
                    .DoubleStrikeThrough = False
                    .Name = "宋体"
                End With
                With .Borders
                    .InsideLineStyle = WdLineStyle.wdLineStyleSingle
                    .OutsideLineStyle = WdLineStyle.wdLineStyleSingle
                    .OutsideLineWidth = WdLineWidth.wdLineWidth100pt

                End With
                .Rows(1).HeadingFormat = True

            End With

        End If
        dialog_table.Dispose()



    End Sub

    Private Sub FontDialog1_Apply(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim tableIndex As Integer = 1
        Dim tableTipIndex As Integer = 1
        Dim table As Table
        Dim tableAfter As Table
        Do While tableIndex <= wdApp.ActiveDocument.Tables.Count
            table = wdApp.ActiveDocument.Tables(tableIndex)
            With table
                '判断table是否跨页,跨页则拆分
                Dim row As Row
                Dim firstRowPage = .Rows(1).Range.Information(WdInformation.wdActiveEndPageNumber)
                Dim lastRowPage = .Rows(.Rows.Count).Range.Information(WdInformation.wdActiveEndPageNumber)
                If lastRowPage = firstRowPage Then
                    tableIndex += 1
                    tableTipIndex += 1
                    Continue Do
                End If
                For Each row In .Rows
                    lastRowPage = row.Range.Information(WdInformation.wdActiveEndPageNumber)
                    If lastRowPage = firstRowPage Then
                        Continue For
                    End If
                    '拆分
                    .Split(row.Index)
                    Exit For
                Next row

                '拆完以后，为后一个表补充表头和表题

                '表头

                tableAfter = wdApp.ActiveDocument.Tables(tableIndex + 1)
                tableAfter.Rows(1).Cells(1).Range.Select()
                Dim unit = Microsoft.Office.Interop.Word.WdUnits.wdLine
                wdApp.Selection.MoveUp(unit, 1)
                lastRowPage = wdApp.Selection.Range.Information(WdInformation.wdActiveEndPageNumber)
                '如果表题还在上一页，那么加一个空行
                If lastRowPage = firstRowPage Then
                    wdApp.Selection.TypeText(Text:="" & vbCrLf)
                End If
                tableContinue(tableTipIndex, "（续）")


                table.Rows(1).Select()
                wdApp.Selection.Copy()

                tableAfter.Rows(1).Select()
                wdApp.Selection.Paste()

                tableAfter.Rows(1).HeadingFormat = True
            End With
            tableIndex += 1

        Loop
    End Sub
    Private Sub tableContinue(tableTipIndex As Integer, tip As String)

        DiyParagraph()
        tableTip(tableTipIndex)
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim LT As Word.ListTemplate = wdApp.ListGalleries(WdListGalleryType.wdNumberGallery).ListTemplates(1)

        wdApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate:=
            LT, ContinuePreviousList:=
            False, ApplyTo:=WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=
            WdDefaultListBehavior.wdWord10ListBehavior)
        With wdApp.Selection.Paragraphs(1)
            .Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            With .Range.Font
                .Size = 10.5
                .Name = "黑体"
            End With
        End With
        wdApp.Selection.TypeText(Text:=tip)

    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim sectionIndex As Integer
        For sectionIndex = 3 To wdApp.ActiveDocument.Sections.Count
            With wdApp.ActiveDocument.Sections(sectionIndex)
                .Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            End With
        Next sectionIndex
    End Sub

    Private Sub Button40_Click(sender As Object, e As RibbonControlEventArgs) Handles Button40.Click
        endLine()
    End Sub
    Private Sub endLine()
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim shpLine As Shape

        'Add a line to the drawing canvas
        '判断当前是否有封底
        '如果有封底，则在倒数第二节的末尾加入终结线
        '如果没有封底，则在倒数第一节的末尾增加终结线
        Dim mySec As Section = wdApp.ActiveDocument.Sections(3)
        mySec.Range.InsertAfter(vbCrLf)
        Dim ss As InlineShape = mySec.Range.Paragraphs(mySec.Range.Paragraphs.Count).Range.InlineShapes.AddHorizontalLineStandard()
        'shpLine = ss.ConvertToShape()

        'shpLine.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        'shpLine.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        With ss.Line
            ' .BeginArrowheadWidth = MsoArrowheadWidth.msoArrowheadWide
            .ForeColor.RGB = RGB(Red:=0, Green:=0, Blue:=0)
            .Weight = 1.5

        End With
        ' shpLine = wdApp.ActiveDocument.Shapes.AddLine(wdApp.CentimetersToPoints(8.7), wdApp.CentimetersToPoints(7.5), wdApp.CentimetersToPoints(12.8), wdApp.CentimetersToPoints(7.5))
        ' shpLine.WrapFormat.Type = WdWrapType.wdWrapFront
        'shpLine.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        'shpLine.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph

        'Add an arrow to the line and sets the color to purple

    End Sub

    Private Sub Button42_Click(sender As Object, e As RibbonControlEventArgs) Handles Button42.Click
        Dim wdApp As Word.Application = Globals.ThisAddIn.Application
        Dim sectionIndex As Integer
        For sectionIndex = 3 To wdApp.ActiveDocument.Sections.Count
            With wdApp.ActiveDocument.Sections(sectionIndex)
                .Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            End With
        Next sectionIndex

    End Sub
End Class
