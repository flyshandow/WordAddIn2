Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim SplitButton1 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Dim SplitButton9 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
        Dim SplitButton3 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
        Dim SplitButton10 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
        Dim SplitButton7 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button32 = Me.Factory.CreateRibbonButton
        Me.Button33 = Me.Factory.CreateRibbonButton
        Me.Button34 = Me.Factory.CreateRibbonButton
        Me.Button76 = Me.Factory.CreateRibbonButton
        Me.Button35 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button28 = Me.Factory.CreateRibbonButton
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.Button36 = Me.Factory.CreateRibbonButton
        Me.Button37 = Me.Factory.CreateRibbonButton
        Me.Button38 = Me.Factory.CreateRibbonButton
        Me.Button39 = Me.Factory.CreateRibbonButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button27 = Me.Factory.CreateRibbonButton
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Group8 = Me.Factory.CreateRibbonGroup
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.SplitButton12 = Me.Factory.CreateRibbonSplitButton
        Me.Button46 = Me.Factory.CreateRibbonButton
        Me.SplitButton14 = Me.Factory.CreateRibbonSplitButton
        Me.Button47 = Me.Factory.CreateRibbonButton
        Me.SplitButton6 = Me.Factory.CreateRibbonSplitButton
        Me.Button43 = Me.Factory.CreateRibbonButton
        Me.SplitButton4 = Me.Factory.CreateRibbonSplitButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.SplitButton13 = Me.Factory.CreateRibbonSplitButton
        Me.Button45 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.SplitButton8 = Me.Factory.CreateRibbonSplitButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.SplitButton21 = Me.Factory.CreateRibbonSplitButton
        Me.Button70 = Me.Factory.CreateRibbonButton
        Me.Button71 = Me.Factory.CreateRibbonButton
        Me.Button72 = Me.Factory.CreateRibbonButton
        Me.SplitButton22 = Me.Factory.CreateRibbonSplitButton
        Me.Button73 = Me.Factory.CreateRibbonButton
        Me.Button74 = Me.Factory.CreateRibbonButton
        Me.Button75 = Me.Factory.CreateRibbonButton
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.itemDescription = Me.Factory.CreateRibbonButton
        Me.connectLine = Me.Factory.CreateRibbonButton
        Me.quote = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button42 = Me.Factory.CreateRibbonButton
        Me.Button31 = Me.Factory.CreateRibbonButton
        Me.Button40 = Me.Factory.CreateRibbonButton
        Me.Button41 = Me.Factory.CreateRibbonButton
        SplitButton1 = Me.Factory.CreateRibbonSplitButton
        SplitButton9 = Me.Factory.CreateRibbonSplitButton
        SplitButton3 = Me.Factory.CreateRibbonSplitButton
        SplitButton10 = Me.Factory.CreateRibbonSplitButton
        SplitButton7 = Me.Factory.CreateRibbonSplitButton
        Me.Tab1.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group8.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitButton1
        '
        SplitButton1.Checked = True
        SplitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        SplitButton1.Image = CType(resources.GetObject("SplitButton1.Image"), System.Drawing.Image)
        SplitButton1.Items.Add(Me.Button7)
        SplitButton1.Items.Add(Me.Button8)
        SplitButton1.Items.Add(Me.Button9)
        SplitButton1.Items.Add(Me.Button10)
        SplitButton1.Label = "新建"
        SplitButton1.Name = "SplitButton1"
        AddHandler SplitButton1.Click, AddressOf Me.SplitButton1_Click
        '
        'Button7
        '
        Me.Button7.Label = "国家计量技术规范"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Button8
        '
        Me.Button8.Label = "国家计量检定规程"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        '
        'Button9
        '
        Me.Button9.Label = "地方计量技术规范"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button10
        '
        Me.Button10.Label = "地方计量检定规程"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'SplitButton9
        '
        SplitButton9.Checked = True
        SplitButton9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        SplitButton9.Image = CType(resources.GetObject("SplitButton9.Image"), System.Drawing.Image)
        SplitButton9.Items.Add(Me.Button32)
        SplitButton9.Items.Add(Me.Button33)
        SplitButton9.Items.Add(Me.Button34)
        SplitButton9.Items.Add(Me.Button76)
        SplitButton9.Items.Add(Me.Button35)
        SplitButton9.Label = "条"
        SplitButton9.Name = "SplitButton9"
        AddHandler SplitButton9.Click, AddressOf Me.SplitButton1_Click
        '
        'Button32
        '
        Me.Button32.Image = CType(resources.GetObject("Button32.Image"), System.Drawing.Image)
        Me.Button32.Label = "条1"
        Me.Button32.Name = "Button32"
        Me.Button32.ShowImage = True
        '
        'Button33
        '
        Me.Button33.Image = CType(resources.GetObject("Button33.Image"), System.Drawing.Image)
        Me.Button33.Label = "条2"
        Me.Button33.Name = "Button33"
        Me.Button33.ShowImage = True
        '
        'Button34
        '
        Me.Button34.Image = CType(resources.GetObject("Button34.Image"), System.Drawing.Image)
        Me.Button34.Label = "条3"
        Me.Button34.Name = "Button34"
        Me.Button34.ShowImage = True
        '
        'Button76
        '
        Me.Button76.Image = CType(resources.GetObject("Button76.Image"), System.Drawing.Image)
        Me.Button76.Label = "条4"
        Me.Button76.Name = "Button76"
        Me.Button76.ShowImage = True
        '
        'Button35
        '
        Me.Button35.Image = CType(resources.GetObject("Button35.Image"), System.Drawing.Image)
        Me.Button35.Label = "条5"
        Me.Button35.Name = "Button35"
        Me.Button35.ShowImage = True
        '
        'SplitButton3
        '
        SplitButton3.Checked = True
        SplitButton3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        SplitButton3.Image = CType(resources.GetObject("SplitButton3.Image"), System.Drawing.Image)
        SplitButton3.Items.Add(Me.Button20)
        SplitButton3.Items.Add(Me.Button21)
        SplitButton3.Items.Add(Me.Button28)
        SplitButton3.Items.Add(Me.Button29)
        SplitButton3.Items.Add(Me.Button30)
        SplitButton3.Label = "无题条"
        SplitButton3.Name = "SplitButton3"
        AddHandler SplitButton3.Click, AddressOf Me.SplitButton1_Click
        '
        'Button20
        '
        Me.Button20.Image = CType(resources.GetObject("Button20.Image"), System.Drawing.Image)
        Me.Button20.Label = "条1"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        '
        'Button21
        '
        Me.Button21.Image = CType(resources.GetObject("Button21.Image"), System.Drawing.Image)
        Me.Button21.Label = "条2"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        '
        'Button28
        '
        Me.Button28.Image = CType(resources.GetObject("Button28.Image"), System.Drawing.Image)
        Me.Button28.Label = "条3"
        Me.Button28.Name = "Button28"
        Me.Button28.ShowImage = True
        '
        'Button29
        '
        Me.Button29.Image = CType(resources.GetObject("Button29.Image"), System.Drawing.Image)
        Me.Button29.Label = "条4"
        Me.Button29.Name = "Button29"
        Me.Button29.ShowImage = True
        '
        'Button30
        '
        Me.Button30.Image = CType(resources.GetObject("Button30.Image"), System.Drawing.Image)
        Me.Button30.Label = "条5"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
        '
        'SplitButton10
        '
        SplitButton10.Checked = True
        SplitButton10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        SplitButton10.Image = CType(resources.GetObject("SplitButton10.Image"), System.Drawing.Image)
        SplitButton10.Items.Add(Me.Button36)
        SplitButton10.Items.Add(Me.Button37)
        SplitButton10.Items.Add(Me.Button38)
        SplitButton10.Items.Add(Me.Button39)
        SplitButton10.Label = "列项"
        SplitButton10.Name = "SplitButton10"
        AddHandler SplitButton10.Click, AddressOf Me.SplitButton1_Click
        '
        'Button36
        '
        Me.Button36.Image = CType(resources.GetObject("Button36.Image"), System.Drawing.Image)
        Me.Button36.Label = "列项1"
        Me.Button36.Name = "Button36"
        Me.Button36.ShowImage = True
        '
        'Button37
        '
        Me.Button37.Image = CType(resources.GetObject("Button37.Image"), System.Drawing.Image)
        Me.Button37.Label = "列项2"
        Me.Button37.Name = "Button37"
        Me.Button37.ShowImage = True
        '
        'Button38
        '
        Me.Button38.Image = CType(resources.GetObject("Button38.Image"), System.Drawing.Image)
        Me.Button38.Label = "字母项"
        Me.Button38.Name = "Button38"
        Me.Button38.ShowImage = True
        '
        'Button39
        '
        Me.Button39.Image = CType(resources.GetObject("Button39.Image"), System.Drawing.Image)
        Me.Button39.Label = "数字项"
        Me.Button39.Name = "Button39"
        Me.Button39.ShowImage = True
        '
        'SplitButton7
        '
        SplitButton7.Checked = True
        SplitButton7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        SplitButton7.Image = CType(resources.GetObject("SplitButton7.Image"), System.Drawing.Image)
        SplitButton7.Items.Add(Me.Button24)
        SplitButton7.Items.Add(Me.Button25)
        SplitButton7.Items.Add(Me.Button26)
        SplitButton7.Items.Add(Me.Button27)
        SplitButton7.Label = "表格"
        SplitButton7.Name = "SplitButton7"
        AddHandler SplitButton7.Click, AddressOf Me.SplitButton1_Click
        '
        'Button24
        '
        Me.Button24.Image = CType(resources.GetObject("Button24.Image"), System.Drawing.Image)
        Me.Button24.Label = "插入表格"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        '
        'Button25
        '
        Me.Button25.Image = CType(resources.GetObject("Button25.Image"), System.Drawing.Image)
        Me.Button25.Label = "标准样式"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Button26
        '
        Me.Button26.Image = CType(resources.GetObject("Button26.Image"), System.Drawing.Image)
        Me.Button26.Label = "跨页拆分"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        '
        'Button27
        '
        Me.Button27.Image = CType(resources.GetObject("Button27.Image"), System.Drawing.Image)
        Me.Button27.Label = "同页合并"
        Me.Button27.Name = "Button27"
        Me.Button27.ShowImage = True
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.Group2)
        Me.Tab2.Groups.Add(Me.Group8)
        Me.Tab2.Groups.Add(Me.Group1)
        Me.Tab2.Groups.Add(Me.Group3)
        Me.Tab2.Groups.Add(Me.Group4)
        Me.Tab2.Groups.Add(Me.Group5)
        Me.Tab2.Groups.Add(Me.Group6)
        Me.Tab2.Groups.Add(Me.Group7)
        Me.Tab2.Label = "标准文档编辑"
        Me.Tab2.Name = "Tab2"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.Separator1)
        Me.Group2.Items.Add(Me.Button1)
        Me.Group2.Label = "联机"
        Me.Group2.Name = "Group2"
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "注册"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "标准更新"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Group8
        '
        Me.Group8.Items.Add(SplitButton1)
        Me.Group8.Label = "新建"
        Me.Group8.Name = "Group8"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.SplitButton12)
        Me.Group1.Items.Add(Me.SplitButton14)
        Me.Group1.Items.Add(Me.SplitButton6)
        Me.Group1.Items.Add(Me.SplitButton4)
        Me.Group1.Items.Add(Me.SplitButton13)
        Me.Group1.Label = "结构要素"
        Me.Group1.Name = "Group1"
        '
        'SplitButton12
        '
        Me.SplitButton12.Image = CType(resources.GetObject("SplitButton12.Image"), System.Drawing.Image)
        Me.SplitButton12.Items.Add(Me.Button46)
        Me.SplitButton12.Label = "扉页"
        Me.SplitButton12.Name = "SplitButton12"
        '
        'Button46
        '
        Me.Button46.Label = "删除目次"
        Me.Button46.Name = "Button46"
        Me.Button46.ShowImage = True
        '
        'SplitButton14
        '
        Me.SplitButton14.Image = CType(resources.GetObject("SplitButton14.Image"), System.Drawing.Image)
        Me.SplitButton14.Items.Add(Me.Button47)
        Me.SplitButton14.Label = "前言"
        Me.SplitButton14.Name = "SplitButton14"
        '
        'Button47
        '
        Me.Button47.Label = "删除目次"
        Me.Button47.Name = "Button47"
        Me.Button47.ShowImage = True
        '
        'SplitButton6
        '
        Me.SplitButton6.Image = CType(resources.GetObject("SplitButton6.Image"), System.Drawing.Image)
        Me.SplitButton6.Items.Add(Me.Button43)
        Me.SplitButton6.Label = "目录"
        Me.SplitButton6.Name = "SplitButton6"
        '
        'Button43
        '
        Me.Button43.Label = "删除目次"
        Me.Button43.Name = "Button43"
        Me.Button43.ShowImage = True
        '
        'SplitButton4
        '
        Me.SplitButton4.Image = CType(resources.GetObject("SplitButton4.Image"), System.Drawing.Image)
        Me.SplitButton4.Items.Add(Me.Button13)
        Me.SplitButton4.Label = "引言"
        Me.SplitButton4.Name = "SplitButton4"
        '
        'Button13
        '
        Me.Button13.Label = "删除引言"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        '
        'SplitButton13
        '
        Me.SplitButton13.Image = CType(resources.GetObject("SplitButton13.Image"), System.Drawing.Image)
        Me.SplitButton13.Items.Add(Me.Button45)
        Me.SplitButton13.Items.Add(Me.Button3)
        Me.SplitButton13.Items.Add(Me.Button16)
        Me.SplitButton13.Items.Add(Me.Button19)
        Me.SplitButton13.Label = "附录"
        Me.SplitButton13.Name = "SplitButton13"
        '
        'Button45
        '
        Me.Button45.Label = "附录"
        Me.Button45.Name = "Button45"
        Me.Button45.ShowImage = True
        '
        'Button3
        '
        Me.Button3.Label = "附录1"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Button16
        '
        Me.Button16.Label = "附录2"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        '
        'Button19
        '
        Me.Button19.Label = "附录3"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button22)
        Me.Group3.Items.Add(Me.Button23)
        Me.Group3.Items.Add(Me.Separator2)
        Me.Group3.Items.Add(SplitButton9)
        Me.Group3.Items.Add(SplitButton3)
        Me.Group3.Items.Add(SplitButton10)
        Me.Group3.Label = "层次样式"
        Me.Group3.Name = "Group3"
        '
        'Button22
        '
        Me.Button22.Image = CType(resources.GetObject("Button22.Image"), System.Drawing.Image)
        Me.Button22.Label = "章"
        Me.Button22.Name = "Button22"
        Me.Button22.ShowImage = True
        '
        'Button23
        '
        Me.Button23.Image = CType(resources.GetObject("Button23.Image"), System.Drawing.Image)
        Me.Button23.Label = "段"
        Me.Button23.Name = "Button23"
        Me.Button23.ShowImage = True
        Me.Button23.SuperTip = "段"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button5)
        Me.Group4.Items.Add(Me.Button4)
        Me.Group4.Items.Add(Me.SplitButton8)
        Me.Group4.Items.Add(Me.Button12)
        Me.Group4.Items.Add(Me.Button11)
        Me.Group4.Label = "标注样式"
        Me.Group4.Name = "Group4"
        '
        'Button5
        '
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Label = "注"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button4
        '
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "注X"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'SplitButton8
        '
        Me.SplitButton8.Image = CType(resources.GetObject("SplitButton8.Image"), System.Drawing.Image)
        Me.SplitButton8.Label = "脚注"
        Me.SplitButton8.Name = "SplitButton8"
        '
        'Button12
        '
        Me.Button12.Image = CType(resources.GetObject("Button12.Image"), System.Drawing.Image)
        Me.Button12.Label = "示例"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        '
        'Button11
        '
        Me.Button11.Image = CType(resources.GetObject("Button11.Image"), System.Drawing.Image)
        Me.Button11.Label = "示例X"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        '
        'Group5
        '
        Me.Group5.Items.Add(SplitButton7)
        Me.Group5.Items.Add(Me.Button15)
        Me.Group5.Items.Add(Me.Button14)
        Me.Group5.Items.Add(Me.Separator3)
        Me.Group5.Items.Add(Me.Button17)
        Me.Group5.Items.Add(Me.Button18)
        Me.Group5.Label = "图表"
        Me.Group5.Name = "Group5"
        '
        'Button15
        '
        Me.Button15.Image = CType(resources.GetObject("Button15.Image"), System.Drawing.Image)
        Me.Button15.Label = "表题"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        '
        'Button14
        '
        Me.Button14.Image = CType(resources.GetObject("Button14.Image"), System.Drawing.Image)
        Me.Button14.Label = "表注"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'Button17
        '
        Me.Button17.Image = CType(resources.GetObject("Button17.Image"), System.Drawing.Image)
        Me.Button17.Label = "图题"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button18
        '
        Me.Button18.Image = CType(resources.GetObject("Button18.Image"), System.Drawing.Image)
        Me.Button18.Label = "图注"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.SplitButton21)
        Me.Group6.Items.Add(Me.SplitButton22)
        Me.Group6.Label = "公式|单位|符号"
        Me.Group6.Name = "Group6"
        '
        'SplitButton21
        '
        Me.SplitButton21.Image = CType(resources.GetObject("SplitButton21.Image"), System.Drawing.Image)
        Me.SplitButton21.Items.Add(Me.Button70)
        Me.SplitButton21.Items.Add(Me.Button71)
        Me.SplitButton21.Items.Add(Me.Button72)
        Me.SplitButton21.Label = "公式"
        Me.SplitButton21.Name = "SplitButton21"
        '
        'Button70
        '
        Me.Button70.Label = "Button6"
        Me.Button70.Name = "Button70"
        Me.Button70.ShowImage = True
        '
        'Button71
        '
        Me.Button71.Label = "Button11"
        Me.Button71.Name = "Button71"
        Me.Button71.ShowImage = True
        '
        'Button72
        '
        Me.Button72.Label = "Button12"
        Me.Button72.Name = "Button72"
        Me.Button72.ShowImage = True
        '
        'SplitButton22
        '
        Me.SplitButton22.Image = CType(resources.GetObject("SplitButton22.Image"), System.Drawing.Image)
        Me.SplitButton22.Items.Add(Me.Button73)
        Me.SplitButton22.Items.Add(Me.Button74)
        Me.SplitButton22.Items.Add(Me.Button75)
        Me.SplitButton22.Label = "单位和符号"
        Me.SplitButton22.Name = "SplitButton22"
        '
        'Button73
        '
        Me.Button73.Label = "Button6"
        Me.Button73.Name = "Button73"
        Me.Button73.ShowImage = True
        '
        'Button74
        '
        Me.Button74.Label = "Button11"
        Me.Button74.Name = "Button74"
        Me.Button74.ShowImage = True
        '
        'Button75
        '
        Me.Button75.Label = "Button12"
        Me.Button75.Name = "Button75"
        Me.Button75.ShowImage = True
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.itemDescription)
        Me.Group7.Items.Add(Me.connectLine)
        Me.Group7.Items.Add(Me.quote)
        Me.Group7.Items.Add(Me.Button6)
        Me.Group7.Items.Add(Me.Button42)
        Me.Group7.Items.Add(Me.Button31)
        Me.Group7.Items.Add(Me.Button40)
        Me.Group7.Items.Add(Me.Button41)
        Me.Group7.Label = "其他"
        Me.Group7.Name = "Group7"
        '
        'itemDescription
        '
        Me.itemDescription.Label = "列项说明"
        Me.itemDescription.Name = "itemDescription"
        '
        'connectLine
        '
        Me.connectLine.Label = "连接线"
        Me.connectLine.Name = "connectLine"
        '
        'quote
        '
        Me.quote.Label = "引用"
        Me.quote.Name = "quote"
        '
        'Button6
        '
        Me.Button6.Label = "单面排"
        Me.Button6.Name = "Button6"
        '
        'Button42
        '
        Me.Button42.Label = "双面排"
        Me.Button42.Name = "Button42"
        '
        'Button31
        '
        Me.Button31.Label = "横页"
        Me.Button31.Name = "Button31"
        '
        'Button40
        '
        Me.Button40.Label = "终结线"
        Me.Button40.Name = "Button40"
        '
        'Button41
        '
        Me.Button41.Label = "检定证书和检定结果"
        Me.Button41.Name = "Button41"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group8.ResumeLayout(False)
        Me.Group8.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group8 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton4 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton13 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button45 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Button32 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button33 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button34 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button35 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button36 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button37 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button38 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button39 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button27 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton21 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button70 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button71 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button72 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton22 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button73 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button74 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button75 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button76 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton8 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents connectLine As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents itemDescription As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents quote As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button28 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button31 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button40 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button41 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents SplitButton12 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button46 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton14 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button47 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton6 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button43 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button42 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
