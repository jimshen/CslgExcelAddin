Partial Class LabItemsRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButtonSearch = Me.Factory.CreateRibbonButton
        Me.ButtonCalc = Me.Factory.CreateRibbonButton
        Me.ButtonSplitSum = Me.Factory.CreateRibbonButton
        Me.ButtonStatistics = Me.Factory.CreateRibbonButton
        Me.ButtonDomain = Me.Factory.CreateRibbonButton
        Me.ButtonPaste = Me.Factory.CreateRibbonButton
        Me.ButtonPinyin = Me.Factory.CreateRibbonButton
        Me.ButtonCalcu = Me.Factory.CreateRibbonButton
        Me.Group1.SuspendLayout()
        Me.Tab1.SuspendLayout()
        Me.Group2.SuspendLayout()
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonSearch)
        Me.Group1.Items.Add(Me.ButtonCalc)
        Me.Group1.Items.Add(Me.ButtonSplitSum)
        Me.Group1.Items.Add(Me.ButtonStatistics)
        Me.Group1.Items.Add(Me.ButtonDomain)
        Me.Group1.Items.Add(Me.ButtonPaste)
        Me.Group1.Label = "实验项目工具"
        Me.Group1.Name = "Group1"
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "业务工具"
        Me.Tab1.Name = "Tab1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButtonPinyin)
        Me.Group2.Label = "其他工具"
        Me.Group2.Name = "Group2"
        '
        'ButtonSearch
        '
        Me.ButtonSearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonSearch.Image = Global.CslgExcelAddIn.My.Resources.Resources.search
        Me.ButtonSearch.Label = "查找项目"
        Me.ButtonSearch.Name = "ButtonSearch"
        Me.ButtonSearch.ShowImage = True
        Me.ButtonSearch.SuperTip = "根据选中单元格中文字搜索实验名称库，可选择对应项目编号插入指定列"
        '
        'ButtonCalc
        '
        Me.ButtonCalc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonCalc.Image = Global.CslgExcelAddIn.My.Resources.Resources.sum
        Me.ButtonCalc.Label = "合并求和"
        Me.ButtonCalc.Name = "ButtonCalc"
        Me.ButtonCalc.ShowImage = True
        Me.ButtonCalc.SuperTip = "对选定单元格（单列）中数值进行求和后放在第一个单元格，其他单元格清空"
        '
        'ButtonSplitSum
        '
        Me.ButtonSplitSum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonSplitSum.Image = Global.CslgExcelAddIn.My.Resources.Resources.split_sum
        Me.ButtonSplitSum.Label = "分列求和"
        Me.ButtonSplitSum.Name = "ButtonSplitSum"
        Me.ButtonSplitSum.ShowImage = True
        '
        'ButtonStatistics
        '
        Me.ButtonStatistics.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonStatistics.Image = Global.CslgExcelAddIn.My.Resources.Resources.statistics
        Me.ButtonStatistics.Label = "项目统计"
        Me.ButtonStatistics.Name = "ButtonStatistics"
        Me.ButtonStatistics.ShowImage = True
        Me.ButtonStatistics.SuperTip = "统计实验项目总个数和各类型实验项目的数量"
        '
        'ButtonDomain
        '
        Me.ButtonDomain.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonDomain.Image = Global.CslgExcelAddIn.My.Resources.Resources.bmzy
        Me.ButtonDomain.Label = "班名查专业"
        Me.ButtonDomain.Name = "ButtonDomain"
        Me.ButtonDomain.ShowImage = True
        '
        'ButtonPaste
        '
        Me.ButtonPaste.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonPaste.Image = Global.CslgExcelAddIn.My.Resources.Resources.paste
        Me.ButtonPaste.Label = "多行粘贴"
        Me.ButtonPaste.Name = "ButtonPaste"
        Me.ButtonPaste.ShowImage = True
        Me.ButtonPaste.SuperTip = "根据剪贴板内容行数在目标单元格下方插入适当数量空行，并将内容粘贴到目标单元格"
        '
        'ButtonPinyin
        '
        Me.ButtonPinyin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonPinyin.Image = Global.CslgExcelAddIn.My.Resources.Resources.Hz
        Me.ButtonPinyin.Label = "汉字转拼音"
        Me.ButtonPinyin.Name = "ButtonPinyin"
        Me.ButtonPinyin.ScreenTip = "转换汉字到拼音，生僻词和多音字请自行检查"
        Me.ButtonPinyin.ShowImage = True
        '
        'ButtonCalcu
        '
        Me.ButtonCalcu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonCalcu.Image = Global.CslgExcelAddIn.My.Resources.Resources.sum
        Me.ButtonCalcu.Label = "合并求和"
        Me.ButtonCalcu.Name = "ButtonCalcu"
        Me.ButtonCalcu.ShowImage = True
        '
        'LabItemsRibbon
        '
        Me.Name = "LabItemsRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()

    End Sub

    Friend WithEvents ButtonCalcu As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCalc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents ButtonPaste As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonStatistics As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSplitSum As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonDomain As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonPinyin As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property LabItemsRibbon() As LabItemsRibbon
        Get
            Return Me.GetRibbon(Of LabItemsRibbon)()
        End Get
    End Property
End Class
