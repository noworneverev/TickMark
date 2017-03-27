Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group10 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.Group9 = Me.Factory.CreateRibbonGroup
        Me.Group8 = Me.Factory.CreateRibbonGroup
        Me.Group17 = Me.Factory.CreateRibbonGroup
        Me.Roman_numerals = Me.Factory.CreateRibbonGroup
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.Group15 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Group18 = Me.Factory.CreateRibbonGroup
        Me.Group14 = Me.Factory.CreateRibbonGroup
        Me.Group12 = Me.Factory.CreateRibbonGroup
        Me.Group16 = Me.Factory.CreateRibbonGroup
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Group11 = Me.Factory.CreateRibbonGroup
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group13 = Me.Factory.CreateRibbonGroup
        Me.PBC = Me.Factory.CreateRibbonButton
        Me.FontName1 = Me.Factory.CreateRibbonButton
        Me.FontName2 = Me.Factory.CreateRibbonButton
        Me.PMingLiU = Me.Factory.CreateRibbonButton
        Me.Arial = Me.Factory.CreateRibbonButton
        Me.BookAntiqua = Me.Factory.CreateRibbonButton
        Me.Calibri = Me.Factory.CreateRibbonButton
        Me.HyperLink = Me.Factory.CreateRibbonButton
        Me.hyperlinkCell = Me.Factory.CreateRibbonButton
        Me.InsertColumn = Me.Factory.CreateRibbonButton
        Me.CapP = Me.Factory.CreateRibbonButton
        Me.AJE = Me.Factory.CreateRibbonButton
        Me.RJE = Me.Factory.CreateRibbonButton
        Me.footed = Me.Factory.CreateRibbonButton
        Me.btnCB = Me.Factory.CreateRibbonButton
        Me.RD = Me.Factory.CreateRibbonButton
        Me.TB = Me.Factory.CreateRibbonButton
        Me.FS = Me.Factory.CreateRibbonButton
        Me.NotMaterial = Me.Factory.CreateRibbonButton
        Me.tick = Me.Factory.CreateRibbonButton
        Me.Tick2 = Me.Factory.CreateRibbonButton
        Me.Tick3 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Minor_Pass = Me.Factory.CreateRibbonButton
        Me.Conclusion = Me.Factory.CreateRibbonButton
        Me.star1 = Me.Factory.CreateRibbonButton
        Me.tri1 = Me.Factory.CreateRibbonButton
        Me.dia1 = Me.Factory.CreateRibbonButton
        Me.star2 = Me.Factory.CreateRibbonButton
        Me.tri2 = Me.Factory.CreateRibbonButton
        Me.Menu3 = Me.Factory.CreateRibbonMenu
        Me.o = Me.Factory.CreateRibbonButton
        Me.OO = Me.Factory.CreateRibbonButton
        Me.squ1 = Me.Factory.CreateRibbonButton
        Me.squ2 = Me.Factory.CreateRibbonButton
        Me.dia2 = Me.Factory.CreateRibbonButton
        Me.divided = Me.Factory.CreateRibbonButton
        Me.alpha = Me.Factory.CreateRibbonButton
        Me.beta = Me.Factory.CreateRibbonButton
        Me.gamma = Me.Factory.CreateRibbonButton
        Me.delta = Me.Factory.CreateRibbonButton
        Me.epsilon = Me.Factory.CreateRibbonButton
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.zeta = Me.Factory.CreateRibbonButton
        Me.eta = Me.Factory.CreateRibbonButton
        Me.theta = Me.Factory.CreateRibbonButton
        Me.pi = Me.Factory.CreateRibbonButton
        Me.mu = Me.Factory.CreateRibbonButton
        Me.rho = Me.Factory.CreateRibbonButton
        Me.sigma = Me.Factory.CreateRibbonButton
        Me.phi = Me.Factory.CreateRibbonButton
        Me.psi = Me.Factory.CreateRibbonButton
        Me.omega = Me.Factory.CreateRibbonButton
        Me.one = Me.Factory.CreateRibbonButton
        Me.two = Me.Factory.CreateRibbonButton
        Me.three = Me.Factory.CreateRibbonButton
        Me.four = Me.Factory.CreateRibbonButton
        Me.five = Me.Factory.CreateRibbonButton
        Me.Menu4 = Me.Factory.CreateRibbonMenu
        Me.six = Me.Factory.CreateRibbonButton
        Me.seven = Me.Factory.CreateRibbonButton
        Me.eight = Me.Factory.CreateRibbonButton
        Me.nine = Me.Factory.CreateRibbonButton
        Me.ten = Me.Factory.CreateRibbonButton
        Me.a = Me.Factory.CreateRibbonButton
        Me.b = Me.Factory.CreateRibbonButton
        Me.c = Me.Factory.CreateRibbonButton
        Me.siga = Me.Factory.CreateRibbonButton
        Me.sigb = Me.Factory.CreateRibbonButton
        Me.sigc = Me.Factory.CreateRibbonButton
        Me.DownApply = Me.Factory.CreateRibbonButton
        Me.RightApply = Me.Factory.CreateRibbonButton
        Me.ArrowBox = Me.Factory.CreateRibbonButton
        Me.MultiIllu = Me.Factory.CreateRibbonButton
        Me.ColorYellow = Me.Factory.CreateRibbonButton
        Me.ColorAqua = Me.Factory.CreateRibbonButton
        Me.ColorLime = Me.Factory.CreateRibbonButton
        Me.ColorSilver = Me.Factory.CreateRibbonButton
        Me.ColorFuchsia = Me.Factory.CreateRibbonButton
        Me.ColorWhite = Me.Factory.CreateRibbonButton
        Me.SheetArrow = Me.Factory.CreateRibbonButton
        Me.Direction1 = Me.Factory.CreateRibbonButton
        Me.Direction2 = Me.Factory.CreateRibbonButton
        Me.Direction3 = Me.Factory.CreateRibbonButton
        Me.Direction4 = Me.Factory.CreateRibbonButton
        Me.Direction5 = Me.Factory.CreateRibbonButton
        Me.Note = Me.Factory.CreateRibbonButton
        Me.Note1 = Me.Factory.CreateRibbonButton
        Me.Note2 = Me.Factory.CreateRibbonButton
        Me.Note3 = Me.Factory.CreateRibbonButton
        Me.Note4 = Me.Factory.CreateRibbonButton
        Me.Note5 = Me.Factory.CreateRibbonButton
        Me.LeftBrace = Me.Factory.CreateRibbonButton
        Me.Preparer2 = Me.Factory.CreateRibbonButton
        Me.Reviewer2 = Me.Factory.CreateRibbonButton
        Me.RedPen = Me.Factory.CreateRibbonButton
        Me.BluePen = Me.Factory.CreateRibbonButton
        Me.BlackPen = Me.Factory.CreateRibbonButton
        Me.Calendar_ToggleButton = Me.Factory.CreateRibbonToggleButton
        Me.CheckBox = Me.Factory.CreateRibbonButton
        Me.ClearCheckBox = Me.Factory.CreateRibbonButton
        Me.RedArrow = Me.Factory.CreateRibbonButton
        Me.DoubleLine = Me.Factory.CreateRibbonButton
        Me.CommaStyle = Me.Factory.CreateRibbonButton
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.ClearHyperLink = Me.Factory.CreateRibbonButton
        Me.ClearPen = Me.Factory.CreateRibbonButton
        Me.ClearAllShape = Me.Factory.CreateRibbonButton
        Me.Quote = Me.Factory.CreateRibbonButton
        Me.Copyright_ToggleButton = Me.Factory.CreateRibbonToggleButton
        Me.Menu5 = Me.Factory.CreateRibbonMenu
        Me.Tab1.SuspendLayout()
        Me.Group10.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group9.SuspendLayout()
        Me.Group8.SuspendLayout()
        Me.Group17.SuspendLayout()
        Me.Roman_numerals.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.Group15.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group18.SuspendLayout()
        Me.Group14.SuspendLayout()
        Me.Group12.SuspendLayout()
        Me.Group16.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group11.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group13.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group10)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group9)
        Me.Tab1.Groups.Add(Me.Group8)
        Me.Tab1.Groups.Add(Me.Group17)
        Me.Tab1.Groups.Add(Me.Roman_numerals)
        Me.Tab1.Groups.Add(Me.Group7)
        Me.Tab1.Groups.Add(Me.Group15)
        Me.Tab1.Groups.Add(Me.Group18)
        Me.Tab1.Groups.Add(Me.Group14)
        Me.Tab1.Groups.Add(Me.Group12)
        Me.Tab1.Groups.Add(Me.Group16)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group11)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group13)
        Me.Tab1.Label = "Tick Mark"
        Me.Tab1.Name = "Tab1"
        '
        'Group10
        '
        Me.Group10.Items.Add(Me.PBC)
        Me.Group10.Label = "PBC"
        Me.Group10.Name = "Group10"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.FontName1)
        Me.Group2.Items.Add(Me.FontName2)
        Me.Group2.Items.Add(Me.PMingLiU)
        Me.Group2.Items.Add(Me.Separator2)
        Me.Group2.Items.Add(Me.Arial)
        Me.Group2.Items.Add(Me.BookAntiqua)
        Me.Group2.Items.Add(Me.Calibri)
        Me.Group2.Label = "Font Name"
        Me.Group2.Name = "Group2"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'Group9
        '
        Me.Group9.Items.Add(Me.CapP)
        Me.Group9.Items.Add(Me.AJE)
        Me.Group9.Items.Add(Me.RJE)
        Me.Group9.Items.Add(Me.footed)
        Me.Group9.Items.Add(Me.btnCB)
        Me.Group9.Items.Add(Me.RD)
        Me.Group9.Items.Add(Me.TB)
        Me.Group9.Items.Add(Me.FS)
        Me.Group9.Items.Add(Me.NotMaterial)
        Me.Group9.Items.Add(Me.tick)
        Me.Group9.Items.Add(Me.Tick2)
        Me.Group9.Items.Add(Me.Tick3)
        Me.Group9.Items.Add(Me.Button1)
        Me.Group9.Items.Add(Me.Minor_Pass)
        Me.Group9.Items.Add(Me.Conclusion)
        Me.Group9.Label = "EWP"
        Me.Group9.Name = "Group9"
        '
        'Group8
        '
        Me.Group8.Items.Add(Me.star1)
        Me.Group8.Items.Add(Me.tri1)
        Me.Group8.Items.Add(Me.dia1)
        Me.Group8.Items.Add(Me.star2)
        Me.Group8.Items.Add(Me.tri2)
        Me.Group8.Items.Add(Me.Menu3)
        Me.Group8.Label = "Symbol"
        Me.Group8.Name = "Group8"
        '
        'Group17
        '
        Me.Group17.Items.Add(Me.alpha)
        Me.Group17.Items.Add(Me.beta)
        Me.Group17.Items.Add(Me.gamma)
        Me.Group17.Items.Add(Me.delta)
        Me.Group17.Items.Add(Me.epsilon)
        Me.Group17.Items.Add(Me.Menu2)
        Me.Group17.Label = "Greek Alphabet"
        Me.Group17.Name = "Group17"
        '
        'Roman_numerals
        '
        Me.Roman_numerals.Items.Add(Me.one)
        Me.Roman_numerals.Items.Add(Me.two)
        Me.Roman_numerals.Items.Add(Me.three)
        Me.Roman_numerals.Items.Add(Me.four)
        Me.Roman_numerals.Items.Add(Me.five)
        Me.Roman_numerals.Items.Add(Me.Menu4)
        Me.Roman_numerals.Label = "Roman Numerals"
        Me.Roman_numerals.Name = "Roman_numerals"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.a)
        Me.Group7.Items.Add(Me.b)
        Me.Group7.Items.Add(Me.c)
        Me.Group7.Items.Add(Me.siga)
        Me.Group7.Items.Add(Me.sigb)
        Me.Group7.Items.Add(Me.sigc)
        Me.Group7.Label = "Sum a b c"
        Me.Group7.Name = "Group7"
        '
        'Group15
        '
        Me.Group15.Items.Add(Me.DownApply)
        Me.Group15.Items.Add(Me.RightApply)
        Me.Group15.Items.Add(Me.ArrowBox)
        Me.Group15.Items.Add(Me.MultiIllu)
        Me.Group15.Label = "Arrow"
        Me.Group15.Name = "Group15"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Calendar_ToggleButton)
        Me.Group3.Items.Add(Me.CheckBox)
        Me.Group3.Items.Add(Me.ClearCheckBox)
        Me.Group3.Items.Add(Me.Separator1)
        Me.Group3.Items.Add(Me.RedArrow)
        Me.Group3.Items.Add(Me.DoubleLine)
        Me.Group3.Items.Add(Me.CommaStyle)
        Me.Group3.Label = "Others"
        Me.Group3.Name = "Group3"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Group18
        '
        Me.Group18.Items.Add(Me.ColorYellow)
        Me.Group18.Items.Add(Me.ColorAqua)
        Me.Group18.Items.Add(Me.ColorLime)
        Me.Group18.Items.Add(Me.ColorSilver)
        Me.Group18.Items.Add(Me.ColorFuchsia)
        Me.Group18.Items.Add(Me.ColorWhite)
        Me.Group18.Label = "Fill Color"
        Me.Group18.Name = "Group18"
        '
        'Group14
        '
        Me.Group14.Items.Add(Me.SheetArrow)
        Me.Group14.Items.Add(Me.Direction1)
        Me.Group14.Items.Add(Me.Direction2)
        Me.Group14.Items.Add(Me.Direction3)
        Me.Group14.Items.Add(Me.Direction4)
        Me.Group14.Items.Add(Me.Direction5)
        Me.Group14.Label = "Direction"
        Me.Group14.Name = "Group14"
        '
        'Group12
        '
        Me.Group12.Items.Add(Me.Note)
        Me.Group12.Items.Add(Me.Note1)
        Me.Group12.Items.Add(Me.Note2)
        Me.Group12.Items.Add(Me.Note3)
        Me.Group12.Items.Add(Me.Note4)
        Me.Group12.Items.Add(Me.Note5)
        Me.Group12.Label = "Note"
        Me.Group12.Name = "Group12"
        '
        'Group16
        '
        Me.Group16.Items.Add(Me.LeftBrace)
        Me.Group16.Items.Add(Me.InsertColumn)
        Me.Group16.Label = "Insert"
        Me.Group16.Name = "Group16"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.HyperLink)
        Me.Group6.Items.Add(Me.hyperlinkCell)
        Me.Group6.Label = "HyperLink"
        Me.Group6.Name = "Group6"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Menu5)
        Me.Group4.Label = "Signature"
        Me.Group4.Name = "Group4"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.RedPen)
        Me.Group5.Items.Add(Me.BluePen)
        Me.Group5.Items.Add(Me.BlackPen)
        Me.Group5.Label = "Marker"
        Me.Group5.Name = "Group5"
        '
        'Group11
        '
        Me.Group11.Items.Add(Me.Menu1)
        Me.Group11.Label = "Be Careful"
        Me.Group11.Name = "Group11"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Quote)
        Me.Group1.Label = "Get Motivated"
        Me.Group1.Name = "Group1"
        '
        'Group13
        '
        Me.Group13.Items.Add(Me.Copyright_ToggleButton)
        Me.Group13.Label = "License"
        Me.Group13.Name = "Group13"
        '
        'PBC
        '
        Me.PBC.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.PBC.Image = CType(resources.GetObject("PBC.Image"), System.Drawing.Image)
        Me.PBC.Label = "PBC"
        Me.PBC.Name = "PBC"
        Me.PBC.OfficeImageId = "P"
        Me.PBC.ScreenTip = "Prepared by client"
        Me.PBC.ShowImage = True
        '
        'FontName1
        '
        Me.FontName1.Image = CType(resources.GetObject("FontName1.Image"), System.Drawing.Image)
        Me.FontName1.Label = "標楷體"
        Me.FontName1.Name = "FontName1"
        Me.FontName1.ScreenTip = "標楷體"
        Me.FontName1.ShowImage = True
        Me.FontName1.ShowLabel = False
        Me.FontName1.Tag = ""
        '
        'FontName2
        '
        Me.FontName2.Image = CType(resources.GetObject("FontName2.Image"), System.Drawing.Image)
        Me.FontName2.Label = "正黑體"
        Me.FontName2.Name = "FontName2"
        Me.FontName2.ScreenTip = "微軟正黑體"
        Me.FontName2.ShowImage = True
        Me.FontName2.ShowLabel = False
        '
        'PMingLiU
        '
        Me.PMingLiU.Image = CType(resources.GetObject("PMingLiU.Image"), System.Drawing.Image)
        Me.PMingLiU.Label = "新細明"
        Me.PMingLiU.Name = "PMingLiU"
        Me.PMingLiU.ScreenTip = "新細明體"
        Me.PMingLiU.ShowImage = True
        Me.PMingLiU.ShowLabel = False
        '
        'Arial
        '
        Me.Arial.Image = CType(resources.GetObject("Arial.Image"), System.Drawing.Image)
        Me.Arial.Label = "Arial"
        Me.Arial.Name = "Arial"
        Me.Arial.ScreenTip = "Arial"
        Me.Arial.ShowImage = True
        Me.Arial.ShowLabel = False
        '
        'BookAntiqua
        '
        Me.BookAntiqua.Image = CType(resources.GetObject("BookAntiqua.Image"), System.Drawing.Image)
        Me.BookAntiqua.Label = "Book Antiqua"
        Me.BookAntiqua.Name = "BookAntiqua"
        Me.BookAntiqua.ScreenTip = "Book Antiqua"
        Me.BookAntiqua.ShowImage = True
        Me.BookAntiqua.ShowLabel = False
        '
        'Calibri
        '
        Me.Calibri.Image = CType(resources.GetObject("Calibri.Image"), System.Drawing.Image)
        Me.Calibri.Label = "Calibri"
        Me.Calibri.Name = "Calibri"
        Me.Calibri.ScreenTip = "Calibri"
        Me.Calibri.ShowImage = True
        Me.Calibri.ShowLabel = False
        '
        'HyperLink
        '
        Me.HyperLink.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.HyperLink.Label = "Cell Link"
        Me.HyperLink.Name = "HyperLink"
        Me.HyperLink.OfficeImageId = "AccessRelinkLists"
        Me.HyperLink.ScreenTip = "Cell Link"
        Me.HyperLink.ShowImage = True
        Me.HyperLink.SuperTip = "Create boxes linked with hyperlinks to the lower right of both of the selected ce" &
    "lls."
        '
        'hyperlinkCell
        '
        Me.hyperlinkCell.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.hyperlinkCell.Label = "Hyperlink"
        Me.hyperlinkCell.Name = "hyperlinkCell"
        Me.hyperlinkCell.OfficeImageId = "BuildHyperlink"
        Me.hyperlinkCell.ScreenTip = "Build Hyperlink"
        Me.hyperlinkCell.ShowImage = True
        Me.hyperlinkCell.SuperTip = "Build hyperlinks on both of the selected cells ."
        '
        'InsertColumn
        '
        Me.InsertColumn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.InsertColumn.Label = "Insert Column"
        Me.InsertColumn.Name = "InsertColumn"
        Me.InsertColumn.OfficeImageId = "AddNewColumnMenu"
        Me.InsertColumn.ScreenTip = "Insert a narrow column "
        Me.InsertColumn.ShowImage = True
        '
        'CapP
        '
        Me.CapP.Image = CType(resources.GetObject("CapP.Image"), System.Drawing.Image)
        Me.CapP.Label = "P"
        Me.CapP.Name = "CapP"
        Me.CapP.ScreenTip = "Agree to prior"
        Me.CapP.ShowImage = True
        Me.CapP.ShowLabel = False
        '
        'AJE
        '
        Me.AJE.Image = CType(resources.GetObject("AJE.Image"), System.Drawing.Image)
        Me.AJE.Label = "AJE"
        Me.AJE.Name = "AJE"
        Me.AJE.ScreenTip = "AJE"
        Me.AJE.ShowImage = True
        Me.AJE.ShowLabel = False
        '
        'RJE
        '
        Me.RJE.Image = CType(resources.GetObject("RJE.Image"), System.Drawing.Image)
        Me.RJE.Label = "RJE"
        Me.RJE.Name = "RJE"
        Me.RJE.ScreenTip = "RJE"
        Me.RJE.ShowImage = True
        Me.RJE.ShowLabel = False
        '
        'footed
        '
        Me.footed.Image = CType(resources.GetObject("footed.Image"), System.Drawing.Image)
        Me.footed.Label = "f"
        Me.footed.Name = "footed"
        Me.footed.ScreenTip = "Footed"
        Me.footed.ShowImage = True
        Me.footed.ShowLabel = False
        '
        'btnCB
        '
        Me.btnCB.Image = CType(resources.GetObject("btnCB.Image"), System.Drawing.Image)
        Me.btnCB.Label = "CB"
        Me.btnCB.Name = "btnCB"
        Me.btnCB.ScreenTip = "CB"
        Me.btnCB.ShowImage = True
        Me.btnCB.ShowLabel = False
        '
        'RD
        '
        Me.RD.Image = CType(resources.GetObject("RD.Image"), System.Drawing.Image)
        Me.RD.Label = "RD"
        Me.RD.Name = "RD"
        Me.RD.ScreenTip = "RD"
        Me.RD.ShowImage = True
        Me.RD.ShowLabel = False
        '
        'TB
        '
        Me.TB.Image = CType(resources.GetObject("TB.Image"), System.Drawing.Image)
        Me.TB.Label = "TB"
        Me.TB.Name = "TB"
        Me.TB.ScreenTip = "Agrees to trial balance"
        Me.TB.ShowImage = True
        Me.TB.ShowLabel = False
        '
        'FS
        '
        Me.FS.Image = CType(resources.GetObject("FS.Image"), System.Drawing.Image)
        Me.FS.Label = "FS"
        Me.FS.Name = "FS"
        Me.FS.ScreenTip = "Forward to F/S"
        Me.FS.ShowImage = True
        Me.FS.ShowLabel = False
        '
        'NotMaterial
        '
        Me.NotMaterial.Image = CType(resources.GetObject("NotMaterial.Image"), System.Drawing.Image)
        Me.NotMaterial.Label = "Not Material"
        Me.NotMaterial.Name = "NotMaterial"
        Me.NotMaterial.ScreenTip = "Not Material"
        Me.NotMaterial.ShowImage = True
        Me.NotMaterial.ShowLabel = False
        '
        'tick
        '
        Me.tick.Image = CType(resources.GetObject("tick.Image"), System.Drawing.Image)
        Me.tick.Label = "✔"
        Me.tick.Name = "tick"
        Me.tick.ScreenTip = "Agrees 1st"
        Me.tick.ShowImage = True
        Me.tick.ShowLabel = False
        '
        'Tick2
        '
        Me.Tick2.Image = CType(resources.GetObject("Tick2.Image"), System.Drawing.Image)
        Me.Tick2.Label = "Tick2"
        Me.Tick2.Name = "Tick2"
        Me.Tick2.ScreenTip = "Agrees 2nd"
        Me.Tick2.ShowImage = True
        Me.Tick2.ShowLabel = False
        '
        'Tick3
        '
        Me.Tick3.Image = CType(resources.GetObject("Tick3.Image"), System.Drawing.Image)
        Me.Tick3.Label = "Tick3"
        Me.Tick3.Name = "Tick3"
        Me.Tick3.ScreenTip = "Agrees 3rd"
        Me.Tick3.ShowImage = True
        Me.Tick3.ShowLabel = False
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "✘"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        Me.Button1.ShowLabel = False
        '
        'Minor_Pass
        '
        Me.Minor_Pass.Image = CType(resources.GetObject("Minor_Pass.Image"), System.Drawing.Image)
        Me.Minor_Pass.Label = "Minor Pass"
        Me.Minor_Pass.Name = "Minor_Pass"
        Me.Minor_Pass.ScreenTip = "Minor Pass"
        Me.Minor_Pass.ShowImage = True
        Me.Minor_Pass.ShowLabel = False
        '
        'Conclusion
        '
        Me.Conclusion.Image = CType(resources.GetObject("Conclusion.Image"), System.Drawing.Image)
        Me.Conclusion.Label = "Conclusion"
        Me.Conclusion.Name = "Conclusion"
        Me.Conclusion.ScreenTip = "Conclusion"
        Me.Conclusion.ShowImage = True
        Me.Conclusion.ShowLabel = False
        '
        'star1
        '
        Me.star1.Image = CType(resources.GetObject("star1.Image"), System.Drawing.Image)
        Me.star1.Label = "☆"
        Me.star1.Name = "star1"
        Me.star1.ShowImage = True
        Me.star1.ShowLabel = False
        '
        'tri1
        '
        Me.tri1.Image = CType(resources.GetObject("tri1.Image"), System.Drawing.Image)
        Me.tri1.Label = "△"
        Me.tri1.Name = "tri1"
        Me.tri1.ShowImage = True
        Me.tri1.ShowLabel = False
        '
        'dia1
        '
        Me.dia1.Image = CType(resources.GetObject("dia1.Image"), System.Drawing.Image)
        Me.dia1.Label = "◇"
        Me.dia1.Name = "dia1"
        Me.dia1.ShowImage = True
        Me.dia1.ShowLabel = False
        '
        'star2
        '
        Me.star2.Image = CType(resources.GetObject("star2.Image"), System.Drawing.Image)
        Me.star2.Label = "★"
        Me.star2.Name = "star2"
        Me.star2.ShowImage = True
        Me.star2.ShowLabel = False
        '
        'tri2
        '
        Me.tri2.Image = CType(resources.GetObject("tri2.Image"), System.Drawing.Image)
        Me.tri2.Label = "▲"
        Me.tri2.Name = "tri2"
        Me.tri2.ShowImage = True
        Me.tri2.ShowLabel = False
        '
        'Menu3
        '
        Me.Menu3.Image = CType(resources.GetObject("Menu3.Image"), System.Drawing.Image)
        Me.Menu3.Items.Add(Me.o)
        Me.Menu3.Items.Add(Me.OO)
        Me.Menu3.Items.Add(Me.squ1)
        Me.Menu3.Items.Add(Me.squ2)
        Me.Menu3.Items.Add(Me.dia2)
        Me.Menu3.Items.Add(Me.divided)
        Me.Menu3.Label = "Menu3"
        Me.Menu3.Name = "Menu3"
        Me.Menu3.ShowImage = True
        Me.Menu3.ShowLabel = False
        '
        'o
        '
        Me.o.Image = CType(resources.GetObject("o.Image"), System.Drawing.Image)
        Me.o.Label = "○"
        Me.o.Name = "o"
        Me.o.ShowImage = True
        Me.o.ShowLabel = False
        '
        'OO
        '
        Me.OO.Image = CType(resources.GetObject("OO.Image"), System.Drawing.Image)
        Me.OO.Label = "◎"
        Me.OO.Name = "OO"
        Me.OO.ShowImage = True
        Me.OO.ShowLabel = False
        '
        'squ1
        '
        Me.squ1.Image = CType(resources.GetObject("squ1.Image"), System.Drawing.Image)
        Me.squ1.Label = "□"
        Me.squ1.Name = "squ1"
        Me.squ1.ShowImage = True
        Me.squ1.ShowLabel = False
        '
        'squ2
        '
        Me.squ2.Image = CType(resources.GetObject("squ2.Image"), System.Drawing.Image)
        Me.squ2.Label = "■"
        Me.squ2.Name = "squ2"
        Me.squ2.ShowImage = True
        Me.squ2.ShowLabel = False
        '
        'dia2
        '
        Me.dia2.Image = CType(resources.GetObject("dia2.Image"), System.Drawing.Image)
        Me.dia2.Label = "◆"
        Me.dia2.Name = "dia2"
        Me.dia2.ShowImage = True
        Me.dia2.ShowLabel = False
        '
        'divided
        '
        Me.divided.Image = CType(resources.GetObject("divided.Image"), System.Drawing.Image)
        Me.divided.Label = "/"
        Me.divided.Name = "divided"
        Me.divided.ShowImage = True
        Me.divided.ShowLabel = False
        '
        'alpha
        '
        Me.alpha.Image = CType(resources.GetObject("alpha.Image"), System.Drawing.Image)
        Me.alpha.Label = "α"
        Me.alpha.Name = "alpha"
        Me.alpha.ScreenTip = "Alpha"
        Me.alpha.ShowImage = True
        Me.alpha.ShowLabel = False
        '
        'beta
        '
        Me.beta.Image = CType(resources.GetObject("beta.Image"), System.Drawing.Image)
        Me.beta.Label = "β"
        Me.beta.Name = "beta"
        Me.beta.ScreenTip = "Beta"
        Me.beta.ShowImage = True
        Me.beta.ShowLabel = False
        '
        'gamma
        '
        Me.gamma.Image = CType(resources.GetObject("gamma.Image"), System.Drawing.Image)
        Me.gamma.Label = "γ"
        Me.gamma.Name = "gamma"
        Me.gamma.ScreenTip = "Gamma"
        Me.gamma.ShowImage = True
        Me.gamma.ShowLabel = False
        '
        'delta
        '
        Me.delta.Image = CType(resources.GetObject("delta.Image"), System.Drawing.Image)
        Me.delta.Label = "delta"
        Me.delta.Name = "delta"
        Me.delta.ScreenTip = "Delta"
        Me.delta.ShowImage = True
        Me.delta.ShowLabel = False
        '
        'epsilon
        '
        Me.epsilon.Image = CType(resources.GetObject("epsilon.Image"), System.Drawing.Image)
        Me.epsilon.Label = "epsilon"
        Me.epsilon.Name = "epsilon"
        Me.epsilon.ScreenTip = "Epsilon"
        Me.epsilon.ShowImage = True
        Me.epsilon.ShowLabel = False
        '
        'Menu2
        '
        Me.Menu2.Image = CType(resources.GetObject("Menu2.Image"), System.Drawing.Image)
        Me.Menu2.Items.Add(Me.zeta)
        Me.Menu2.Items.Add(Me.eta)
        Me.Menu2.Items.Add(Me.theta)
        Me.Menu2.Items.Add(Me.pi)
        Me.Menu2.Items.Add(Me.mu)
        Me.Menu2.Items.Add(Me.rho)
        Me.Menu2.Items.Add(Me.sigma)
        Me.Menu2.Items.Add(Me.phi)
        Me.Menu2.Items.Add(Me.psi)
        Me.Menu2.Items.Add(Me.omega)
        Me.Menu2.Label = "Others"
        Me.Menu2.Name = "Menu2"
        Me.Menu2.ScreenTip = "Others"
        Me.Menu2.ShowImage = True
        Me.Menu2.ShowLabel = False
        '
        'zeta
        '
        Me.zeta.Image = CType(resources.GetObject("zeta.Image"), System.Drawing.Image)
        Me.zeta.Label = "zeta"
        Me.zeta.Name = "zeta"
        Me.zeta.ShowImage = True
        Me.zeta.ShowLabel = False
        '
        'eta
        '
        Me.eta.Image = CType(resources.GetObject("eta.Image"), System.Drawing.Image)
        Me.eta.Label = "eta"
        Me.eta.Name = "eta"
        Me.eta.ShowImage = True
        Me.eta.ShowLabel = False
        '
        'theta
        '
        Me.theta.Image = CType(resources.GetObject("theta.Image"), System.Drawing.Image)
        Me.theta.Label = "theta"
        Me.theta.Name = "theta"
        Me.theta.ShowImage = True
        Me.theta.ShowLabel = False
        '
        'pi
        '
        Me.pi.Image = CType(resources.GetObject("pi.Image"), System.Drawing.Image)
        Me.pi.Label = "pi"
        Me.pi.Name = "pi"
        Me.pi.ShowImage = True
        Me.pi.ShowLabel = False
        '
        'mu
        '
        Me.mu.Image = CType(resources.GetObject("mu.Image"), System.Drawing.Image)
        Me.mu.Label = "mu"
        Me.mu.Name = "mu"
        Me.mu.ShowImage = True
        Me.mu.ShowLabel = False
        '
        'rho
        '
        Me.rho.Image = CType(resources.GetObject("rho.Image"), System.Drawing.Image)
        Me.rho.Label = "rho"
        Me.rho.Name = "rho"
        Me.rho.ShowImage = True
        Me.rho.ShowLabel = False
        '
        'sigma
        '
        Me.sigma.Image = CType(resources.GetObject("sigma.Image"), System.Drawing.Image)
        Me.sigma.Label = "∑"
        Me.sigma.Name = "sigma"
        Me.sigma.ShowImage = True
        Me.sigma.ShowLabel = False
        '
        'phi
        '
        Me.phi.Image = CType(resources.GetObject("phi.Image"), System.Drawing.Image)
        Me.phi.Label = "phi"
        Me.phi.Name = "phi"
        Me.phi.ShowImage = True
        Me.phi.ShowLabel = False
        '
        'psi
        '
        Me.psi.Image = CType(resources.GetObject("psi.Image"), System.Drawing.Image)
        Me.psi.Label = "psi"
        Me.psi.Name = "psi"
        Me.psi.ShowImage = True
        Me.psi.ShowLabel = False
        '
        'omega
        '
        Me.omega.Image = CType(resources.GetObject("omega.Image"), System.Drawing.Image)
        Me.omega.Label = "omega"
        Me.omega.Name = "omega"
        Me.omega.ShowImage = True
        Me.omega.ShowLabel = False
        '
        'one
        '
        Me.one.Image = CType(resources.GetObject("one.Image"), System.Drawing.Image)
        Me.one.Label = "Ⅰ"
        Me.one.Name = "one"
        Me.one.ScreenTip = "Roman 1"
        Me.one.ShowImage = True
        Me.one.ShowLabel = False
        '
        'two
        '
        Me.two.Image = CType(resources.GetObject("two.Image"), System.Drawing.Image)
        Me.two.Label = "Ⅱ"
        Me.two.Name = "two"
        Me.two.ScreenTip = "Roman 2"
        Me.two.ShowImage = True
        Me.two.ShowLabel = False
        '
        'three
        '
        Me.three.Image = CType(resources.GetObject("three.Image"), System.Drawing.Image)
        Me.three.Label = "Ⅲ"
        Me.three.Name = "three"
        Me.three.ScreenTip = "Roman 3"
        Me.three.ShowImage = True
        Me.three.ShowLabel = False
        '
        'four
        '
        Me.four.Image = CType(resources.GetObject("four.Image"), System.Drawing.Image)
        Me.four.Label = "Ⅳ"
        Me.four.Name = "four"
        Me.four.ScreenTip = "Roman 4"
        Me.four.ShowImage = True
        Me.four.ShowLabel = False
        '
        'five
        '
        Me.five.Image = CType(resources.GetObject("five.Image"), System.Drawing.Image)
        Me.five.Label = "Ⅴ"
        Me.five.Name = "five"
        Me.five.ScreenTip = "Roman 5"
        Me.five.ShowImage = True
        Me.five.ShowLabel = False
        '
        'Menu4
        '
        Me.Menu4.Image = CType(resources.GetObject("Menu4.Image"), System.Drawing.Image)
        Me.Menu4.Items.Add(Me.six)
        Me.Menu4.Items.Add(Me.seven)
        Me.Menu4.Items.Add(Me.eight)
        Me.Menu4.Items.Add(Me.nine)
        Me.Menu4.Items.Add(Me.ten)
        Me.Menu4.Label = "Menu4"
        Me.Menu4.Name = "Menu4"
        Me.Menu4.ScreenTip = "Roman numerals"
        Me.Menu4.ShowImage = True
        Me.Menu4.ShowLabel = False
        '
        'six
        '
        Me.six.Image = CType(resources.GetObject("six.Image"), System.Drawing.Image)
        Me.six.Label = "Ⅵ"
        Me.six.Name = "six"
        Me.six.ScreenTip = "Roman 6"
        Me.six.ShowImage = True
        Me.six.ShowLabel = False
        '
        'seven
        '
        Me.seven.Image = CType(resources.GetObject("seven.Image"), System.Drawing.Image)
        Me.seven.Label = "Ⅶ"
        Me.seven.Name = "seven"
        Me.seven.ScreenTip = "Roman 7"
        Me.seven.ShowImage = True
        Me.seven.ShowLabel = False
        '
        'eight
        '
        Me.eight.Image = CType(resources.GetObject("eight.Image"), System.Drawing.Image)
        Me.eight.Label = "Ⅷ"
        Me.eight.Name = "eight"
        Me.eight.ScreenTip = "Roman 8"
        Me.eight.ShowImage = True
        Me.eight.ShowLabel = False
        '
        'nine
        '
        Me.nine.Image = CType(resources.GetObject("nine.Image"), System.Drawing.Image)
        Me.nine.Label = "Ⅸ"
        Me.nine.Name = "nine"
        Me.nine.ScreenTip = "Roman 9"
        Me.nine.ShowImage = True
        Me.nine.ShowLabel = False
        '
        'ten
        '
        Me.ten.Image = CType(resources.GetObject("ten.Image"), System.Drawing.Image)
        Me.ten.Label = "Ⅹ"
        Me.ten.Name = "ten"
        Me.ten.ScreenTip = "Roman 10"
        Me.ten.ShowImage = True
        Me.ten.ShowLabel = False
        '
        'a
        '
        Me.a.Image = CType(resources.GetObject("a.Image"), System.Drawing.Image)
        Me.a.Label = "a"
        Me.a.Name = "a"
        Me.a.ScreenTip = "a"
        Me.a.ShowImage = True
        Me.a.ShowLabel = False
        '
        'b
        '
        Me.b.Image = CType(resources.GetObject("b.Image"), System.Drawing.Image)
        Me.b.Label = "b"
        Me.b.Name = "b"
        Me.b.ScreenTip = "b"
        Me.b.ShowImage = True
        Me.b.ShowLabel = False
        '
        'c
        '
        Me.c.Image = CType(resources.GetObject("c.Image"), System.Drawing.Image)
        Me.c.Label = "c"
        Me.c.Name = "c"
        Me.c.ScreenTip = "c"
        Me.c.ShowImage = True
        Me.c.ShowLabel = False
        '
        'siga
        '
        Me.siga.Image = CType(resources.GetObject("siga.Image"), System.Drawing.Image)
        Me.siga.Label = "∑a"
        Me.siga.Name = "siga"
        Me.siga.ScreenTip = "Sum a"
        Me.siga.ShowImage = True
        Me.siga.ShowLabel = False
        '
        'sigb
        '
        Me.sigb.Image = CType(resources.GetObject("sigb.Image"), System.Drawing.Image)
        Me.sigb.Label = "∑b"
        Me.sigb.Name = "sigb"
        Me.sigb.ScreenTip = "Sum "
        Me.sigb.ShowImage = True
        Me.sigb.ShowLabel = False
        '
        'sigc
        '
        Me.sigc.Image = CType(resources.GetObject("sigc.Image"), System.Drawing.Image)
        Me.sigc.Label = "∑c"
        Me.sigc.Name = "sigc"
        Me.sigc.ScreenTip = "Sum c"
        Me.sigc.ShowImage = True
        Me.sigc.ShowLabel = False
        '
        'DownApply
        '
        Me.DownApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DownApply.Image = CType(resources.GetObject("DownApply.Image"), System.Drawing.Image)
        Me.DownApply.Label = "Down Arrow"
        Me.DownApply.Name = "DownApply"
        Me.DownApply.ScreenTip = "Select a range to draw a down arrow"
        Me.DownApply.ShowImage = True
        '
        'RightApply
        '
        Me.RightApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.RightApply.Image = CType(resources.GetObject("RightApply.Image"), System.Drawing.Image)
        Me.RightApply.Label = "Right Arrow"
        Me.RightApply.Name = "RightApply"
        Me.RightApply.ScreenTip = "Select a range to draw a right arrow"
        Me.RightApply.ShowImage = True
        '
        'ArrowBox
        '
        Me.ArrowBox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ArrowBox.Image = CType(resources.GetObject("ArrowBox.Image"), System.Drawing.Image)
        Me.ArrowBox.Label = "Arrow Textbox"
        Me.ArrowBox.Name = "ArrowBox"
        Me.ArrowBox.ScreenTip = "Select a cell to add a connecting arrow and a textbox to the corner"
        Me.ArrowBox.ShowImage = True
        '
        'MultiIllu
        '
        Me.MultiIllu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MultiIllu.Image = CType(resources.GetObject("MultiIllu.Image"), System.Drawing.Image)
        Me.MultiIllu.Label = "Range Arrow"
        Me.MultiIllu.Name = "MultiIllu"
        Me.MultiIllu.ScreenTip = "Select a range to draw a connecting arrow"
        Me.MultiIllu.ShowImage = True
        '
        'ColorYellow
        '
        Me.ColorYellow.Image = CType(resources.GetObject("ColorYellow.Image"), System.Drawing.Image)
        Me.ColorYellow.Label = "ColorYellow"
        Me.ColorYellow.Name = "ColorYellow"
        Me.ColorYellow.OfficeImageId = "ColorYellow"
        Me.ColorYellow.ScreenTip = "Yellow"
        Me.ColorYellow.ShowImage = True
        Me.ColorYellow.ShowLabel = False
        '
        'ColorAqua
        '
        Me.ColorAqua.Image = CType(resources.GetObject("ColorAqua.Image"), System.Drawing.Image)
        Me.ColorAqua.Label = "ColorAqua"
        Me.ColorAqua.Name = "ColorAqua"
        Me.ColorAqua.OfficeImageId = "ColorAqua"
        Me.ColorAqua.ScreenTip = "Aqua"
        Me.ColorAqua.ShowImage = True
        Me.ColorAqua.ShowLabel = False
        '
        'ColorLime
        '
        Me.ColorLime.Image = CType(resources.GetObject("ColorLime.Image"), System.Drawing.Image)
        Me.ColorLime.Label = "ColorLime"
        Me.ColorLime.Name = "ColorLime"
        Me.ColorLime.ScreenTip = "Red"
        Me.ColorLime.ShowImage = True
        Me.ColorLime.ShowLabel = False
        '
        'ColorSilver
        '
        Me.ColorSilver.Image = CType(resources.GetObject("ColorSilver.Image"), System.Drawing.Image)
        Me.ColorSilver.Label = "ColorSilver"
        Me.ColorSilver.Name = "ColorSilver"
        Me.ColorSilver.ScreenTip = "#ffffc4"
        Me.ColorSilver.ShowImage = True
        Me.ColorSilver.ShowLabel = False
        '
        'ColorFuchsia
        '
        Me.ColorFuchsia.Image = CType(resources.GetObject("ColorFuchsia.Image"), System.Drawing.Image)
        Me.ColorFuchsia.Label = "ColorFuchsia"
        Me.ColorFuchsia.Name = "ColorFuchsia"
        Me.ColorFuchsia.ScreenTip = "#c4c4ff"
        Me.ColorFuchsia.ShowImage = True
        Me.ColorFuchsia.ShowLabel = False
        '
        'ColorWhite
        '
        Me.ColorWhite.Image = CType(resources.GetObject("ColorWhite.Image"), System.Drawing.Image)
        Me.ColorWhite.Label = "ColorWhite"
        Me.ColorWhite.Name = "ColorWhite"
        Me.ColorWhite.ScreenTip = "Remove color fill"
        Me.ColorWhite.ShowImage = True
        Me.ColorWhite.ShowLabel = False
        '
        'SheetArrow
        '
        Me.SheetArrow.Label = "SheetArrow"
        Me.SheetArrow.Name = "SheetArrow"
        Me.SheetArrow.OfficeImageId = "EastAsianEditingMarks"
        Me.SheetArrow.ScreenTip = "Create two arrows pointing to each other"
        Me.SheetArrow.ShowImage = True
        Me.SheetArrow.ShowLabel = False
        '
        'Direction1
        '
        Me.Direction1.Image = CType(resources.GetObject("Direction1.Image"), System.Drawing.Image)
        Me.Direction1.Label = "Direction1"
        Me.Direction1.Name = "Direction1"
        Me.Direction1.ScreenTip = "Create two boxes pointing to each other ⑴"
        Me.Direction1.ShowImage = True
        Me.Direction1.ShowLabel = False
        '
        'Direction2
        '
        Me.Direction2.Image = CType(resources.GetObject("Direction2.Image"), System.Drawing.Image)
        Me.Direction2.Label = "Direction2"
        Me.Direction2.Name = "Direction2"
        Me.Direction2.ScreenTip = "Create two boxes pointing to each other ⑵"
        Me.Direction2.ShowImage = True
        Me.Direction2.ShowLabel = False
        '
        'Direction3
        '
        Me.Direction3.Image = CType(resources.GetObject("Direction3.Image"), System.Drawing.Image)
        Me.Direction3.Label = "Direction3"
        Me.Direction3.Name = "Direction3"
        Me.Direction3.ScreenTip = "Create two boxes pointing to each other ⑶"
        Me.Direction3.ShowImage = True
        Me.Direction3.ShowLabel = False
        '
        'Direction4
        '
        Me.Direction4.Image = CType(resources.GetObject("Direction4.Image"), System.Drawing.Image)
        Me.Direction4.Label = "Direction4"
        Me.Direction4.Name = "Direction4"
        Me.Direction4.ScreenTip = "Create two boxes pointing to each other ⑷"
        Me.Direction4.ShowImage = True
        Me.Direction4.ShowLabel = False
        '
        'Direction5
        '
        Me.Direction5.Image = CType(resources.GetObject("Direction5.Image"), System.Drawing.Image)
        Me.Direction5.Label = "Direction5"
        Me.Direction5.Name = "Direction5"
        Me.Direction5.ScreenTip = "Create two boxes pointing to each other ⑸ "
        Me.Direction5.ShowImage = True
        Me.Direction5.ShowLabel = False
        '
        'Note
        '
        Me.Note.Image = CType(resources.GetObject("Note.Image"), System.Drawing.Image)
        Me.Note.Label = "<Note>"
        Me.Note.Name = "Note"
        Me.Note.ScreenTip = "<Note>"
        Me.Note.ShowImage = True
        Me.Note.ShowLabel = False
        '
        'Note1
        '
        Me.Note1.Image = CType(resources.GetObject("Note1.Image"), System.Drawing.Image)
        Me.Note1.Label = "<Note1>"
        Me.Note1.Name = "Note1"
        Me.Note1.ScreenTip = "<N1>"
        Me.Note1.ShowImage = True
        Me.Note1.ShowLabel = False
        '
        'Note2
        '
        Me.Note2.Image = CType(resources.GetObject("Note2.Image"), System.Drawing.Image)
        Me.Note2.Label = "<Note2>"
        Me.Note2.Name = "Note2"
        Me.Note2.ScreenTip = "<N2>"
        Me.Note2.ShowImage = True
        Me.Note2.ShowLabel = False
        '
        'Note3
        '
        Me.Note3.Image = CType(resources.GetObject("Note3.Image"), System.Drawing.Image)
        Me.Note3.Label = "<Note3>"
        Me.Note3.Name = "Note3"
        Me.Note3.ScreenTip = "<N3>"
        Me.Note3.ShowImage = True
        Me.Note3.ShowLabel = False
        '
        'Note4
        '
        Me.Note4.Image = CType(resources.GetObject("Note4.Image"), System.Drawing.Image)
        Me.Note4.Label = "<Note4>"
        Me.Note4.Name = "Note4"
        Me.Note4.ScreenTip = "<N4>"
        Me.Note4.ShowImage = True
        Me.Note4.ShowLabel = False
        '
        'Note5
        '
        Me.Note5.Image = CType(resources.GetObject("Note5.Image"), System.Drawing.Image)
        Me.Note5.Label = "<Note5>"
        Me.Note5.Name = "Note5"
        Me.Note5.ScreenTip = "<N5>"
        Me.Note5.ShowImage = True
        Me.Note5.ShowLabel = False
        '
        'LeftBrace
        '
        Me.LeftBrace.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.LeftBrace.Label = "Left Brace"
        Me.LeftBrace.Name = "LeftBrace"
        Me.LeftBrace.OfficeImageId = "ShapeLeftBrace"
        Me.LeftBrace.ScreenTip = "Entry brace"
        Me.LeftBrace.ShowImage = True
        '
        'Preparer2
        '
        Me.Preparer2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Preparer2.Label = "Preparer"
        Me.Preparer2.Name = "Preparer2"
        Me.Preparer2.OfficeImageId = "SignatureLineInsert"
        Me.Preparer2.ScreenTip = "Preparer"
        Me.Preparer2.ShowImage = True
        Me.Preparer2.SuperTip = "To set up your name, go to File > Options > General > Personalize your copy of Mi" &
    "crosoft office."
        '
        'Reviewer2
        '
        Me.Reviewer2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Reviewer2.Label = "Reviewer"
        Me.Reviewer2.Name = "Reviewer2"
        Me.Reviewer2.OfficeImageId = "SignatureLineInsert"
        Me.Reviewer2.ScreenTip = "Reviewer"
        Me.Reviewer2.ShowImage = True
        Me.Reviewer2.SuperTip = "To set up your name, go to File > Options > General > Personalize your copy of Mi" &
    "crosoft office."
        '
        'RedPen
        '
        Me.RedPen.Image = CType(resources.GetObject("RedPen.Image"), System.Drawing.Image)
        Me.RedPen.Label = "Red"
        Me.RedPen.Name = "RedPen"
        Me.RedPen.OfficeImageId = "PencilTool"
        Me.RedPen.ScreenTip = "Red Marker"
        Me.RedPen.ShowImage = True
        '
        'BluePen
        '
        Me.BluePen.Image = CType(resources.GetObject("BluePen.Image"), System.Drawing.Image)
        Me.BluePen.Label = "Blue"
        Me.BluePen.Name = "BluePen"
        Me.BluePen.OfficeImageId = "PencilTool"
        Me.BluePen.ScreenTip = "Blue Marker"
        Me.BluePen.ShowImage = True
        '
        'BlackPen
        '
        Me.BlackPen.Image = CType(resources.GetObject("BlackPen.Image"), System.Drawing.Image)
        Me.BlackPen.Label = "Black"
        Me.BlackPen.Name = "BlackPen"
        Me.BlackPen.OfficeImageId = "PencilTool"
        Me.BlackPen.ScreenTip = "Black Marker"
        Me.BlackPen.ShowImage = True
        '
        'Calendar_ToggleButton
        '
        Me.Calendar_ToggleButton.Image = CType(resources.GetObject("Calendar_ToggleButton.Image"), System.Drawing.Image)
        Me.Calendar_ToggleButton.Label = "Calendar"
        Me.Calendar_ToggleButton.Name = "Calendar_ToggleButton"
        Me.Calendar_ToggleButton.ScreenTip = "Calendar"
        Me.Calendar_ToggleButton.ShowImage = True
        Me.Calendar_ToggleButton.ShowLabel = False
        '
        'CheckBox
        '
        Me.CheckBox.Label = "CheckBox"
        Me.CheckBox.Name = "CheckBox"
        Me.CheckBox.OfficeImageId = "ActiveXCheckBox"
        Me.CheckBox.ScreenTip = "Add a checkbox"
        Me.CheckBox.ShowImage = True
        Me.CheckBox.ShowLabel = False
        '
        'ClearCheckBox
        '
        Me.ClearCheckBox.Image = CType(resources.GetObject("ClearCheckBox.Image"), System.Drawing.Image)
        Me.ClearCheckBox.Label = "清除方塊"
        Me.ClearCheckBox.Name = "ClearCheckBox"
        Me.ClearCheckBox.OfficeImageId = "RemoveAttach"
        Me.ClearCheckBox.ScreenTip = "Remove a checkbox"
        Me.ClearCheckBox.ShowImage = True
        Me.ClearCheckBox.ShowLabel = False
        '
        'RedArrow
        '
        Me.RedArrow.Label = "紅箭頭"
        Me.RedArrow.Name = "RedArrow"
        Me.RedArrow.OfficeImageId = "DrawArrow"
        Me.RedArrow.ScreenTip = "Draw an arrow to a cell you select"
        Me.RedArrow.ShowImage = True
        Me.RedArrow.ShowLabel = False
        '
        'DoubleLine
        '
        Me.DoubleLine.Label = "雙底線"
        Me.DoubleLine.Name = "DoubleLine"
        Me.DoubleLine.OfficeImageId = "BorderDoubleBottom"
        Me.DoubleLine.ScreenTip = "Bottom double border"
        Me.DoubleLine.ShowImage = True
        Me.DoubleLine.ShowLabel = False
        '
        'CommaStyle
        '
        Me.CommaStyle.Label = "Comma"
        Me.CommaStyle.Name = "CommaStyle"
        Me.CommaStyle.OfficeImageId = "CommaStyle"
        Me.CommaStyle.ScreenTip = "Comma Style"
        Me.CommaStyle.ShowImage = True
        Me.CommaStyle.ShowLabel = False
        Me.CommaStyle.SuperTip = "Format with a thousands separator."
        '
        'Menu1
        '
        Me.Menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Image = CType(resources.GetObject("Menu1.Image"), System.Drawing.Image)
        Me.Menu1.Items.Add(Me.ClearHyperLink)
        Me.Menu1.Items.Add(Me.ClearPen)
        Me.Menu1.Items.Add(Me.ClearAllShape)
        Me.Menu1.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu1.Label = "Remove"
        Me.Menu1.Name = "Menu1"
        Me.Menu1.ShowImage = True
        '
        'ClearHyperLink
        '
        Me.ClearHyperLink.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ClearHyperLink.Image = CType(resources.GetObject("ClearHyperLink.Image"), System.Drawing.Image)
        Me.ClearHyperLink.Label = "Remove HyperLink"
        Me.ClearHyperLink.Name = "ClearHyperLink"
        Me.ClearHyperLink.ScreenTip = "Remove all hyperLink boxes"
        Me.ClearHyperLink.ShowImage = True
        '
        'ClearPen
        '
        Me.ClearPen.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ClearPen.Image = CType(resources.GetObject("ClearPen.Image"), System.Drawing.Image)
        Me.ClearPen.Label = "Remove Marker"
        Me.ClearPen.Name = "ClearPen"
        Me.ClearPen.ScreenTip = "Remove all markers"
        Me.ClearPen.ShowImage = True
        '
        'ClearAllShape
        '
        Me.ClearAllShape.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ClearAllShape.Image = CType(resources.GetObject("ClearAllShape.Image"), System.Drawing.Image)
        Me.ClearAllShape.Label = "Remove All"
        Me.ClearAllShape.Name = "ClearAllShape"
        Me.ClearAllShape.OfficeImageId = "RemoveAttach"
        Me.ClearAllShape.ScreenTip = "Remove all shapes"
        Me.ClearAllShape.ShowImage = True
        '
        'Quote
        '
        Me.Quote.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Quote.Image = CType(resources.GetObject("Quote.Image"), System.Drawing.Image)
        Me.Quote.Label = "Quote"
        Me.Quote.Name = "Quote"
        Me.Quote.ScreenTip = "Get motivated every day!"
        Me.Quote.ShowImage = True
        '
        'Copyright_ToggleButton
        '
        Me.Copyright_ToggleButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Copyright_ToggleButton.Image = CType(resources.GetObject("Copyright_ToggleButton.Image"), System.Drawing.Image)
        Me.Copyright_ToggleButton.Label = "Copyright"
        Me.Copyright_ToggleButton.Name = "Copyright_ToggleButton"
        Me.Copyright_ToggleButton.ScreenTip = "Copyright"
        Me.Copyright_ToggleButton.ShowImage = True
        '
        'Menu5
        '
        Me.Menu5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu5.Items.Add(Me.Preparer2)
        Me.Menu5.Items.Add(Me.Reviewer2)
        Me.Menu5.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Menu5.Label = "Signature"
        Me.Menu5.Name = "Menu5"
        Me.Menu5.OfficeImageId = "SignatureLineInsert"
        Me.Menu5.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group10.ResumeLayout(False)
        Me.Group10.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group9.ResumeLayout(False)
        Me.Group9.PerformLayout()
        Me.Group8.ResumeLayout(False)
        Me.Group8.PerformLayout()
        Me.Group17.ResumeLayout(False)
        Me.Group17.PerformLayout()
        Me.Roman_numerals.ResumeLayout(False)
        Me.Roman_numerals.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.Group15.ResumeLayout(False)
        Me.Group15.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group18.ResumeLayout(False)
        Me.Group18.PerformLayout()
        Me.Group14.ResumeLayout(False)
        Me.Group14.PerformLayout()
        Me.Group12.ResumeLayout(False)
        Me.Group12.PerformLayout()
        Me.Group16.ResumeLayout(False)
        Me.Group16.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group11.ResumeLayout(False)
        Me.Group11.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group13.ResumeLayout(False)
        Me.Group13.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents CapP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents FontName1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FontName2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PBC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents CheckBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ClearCheckBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents RedPen As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BlackPen As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ClearAllShape As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents RedArrow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents HyperLink As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SheetArrow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents footed As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tick As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DoubleLine As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents a As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents b As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents c As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents siga As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents sigb As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents sigc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group9 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AJE As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents RJE As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group8 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents OO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents star1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents star2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents o As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tri1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tri2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents dia1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents squ1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents sigma As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents dia2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents squ2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents divided As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents alpha As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents beta As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents gamma As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CommaStyle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents InsertColumn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group10 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group11 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Preparer2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Reviewer2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group12 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Note As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Note1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Note2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Note3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Note4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Note5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Copyright_ToggleButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ClearHyperLink As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Quote As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group13 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ClearPen As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group14 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Direction1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Direction2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Direction3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Direction4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Direction5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents hyperlinkCell As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Calendar_ToggleButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group15 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents LeftBrace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DownApply As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents RightApply As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PMingLiU As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Arial As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BookAntiqua As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Calibri As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MultiIllu As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BluePen As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Group16 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group17 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents delta As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents epsilon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents zeta As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents eta As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents theta As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents mu As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents pi As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents rho As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents phi As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents psi As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents omega As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Roman_numerals As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents one As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents two As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents three As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents four As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents five As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents six As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents seven As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents eight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents nine As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ten As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Menu4 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Group18 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ColorYellow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ColorAqua As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ColorLime As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ColorSilver As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ColorFuchsia As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ColorWhite As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents RD As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tick2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tick3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FS As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents NotMaterial As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Minor_Pass As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Conclusion As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ArrowBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu5 As Microsoft.Office.Tools.Ribbon.RibbonMenu
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
