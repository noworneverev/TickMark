'-------------------------------------------------------------------------
' Created by : Liao Yan Ying
'Date :  03.25.2017
'Purpose: Tick marks For auditors
'Version: 1.0.0
'Copyright:This addin I created are licensed under 
'          the Creative Commons Attribution 3.0 License, 
'          which means you can use them for personal stuff,
'          use them for commercial stuff, And change them 
'          however you Like. In exchange, just give Me credit 
'          For the design And tell your friends/colleagues about it:)

'-------------------------------------------------------------------------
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub CapP_Click(sender As Object, e As RibbonControlEventArgs) Handles CapP.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "P"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .Font.Italic = True
        End With
    End Sub
    Private Sub footed_Click(sender As Object, e As RibbonControlEventArgs) Handles footed.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "f"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .Font.Italic = True
        End With
    End Sub
    Private Sub tick_Click(sender As Object, e As RibbonControlEventArgs) Handles tick.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "✔"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .Font.Italic = True
        End With
    End Sub
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "✗"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .Font.Italic = True
        End With
    End Sub

    Private Sub FontName1_Click(sender As Object, e As RibbonControlEventArgs) Handles FontName1.Click
        Globals.ThisAddIn.Application.Selection.Font.Name = "標楷體"
    End Sub

    Private Sub FontName2_Click(sender As Object, e As RibbonControlEventArgs) Handles FontName2.Click
        Globals.ThisAddIn.Application.Selection.Font.Name = "微軟正黑體"
    End Sub

    Private Sub PBC_Click(sender As Object, e As RibbonControlEventArgs) Handles PBC.Click
        Dim shape As Object
        shape = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 50, 25)
        shape.TextFrame2.TextRange.text = "PBC"
        shape.line.forecolor.rgb = RGB(102, 178, 255)
        With shape.textframe2.textrange.font
            .size = 16
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
    End Sub

    Private Sub Calendar_Click(sender As Object, e As RibbonControlEventArgs)
        Dim ctpCal As Microsoft.Office.Tools.CustomTaskPane =
 Globals.ThisAddIn.CustomTaskPanes.Add(New CalendarTaskPaneControl, "Calendar")
        ctpCal.Width = 232
        ctpCal.Visible = True
    End Sub


    Private Sub CheckBox_Click(sender As Object, e As RibbonControlEventArgs) Handles CheckBox.Click
        Dim checkbox As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        checkbox = Globals.ThisAddIn.Application.ActiveSheet.checkboxes.add _
            (Left, Top, Width, Height)
        With checkbox
            .Caption = ""
        End With
    End Sub

    Private Sub ClearCheckBox_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearCheckBox.Click
        Dim sht As Excel.Worksheet
        sht = Globals.ThisAddIn.Application.ActiveSheet
        For Each cb In sht.CheckBoxes
            If cb.TopLeftCell.Address = Globals.ThisAddIn.Application.ActiveCell.Address Then cb.Delete
        Next
    End Sub

    Private Sub Submit_Click(sender As Object, e As RibbonControlEventArgs)
        Dim username As String
        Dim submit As String
        Dim time As Date = Now

        username = Environment.UserName
        submit = "Prepared by" & " " & username & " " & "on " & time.Month _
                                                              & "/" & time.Day _
                                                              & "/" & time.Year

        With Globals.ThisAddIn.Application.ActiveCell
            .Value = submit
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub Review_Click(sender As Object, e As RibbonControlEventArgs)
        Dim username As String
        Dim review As String
        Dim time As Date = Now

        username = Environment.UserName

        review = "Reviewed by" & " " & username & " " & "on " & time.Month _
                                                             & "/" & time.Day _
                                                             & "/" & time.Year
        With Globals.ThisAddIn.Application.ActiveCell
            .Value = review
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub RedPen_Click(sender As Object, e As RibbonControlEventArgs) Handles RedPen.Click
        Dim connector As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        connector = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 15 + Width, Top + 2, Left - 2.5 + Width, Top + 9.25)
        With connector
            .Name = "RedPen"
            .line.weight = 2
            .line.ForeColor.RGB = RGB(255, 0, 0)
        End With
    End Sub

    Private Sub BluePen_Click(sender As Object, e As RibbonControlEventArgs) Handles BluePen.Click
        Dim connector As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        connector = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 15 + Width, Top - 1, Left - 2.5 + Width, Top + 6.25)
        With connector
            .Name = "BluePen"
            .line.weight = 2
            .line.ForeColor.RGB = RGB(0, 0, 255)
        End With
    End Sub

    Private Sub BlackPen_Click(sender As Object, e As RibbonControlEventArgs) Handles BlackPen.Click
        Dim connector As Object
        Dim Left As Double = Globals.ThisAddIn.Application.ActiveCell.Left
        Dim Top As Double = Globals.ThisAddIn.Application.ActiveCell.Top
        Dim Width As Double = Globals.ThisAddIn.Application.ActiveCell.Width
        Dim Height As Double = Globals.ThisAddIn.Application.ActiveCell.Height

        connector = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 15 + Width, Top - 6, Left - 2.5 + Width, Top + 1.25)
        With connector
            .Name = "BlackPen"
            .line.weight = 2
            .line.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End Sub

    Private Sub ClearAllShape_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearAllShape.Click
        Dim sht As Excel.Worksheet
        sht = Globals.ThisAddIn.Application.ActiveSheet
        sht.Shapes.SelectAll()
        Globals.ThisAddIn.Application.Selection.Delete()
    End Sub

    Private Sub RedArrow_Click(sender As Object, e As RibbonControlEventArgs) Handles RedArrow.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connector As Object
        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing an arrow", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If
        Dim TargetLeft As Double = location.left
        Dim TargetTop As Double = location.Top
        Dim TargetHeight As Double = location.Height


        connector = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + 4, Top, TargetLeft + 4, TargetTop + TargetHeight)
        With connector
            .line.weight = 1
            .line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
        End With
    End Sub

    Private Sub HyperLink_Click(sender As Object, e As RibbonControlEventArgs) Handles HyperLink.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        'Origin Sheet
        Dim strBeginningSheetName As String = app.Activesheet.name
        Dim BeginningSheet As Object = Globals.ThisAddIn.Application.ActiveSheet
        'Store the Original sheet & cell location
        Dim OriSheetCellAddress As String = appCell.address

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height
        Dim connector1 As Excel.Shape
        Dim connector2 As Excel.Shape

        Dim location As Object
        location = app.inputbox("Select a cell", "Creating a HyperLink", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If
        Dim TargetLeft As Double = location.left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height

        'Get the Target's sheet name
        Dim strDestinationSheetName As String = location.Parent.Name

        'Create the textbox in the beginning sheet
        Dim shapeBeg As Object
        shapeBeg = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width * 0.796, Top + 7, 100, Height)
        shapeBeg.TextFrame2.TextRange.text = strDestinationSheetName
        With shapeBeg.textframe2.textrange.font
            .size = 9
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        With shapeBeg
            .Name = "box1"
            .fill.visible = Microsoft.Office.Core.MsoTriState.msoFalse
            .line.visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End With
        shapeBeg.TextFrame.autosize = True

        'Add the target address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=shapeBeg,
             Address:="", SubAddress:=strDestinationSheetName & "!" & location.address,
             TextToDisplay:=strDestinationSheetName)

        'Add the connector1 to point to the textbox
        connector1 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width * 0.796, Top + 12, Left + Width * 0.796 + 6, Top + 14)
        With connector1
            .Name = "line1"
            .line.weight = 1
            .line.ForeColor.RGB = RGB(255, 0, 0)
        End With

        'Create the hyperlink next to the Target Cell in Target Sheet
        'Select the target sheet first
        app.Sheets(strDestinationSheetName).Activate
        Dim shape As Object
        shape = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 40, TargetTop - 10, 100, TargetHeight)
        shape.TextFrame2.TextRange.text = strBeginningSheetName
        With shape.textframe2.textrange.font
            .size = 9
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        With shape
            .Name = "box2"
            .fill.visible = Microsoft.Office.Core.MsoTriState.msoFalse
            .line.visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End With
        shape.TextFrame.autosize = True
        'Add the origin address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=shape,
             Address:="", SubAddress:=strBeginningSheetName & "!" & OriSheetCellAddress,
             TextToDisplay:=strBeginningSheetName)

        'Add the connector2 to point to the textbox
        connector2 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + 10, TargetTop + 5, TargetLeft + 1, TargetTop + 1)
        With connector2
            .Name = "line2"
            .line.weight = 1
            .line.ForeColor.RGB = RGB(255, 0, 0)
        End With

        app.Sheets(strBeginningSheetName).Activate
    End Sub

    Private Sub SheetArrow_Click(sender As Object, e As RibbonControlEventArgs) Handles SheetArrow.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting arrows", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height


        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left, Top + Height, Left - 5, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop, TargetLeft + TargetWidth + 5, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left, Top, Left - 5, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight, TargetLeft + TargetWidth + 5, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left, Top + Height, Left - 13, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight, TargetLeft + TargetWidth + 13, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height, Left + Width + 13, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft, TargetTop + TargetHeight, TargetLeft - 13, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width, Top + Height, Left + Width + 5, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft, TargetTop, TargetLeft - 5, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width, Top, Left + Width + 5, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft, TargetTop + TargetHeight, TargetLeft - 5, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left = targetleft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top, Left + Width, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight, TargetLeft + TargetWidth, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        ElseIf (Left = targetleft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height, Left + Width, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop, TargetLeft + TargetWidth, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen
                .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval
            End With
        End If

    End Sub

    Private Sub DoubleLine_Click(sender As Object, e As RibbonControlEventArgs) Handles DoubleLine.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        With app.selection.borders(Excel.XlBordersIndex.xlEdgeBottom)
            .linestyle = Excel.XlLineStyle.xlDouble
        End With
    End Sub

    Private Sub a_Click(sender As Object, e As RibbonControlEventArgs) Handles a.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "a"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub b_Click(sender As Object, e As RibbonControlEventArgs) Handles b.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "b"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub c_Click(sender As Object, e As RibbonControlEventArgs) Handles c.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "c"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub siga_Click(sender As Object, e As RibbonControlEventArgs) Handles siga.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "∑a="
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub sigb_Click(sender As Object, e As RibbonControlEventArgs) Handles sigb.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "∑b="
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub sigc_Click(sender As Object, e As RibbonControlEventArgs) Handles sigc.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        With appCell
            .Value = begValue & "∑c="
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Calisto MT"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub AJE_Click(sender As Object, e As RibbonControlEventArgs) Handles AJE.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "AJE<>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub RJE_Click(sender As Object, e As RibbonControlEventArgs) Handles RJE.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "RJE<>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub o_Click(sender As Object, e As RibbonControlEventArgs) Handles o.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "○"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub OO_Click(sender As Object, e As RibbonControlEventArgs) Handles OO.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "◎"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub star1_Click(sender As Object, e As RibbonControlEventArgs) Handles star1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "☆"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub star2_Click(sender As Object, e As RibbonControlEventArgs) Handles star2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "★"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub tri1_Click(sender As Object, e As RibbonControlEventArgs) Handles tri1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "△"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub tri2_Click(sender As Object, e As RibbonControlEventArgs) Handles tri2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "▲"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub dia1_Click(sender As Object, e As RibbonControlEventArgs) Handles dia1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "◇"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub dia2_Click(sender As Object, e As RibbonControlEventArgs) Handles dia2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "◆"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub squ1_Click(sender As Object, e As RibbonControlEventArgs) Handles squ1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "□"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub squ2_Click(sender As Object, e As RibbonControlEventArgs) Handles squ2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "■"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub sigma_Click(sender As Object, e As RibbonControlEventArgs) Handles sigma.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "∑"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub divided_Click(sender As Object, e As RibbonControlEventArgs) Handles divided.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "/"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Book Antiqua"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub alpha_Click(sender As Object, e As RibbonControlEventArgs) Handles alpha.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "α"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub beta_Click(sender As Object, e As RibbonControlEventArgs) Handles beta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "β"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub gamma_Click(sender As Object, e As RibbonControlEventArgs) Handles gamma.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "γ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub CommaStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles CommaStyle.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim number As Double = app.Activecell.value
        If number >= 0 Then
            app.Selection.NumberFormatLocal = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
        Else
            app.Selection.NumberFormatLocal = "#,##0_);(#,##0)"
        End If
    End Sub

    Private Sub InsertColumn_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertColumn.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell
        app.selection.entirecolumn.Offset(0, 1).insert
        appCell.Offset(0, 1).Select
        app.selection.columnwidth = 2

    End Sub

    Private Sub Preparer2_Click(sender As Object, e As RibbonControlEventArgs) Handles Preparer2.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim username As String
        Dim submit As String
        Dim time As Date = Now

        username = app.UserName
        submit = "Prepared by" & " " & username & " " & "on " & time.Month _
                                                              & "/" & time.Day _
                                                              & "/" & time.Year

        With app.ActiveCell
            .Value = submit
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub Reviewer2_Click(sender As Object, e As RibbonControlEventArgs) Handles Reviewer2.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim username As String
        Dim review As String
        Dim time As Date = Now

        username = app.UserName
        review = "Reviewed by" & " " & username & " " & "on " & time.Month _
                                                              & "/" & time.Day _
                                                              & "/" & time.Year

        With app.ActiveCell
            .Value = review
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 12
            .Font.Name = "Arial"
            .Font.Italic = True
        End With
    End Sub

    Private Sub Note_Click(sender As Object, e As RibbonControlEventArgs) Handles Note.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<Note>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note1_Click(sender As Object, e As RibbonControlEventArgs) Handles Note1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N1>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note2_Click(sender As Object, e As RibbonControlEventArgs) Handles Note2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N2>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note3_Click(sender As Object, e As RibbonControlEventArgs) Handles Note3.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N3>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note4_Click(sender As Object, e As RibbonControlEventArgs) Handles Note4.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N4>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Note5_Click(sender As Object, e As RibbonControlEventArgs) Handles Note5.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "<N5>"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Narrow"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Copyright_ToggleButton_Click(sender As Object, e As RibbonControlEventArgs) Handles Copyright_ToggleButton.Click
        Globals.ThisAddIn.TaskPane.Visible =
    TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
    End Sub

    Private Sub ClearHyperLink_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearHyperLink.Click

        Dim app As Object = Globals.ThisAddIn.Application
        Dim sht As Excel.Worksheet = Nothing
        Dim shp As Excel.Shape = Nothing

        sht = app.ActiveWorkbook.ActiveSheet
        For Each shp In sht.Shapes
            If shp.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox And (shp.Name = "box1" Or shp.Name = "box2") Then
                shp.Select()
                shp.Delete()
            ElseIf shp.Type = Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight And (shp.Name = "line1" Or shp.Name = "line2") Then
                shp.Select()
                shp.Delete()
            End If
        Next

    End Sub

    Private Sub Quote_Click(sender As Object, e As RibbonControlEventArgs) Handles Quote.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value
        Dim Quote As String = Nothing
        Randomize()
        ' Generate random value between 1 and 900.
        Dim value As Integer = Int((900 * Rnd()) + 1)
        Select Case value
            Case 1
                Quote = "The bird who dares to fall is the bird who learns to fly."
                Exit Select
            Case 2
                Quote = "Work hard in silence. Let success be your noise."
                Exit Select
            Case 3
                Quote = "One of the hardest decisions you'll ever face in life is choosing whether to walk away or try harder. - Ziad K. Abdelnour"
                Exit Select
            Case 4
                Quote = "As long as you think your past is bad you must be improving. - Louis C.K."
                Exit Select
            Case 5
                Quote = "And if i asked you to name all the things that you love, how long would it take for you to name yourself."
                Exit Select
            Case 6
                Quote = "Behind every successful person, there is a lot of unsuccessful years."
                Exit Select
            Case 7
                Quote = "You know, sometimes all you need is twenty seconds of insane courage. Just literally twenty seconds of just embarrassing bravery and I promise you, something great will come of it."
                Exit Select
            Case 8
                Quote = "If you want something you've never had, then you've got to do something you've never done."
                Exit Select
            Case 9
                Quote = "Remember, time you enjoy wasting is not time wasted."
                Exit Select
            Case 10
                Quote = "I will not let my sense of pride keep me from appreciating an act of kindness."
                Exit Select
            Case 11
                Quote = "The more you learn, the less you feel like you know."
                Exit Select
            Case 12
                Quote = "The greatest pleasure in life is doing what people say you cannot do."
                Exit Select
            Case 13
                Quote = "I'd always end up broken down on the highway. When I stood there trying to flag someone down, nobody stopped. But when I pushed my own car, other drivers would get out and push with me. If you want help, help youreself - people like to see that. - Chris Rock"
                Exit Select
            Case 14
                Quote = "If it is important to you, you will find a way. If not, you'll find an excuse."
                Exit Select
            Case 15
                Quote = "Some people died at 25 and aren't burned until 75."
                Exit Select
            Case 16
                Quote = "Listen, smile, agree, and then do whatever the fuck you were gonna do anyway."
                Exit Select
            Case 17
                Quote = "When life gives you lemons, EAT THEM WHOLE. Seriously. Just choke them all down... skin, pulp, seeds and all and don't break eye contact. Maybe life will stop being suck an asshole if yhou show it that youre done fucking around."
                Exit Select
            Case 18
                Quote = "BE BRAVE. Even if you're not, pretend to be."
                Exit Select
            Case 19
                Quote = "We can't be everything to everyone, but we can be something to someone ... even a lot of someones."
                Exit Select
            Case 20
                Quote = "Generosity is giving without expecting anything in return."
                Exit Select
            Case 21
                Quote = "When we have a clear sense of where we are going, we are flexible in how we get there."
                Exit Select
            Case 22
                Quote = "Dream big. Start small. But most of all … Start."
                Exit Select
            Case 23
                Quote = "Your time is limited, so don't waste it living someone else's life.  -Steve Jobs"
                Exit Select
            Case 24
                Quote = "Don't be trapped by dogma - which is living with the results of other people's thinking.  -Steve Jobs"
                Exit Select
            Case 25
                Quote = "Don't let the noise of other's opinions drown out your own inner voice.  -Steve Jobs"
                Exit Select
            Case 26
                Quote = "And most important, have the courage to follow your heart and intuition. They somehow already know what you truly want to become. Everything else is secondary.  -Steve Jobs"
                Exit Select
            Case 27
                Quote = "If you don't follow your heart, you might spend the rest of your time wishing you had."
                Exit Select
            Case 28
                Quote = "Don't die before you're dead."
                Exit Select
            Case 29
                Quote = "You're not a loser. You know what a loser is? A real loser is somebody that's so afraid of not winning, they don't even try."
                Exit Select
            Case 30
                Quote = "Life has no smooth road for any us. As we go down it, we need to remember that happiness is a talent we develop, not an object we seek. It's the ability to bounce back from life's inevitable setbacks. Some people are crushed by misfortune. Others grow because of it."
                Exit Select
            Case 31
                Quote = "IT'S OKAY TO HAVE A BAD DAY."
                Exit Select
            Case 32
                Quote = "A ship in the harbor is safe, but that's not what ships are for."
                Exit Select
            Case 33
                Quote = "I understand there's a guy inside me who wants to lay in bed, smoke weed all day, and watch cartoons and old movies. My whole life is a series of stratagems to avoid, and outwit, that guy. - Anthony Bourdain"
                Exit Select
            Case 34
                Quote = "We all must suffer one of two things: the pain of discipline or the pain of regret."
                Exit Select
            Case 35
                Quote = "Stop being afraid of what could go wrong, and start being excited about what could go right."
                Exit Select
            Case 36
                Quote = "It gets easier. Every day, it gets a little easier. But you gotta do it every day. That's the hard part. But it does get easier."
                Exit Select
            Case 37
                Quote = "You have survived every single bad day so far."
                Exit Select
            Case 38
                Quote = "Please pay sttention very carefully, because this is the truest thing a stranger will ever say to you: In the face of such hopelessness as our eventual, unavoidable death, there is little sense in not at least TRYING to accomplish all your wildest dreams in life. - Kevin Smith"
                Exit Select
            Case 39
                Quote = "Some say they stay up late because they're delaying tomorrow. I stay up late Because I'm not done with today."
                Exit Select
            Case 40
                Quote = "I know people who graduated college at 21 and didn't get a salary job until they were 27. I know people who have children and are single, I know people who are married and had to wait 8-10 years to be parents. I know people who are in a relationship and love someone else, I know people who love each other and aren't together, there are people waiting to love and be loved. My point is, everything in life happens according to our time, our clock. You may look at your friends and some may seem to be ahead or behind you, but they are not, they are living according to the pace of their clock, so be patient. You are not falling behind, it's just not your time."
                Exit Select
            Case 41
                Quote = "To the world you may be one person, but to one person you may be the world."
                Exit Select
            Case 42
                Quote = "For what it's worth: It's never too late to be whoever you want to be. I hope you live a life you're proud of, and if you find that you're not, I hope you have the strength to start over."
                Exit Select
            Case 43
                Quote = "Just because you took longer than others, doesn't mean you failed. Remember that."
                Exit Select
            Case 44
                Quote = "When something is important enough, you do it even if the odds are not in your favor. - Elon Musk"
                Exit Select
            Case 45
                Quote = "Sometimes you just have to chuck it in the fuck it bucket and move on."
                Exit Select
            Case 46
                Quote = "You have $86,400 in your account and someone stole $10 from you, would you be upset and throw all of the $86,390 away in hopes of getting back at the person that took your $10? Or move on and live? Right, move on and live. See, we have 86,400 secs in every day so don't let someone's negative 10 seconds ruin the rest of the 86,390. Don't sweat the small stuff, life is bigger than that."
                Exit Select
            Case 47
                Quote = "Service to others is the rent you pay for your room here on Earth. - Muhammad Ali"
                Exit Select
            Case 48
                Quote = "Don't worry if people don't like you. Most people are struggling to like themselves."
                Exit Select
            Case 49
                Quote = "Don't cling to a mistake just because you spent a lot of time making it."
                Exit Select
            Case 50
                Quote = "And so rock bottom became the solid foundation on which I rebuilt my life. - J.K. Rowling"
                Exit Select
            Case 51
                Quote = "If it falls your lot to be a street sweeper, sweep streets like Michelangelo painted pictures, sweep streets like Beethoven composed music, sweep streets like Leontyne Price sings before the Metropolitan Opera. Sweep streets like Shakespeare wrote poetry. Sweep streets so well that all the hosts of heaven and earth will have to pause and say: Here lived a great street sweeper who swept his job well. If you can’t be a pine at the top of the hill, be a shrub in the valley. Be be the best little shrub on the side of the hill. - Dr. Martin Luther King Jr"
                Exit Select
            Case 52
                Quote = "One day or day one. You decide."
                Exit Select
            Case 53
                Quote = "Dude, suckin' at something is the first step to being sorta good at something."
                Exit Select
            Case 54
                Quote = "The pessimist complains about the wind; the optimist expects it to change; the realist adjusts the sails. - William A. Ward"
                Exit Select
            Case 55
                Quote = "I've missed more the 9,000 shots in my career. I've lost almost 300 games. 26 times I've been trusted to take the game winning shot and missed. I've failed over and over and over again in my life. And that is why I succeed. - Michael Jordan"
                Exit Select
            Case 56
                Quote = "Don't think about what can happen in a month . Don't think about what can happen in a year. Just focus on the 24 hours in front of you and do what you can to get closer to where you want to be."
                Exit Select
            Case 57
                Quote = "The human being is meant to bear the burden of 24 hours -- no more, no less."
                Exit Select
            Case 58
                Quote = "Happiness is not something you can pursue - but instead the byproduct of doing the right thing."
                Exit Select
            Case 59
                Quote = "The world's idea of success is total shit."
                Exit Select
            Case 60
                Quote = "If you can survive disappointment, then nothing can beat you. - Louis C.K."
                Exit Select
            Case 61
                Quote = "A river cuts through rock, not because of its power, but because of its persistence. - Jim Watkins"
                Exit Select
            Case 62
                Quote = "Talent is nothing more than a pursued interest - Bob Ross"
                Exit Select
            Case 63
                Quote = "When it feels scary to jump, that is exactly when you jump. Otherwise, you end up staying in the same place your whole life."
                Exit Select
            Case 64
                Quote = "If I quit now, I will soon be back to where I started, and when I started, I was desperately wishing to be where I am now."
                Exit Select
            Case 65
                Quote = "The happiest people don't have the best of everything, they just make the best of everything."
                Exit Select
            Case 66
                Quote = "In case no one has told you today. Or when you're needed to hear it. You are loved. You are beautiful. You are important. You are forgiven. "
                Exit Select
            Case 67
                Quote = "We all have two fates. One Is the result of whatever mess you were born into, And one Is the result of deciding that's just not fucking good enough."
                Exit Select
            Case 68
                Quote = "Stop waiting for Friday, for summer, for someone to fall in love with you, for life. Happiness is achieved when you stop waiting for it and make the most of the moment you're in right now."
                Exit Select
            Case 69
                Quote = "If you chase two rabbits, you will not catch either one."
                Exit Select
            Case 70
                Quote = "Don't settle. Don't finish bad books. If you don't like the menu, leave the restaurant. If you're not on the right path, get off."
                Exit Select
            Case 71
                Quote = "Be decisive. Right or wrong, make a decision. The road of life is paved with flat squirrels who couldn't make a decision."
                Exit Select
            Case 72
                Quote = "If you don't design your own life plan, chances are you'll fall into someone else's plan. And guess what they have planned for you? Not much. - Jim Rohn"
                Exit Select
            Case 73
                Quote = "I never lose. I either win or learn. - Nelson Mandela"
                Exit Select
            Case 74
                Quote = "I always wonder why birds stay in the same place when they can fly anywhere on the earth. Then I ask myself the same question. - Harun Yahya"
                Exit Select
            Case 75
                Quote = "I wonder if we might pledge ourselves to remember what life is really all about - not to be afraid that we're less flashy than the next, not to worry that our influence is not that of a tornado, but rather that of a grain of sand in an oyster! Do we have that kind of patience? - Fred Rogers"
                Exit Select
            Case 76
                Quote = "These mountains that you are carrying you were only supposed to climb."
                Exit Select
            Case 77
                Quote = "The same boiling water that softens the potato, hardens the egg. It's about what you're made of, not the circumstances."
                Exit Select
            Case 78
                Quote = "This is your life and you're writing your story. Don’t ever let someone else hold the pen."
                Exit Select
            Case 79
                Quote = "I constantly get out of my comfort zone. Looking cool is the easiest way to mediocrity. The coolest guy in my high school ended up working at a car wash. Once you push yourself into something new, and whole new world of opportunities opens up. But you might get hurt. In fact you will get hurt. But amazingly when you heal – you are somewhere you’ve never been. - Terry Crews"
                Exit Select
            Case 80
                Quote = "Make it happen now, not tomorrow. Tomorrow is a loser's excuse."
                Exit Select
            Case 81
                Quote = "Whatever you're thinking, think bigger. - Tony Hsieh"
                Exit Select
            Case 82
                Quote = "You never fail until you stop trying. - Albert Einstein"
                Exit Select
            Case 83
                Quote = "If you want to live a happy life, tie it to a goal, not to people or things. - Albert Einstein"
                Exit Select
            Case 84
                Quote = "A clever person solves a problem. A wise person avoids it. - Albert Einstein"
                Exit Select
            Case 85
                Quote = "Once we accept our limits, we go beyond them. - Albert Einstein"
                Exit Select
            Case 86
                Quote = "Everybody is a genius. But if you judge a fish by its ability to climb a tree, it will live its whole life believing that it is stupid. - Albert Einstein"
                Exit Select
            Case 87
                Quote = "In the middle of difficulty lies opportunity. - Albert Einstein"
                Exit Select
            Case 88
                Quote = "Life is like riding a bicycle. To keep your balance, you must keep moving. - Albert Einstein"
                Exit Select
            Case 89
                Quote = "I must be willing to give up what I am  in order to become what I will be. - Albert Einstein"
                Exit Select
            Case 90
                Quote = "I have no special talents. I am only passionately curious. - Albert Einstein"
                Exit Select
            Case 91
                Quote = "Anyone who has never made a mistake has never tried anything new. - Albert Einstein"
                Exit Select
            Case 92
                Quote = "Common sense is what tells us the earth is flat. - Albert Einstein"
                Exit Select
            Case 93
                Quote = "Try not to become a man of success. Rather become a man of value. - Albert Einstein"
                Exit Select
            Case 94
                Quote = "The mind that opens to a new idea never returns to its original size."
                Exit Select
            Case 95
                Quote = "Life is like riding a bicycle. To keep your balance, you must keep moving. - Albert Einstein"
                Exit Select
            Case 96
                Quote = "Life begins at the end of your comfort zone. - Neale Donald Walsch"
                Exit Select
            Case 97
                Quote = "Two things define you: Your patience when you have nothing and your attitude when you have everything."
                Exit Select
            Case 98
                Quote = "Never give up on a dream just because of the time it will take to accomplish it. The time will pass anyway. - Earl Nightingale"
                Exit Select
            Case 99
                Quote = "If people are not laughing at your goals, your goals are too small. - Azim Premji"
                Exit Select
            Case 100
                Quote = "You don't have to be great to start, but you have to start to be great. - Zig Ziglar"
                Exit Select
            Case 101
                Quote = "I am not what happened to me, I am what I choose to become. - Carl Gustav Jung"
                Exit Select
            Case 102
                Quote = "An entire sea of water can't sink a ship unless it gets inside the ship. Similarly, the negativity of the world can't put you down unless you allow it to get inside you. - Goi Nasu"
                Exit Select
            Case 103
                Quote = "Great minds discuss ideas. Average minds discuss events. Small minds discuss people. - Eleanor Roosevelt"
                Exit Select
            Case 104
                Quote = "It does not matter how slowly you go as long as you do not stop."
                Exit Select
            Case 105
                Quote = "Hard work beats talent when talent doesn't work hard. - Tim Notke"
                Exit Select
            Case 106
                Quote = "Doubt kills more dreams than failure ever will. - Suzy Kassem"
                Exit Select
            Case 107
                Quote = "Knowing is not enough, we must apply. Willing is not enough, we must do. - Bruce Lee"
                Exit Select
            Case 108
                Quote = "The first step to getting anywhere is deciding you're no longer willing to stay where you are."
                Exit Select
            Case 109
                Quote = "A year from now you may wish you had started today. - Karen Lamb"
                Exit Select
            Case 110
                Quote = "The activity you're most avoding contains your biggest opportunity. - Robin S. Sharma"
                Exit Select
            Case 111
                Quote = "The master has failed more times than the beginner has even tried. - Stephen McCranie"
                Exit Select
            Case 112
                Quote = "The world is changed by your example, not by your opinion.- Paulo Coelho"
                Exit Select
            Case 113
                Quote = "I hated every minute of training, but I said, 'Don't quit. Suffer now and live the rest of your life as a champion'. - Muhammad Ali"
                Exit Select
            Case 114
                Quote = "If you don't make the time to work on creating the life you want, you're eventually going to be forced to spend a lot of time dealing with a life you don't want. - Kevin Ngo"
                Exit Select
            Case 115
                Quote = "When someone tells me ""no,"" it doesn't mean I can't do it, it simply means I can't do it with them. - Karen E. Quinones Miller"
                Exit Select
            Case 116
                Quote = "When you want to succeed as bad as you want to breathe, then you'll be successful. - Eric Thomas"
                Exit Select
            Case 117
                Quote = "Today I will do what others won't so tomorrow I can what others can't. - Jerry Rice"
                Exit Select
            Case 118
                Quote = "Never confuse a single defeat with a final defeat. - F. Scott Fitzgerald"
                Exit Select
            Case 119
                Quote = "Turn your face to the sun and the shadows fall behind you. - Maori Proverb"
                Exit Select
            Case 120
                Quote = "I don't count my sit-ups. I only start counting when it starts hurting. When I feel pain, that's when I start counting, because that's when it really counts. - Muhammad Ali"
                Exit Select
            Case 121
                Quote = "Just stick with it. What seems so hard now will one day be your warm up."
                Exit Select
            Case 122
                Quote = "If you don't build your dream someone will hire you to help build theirs. - Tony Gaskins"
                Exit Select
            Case 123
                Quote = "What would you do if you weren't afraid? - Sheryl Sandberg"
                Exit Select
            Case 124
                Quote = "Courage doesn't always roar. Sometimes courage is the little voice at the end of the day that says I'll try again tomorrow."
                Exit Select
            Case 125
                Quote = "Do what you can, with what you have, where you are. - Theodore Roosevelt"
                Exit Select
            Case 126
                Quote = "If you think you are too small to make a difference, try sleeping with a mosquito. - Dalai Lama XIV"
                Exit Select
            Case 127
                Quote = "How we spend our days is, of course, how we spend our lives. - Annie Dillard"
                Exit Select
            Case 128
                Quote = "I already know what giving up feels like. I want to see what happens if I don't."
                Exit Select
            Case 129
                Quote = "Stop letting people who do so little for you control so much of your mind, feeling s and emotions. - Will Smith"
                Exit Select
            Case 130
                Quote = "The habits that took years to build, do not take a day to change. - Susan Powter"
                Exit Select
            Case 131
                Quote = "Worry a little bit every day and in a lifetime you will lose a couple of years. If something is wrong, fix it if you can. But train yourself not to worry: Worry never fixes anything. - Ernest Hemingway"
                Exit Select
            Case 132
                Quote = "It's not the load that breaks you down, it's the way you carry it. - Lou Holtz"
                Exit Select
            Case 133
                Quote = "Always be a first rate version of yourself and not a second rate version of someone else. - Judy Garland"
                Exit Select
            Case 134
                Quote = "Better to be the one who smiled than the one who didn't smile back. - Mari Gayatri Stein"
                Exit Select
            Case 135
                Quote = "I'm too busy working on my own grass to notice if yours is greener."
                Exit Select
            Case 136
                Quote = "Discipline is the bridge between goals and accomplishment. - Jim Rohn"
                Exit Select
            Case 137
                Quote = "If you can't do great things, do small things in a great way. - Napoleon Hill"
                Exit Select
            Case 138
                Quote = "A goal is a dream with a deadline. - Napoleon Hill"
                Exit Select
            Case 139
                Quote = "Motivation is what gets you started. Habit is what keeps you going."
                Exit Select
            Case 140
                Quote = "Success is the sum of small efforts, repeated day in and day out. - Robert Collier"
                Exit Select
            Case 141
                Quote = "A quitter never wins and a winner never quits. - Napoleon Hill"
                Exit Select
            Case 142
                Quote = "You miss 100% of the shots you don't take. - Wayne Gretzky"
                Exit Select
            Case 143
                Quote = "Fall down seven times, get up eight. - Japanese Proverb"
                Exit Select
            Case 144
                Quote = "Don't wish it were easier. Wish you were better. - Jim Rohn"
                Exit Select
            Case 145
                Quote = "Screw it. Let's do it. - Richard Branson"
                Exit Select
            Case 146
                Quote = "Nothing is particularly hard if you divide it into small jobs. - Henry Ford"
                Exit Select
            Case 147
                Quote = "Formal education will make you a living. Self-education will make you a fortune. - Jim Rohn"
                Exit Select
            Case 148
                Quote = "What is not started today is never finished tomorrow. - Johann Wolfgang von Goethe"
                Exit Select
            Case 149
                Quote = "I have not failed. I've just found 10,000 ways that won't work. - Thomas A. Edison"
                Exit Select
            Case 150
                Quote = "The secret to getting ahead is getting started.- Mark Twain"
                Exit Select
            Case 151
                Quote = "The best way to predict the future is to create it. - Peter Ducker"
                Exit Select
            Case 152
                Quote = "Go Big, or Go Home. - Eliza Dushku"
                Exit Select
            Case 153
                Quote = "Quality means doing it right when no one is looking. - Henry Ford"
                Exit Select
            Case 154
                Quote = "Success is not final, failure is not fatal: it is the courage to continue that counts. - Winston Churchill"
                Exit Select
            Case 155
                Quote = "Wherever you go, go with all your heart. - Confucius"
                Exit Select
            Case 156
                Quote = "If you are going through hell, keep going. - Winston Churchill"
                Exit Select
            Case 157
                Quote = "If you aren't making waves, you aren't kicking hard enough."
                Exit Select
            Case 158
                Quote = "Find what you love and let it kill you. - Charles Bukowski"
                Exit Select
            Case 159
                Quote = "The true entrepreneur is a doer, not a dreamer. - Nolan Bushnell"
                Exit Select
            Case 160
                Quote = "When you have exhausted all possibilities, remember this - you haven't. - Thomas A. Edison"
                Exit Select
            Case 161
                Quote = "The way to get started is to quit talking and start doing. - Walt Disney"
                Exit Select
            Case 162
                Quote = "Ideas are easy. Implementation is hard. - Guy Kawasaki"
                Exit Select
            Case 163
                Quote = "It's hard to beat a person who never gives up. - Babe Ruth"
                Exit Select
            Case 164
                Quote = "If everything seems under control, you're not going fast enough. - Mario Andretti"
                Exit Select
            Case 165
                Quote = "Always deliver more than expected. - Larry Page"
                Exit Select
            Case 166
                Quote = "Failure is simply an opportunity to begin again, this time more intelligently. - Henry Ford"
                Exit Select
            Case 167
                Quote = "Always wake up with a smile knowing that today you are going to have fun accomplishing what others are too afraid to do."
                Exit Select
            Case 168
                Quote = "When you find an idea that you just can't stop thinking about, that's probably a good one to pursue.- Josh James"
                Exit Select
            Case 169
                Quote = "Forget about your competitors, just focus on your customers. - Jack Ma"
                Exit Select
            Case 170
                Quote = "Stay hungry, Stay foolish. - Steve Jobs"
                Exit Select
            Case 171
                Quote = "High expectations are the key to everything. - Sam Walton"
                Exit Select
            Case 172
                Quote = "A lot of times, people don't know what they want until you show it to them. - Steve Jobs"
                Exit Select
            Case 173
                Quote = "Don't worry about failure; you only have to be right once. - Drew Houston"
                Exit Select
            Case 174
                Quote = "A pessimist sees the difficulty in every opportunity; an optimist sees the opportunity in every difficulty. - Winston Churchill"
                Exit Select
            Case 175
                Quote = "All things are difficult before they are easy. - Thomas Fuller"
                Exit Select
            Case 176
                Quote = "Life is 10% what happens to you and 90% how you react to it. - Charles R. Swindoll"
                Exit Select
            Case 177
                Quote = "Either you run the day or the day runs you. - Jim Rohn"
                Exit Select
            Case 178
                Quote = "If opportunity doesn't knock, build a door. - Milton Berle"
                Exit Select
            Case 179
                Quote = "You cannot have a positive life and a negative mind. - Joyce Meyer"
                Exit Select
            Case 180
                Quote = "Always turn a negative situation into a positive situation. - Michael Jordan"
                Exit Select
            Case 181
                Quote = "We are all here for some special reason. Stop being a prisoner of your past. Become the architect of your future. - Robin S. Sharma"
                Exit Select
            Case 182
                Quote = "The best revenge is massive success. - Frank Sinatra"
                Exit Select
            Case 183
                Quote = "If we're growing, we're always going to be out of our comfort zone. - John C. Maxwell"
                Exit Select
            Case 184
                Quote = "Hope is a waking dream. - Aristotle"
                Exit Select
            Case 185
                Quote = "The difference between a successful person and others is not a lack of strength, not a lack of knowledge, but rather a lack in will. - Vince Lombardi"
                Exit Select
            Case 186
                Quote = "There are always flowers for those who want to see them. - Henri Matisse"
                Exit Select
            Case 187
                Quote = "Even if you are on the right track, you'll get run over if you just sit there. - Will Rogers"
                Exit Select
            Case 188
                Quote = "Nurture your mind with great thoughts, for you will never go any higher than you think. - Benjamin Disraeli"
                Exit Select
            Case 189
                Quote = "I am the greatest, I said that even before I knew I was. - Muhammad Ali"
                Exit Select
            Case 190
                Quote = "The people who are crazy enough to think they can change the world are the ones who do. - Steve Jobs"
                Exit Select
            Case 191
                Quote = "The only way to do great work is to love what you do. If you haven't found it yet, keep looking. Don't settle. - Steve Jobs"
                Exit Select
            Case 192
                Quote = "The journey is the reward. - Steve Jobs"
                Exit Select
            Case 193
                Quote = "Focusing is about saying No. - Steve Jobs"
                Exit Select
            Case 194
                Quote = "Deciding what not to do is as important as deciding what to do. - Steve Jobs"
                Exit Select
            Case 195
                Quote = "Let's go invent tomorrow instead of worrying about what happened yesterday. - Steve Jobs"
                Exit Select
            Case 196
                Quote = "Why join the navy if you can be a pirate? - Steve Jobs"
                Exit Select
            Case 197
                Quote = "If you focus on success, you'll have stress. But if you pursue excellence, success will be guaranteed. - Deepok Chopra"
                Exit Select
            Case 198
                Quote = "In order to succeed, your desire for success should be greater than your fear of failure. - Bill Cosby"
                Exit Select
            Case 199
                Quote = "Successful people have libraries. The rest have big screen TVs. - Jim Rohn"
                Exit Select
            Case 200
                Quote = "Action is the foundational key to all success. - Pablo Picasso"
                Exit Select
            Case 201
                Quote = "It is not where you start but how high you aim that matters for success. - Nelson Mandela"
                Exit Select
            Case 202
                Quote = "Success is nothing more than a few simple disciplines, practiced every day. - Jim Rohn"
                Exit Select
            Case 203
                Quote = "Success is doing ordinary things extraordinarily well. - Jim Rohn"
                Exit Select
            Case 204
                Quote = "Successful people don't have fewer problems. They have determined that nothing will stop them from going forward. - Ben Carson"
                Exit Select
            Case 205
                Quote = "You only live once, but if you do it right, once is enough. - Mae West"
                Exit Select
            Case 206
                Quote = "To live is the rarest thing in the world. Most people exist, that is all."
                Exit Select
            Case 207
                Quote = "What is the point of being alive if you don't at least try to do something remarkable? - John Green"
                Exit Select
            Case 208
                Quote = "The two most important days in your life are the day you are born and the day you find out why. - Mark Twain"
                Exit Select
            Case 209
                Quote = "Do not pray for an easy life, pray for the strength to endure a difficult one. - Bruce Lee"
                Exit Select
            Case 210
                Quote = "Be happy, but never satisfied. - Bruce Lee"
                Exit Select
            Case 211
                Quote = "Your ""I CAN"" is more important than you IQ. - Robin S. Sharma"
                Exit Select
            Case 212
                Quote = "Small daily improvements over time create stunning results. - Robin S. Sharma"
                Exit Select
            Case 213
                Quote = "It's better to be a lion for a day than a sheep all your life. - Elizabeth Kenny"
                Exit Select
            Case 214
                Quote = "Believe you can and you're halfway there. - Theodore Roosevelt"
                Exit Select
            Case 215
                Quote = "Anyone can give up; it is the easiest thing in the world to do. But to hold it together when everyone would expect you to fall apart, now that is true strength. - Chris Bradford"
                Exit Select
            Case 216
                Quote = "If not now, when? - Eckhart Tolle"
                Exit Select
            Case 217
                Quote = "This, too, will pass. - Eckhart Tolle"
                Exit Select
            Case 218
                Quote = "If you can't fly then run, if you can't run then walk, if you can't walk then crawl, but whatever you do you have to keep moving forward. - Martin Luther King Jr."
                Exit Select
            Case 219
                Quote = "I am a slow walker, but I never walk back. - Abraham Lincoln"
                Exit Select
            Case 220
                Quote = "Courage is not having the strength to go on; it is going on when you don't have the strength. - Theodore Roosevelt"
                Exit Select
            Case 221
                Quote = "With ordinary talent and extraordinary perseverance, all things are attainable. - Thomas Fowell Buxton"
                Exit Select
            Case 222
                Quote = "I could either watch it happen or be a part of it. - Elon Musk"
                Exit Select
            Case 223
                Quote = "We are all in the gutter, but some of us looking at the stars. - Oscar Wilde"
                Exit Select
            Case 224
                Quote = "A good friend will always stab you in the front. - Oscar Wilde"
                Exit Select
            Case 225
                Quote = "Never try to teach a pig to sing; it wastes your time and it annoys the pig. - Robert A. Heinlein"
                Exit Select
            Case 226
                Quote = "I need a six month vacation twice a year. - Anonymous"
                Exit Select
            Case 227
                Quote = "The best things in life are free. The second best are very expensive. - Coco Chanel"
                Exit Select
            Case 228
                Quote = "Dear Karma, I have a list of people that you missed."
                Exit Select
            Case 229
                Quote = "If at first you don't succeed, destroy all evidence that you tried. - Steven Wright"
                Exit Select
            Case 230
                Quote = "The trouble with being in the rat race is that even if you win, you're still a rat. - Lily Tomlin"
                Exit Select
            Case 231
                Quote = "The road to success is dotted with many tempting parking spaces. - Will Rogers"
                Exit Select
            Case 232
                Quote = "Always listen To your heart, because even though it's on your left side, it’s always right. - Nicholas Sparks"
                Exit Select
            Case 233
                Quote = "The more anger towards the past you carry in your heart, the less capable you are of loving in the present. - Barbara De Angelis"
                Exit Select
            Case 234
                Quote = "Life is never boring, but some people choose to be bored. － Wayne Dyer"
                Exit Select
            Case 235
                Quote = "It is not because things are difficult that we do not dare; it is because we do not dare that they are difficult.  － Seneca the Younger"
                Exit Select
            Case 236
                Quote = "I have lived my life according to this principle: If I'm afraid of it, then I must do it. - Erica Jong"
                Exit Select
            Case 237
                Quote = "The first recipe for happiness is: Avoid too lengthy meditations on the past. - Andre Maurois"
                Exit Select
            Case 238
                Quote = "You cannot find peace by avoiding life. - Virginia Woolf"
                Exit Select
            Case 239
                Quote = "Your big opportunity may be right where you are now. - Napoleon Hill"
                Exit Select
            Case 240
                Quote = "Your level of belief in yourself will inevitably manifest itself in whatever you do. - Les Brown"
                Exit Select
            Case 241
                Quote = "We learn wisdom from failure much more than success. We often discover what we will do, by finding out what we will not do. - Samuel Smiles"
                Exit Select
            Case 242
                Quote = "If you change the way you look at things, the things you look at change. - Dr. Wayne Dyer"
                Exit Select
            Case 243
                Quote = "Our minds can shape the way a thing will be because we act according to our expectations. - Federico Fellini"
                Exit Select
            Case 244
                Quote = "What you become is far more important than what you get. What you get will be influenced by what you become. - Jim Rohn"
                Exit Select
            Case 245
                Quote = "We can do anything we want to do if we stick to it long enough. - Helen Keller"
                Exit Select
            Case 246
                Quote = "The fool needs company, the wise solitude. -Unknown"
                Exit Select
            Case 247
                Quote = "A constant struggle, a ceaseless battle to bring success from inhospitable surroundings, is the price of all great achievements. -Marden"
                Exit Select
            Case 248
                Quote = "The universe is full of magical things patiently waiting for our wits to grow sharper. -Eden Phillpotts"
                Exit Select
            Case 249
                Quote = "You must do the thing you think you cannot do. -Eleanor Roosevelt"
                Exit Select
            Case 250
                Quote = "When your desires are strong enough you will appear to possess superhuman powers to achieve. -Napoleon Hill"
                Exit Select
            Case 251
                Quote = "You are the only person on earth who can use your ability. -Zig Ziglar"
                Exit Select
            Case 252
                Quote = "Remember, no effort that we make to attain something beautiful is ever lost. -Helen Keller"
                Exit Select
            Case 253
                Quote = "Becoming a star may not be your destiny, but being the best that you can be is a goal that you can set for yourselves. -Bryan Lindsay"
                Exit Select
            Case 254
                Quote = "A soul without a high aim is like a ship without a rudder. -Eileen Caddy"
                Exit Select
            Case 255
                Quote = "Do not anticipate trouble, or worry about what may never happen. Keep in the sunlight. -Benjamin Franklin"
                Exit Select
            Case 256
                Quote = "To lose patience is to lose the battle. -Mahatma Gandhi"
                Exit Select
            Case 257
                Quote = "Kindness is a hard thing to give away. It keeps coming back to the giver. -Ralph Scott"
                Exit Select
            Case 258
                Quote = "Give the world the best that you have, and the best will come back to you. -Madeline Bridges"
                Exit Select
            Case 259
                Quote = "Opportunity is missed by most people because it is dressed in overalls and looks like work. -Thomas Edison"
                Exit Select
            Case 260
                Quote = "The secret of discipline is motivation. When a man is sufficiently motivated, discipline will take care of itself. -Alexander Paterson"
                Exit Select
            Case 261
                Quote = "Knowing others is intelligence; knowing yourself is true wisdom. Mastering others is strength, mastering yourself is true power. -Lao-Tzu"
                Exit Select
            Case 262
                Quote = "Failure doesn't mean you are a failure...it just means you haven't succeeded yet. -Robert Schuller"
                Exit Select
            Case 263
                Quote = "Everything that happens happens as it should, and if you observe carefully, you will find this to be so. -Marcus Antoninus"
                Exit Select
            Case 264
                Quote = "It takes a great man to give sound advice tactfully, but a greater to accept it graciously. -J.C. Macaulay"
                Exit Select
            Case 265
                Quote = "Your goal should be out of reach but not out of sight. -Anita DeFrantz"
                Exit Select
            Case 266
                Quote = "You don't become enormously successful without encountering and overcoming a number of extremely challenging problems. -Mark Victor Hansen"
                Exit Select
            Case 267
                Quote = "Start by doing what's necessary, then what's possible and suddenly you are doing the impossible. -St. Francis of Assisi"
                Exit Select
            Case 268
                Quote = "Mistakes and errors are the discipline through which we advance. -William Ellery Channing"
                Exit Select
            Case 269
                Quote = "SUCCESS IS NO ACCIDENT. It is hard work, perseverance, learning, studying, sacrifice and most of all, love of what you are doing. -Pele"
                Exit Select
            Case 270
                Quote = "Today is the first blank page of a 365 page book. Make it a good one."
                Exit Select
            Case 271
                Quote = "It's a terrible thing, I think, in life to wait until you're ready, I have this feeling now that actually no one is ever ready to do anything. There is almost no such thing as ready. There is only now. And you may as well do it now. Generally speaking, now is as good a time as any. -Hugh Laurie"
                Exit Select
            Case 272
                Quote = "Ever loved someone so much, you would do anything for them? Yeah, well make that someone yourself and do whatever the hell you want. -Harvey Specter"
                Exit Select
            Case 273
                Quote = "Sometimes the people with the greatest potential often take the longest to find their path because their sensitivity is a double edged sword - it lives at the heart of their brilliance, but it also makes them more susceptible to life's pains. Good thing we aren't being penalized for handing in our purpose late. The soul doesn't know a thing about deadlines. -Jeff Brown"
                Exit Select
            Case 274
                Quote = "Sadly, the only way some people will learn to appreciate you is by losing you."
                Exit Select
            Case 275
                Quote = "Nobody is superior, nobody is inferior, but nobody is equal either. People are simply unique, incomparable. YOU ARE YOU. I AM I. -osho"
                Exit Select
            Case 276
                Quote = "Follow your heart, but take your brain with you."
                Exit Select
            Case 277
                Quote = "Why do people hide love and express hatred so openly?"
                Exit Select
            Case 278
                Quote = "One of the best lessons you can learn in life is to master how to remain calm."
                Exit Select
            Case 279
                Quote = "I don't think inside the box, I don't think outside the box either. I don't even know where the box is."
                Exit Select
            Case 280
                Quote = "So many people worry about their physical appearance and material possessions, that they completely disregard their shitty personality."
                Exit Select
            Case 281
                Quote = "I've only met about 3 or 4 people That understand me. Everyone else assumes I'm either angry, sarcastic or just an asshole."
                Exit Select
            Case 282
                Quote = "Tell me not to do something and I will do it twice and take pictures."
                Exit Select
            Case 283
                Quote = "If you can't handle stress, you can't handle success."
                Exit Select
            Case 284
                Quote = "It's okay to be a little fucked up in the head, we all are. It's only when you're fucked up in the heart that makes you a piece of shit."
                Exit Select
            Case 285
                Quote = "Don't quit. You're already in pain; you're already hurt. Get a reward from it."
                Exit Select
            Case 286
                Quote = "One person can make a difference, and everyone should try. - JFK"
                Exit Select
            Case 287
                Quote = "If it excites you and scares you at the same time, it might be a good thing to try."
                Exit Select
            Case 288
                Quote = "I still love the people I've loved, even if I cross the street to avoid them. -Uma Thurman"
                Exit Select
            Case 289
                Quote = "Love is not a reason to tolerate disrespect."
                Exit Select
            Case 290
                Quote = "Be the Reason Someone Believes in the Goodness of People. -Lori Deschene"
                Exit Select
            Case 291
                Quote = "Relax and trust the timing of your life. You will figure out your career. You will find the right relationship. You will become the person you always wanted to be. Just don't forget to appreciate who you are now. - Ruben Chavez"
                Exit Select
            Case 292
                Quote = "You have to get to a point where your mood doesn't shift based on the insignificant actions of someone else."
                Exit Select
            Case 293
                Quote = "Nothing compares to the stomach ache you get from laughing with your best friends."
                Exit Select
            Case 294
                Quote = "I like my bed more than I like most people."
                Exit Select
            Case 295
                Quote = "Life is a soup and I'm a fucking fork."
                Exit Select
            Case 296
                Quote = "Never assume that loud is strong and quiet is weak."
                Exit Select
            Case 297
                Quote = "Better an oops than a what if."
                Exit Select
            Case 298
                Quote = "Stay low key. Not everyone needs to know everything about you."
                Exit Select
            Case 299
                Quote = "When you see something beautiful in someone, tell them. It may take seconds to say, but for them, it could last a lifetime."
                Exit Select
            Case 300
                Quote = "It is easier to build strong children than to repair broken adults.. - F. Douglass"
                Exit Select
            Case 301
                Quote = "I am currently under construction. Thank you for your patience."
                Exit Select
            Case 302
                Quote = "Three C's in life: Choices, Chances, Changes. You must make a choice to take a chance or your life will never change."
                Exit Select
            Case 303
                Quote = "Remember why you started."
                Exit Select
            Case 304
                Quote = "Real love doesn't meet you at your best. It meets you in your mess. - J.S. PARK"
                Exit Select
            Case 305
                Quote = "You can't always have a good day but you can always face a bad day with a good attitude."
                Exit Select
            Case 306
                Quote = "People who repeatedly attack your confidence and self-esteem are quite aware of your potential, even if you are not. - Wayne Gerard Trotman"
                Exit Select
            Case 307
                Quote = "I'm a very patient person and I give plenty of second chances, but I'm not a saint, I have my limits."
                Exit Select
            Case 308
                Quote = "Not everyday is a good day, live anyway. Not everyone will tell you the truth, be honest anyway. Not all you love will love you back, love anyway. Not all deals are fair, play fair anyway."
                Exit Select
            Case 309
                Quote = "People say a lot. So, I watch what they do. -via Quotes 'nd Notes"
                Exit Select
            Case 310
                Quote = "Choose people who choose you."
                Exit Select
            Case 311
                Quote = "I stopped explaining myself when I realised people only understand from their level of perception."
                Exit Select
            Case 312
                Quote = "Believe me there is somebody out there who will fall in love with your kind of crazy."
                Exit Select
            Case 313
                Quote = "Train your mind to see the good in every situation."
                Exit Select
            Case 314
                Quote = "Let them judge you. Let them misunderstand you. Let them gossip about you. Their opinions aren't your problem. You stay kind, committed to love, and free in your authenticity. No matter what they do or say, don't you dare doubt your worth or the beauty of your truth. Just keep on shining like you do. - Scott Stabile"
                Exit Select
            Case 315
                Quote = "You can meet somebody tomorrow who has better intentions for you than someone you've known forever. Time means nothing, character does."
                Exit Select
            Case 316
                Quote = "Everything is temporary."
                Exit Select
            Case 317
                Quote = "I notice everything, but I keep my mouth shut."
                Exit Select
            Case 318
                Quote = "When it hurts - observe. Life is trying to teach you something. -Anita Krizzan"
                Exit Select
            Case 319
                Quote = "Don't underestimate me because I'm quiet. I know more than I say, think more than I speak and observe more than you know. - Michaela Chung"
                Exit Select
            Case 320
                Quote = "If some one corrects you, and you feel offended, then you have an ego problem. - Nouman Ali Khan"
                Exit Select
            Case 321
                Quote = "If you're helping someone and expecting something in return, you're doing business not kindness."
                Exit Select
            Case 322
                Quote = "Life is like a book. Some chapters are sad, some are happy, and some are exciting, but if you never turn the page, you will never know what the next chapter has in store for you."
                Exit Select
            Case 323
                Quote = "I don't care how long it takes me, I'm going somewhere beautiful."
                Exit Select
            Case 324
                Quote = "There's a time to be nice, and there's a time to say, ""I've had enough of your bullshit."""
                Exit Select
            Case 325
                Quote = "Life tip: when nothing goes right, go to sleep."
                Exit Select
            Case 326
                Quote = "Always Love your friends from your Heart not from your mood or need."
                Exit Select
            Case 327
                Quote = "The soul always knows what to do to heal itself, The challenge is to silence the mind. - Caroline Myss"
                Exit Select
            Case 328
                Quote = "Never beg someone to be in your life. If you text, call, visit and still get ignored, walk away. It's called 'SELF-RESPECT'. - Steve Wentworth"
                Exit Select
            Case 329
                Quote = "I love people who are direct even if you have a difference of opinions with them, at least you know where they stand and they don't play games."
                Exit Select
            Case 330
                Quote = "At age 20, we worry about what others think of us. At age 40, we don't care what they think of us. At age 60, we discover they haven't been thinking of us at all. - Ann Landers"
                Exit Select
            Case 331
                Quote = "Always remember someone's effort is a reflection of their interest in you."
                Exit Select
            Case 332
                Quote = "Talk with people who make you see the world differently."
                Exit Select
            Case 333
                Quote = "To all the people who are loving and kind to me. Thank you for the sunshine you bring into my life. - Brigitte Nicole"
                Exit Select
            Case 334
                Quote = "When you're wrong admit it. When you're right, be quiet. - Ogden Nash"
                Exit Select
            Case 335
                Quote = "Some of you secretly don't like me, but I secretly know and don't give a shit."
                Exit Select
            Case 336
                Quote = "The biggest communication problem is we do not listen to understand. We listen to reply."
                Exit Select
            Case 337
                Quote = "My words will either attract a strong mind or offend a weak one."
                Exit Select
            Case 338
                Quote = "Loneliness is dangerous. It's addicting. Once you see how peaceful it is, you don't want to deal with people. - Hedonist Poet"
                Exit Select
            Case 339
                Quote = "You need you, more than you need them. Trust me."
                Exit Select
            Case 340
                Quote = "If there's even a slight chance at getting something that will make you happy, RISK IT. Life's too short and happiness is too rare. - A.R. Lucas"
                Exit Select
            Case 341
                Quote = """too busy"" is a myth. People make time for the things that are really important to them."
                Exit Select
            Case 342
                Quote = "Some people are just beautifully wrapped boxes of bullshit."
                Exit Select
            Case 343
                Quote = "Happiness is … meeting an old friend after a long time and feeling that nothing has changed."
                Exit Select
            Case 344
                Quote = "When something good happens, travel to celebrate. If something bad happens, travel to forget it. If nothing happens, travel to make something happen."
                Exit Select
            Case 345
                Quote = "Strangers can become best friends just as easy as best friends can become strangers."
                Exit Select
            Case 346
                Quote = "The most dangerous heart disease: strong memory. - Nizar Qabbani"
                Exit Select
            Case 347
                Quote = "I am a very private person, yet I am an open book. If you do not ask… I will not tell."
                Exit Select
            Case 348
                Quote = "Don't be so quick to believe what you hear, because lies spread quicker than the truth."
                Exit Select
            Case 349
                Quote = "I no longer look for the good in people, I search for the real… because while good is often dressed in fake clothing, real is naked and proud no matter the scars. - Chishala Lishomwa"
                Exit Select
            Case 350
                Quote = "I like people I can have comfortable silences with."
                Exit Select
            Case 351
                Quote = "Technology has connected the world and disconnected all the humans."
                Exit Select
            Case 352
                Quote = "You don't have to be positive all the time. It's perfectly okay to feel sad, angry, annoyed, frustrated, scared, or anxious. Having feelings doesn't make you 'negative person'. It makes you human. - Lori Deschene"
                Exit Select
            Case 353
                Quote = "Sometimes you gotta remember, everyone wasn't raised like you."
                Exit Select
            Case 354
                Quote = "Half the world is starving and the other half is trying to lose weight. - Roseanne Barr"
                Exit Select
            Case 355
                Quote = "We all need someone to talk to, someone who listens, someone who understands."
                Exit Select
            Case 356
                Quote = "I have no time to battle egos and small minds."
                Exit Select
            Case 357
                Quote = "Sometimes… when you give a fuck. That fuck, fucks you up."
                Exit Select
            Case 358
                Quote = "No reason to stay, is a good reason to go."
                Exit Select
            Case 359
                Quote = "Smile. No one cares how you feel."
                Exit Select
            Case 360
                Quote = "I'm becoming more silent these days. I'm speaking less and less in public. But my eyes, god damn, my eyes see everything."
                Exit Select
            Case 361
                Quote = "Magic happens when you don't give up, even though you want to. The universe always falls in love with a stubborn heart. - JmStorm"
                Exit Select
            Case 362
                Quote = "Someday we will find what we are looking for. Or maybe we won't, Maybe we will find something much greater than that."
                Exit Select
            Case 363
                Quote = "Don't be the reason someone feels insecure. Be the reason someone feels seen, heard, and supported by the whole universe. - Cleo Wade"
                Exit Select
            Case 364
                Quote = "The world is going to judge you no matter what you do, so live your life the way you fucking want to."
                Exit Select
            Case 365
                Quote = "I love the 3 am version of people. Vulnerable. Honest. Real."
                Exit Select
            Case 366
                Quote = "I'm Cold As Ice. But In The Right Hands, I'll Melt."
                Exit Select
            Case 367
                Quote = "Everything comes to you. In the right moment. Be patient. Be grateful."
                Exit Select
            Case 368
                Quote = "Hope but never expect. Look forward but never wait."
                Exit Select
            Case 369
                Quote = "When you start looking at people's hearts instead of their faces, life becomes clear."
                Exit Select
            Case 370
                Quote = "Between what is said and not meant, and what is meant and not said, most of love is lost. - Kahlil Gibran"
                Exit Select
            Case 371
                Quote = "Patience is when you're supposed to get mad, but you choose to understand."
                Exit Select
            Case 372
                Quote = "We meet everyone for a reason. Either they're a blessing or a lesson."
                Exit Select
            Case 373
                Quote = "Fake people have an image to maintain. Real people just don't give it a fuck."
                Exit Select
            Case 374
                Quote = "Be alone. Eat alone, take yourself on dates, sleep alone. In the midst of this you will learn about yourself. You will grow, you will figure out what inspires you, you will curate your own dreams, your own beliefs, your own stunning clarity, and when you do meet the person who makes your cells dance, you will be sure of it, because you are sure of yourself. - Bianca Sparacino"
                Exit Select
            Case 375
                Quote = "You don't always need a logical reason for doing everything in your life. Do it because you want to, because it's fun, because it makes you HAPPY."
                Exit Select
            Case 376
                Quote = "Don't worry if you're not where you want to be. Great things take time."
                Exit Select
            Case 377
                Quote = "Say what you feel. It's not being rude, It's called being real."
                Exit Select
            Case 378
                Quote = "I don’t have a short temper, I just have a quick reaction to bullshit."
                Exit Select
            Case 379
                Quote = "Don't climb mountains so that people can see you. Climb mountains so that you can see the world."
                Exit Select
            Case 380
                Quote = "The problem with the world is that the intelligent people are full of doubts, while the stupid ones are full of confidence. - Charles Bukowski"
                Exit Select
            Case 381
                Quote = "When you die they won't remember your car or house. They will remember who you were. Be a good human, not a good materialist."
                Exit Select
            Case 382
                Quote = "I love it when someone's laugh is funnier than the joke."
                Exit Select
            Case 383
                Quote = "You never know what people are going through and sometimes the people with the biggest smiles are struggling the most, so be kind."
                Exit Select
            Case 384
                Quote = "Don't confuse my personality with my attitude… My personality is who I am. My attitude depends on who you are…"
                Exit Select
            Case 385
                Quote = "We mature with the damage, not with the years."
                Exit Select
            Case 386
                Quote = "Who are you?' 'Demon to some. Angel to others.'"
                Exit Select
            Case 387
                Quote = "Some people think I'm unhappy, but I'm not. I just appreciate silence in a world that never stops talking."
                Exit Select
            Case 388
                Quote = "If everybody likes you, you have a serious problem."
                Exit Select
            Case 389
                Quote = "Don't allow someone to treat you poorly just because you love them."
                Exit Select
            Case 390
                Quote = "Behind every successful woman is Herself."
                Exit Select
            Case 391
                Quote = "I was quiet, but I was not blind. - Jane Austen"
                Exit Select
            Case 392
                Quote = "So, if you are too tired to speak, sit next to me, because I, too, am fluent in silence. - R. Arnold"
                Exit Select
            Case 393
                Quote = "Don't study me.. you won't graduate."
                Exit Select
            Case 394
                Quote = "When I was 17, I used to admire people with luxuries. Now, I admire people with inner-peace."
                Exit Select
            Case 395
                Quote = "Being both soft and strong is a combination very few have mastered."
                Exit Select
            Case 396
                Quote = "Stay close to anything that makes you glad you are alive."
                Exit Select
            Case 397
                Quote = "Beautiful faces are everywhere but beautiful minds are hard to find."
                Exit Select
            Case 398
                Quote = "We are not given a good life or a bad life. We have a life. It's up to us to make it good or bad. - Ward Foley"
                Exit Select
            Case 399
                Quote = "I want to remember that no one is going to make my dreams come true for me. It is my job to get up everyday and work towards things that are deepest in my heart and to enjoy every step of the journey rather than wishing I was already where I wanted to end up."
                Exit Select
            Case 400
                Quote = "What consumes your mind, controls your life."
                Exit Select
            Case 401
                Quote = "About me: I can be mean as fuck. Sweet as candy. Cold as ice. Evil as hell. Or loyal like a soldier. It all depends on you."
                Exit Select
            Case 402
                Quote = "Best therapy sometimes is a drive and music."
                Exit Select
            Case 403
                Quote = "I hope I never get tired of the night sky, of thunderstorms, of watching cream make galaxies in my coffee. I hope I never grow to someone who can no longer see the small beautiful things."
                Exit Select
            Case 404
                Quote = "I love places that make you realize how tiny you and YOUR problems are."
                Exit Select
            Case 405
                Quote = "Sometimes people hurt you and act like you hurt them."
                Exit Select
            Case 406
                Quote = "I am a different person to different people. Annoying to one. Talented to another. Quiet to a few. Unknown to a lot. But who am I, to me? - dream-jackson"
                Exit Select
            Case 407
                Quote = "Let things come and go. The things that are meant to stay will stay."
                Exit Select
            Case 408
                Quote = "Anyone can say they care.. But watch their actions, not their words."
                Exit Select
            Case 409
                Quote = "Sometimes things have to go very wrong before they can be right."
                Exit Select
            Case 410
                Quote = "The best thing about the worst time of your life is that you get to see the true colors of everyone."
                Exit Select
            Case 411
                Quote = "The best is yet to come. Be patient."
                Exit Select
            Case 412
                Quote = "Not everyone deserves to know the real you. Let them criticize who they think you are."
                Exit Select
            Case 413
                Quote = "Our character is not defined by the battles we win or lose, but by the battles we dare to fight."
                Exit Select
            Case 414
                Quote = "You will never understand until it happens to you."
                Exit Select
            Case 415
                Quote = "Once you feel you're avoided by someone, Never disturb them again."
                Exit Select
            Case 416
                Quote = "Sometimes you need bad things to happen to inspire you to change and grow."
                Exit Select
            Case 417
                Quote = "Never fuck with someone who is not afraid to be alone. You will lose every single time."
                Exit Select
            Case 418
                Quote = "Be the same person privately, publically and personally."
                Exit Select
            Case 419
                Quote = "The broken will always be able to love harder than most. Once you've been in the dark, you learn to appreciate everything that shines."
                Exit Select
            Case 420
                Quote = "It's hard to find a friend who's cute, loving, generous, sexy, caring, and smart. My advice to all my friends. Don't lose me."
                Exit Select
            Case 421
                Quote = "Never let somebody waste your time, twice."
                Exit Select
            Case 422
                Quote = "Not giving a fuck is better than revenge."
                Exit Select
            Case 423
                Quote = "There are friends, there is family, and then there are friends that become family."
                Exit Select
            Case 424
                Quote = "Deep in my heart I know I am a loner. I have tried to blend in with the world and be sociable, but the more people I meet the more disappointed I am, so I've learned to enjoy myself, my family and a few good friends. - Steven Aitchison"
                Exit Select
            Case 425
                Quote = "Fuck your shoe collection. Show me your book collection."
                Exit Select
            Case 426
                Quote = "The ego wants quantity but the soul wants quality."
                Exit Select
            Case 427
                Quote = "The key to a woman's heart is hidden in her playlist."
                Exit Select
            Case 428
                Quote = "I've got a good heart but this mouth."
                Exit Select
            Case 429
                Quote = "And if today, all you did was hold yourself together, I'm proud of you."
                Exit Select
            Case 430
                Quote = "Fuck excuses, learn to admit when you fuck up."
                Exit Select
            Case 431
                Quote = "Care less, you'll be less stressed."
                Exit Select
            Case 432
                Quote = "Do yourself a favor and learn how to walk away. When a connection starts to fade, Learn how to let it go. When a person starts to mistreat you, learn how to move on.. To something and someone better. Don't waste your energy trying to force something that isn't meant to be.. Because the truth is.. for every one person who doesn't value you there are tons more waiting to love you better. Do better. - Reyna Biddy"
                Exit Select
            Case 433
                Quote = "They keep saying that beautiful is something a girl needs to be. But honestly? Forget that. Don't be beautiful. Be angry, be intelligent, be witty, be klutzy, be interesting, be funny, be adventurous, be crazy, be talented - there are an eternity of other things to be other than beautiful. And what is beautiful anyway but a set of letters strung together to make a word? Be your own is so much more important than anything beautiful, ever. - Nikita Gill"
                Exit Select
            Case 434
                Quote = "I'm both: introvert and extrovert. I like people, but I need to be alone. I'll go out, vibe and meet new people but it has an expiration, because I have to recharge. If I don't find the valuable alone time I need to recharge I cannot be my highest self. - Sylvester McNutt III"
                Exit Select
            Case 435
                Quote = "Don't admire people too much. They might disappoint you."
                Exit Select
            Case 436
                Quote = "Be patient. Sometimes you have to go through the worst to get to the best."
                Exit Select
            Case 437
                Quote = "A friend is someone who listens to your bullshit, tells you that it is bullshit and listens some more."
                Exit Select
            Case 438
                Quote = "Listen to many, trust few."
                Exit Select
            Case 439
                Quote = "But luxury has never appealed to me, I like simply things, books, being alone, or with somebody who understands. - Daphne du Maurier"
                Exit Select
            Case 440
                Quote = "Don't underestimate the healing power of these three things… Music. The Ocean. Stars."
                Exit Select
            Case 441
                Quote = "Sometimes you just need an adventure to cleanse the bitter taste of life from your soul."
                Exit Select
            Case 442
                Quote = "Take off the mask when you are speaking to me."
                Exit Select
            Case 443
                Quote = "If you can't be kind, be quiet."
                Exit Select
            Case 444
                Quote = "I love straightforward people. The lack of drama makes life so much easier."
                Exit Select
            Case 445
                Quote = "Never be a prisoner of your past, it was just a lesson not a life sentence."
                Exit Select
            Case 446
                Quote = "Notice the people who make an effort to stay in your life."
                Exit Select
            Case 447
                Quote = "It's the maybes that will kill you."
                Exit Select
            Case 448
                Quote = "Really is the mood for a long drive with no real destination."
                Exit Select
            Case 449
                Quote = "I'm nice as fuck. So if you see me being mean to someone, they earned that shit."
                Exit Select
            Case 450
                Quote = "I'm too insane to explain and you're too normal to understand."
                Exit Select
            Case 451
                Quote = "Classy is when you have a lot to say but you choose to remain silent in front of fools.'"
                Exit Select
            Case 452
                Quote = "Q: Do you hate people? A: I don't hate them… I just feel better when they're not around. - Charles Bukowski"
                Exit Select
            Case 453
                Quote = "Apology Accepted. Trust Denied. - Malak Tamer"
                Exit Select
            Case 454
                Quote = "If we wait until we're ready, we'll be waiting for the rest of our lives."
                Exit Select
            Case 455
                Quote = "It's amazing how 3 minutes with the wrong person feels like an eternity, yet 3 hours with the right one, feels like only a moment."
                Exit Select
            Case 456
                Quote = "I wish people could drink their words and realize how bitter they taste."
                Exit Select
            Case 457
                Quote = "Never discourage anyone who continually makes progress, no matter how slow."
                Exit Select
            Case 458
                Quote = "You will never understand the damage you did to someone until the same thing is done to you. That's why I'm here. - Karma."
                Exit Select
            Case 459
                Quote = "Don't consider my kindness as my weakness. The beast in me is sleeping, not dead."
                Exit Select
            Case 460
                Quote = "I hate when people ask me 'Why are you so quiet?' Because I am. That's how I function. I don’t ask others 'Why are you so noisy? Why do talk so much?' That's rude."
                Exit Select
            Case 461
                Quote = """You're gonna be happy"" said the life ""but first I'll make you strong."""
                Exit Select
            Case 462
                Quote = "Sometimes the hardest battle is against yourself."
                Exit Select
            Case 463
                Quote = "Call me crazy but I love to see people happy and succeeding. Life is a journey, not a competition."
                Exit Select
            Case 464
                Quote = "And those who are heartless once cared too much."
                Exit Select
            Case 465
                Quote = "The things left unsaid stay with us forever."
                Exit Select
            Case 466
                Quote = "Sometimes, I just want someone to hug me and say, ""I know it's hard. You are going to be okay. Here is chocolate and 6 million dollars."""
                Exit Select
            Case 467
                Quote = "We are often let down by the most trusted people and loved by the most unexpected ones. Some make us cry for things that we haven't done, while others ignore our faults and just see our smile. Some leave us when we need them the most, while some stay with us even when ask them to leave. The world is a mixture of people. We just need to know which hand to shake and which hand to hold! After all that's life, learning to hold on and learning to let go."
                Exit Select
            Case 468
                Quote = "I am built from every mistake I have ever made."
                Exit Select
            Case 469
                Quote = "The key to happiness is low expectations. Lower. Nope, even lower. There you go."
                Exit Select
            Case 470
                Quote = "Honestly, my goal is to build a life, and career, where I'm not constantly waiting for the weekend. I don't want to live that way, where I hate five days of the week because I hate my life and job so much, that the only relief I get is Saturday and Sunday. I want to enjoy my life, and not wish it away every week. I want each day to matter to me, in some way, even some small way. I want to like my life, all of it, not just my life on the weekend."
                Exit Select
            Case 471
                Quote = "Sometimes we have to stop being scared and just go for it. Either it will work or it won't. That's life."
                Exit Select
            Case 472
                Quote = "When you fully trust someone without any doubt, you finally get one of two results: A person for life or A lesson for life.."
                Exit Select
            Case 473
                Quote = "Never say ""That won't happen to me."" Life has a funny way of proving us wrong."
                Exit Select
            Case 474
                Quote = "Always be in love with a soul, not a face."
                Exit Select
            Case 475
                Quote = "It takes two people to build a relationship. Not one person chasing after the other."
                Exit Select
            Case 476
                Quote = "A private life is a happy life."
                Exit Select
            Case 477
                Quote = "If you want it, go for it. Take a risk. Don't always play it safe or you'll die wondering."
                Exit Select
            Case 478
                Quote = """Judge Me When you are perfect."""
                Exit Select
            Case 479
                Quote = "Admit it. You're not like the others. And that's not just okay, it's fucking beautiful."
                Exit Select
            Case 480
                Quote = "Do good for others. It will come back in unexpected ways. - Karen Salmansohn"
                Exit Select
            Case 481
                Quote = "Keep your distance from people who will never admit they are wrong and who always try to make you feel like it's your fault."
                Exit Select
            Case 482
                Quote = "If someone drunk texts you, appreciate it, they're thinking of you when they can barely think straight."
                Exit Select
            Case 483
                Quote = "The sun watches what I do, but the moon knows all my secrets. - J.M. Wonderland"
                Exit Select
            Case 484
                Quote = "I have more conversations in my head than I do in real life."
                Exit Select
            Case 485
                Quote = "If overthinking situations burned calories, I'd be dead."
                Exit Select
            Case 486
                Quote = "Sometimes you have to play the role of a fool to fool the fool who thinks they are fooling you."
                Exit Select
            Case 487
                Quote = "You see a person's true colors when you are no longer beneficial to their life."
                Exit Select
            Case 488
                Quote = "I'm not impressed by money, social status or job title. I'm impressed by the way someone treats other human beings."
                Exit Select
            Case 489
                Quote = "Let it all go. See what stays."
                Exit Select
            Case 490
                Quote = "If you want to be trusted, be honest."
                Exit Select
            Case 491
                Quote = "Being ""raised right"" doesn't mean you don't drink, party, and smoke. Being raised right is how you treat people, your manners & respect."
                Exit Select
            Case 492
                Quote = "Silence isn't empty, it's full of answers."
                Exit Select
            Case 493
                Quote = "Go for it. Whether it ends good or bad, it was an experience."
                Exit Select
            Case 494
                Quote = "When you are dead, you won't even know that you are dead. It's a pain only felt by others. Same thing when you are stupid."
                Exit Select
            Case 495
                Quote = "I don't like forced conversations, forced friendships, forced interactions. I simply do not force things. If we do not vibe, we don't vibe."
                Exit Select
            Case 496
                Quote = "The world is full of monsters with friendly faces & angels full of Scars."
                Exit Select
            Case 497
                Quote = "My favourite kind of friendship is one where there's a mutual understanding of the fact that we both have our own lives so we won't be able to talk or hang out all the time but when we talk or hang out it's like picking up right where we left off. - starksfell"
                Exit Select
            Case 498
                Quote = "Just as general note, You should eliminate any thought that there is an expectation that you do anything by any age.. You don't have to be married with kids by 25.. It's ok to be 16 and never been kissed.. There's nothing wrong with you if you haven't graduated from college by 22.. You're not a failure because you don't have your dream job at 30.. There are no rules to life. You don't get special points for achieving certain things by a deadline. Just go at your own speed. It's not a race."
                Exit Select
            Case 499
                Quote = "Actions prove who someone is, words just prove who they pretend to be."
                Exit Select
            Case 500
                Quote = "Sometimes having coffee with your best friend, is all the therapy you need."
                Exit Select
            Case 501
                Quote = "Character is how you treat someone who can do nothing for you. - Johann Wolfgang von Goethe"
                Exit Select
            Case 502
                Quote = "Sometimes the greatest adventure is simply a conversation."
                Exit Select
            Case 503
                Quote = "I am terribly sorry if you don't like my harsh honesty. But, I don't like your sugar-coated bullshit either."
                Exit Select
            Case 504
                Quote = "How he treats you is how he feels about you."
                Exit Select
            Case 505
                Quote = "Don't make excuse for mean people. You can't put flowers in an asshole and call it a vase."
                Exit Select
            Case 506
                Quote = "But darling… In the end, you have to be your own hero, because everyone is busy trying to save themselves."
                Exit Select
            Case 507
                Quote = "Tomorrow, is the first blank page of a 365 page book. Write a good one. - Brad Paisley"
                Exit Select
            Case 508
                Quote = "Travel and tell no one, live a true love story and tell no one, live happily and tell no one, people ruin beautiful things. - Kahlil Gibran"
                Exit Select
            Case 509
                Quote = "Prove yourself to yourself not others."
                Exit Select
            Case 510
                Quote = "No more expectations, just gonna go with the flow and whatever happens, happens."
                Exit Select
            Case 511
                Quote = "I wish my life had background music so I could understand what the hell is going on."
                Exit Select
            Case 512
                Quote = """why Do you only have Like 5 friends?"" Me: ""quality Not quantity."""
                Exit Select
            Case 513
                Quote = "A soulmate is someone who appreciates your level of weird."
                Exit Select
            Case 514
                Quote = "Falling in love is easy. Having sex is easier. But bumping into someone that can spark your soul, that shit is rare."
                Exit Select
            Case 515
                Quote = "Stop overthinking and quit making up problems that don’t exist."
                Exit Select
            Case 516
                Quote = "I don't get ""disappointed"" anymore, I'm just like aw again? ok lol"
                Exit Select
            Case 517
                Quote = "Sometimes, happy memories hurt the most."
                Exit Select
            Case 518
                Quote = "One bad chapter doesn't mean your story is over."
                Exit Select
            Case 519
                Quote = "Being too nice can be a dangerous thing sometimes."
                Exit Select
            Case 520
                Quote = "People will try to tell you who you are. Don't believe that shit."
                Exit Select
            Case 521
                Quote = "Stop chasing the wrong one. The right one won't run. - Alfa"
                Exit Select
            Case 522
                Quote = "The greatest gifts you can give someone is your time, your love, and your attention. - Joel Osteen."
                Exit Select
            Case 523
                Quote = "Never make fun of someone who speaks broken English. It means they know another language. - H. Jackson Brown"
                Exit Select
            Case 524
                Quote = "The problem is, you think you have time."
                Exit Select
            Case 525
                Quote = "We are all in the same game, just different levels, dealing with the same hell, just different devils."
                Exit Select
            Case 526
                Quote = "I'm the nicest rude person ever."
                Exit Select
            Case 527
                Quote = "there's a message in the way a person treats you just listen… - r.h. Sin"
                Exit Select
            Case 528
                Quote = "Appreciate your rude/blunt friend.. They're always the realist."
                Exit Select
            Case 529
                Quote = "It's okay if you don't like me, not everyone has good taste."
                Exit Select
            Case 530
                Quote = "Don't kill people with kindness because not everyone deserves your kindness. Kill people with silence, because not everyone deserves your attention."
                Exit Select
            Case 531
                Quote = "Life is so much simpler when you stop explaining yourself to people and just do what works for you."
                Exit Select
            Case 532
                Quote = "Do the right thing, even when no one is watching. It's called integrity."
                Exit Select
            Case 533
                Quote = "To all the girls who no longer believe in fairy tales or happy ending: You are the writer of this story. Chin up and straighten your crown, you're the queen of this kingdom and only you know know to rule it. - B. Devine"
                Exit Select
            Case 534
                Quote = "Remember that sometimes not getting what you want is a wonderful stroke of luck. - Dalai Lama"
                Exit Select
            Case 535
                Quote = "Three stage of life: 1. Birth 2. What the fuck is this 3. Death"
                Exit Select
            Case 536
                Quote = "My goal in 2017 is to accomplish the goals I set in 2016 which I should have done in 2015 because I made a promise in 2014 which I planned in 2013."
                Exit Select
            Case 537
                Quote = "I am strong, but I am tired."
                Exit Select
            Case 538
                Quote = """A year changes you a lot."""
                Exit Select
            Case 539
                Quote = "If you don't like me. Please don't pretend that you do. Ever."
                Exit Select
            Case 540
                Quote = "Once you figure out what respect tastes like, you will always choose it over attention. - pink"
                Exit Select
            Case 541
                Quote = "And in the end all I learned was how to be strong alone. - Highway Heart"
                Exit Select
            Case 542
                Quote = "People who show you new music are important."
                Exit Select
            Case 543
                Quote = """What's the most important thing you've done this year?"" ""SURVIVED"""
                Exit Select
            Case 544
                Quote = "I care. I always care. This is my problem."
                Exit Select
            Case 545
                Quote = "The size of your audience doesn't matter. Keep up the good work."
                Exit Select
            Case 546
                Quote = "and sometimes, just sometimes, when people say forever… they mean it."
                Exit Select
            Case 547
                Quote = "I act different around certain people. It's not because I'm fake. It's because I have a different comfort zone around certain people."
                Exit Select
            Case 548
                Quote = "Grow through what you go through."
                Exit Select
            Case 549
                Quote = "You have to die a few times before you can really live. - Charles Bukowski"
                Exit Select
            Case 550
                Quote = "Don't let people know too much about you."
                Exit Select
            Case 551
                Quote = "Be so busy improving yourself that you have no time to criticize others."
                Exit Select
            Case 552
                Quote = "So tell me, where shall I go? To the left, where nothing's right? Or to the right, where nothing's left?"
                Exit Select
            Case 553
                Quote = "IF YOU WANT TO BE POWERFUL EDUCATE YOURSELF."
                Exit Select
            Case 554
                Quote = "Know your worth, then add tax."
                Exit Select
            Case 555
                Quote = "Some people are a human version of migraine."
                Exit Select
            Case 556
                Quote = "I am definitely not the same person I was when this year started."
                Exit Select
            Case 557
                Quote = "Decide what it is you want. Write that shit down. Make a fucking plan. And… work on it. Every. Single. Day."
                Exit Select
            Case 558
                Quote = "The older I get, the more I understand that it's okay to live a life others don't understand."
                Exit Select
            Case 559
                Quote = "Sinners judging sinners for sinning differently."
                Exit Select
            Case 560
                Quote = "If you like someone, set them free. If they comeback, it means nobody liked them. Set them free again."
                Exit Select
            Case 561
                Quote = "No matter how good your heart is, eventually you have to start treating people the way they treat you…."
                Exit Select
            Case 562
                Quote = "Life has taught me that you can't control someone's loyalty. No matter how good you are to them, doesn't mean that they will treat you the same. - Trent Shelton"
                Exit Select
            Case 563
                Quote = "I just woke up one day and decided I didn't want to feel like that anymore, so I changed."
                Exit Select
            Case 564
                Quote = "A queen will always turn pain into power."
                Exit Select
            Case 565
                Quote = "Less friends, less bullshit. Keep your circle small."
                Exit Select
            Case 566
                Quote = "The nicer you are, the easier you're hurt. So, just be an asshole."
                Exit Select
            Case 567
                Quote = "I gave you $10, He gave you $20. You felt that he was better just because he gave you more. But he had $200 dollars, and all I had was $10."
                Exit Select
            Case 568
                Quote = "My favorite thing is when people remember little things I told them. Like seriously? You actually listened to me thank you."
                Exit Select
            Case 569
                Quote = "Be a good person, but don't waste time to prove it."
                Exit Select
            Case 570
                Quote = "I forgive people by forgetting about them."
                Exit Select
            Case 571
                Quote = "Who am I to judge another, when I walk imperfectly."
                Exit Select
            Case 572
                Quote = "@best friend, I don't tell you everyday how thankful I am for you. But just know, deep down inside, I am truly blessed to have you in my life and to share so many memories with you."
                Exit Select
            Case 573
                Quote = "You deserve someone who is terrified to lose you. - r.h. Sin"
                Exit Select
            Case 574
                Quote = "Life is the most difficult exam. Many people fail because they try to copy others, not realizing that everyone has a different question paper."
                Exit Select
            Case 575
                Quote = "What's coming is better than what's gone."
                Exit Select
            Case 576
                Quote = "Working hard for something we don't care about is called stress; Working hard for something we love is called passion. -Simon Sinek"
                Exit Select
            Case 577
                Quote = "In the end, we all just want someone that chooses us.. Over everyone else, under any circumstances."
                Exit Select
            Case 578
                Quote = "Never be afraid to say what you really feel."
                Exit Select
            Case 579
                Quote = "What is the difference between I like you & I love you? Beautifully answered by Buddha: When you like a flower, you just pluck it. But when you love a flower, you water it daily..! One who understands this, understands life..."
                Exit Select
            Case 580
                Quote = "I asked my heart, why I can't sleep at night? Heart replied ""because you slept In the afternoon, don't act like you're in love."""
                Exit Select
            Case 581
                Quote = "People keep telling me the right guy will come along. But I think mine got hit by a bus or something."
                Exit Select
            Case 582
                Quote = "A bird sitting on a tree is never afraid of the branch breaking, because her trust is not on the branch but on it's own wings. Always believe in yourself."
                Exit Select
            Case 583
                Quote = "Before you diagnose yourself with depression or low self-esteem, first make sure you are not, in fact, surrounded by assholes. - Sigmund Freud"
                Exit Select
            Case 584
                Quote = "I have seen ugly souls masked by pretty faces, so don't tell me beauty is solely physical."
                Exit Select
            Case 585
                Quote = "No matter how good you are, you can always be replaced."
                Exit Select
            Case 586
                Quote = "Learn How To: Get fun without drink. Talk without cell phone. Dream without drugs. Smile without selfies. Love without conditions."
                Exit Select
            Case 587
                Quote = "Soon, when all is well, you're going to look back on this period of your life and be so glad that you never gave up."
                Exit Select
            Case 588
                Quote = "Things left unsaid stay with us forever."
                Exit Select
            Case 589
                Quote = "You come home, make some tea, sit down in your armchair, and all around there's silence. Everyone decides for themselves whether that's loneliness or freedom."
                Exit Select
            Case 590
                Quote = "Grades don't measure intelligence and age doesn't define maturity."
                Exit Select
            Case 591
                Quote = "Sometimes there is no next time, no second chance, no time out. Sometimes it's now or never."
                Exit Select
            Case 592
                Quote = "The biggest lesson I've learned this year is that no one is really your friend or truly loves you until they've seen every dark shadow in side you.. and stayed."
                Exit Select
            Case 593
                Quote = "Educate yourself. When a question about a certain topic pops up, google it. Watch movies and documentaries. When something sparks your interest, read about it. Read read read. Study, learn, stimulate your brain. Don't just rely on the school system, educate that beautiful mind of yours."
                Exit Select
            Case 594
                Quote = "It's a beautiful feeling when someone tells you ""I wish I knew you earlier""."
                Exit Select
            Case 595
                Quote = "Autumn shows us how beautiful it is to let things go."
                Exit Select
            Case 596
                Quote = "When trust is broken, sorry means nothing."
                Exit Select
            Case 597
                Quote = "You are not too old and it is not too late. - R. M. Rilke"
                Exit Select
            Case 598
                Quote = "Be a better you, for you."
                Exit Select
            Case 599
                Quote = "The art of knowing is knowing what to ignore. - Rumi"
                Exit Select
            Case 600
                Quote = "We must take adventures in order to know where we truly belong."
                Exit Select
            Case 601
                Quote = "Your heart knows things that your mind can't explain."
                Exit Select
            Case 602
                Quote = "When was the last time you did something for the first time?"
                Exit Select
            Case 603
                Quote = "I've learned more from pain than I could've ever learned from pleasure."
                Exit Select
            Case 604
                Quote = "You can never be happy if you're always afraid to let go of what's comfortable, familiar. Sometimes, those are the things that hurt us."
                Exit Select
            Case 605
                Quote = "Don't educate your children to be rich. Educate them to be happy, so they know the value of things, not the price."
                Exit Select
            Case 606
                Quote = "It's better to look back on life and say: ""I can't believe I did that."" than to look back and say: ""I wish I did that."""
                Exit Select
            Case 607
                Quote = "I love you take 3 seconds to say, 3 hours to explain, and a lifetime to prove."
                Exit Select
            Case 608
                Quote = "Always help someone. You might be the only one that does."
                Exit Select
            Case 609
                Quote = "Seeing someone read a book you love is seeing a book recommend a person."
                Exit Select
            Case 610
                Quote = "Never hope for it more than you work for it."
                Exit Select
            Case 611
                Quote = "Don’t be afraid to give up the GOOD to go for the GREAT. - John D. Rockefeller"
                Exit Select
            Case 612
                Quote = "I am proud of my heart. It has been stabbed, broken and torn apart yet still beats."
                Exit Select
            Case 613
                Quote = "You can never cross the ocean unless you have the courage to lose sight of the shore."
                Exit Select
            Case 614
                Quote = "Better to have an enemy who slaps you in the face than a friend who stabs you in the back."
                Exit Select
            Case 615
                Quote = "COLLECT MOMENTS NOT THINGS."
                Exit Select
            Case 616
                Quote = "Don't let the fear of falling keep you from flying."
                Exit Select
            Case 617
                Quote = "Be stubborn about your goals, but flexible about your methods."
                Exit Select
            Case 618
                Quote = "Your are going to want to give up. Don't."
                Exit Select
            Case 619
                Quote = "It's just a bad day not a bad life."
                Exit Select
            Case 620
                Quote = "Please don't expect me to always be good and kind and loving. There are times when I will be cold and thoughtless and hard to understand."
                Exit Select
            Case 621
                Quote = "never give up, great things take time."
                Exit Select
            Case 622
                Quote = "Speak your mind even if your voice shakes. - Maggie Kuhn"
                Exit Select
            Case 623
                Quote = "I have decided to be happy, because it is good for my health."
                Exit Select
            Case 624
                Quote = "COURAGE is not the absence of fear, but rather the judgement that something else is more important than fear. -Ambrose Redmoon"
                Exit Select
            Case 625
                Quote = "A broken heart is what changes people."
                Exit Select
            Case 626
                Quote = "How beautiful a day can be when kindness touches it."
                Exit Select
            Case 627
                Quote = "Because of you, I laugh a little harder, cry a little less, and smile a lot more."
                Exit Select
            Case 628
                Quote = "You never know how strong you really are until being strong is the only choice you have."
                Exit Select
            Case 629
                Quote = "Fuck people who play with other people's feelings"
                Exit Select
            Case 630
                Quote = "Ask yourself if what you're doing today is getting you closer to where you want to be tomorrow."
                Exit Select
            Case 631
                Quote = "As we grow up, we realize it becomes less important to have more friends and more important to have real ones."
                Exit Select
            Case 632
                Quote = "I was born to make mistakes, not to fake perfection."
                Exit Select
            Case 633
                Quote = "She turned her can'ts into cans and her dreams into plans."
                Exit Select
            Case 634
                Quote = "Happiness can be found in the darkest times, if one only remembers to turn on the light."
                Exit Select
            Case 635
                Quote = "Don't rush anything, when the time is right, it will happen."
                Exit Select
            Case 636
                Quote = "Always be thankful. Life could be worse."
                Exit Select
            Case 637
                Quote = "The cost of not following your heart, is spending the rest of your life wishing you had."
                Exit Select
            Case 638
                Quote = "What comes easy, won't last. What lasts, won't come easy."
                Exit Select
            Case 639
                Quote = "I CAN AND I WILL. Watch me."
                Exit Select
            Case 640
                Quote = "Never let your fear decide your future."
                Exit Select
            Case 641
                Quote = "Be a voice, not an echo."
                Exit Select
            Case 642
                Quote = "Good people go through the most bullshit."
                Exit Select
            Case 643
                Quote = "Are you really happy or just really comfortable?"
                Exit Select
            Case 644
                Quote = "Do you ever feel like you aren't even friends with some of your friends?"
                Exit Select
            Case 645
                Quote = "You only fail when you stop trying."
                Exit Select
            Case 646
                Quote = "Sometimes someone says something really small and it just fits into this empty place in your heart."
                Exit Select
            Case 647
                Quote = "I don't want to have the world's attention. Yours is enough."
                Exit Select
            Case 648
                Quote = "I find pieces of you in every song I listen to."
                Exit Select
            Case 649
                Quote = "She believed she could, so she did."
                Exit Select
            Case 650
                Quote = "Some people make your laugh a little louder. Your smile a little brighter. And your life a little better."
                Exit Select
            Case 651
                Quote = "Our small stupid conversations mean more to me than you'll ever know."
                Exit Select
            Case 652
                Quote = "May you always do what you are afraid to do."
                Exit Select
            Case 653
                Quote = "To laugh often and much; To win the respect of intelligent people and the affection of children; To earn the appreciation of honest critics and endure the betrayal of false friends; To appreciate beauty, to find the best in others; To leave the world a bit better, whether by a healthy child, a garden patch, or a redeemed social condition; To know even one life has breathed easier because you have lived. This is to have succeeded. - Ralph Waldo Emerson "
                Exit Select
            Case 654
                Quote = "Every day may not be good, but there's something good in every day."
                Exit Select
            Case 655
                Quote = "Everything will be okay in the end. If it's not okay, it's not the end. - John Lennon"
                Exit Select
            Case 656
                Quote = "The windshield to life is big and the rear view mirror is small. Using the rear view mirror once in a while will help you not to engage in another accident, but focusing on it will make you smash into what ever is in front of you."
                Exit Select
            Case 657
                Quote = "Talent is a pursued interest. In other words, anything that you're willing to practice, you can do."
                Exit Select
            Case 658
                Quote = "No matter how slow you go, you are still lapping everybody on the couch."
                Exit Select
            Case 659
                Quote = "Ideas not coupled with action never become bigger than the brain cells they occupied."
                Exit Select
            Case 660
                Quote = "To know what you want, to understand why you're doing it, to dedicate every breath in your body to achieve… If you feel you have something to give, if you feel that your particular talent is worth developing, is worth caring for then there's nothing you can't achieve. - Kevin Spacey"
                Exit Select
            Case 661
                Quote = "Listen. I wish I could tell you it gets better. But it doesn't get better. You get better. - Joan Rivers"
                Exit Select
            Case 662
                Quote = "Work until you no longer have to introduce yourself."
                Exit Select
            Case 663
                Quote = "No one's really keeping track of how many times you screw up… So chill the fuck out."
                Exit Select
            Case 664
                Quote = "No such thing as spare time, no such thing as free time, no such thing as down time. All you got is life time. Go. - Henry Rollins"
                Exit Select
            Case 665
                Quote = "A habit cannot be tossed out the window; it must be coaxed down the stairs a step at a time. - Mark Twain"
                Exit Select
            Case 666
                Quote = "Sometimes you will never know the value of a moment until it becomes a memory. - Dr. Seuss"
                Exit Select
            Case 667
                Quote = "The less you respond to negative people, the more peaceful your life will become."
                Exit Select
            Case 668
                Quote = "Just because you don't look like somebody who think is attractive doesn't mean you aren't attractive. Flowers are pretty but so christmas lights and they look nothing like."
                Exit Select
            Case 669
                Quote = "1.01^365 = 37.8; 0.99^365 = 0.03"
                Exit Select
            Case 670
                Quote = "If you let people's perception of you dictate your behavior, you will never grow as a person."
                Exit Select
            Case 671
                Quote = "Fuck motivation. It's a fickle and unreliable little dickfuck and it isn't worth your time. Better to cultivate discipline than to rely on motivation. force yourself to do things. force yourself to get up out of bed and practice. Force yourself to work. Motivation is fleeting and it’s easy to rely on because it requires no concentrated effort to get. Motivation comes to you, and you don’t have to chase after it. Discipline is reliable, motivation is fleeting. The question isn’t how to keep yourself motivated. It’s how to train yourself to work without it."
                Exit Select
            Case 672
                Quote = "Forgive other not because they deserve forgiveness, but because you deserve peace."
                Exit Select
            Case 673
                Quote = "The whole idea of motivation is a trap. Forget motivation. Just do it. Exercise, lose weight, test your blood sugar, or whatever. Do it without motivation. And then, guess what? After you start doing the thing, that's when the motivation comes and makes it easy for you to keep on doing it. - John Maxwell"
                Exit Select
            Case 674
                Quote = "It's better to shit your pants than to die of constipation. - My dad on whether or not I should ask out a girl when I was younger."
                Exit Select
            Case 675
                Quote = "They tried to bury us. They didn't know we were seeds. -Mexican Proverb"
                Exit Select
            Case 676
                Quote = "Your focus determines your reality."
                Exit Select
            Case 677
                Quote = "The pessimist looks down, and hits his head. The optimist looks up, and loses his footing. The realist looks forward, and adjusts his path accordingly. - King Ezekiel "
                Exit Select
            Case 678
                Quote = "Success is accepting yourself, enjoying what you do, and loving how you do it."
                Exit Select
            Case 679
                Quote = "If you get tired learn to rest, Not to quit.."
                Exit Select
            Case 680
                Quote = "Success isn't owned, it's leased. And rent is due every day. - J.J. Watt"
                Exit Select
            Case 681
                Quote = "If there's no wind, row. - Latin Proverb"
                Exit Select
            Case 682
                Quote = "There is a crack in everything. That's how the light gets in. -Leonard Cohen"
                Exit Select
            Case 683
                Quote = "Don't look back. You are not going that way."
                Exit Select
            Case 684
                Quote = "The trouble with most of us is that we would rather be ruined by praise than saved by criticism. - Norman Vincent Peal"
                Exit Select
            Case 685
                Quote = "Just because my path is different doesn't mean I'm lost."
                Exit Select
            Case 686
                Quote = "You will never influence the world by trying to be like it."
                Exit Select
            Case 687
                Quote = "There is a saying: Yesterday is history, Tomorrow is mystery, But today is a gift. That is why it is called the ""present""."
                Exit Select
            Case 688
                Quote = "The best time to plant a tree is 20 years ago. The second best time is now."
                Exit Select
            Case 689
                Quote = "If video games have taught me anything, it's that if you encounter enemies then you're going the right way."
                Exit Select
            Case 690
                Quote = "Not every place you fit in is where you belong."
                Exit Select
            Case 691
                Quote = "No-one really feels self-confident deep down because it's an artificial idea. Really, people aren't that worried about what you're doing or what you're saying, so you can drift around the world relatively anonymously: you must not feel persecuted and examined. Liberate yourself from that idea that people are watching you. - Russell Brand"
                Exit Select
            Case 692
                Quote = "The world is against me. It wouldn't be fair otherwise."
                Exit Select
            Case 693
                Quote = "Give me six hours to chop down a tree and I will spend the first four sharpening the axe. - Abraham Lincoln"
                Exit Select
            Case 694
                Quote = "It doesn't make sense to hire smart people and tell them what to do; we hire smart people so they can tell us what to do. - Steve Jobs"
                Exit Select
            Case 695
                Quote = "Take the bricks your enemies throw at you and build your castle."
                Exit Select
            Case 696
                Quote = "Realize that sleeping on a futon when you're 30 is not the worst thing. You know what's worse, sleeping in a king bed next to a wife you're not really in love with but for some reason you married, and you got a couple kids, and you got a job you hate. You'll be laying there fantasizing about sleeping on a futon. There's no risk when you go after a dream. There's a tremendous amount to risk to playing it safe. - Bill Burr"
                Exit Select
            Case 697
                Quote = "If you can dream it, you can achieve it. "
                Exit Select
            Case 698
                Quote = "If it's not impossible, there must be a way to do it."
                Exit Select
            Case 699
                Quote = "Climb mountains not so the world can see you, but so you can see the world."
                Exit Select
            Case 700
                Quote = "Fear is wisdom in the face of danger. It is nothing to ashamed of."
                Exit Select
            Case 701
                Quote = "I've had a lot of worries in my life, most of which never happened. - Mark Twain"
                Exit Select
            Case 702
                Quote = "AT THE END OF THE DAY I say to myself, 'Did I make a difference?' I hope the answer is always yes. - Lenny Robinson"
                Exit Select
            Case 703
                Quote = "Be willing to walk alone. Many who started with you won't finish with you."
                Exit Select
            Case 704
                Quote = "The best people possess a feeling for beauty, the courage to take risks, the discipline to tell the truth, the capacity for sacrifice. Ironically, their virtues make them vulnerable; they are often wounded, sometimes destroyed. - Ernest Hemingway"
                Exit Select
            Case 705
                Quote = "Do something today that your future self will thank you for."
                Exit Select
            Case 706
                Quote = "I still have a long way to go. But I'm already so far from where I used to be. And I'm proud of that."
                Exit Select
            Case 707
                Quote = "Worrying is like praying for something that you don't what to happen. - Robert Downey Jr."
                Exit Select
            Case 708
                Quote = "Don't let anyone tell you you're too young to accomplish something. A BABY SHARK IS STILL A FUCKING SHARK."
                Exit Select
            Case 709
                Quote = "Your value does not decrease based on someone's inability to see your worth."
                Exit Select
            Case 710
                Quote = "The truth is, everyone is goint to hurt you. You just got to find the ones worth suffering for. - Bob Marley"
                Exit Select
            Case 711
                Quote = "Only in the darkness, you are able to see the stars. - Martin Luther King"
                Exit Select
            Case 712
                Quote = "Never tell your problems to anyone… 20% don't care and the other 80% are glad you have them. - Lou Holtz"
                Exit Select
            Case 713
                Quote = "Vision without execution is just hallucination. - Henry Ford"
                Exit Select
            Case 714
                Quote = "Some people feel the rain. Others just get wet. - Bob Marley"
                Exit Select
            Case 715
                Quote = "Not everything that can be counted counts, and not everything that counts can be counted. - Albert Einstein"
                Exit Select
            Case 716
                Quote = "A superior man is modest in his speech, but exceeds in his actions. - Confucius"
                Exit Select
            Case 717
                Quote = "What you're thinking is what you're becoming. - Muhammad Ali"
                Exit Select
            Case 718
                Quote = "Only I can change my life. No one can do it for me. - Carol Burnett"
                Exit Select
            Case 719
                Quote = "Men for the sake of getting a living forget to live. - Margaret Fuller"
                Exit Select
            Case 720
                Quote = "Education is the most powerful weapon which you can use to change the world.- Nelson Mandela"
                Exit Select
            Case 721
                Quote = "Nothing is impossible, the word itself says 'I'm possible'! - Audrey Hepburn"
                Exit Select
            Case 722
                Quote = "Do not take life too seriously. You will never get out of it alive. -Elbert Hubbard"
                Exit Select
            Case 723
                Quote = "Quality is not an act, it is a habit. - Aristotle"
                Exit Select
            Case 724
                Quote = "Our greatest weakness lies in giving up. The most certain way to succeed is always to try just one more time. - Thomas A. Edison"
                Exit Select
            Case 725
                Quote = "It always seems impossible until its done. - Nelson Mandela"
                Exit Select
            Case 726
                Quote = "Accept the challenges so that you can feel the exhilaration of victory. - George S. Patton"
                Exit Select
            Case 727
                Quote = "Setting goals is the first step in turning the invisible into the visible. - Tony Robbins"
                Exit Select
            Case 728
                Quote = "Failure will never overtake me if my determination to succeed is strong enough. - Og Mandino"
                Exit Select
            Case 729
                Quote = "A creative man is motivated by the desire to achieve, not by the desire to beat others. - Ayn Rand"
                Exit Select
            Case 730
                Quote = "Perseverance is not a long race; it is many short races one after the other. - Walter Elliot"
                Exit Select
            Case 731
                Quote = "Don't watch the clock; do what it does. Keep going. -Sam Levenson"
                Exit Select
            Case 732
                Quote = "You are never too old to set another goal or to dream a new dream."
                Exit Select
            Case 733
                Quote = "A good plan violently executed now is better than a perfect plan executed next week. - George S. Patton"
                Exit Select
            Case 734
                Quote = "We aim above the mark to hit the mark. - Ralph Waldo Emerson"
                Exit Select
            Case 735
                Quote = "Thing s do not happen. Things are made to happen. - John F. Kennedy"
                Exit Select
            Case 736
                Quote = "The ultimate aim of the ego is not to see something, but to be something. - Muhammad Iqbal"
                Exit Select
            Case 737
                Quote = "Do something wonderful, people may imitate it. - Albert Schweitzer"
                Exit Select
            Case 738
                Quote = "Deserve your dream. - Octavio Paz"
                Exit Select
            Case 739
                Quote = "When one must, one can. - Charlotte Whitton"
                Exit Select
            Case 740
                Quote = "You have to make it happen. - Denis Diderot"
                Exit Select
            Case 741
                Quote = "How bad do you want it?"
                Exit Select
            Case 742
                Quote = "Many are called but few get up. - Oliver Herford"
                Exit Select
            Case 743
                Quote = "Who seeks shall find. - Sophocles"
                Exit Select
            Case 744
                Quote = "There is nothing deep down inside us except what we have put there ourselves. - Richard Rorty"
                Exit Select
            Case 745
                Quote = "Losers quit when they're tired. Winners quit when they've won."
                Exit Select
            Case 746
                Quote = "Mistakes are proof that you are trying."
                Exit Select
            Case 747
                Quote = "You have to fight some of the bad days to earn some of the best days of your life."
                Exit Select
            Case 748
                Quote = "If you can't stop thinking about it, don't stop working for it."
                Exit Select
            Case 749
                Quote = "Practice like you've never won. Perform like you've never lost."
                Exit Select
            Case 750
                Quote = "Nothing changes if nothing changes."
                Exit Select
            Case 751
                Quote = "Only the weak give up. No one said it was fucking easy."
                Exit Select
            Case 752
                Quote = "Every accomplishment STARTS with the Decision to TRY."
                Exit Select
            Case 753
                Quote = "Your future is created by what you do today not tomorrow."
                Exit Select
            Case 754
                Quote = "Sometimes, it takes a good fall to really know where you stand."
                Exit Select
            Case 755
                Quote = "I do it because I can. I can because I want to. I want to because you said I couldn't."
                Exit Select
            Case 756
                Quote = "Being the richest man in the cemetery doesn't matter to me. Going to bed at night saying we've done something wonderful, that's what matters to me. - Steve Jobs"
                Exit Select
            Case 757
                Quote = "When you feel like quitting, think about why you started."
                Exit Select
            Case 758
                Quote = "A little progress each day adds up to big results."
                Exit Select
            Case 759
                Quote = "Don't stop when you are tired. STOP when you are DONE!"
                Exit Select
            Case 760
                Quote = "Every single person on the planet has a story. Don't judge people before you truly know them. The truth might surprise you."
                Exit Select
            Case 761
                Quote = "Don't call it a dream. Call it a plan."
                Exit Select
            Case 762
                Quote = "It is during our failures that we discover our true desire for success. - Kevin Ngo"
                Exit Select
            Case 763
                Quote = "Your life is your message to the world. Make sure its inspiring."
                Exit Select
            Case 764
                Quote = "When you're about to give up, remember those who said you're not good enough."
                Exit Select
            Case 765
                Quote = "Life is tough, my darling, but so are you. - Stephanie Bennett Henry"
                Exit Select
            Case 766
                Quote = "YESTERDAY YOU SAID TOMORROW."
                Exit Select
            Case 767
                Quote = "Difficult roads often lead to beautiful destinations."
                Exit Select
            Case 768
                Quote = "Don't tell me the sky's the limit when there are footprints on the moon."
                Exit Select
            Case 769
                Quote = "Prove them wrong."
                Exit Select
            Case 770
                Quote = "We become what we think about. - Earl Nightingale"
                Exit Select
            Case 771
                Quote = "A lot of problems in the world would disappear if we talk to each other instead of about each other."
                Exit Select
            Case 772
                Quote = "Don't tell people your dreams. Show them!"
                Exit Select
            Case 773
                Quote = "Never stop doing little things for others. Sometimes, those little things occupy the biggest part of their heart."
                Exit Select
            Case 774
                Quote = "Be who you are and say what you feel, because those who mind don't matter and those who matter don't mind. - Dr. Seuss"
                Exit Select
            Case 775
                Quote = "I can. I will. End of story."
                Exit Select
            Case 776
                Quote = "Put your heart, mind, and soul into even your smallest acts. This is the secret of success. - Swami Sivananda"
                Exit Select
            Case 777
                Quote = "Good, better, best. Never let it rest. 'Til your good is better and your better is best. -St. Jerome"
                Exit Select
            Case 778
                Quote = "Why cheat? If you're not happy just leave!"
                Exit Select
            Case 779
                Quote = "Some walks you have to take alone."
                Exit Select
            Case 780
                Quote = "Not everyone you lose is a loss."
                Exit Select
            Case 781
                Quote = "If you lose someone, but find yourself, you won."
                Exit Select
            Case 782
                Quote = "People ask me why is it so hard to trust people. I ask why is it so hard to keep a promise."
                Exit Select
            Case 783
                Quote = "Having a soulmate isn't always about love, you can also find one in a friend."
                Exit Select
            Case 784
                Quote = "There are all these moments you think you won't survive. And then you survive. - David Levithan"
                Exit Select
            Case 785
                Quote = "Judge me by the people I avoid."
                Exit Select
            Case 786
                Quote = "When your past calls… Don't answer. It has nothing new to say."
                Exit Select
            Case 787
                Quote = "Not everyone is meant to be in your future.  Some people are just passing through to teach you lessons in life."
                Exit Select
            Case 788
                Quote = "If my eyes could show my soul, everyone would cry when they saw me smile. - Kurt Cobain"
                Exit Select
            Case 789
                Quote = "If you still speak about it, you still care about it."
                Exit Select
            Case 790
                Quote = "And they lived happily ever after. Separately."
                Exit Select
            Case 791
                Quote = "I distance myself from people for a reason."
                Exit Select
            Case 792
                Quote = "You can't open up the story of my life and just go to page 738 and think you know me."
                Exit Select
            Case 793
                Quote = "Some people don't change. They just find new ways to lie."
                Exit Select
            Case 794
                Quote = "The deepest pain I ever felt was denying my own feelings to make everyone else comfortable. - Nicole Lyons"
                Exit Select
            Case 795
                Quote = "If you have the courage to make it through a lonely night with nothing but your self destructive thoughts to keep you company, darling, you have the courage to make it through anything."
                Exit Select
            Case 796
                Quote = "Always remember that behind every strong and independent woman, there are days when she was alone and helpless. There are lessons she has learnt from life and there are stories of battles and struggles which she has fought alone. Beneath the shield of confidence and strength there is a plethora of sadness and pain which she has endured. - Aarti Khurana"
                Exit Select
            Case 797
                Quote = "I'm not upset that you lied to me, I'm upset that from now on I can't believe you. - Friedrich Nietzsche."
                Exit Select
            Case 798
                Quote = "I was told I was dangerous… I asked why? They said ""Because you don't need anyone."" That's when I smiled."
                Exit Select
            Case 799
                Quote = "Softness is not weakness. It takes courage to stay delicate in a world this cruel. - Beau Taplin."
                Exit Select
            Case 800
                Quote = "Distance doesn't separate people.. Silence does.."
                Exit Select
            Case 801
                Quote = "Damaged people are dangerous. They know how to make hell feel like home."
                Exit Select
            Case 802
                Quote = "I'm not afraid of werewolves or vampires or haunted hotels, I'm afraid of what real human beings do to other real human beings. - Walter Jon Williams"
                Exit Select
            Case 803
                Quote = "3 things to keep private. 1. Love life. 2. Income. 3. Next move."
                Exit Select
            Case 804
                Quote = "I'm close to very vew people, but those few people mean everything to me."
                Exit Select
            Case 805
                Quote = "The first step of change is to become aware of your own bullshit."
                Exit Select
            Case 806
                Quote = "If you see something beautiful in someone, speak it. - Ruthie Lindsey"
                Exit Select
            Case 807
                Quote = "You're always one decision away from a totally different life."
                Exit Select
            Case 808
                Quote = "If it's destroying you, then it isn't love, my dear."
                Exit Select
            Case 809
                Quote = "The most beautiful people you will ever meet aren't always the ones who catch your eye first. No, the most beautiful people are the ones that can never be figured out. The ones you could talk with for hours and still have a millions things to ask. The people who have minds so lovely and special, you can't help but fall in love with them..."
                Exit Select
            Case 810
                Quote = "It's all about the first person you wanna tell good news to."
                Exit Select
            Case 811
                Quote = "I have a special skill of feeling too much when I shouldn't, and feeling nothing when I should."
                Exit Select
            Case 812
                Quote = "Things end. People change. And you know what? Life goes on. - Elizabeth Scott"
                Exit Select
            Case 813
                Quote = "If I'm wrong, educate me. Don't belittle me."
                Exit Select
            Case 814
                Quote = "The older I get, the more I appreciate being home doing absolutely nothing."
                Exit Select
            Case 815
                Quote = "The person you will be in 5 years is based on the books you read and the people you surround yourself with today."
                Exit Select
            Case 816
                Quote = "The things you hide in your heart… eats you alive…"
                Exit Select
            Case 817
                Quote = "Just because I'mj silent doesn't mean I agree with you. I'm just being polite. That, or I'm actually too stunned by your stupidity to even respond."
                Exit Select
            Case 818
                Quote = "I've never met a strong person with an easy past."
                Exit Select
            Case 819
                Quote = "A heart that always understands also gets tired."
                Exit Select
            Case 820
                Quote = "You will search for me in another person. I promise."
                Exit Select
            Case 821
                Quote = "Never forget the people who take time out of their day to check up on you."
                Exit Select
            Case 822
                Quote = "Shit happens. Every day. To everyone. The difference is in how people deal with it."
                Exit Select
            Case 823
                Quote = "She's stuck between who she is, who she wants to be, and who she should be."
                Exit Select
            Case 824
                Quote = "The cure for anything is salt water: sweat, tears or the sea. - Isak Dinesen"
                Exit Select
            Case 825
                Quote = "Sleep doesn't help if it's your soul that's tired."
                Exit Select
            Case 826
                Quote = "We ignored truths for temporary happiness."
                Exit Select
            Case 827
                Quote = "I have seen beauty in people who were called ugly and I've seen the devil in the most angelic faces. - Conny Cernik"
                Exit Select
            Case 828
                Quote = "If it's still in your mind, it is worth taking the risk. - Paulo Coelho, The Alchemist"
                Exit Select
            Case 829
                Quote = "I don't regret the things I did wrong. I regret the good things I did for the wrong people."
                Exit Select
            Case 830
                Quote = "People aren't ignoring you. They are busy with their lives. And the way to stop feeling ignored is to get busy with yours."
                Exit Select
            Case 831
                Quote = "Kill the part of you that believes it can't survive without someone else. - Sade Andria Zabala"
                Exit Select
            Case 832
                Quote = "oh sorry, I forgot. I only exist when you need something."
                Exit Select
            Case 833
                Quote = "When I got enough confidence, the stage was gone. When I was sure of losing, I won. When I needed people the most, they left me. When I learnt to dry my tears, I found a shoulder to cry on. And when I mastered the art of hating, somebody started loving me. - William Shakespeare"
                Exit Select
            Case 834
                Quote = "People always choose the wrong person first, and then when the right person arrives, they just stop trusting people."
                Exit Select
            Case 835
                Quote = "Someday, all the love you've given away, will find its way back to you, and it will finally stay. - Drewniverses"
                Exit Select
            Case 836
                Quote = "Life is too short to be normal. Stay weird."
                Exit Select
            Case 837
                Quote = "No one is going to stand up at your funeral and say ""She had a really expensive couch And great shoes."" Don't make life about stuff."
                Exit Select
            Case 838
                Quote = "In life, what you really want; will never come easy."
                Exit Select
            Case 839
                Quote = "When I was a kid, I wanted to be older… This shit was not what I expected."
                Exit Select
            Case 840
                Quote = "If I won the award for laziness, I would send somebody to pick it up for me."
                Exit Select
            Case 841
                Quote = "Some people are like clouds. When they go away, it's a brighter day."
                Exit Select
            Case 842
                Quote = "When nothing is going right, go left."
                Exit Select
            Case 843
                Quote = "A best friend is like a four leaf clover, hard to find, lucky to have."
                Exit Select
            Case 844
                Quote = "Maybe if we tell people the brain is an app, they'll start using it."
                Exit Select
            Case 845
                Quote = "When you wake up at 6 in the morning, you close your eyes for 5 minutes and it's already 6:45. When you're at work and it's 2:30, you close your eyes for 5 minutes and it's 2:31."
                Exit Select
            Case 846
                Quote = "If people are talking behind your back, be happy that you are the one in front."
                Exit Select
            Case 847
                Quote = "Life is not about how you survive the storm, it's about how you dance in the rain."
                Exit Select
            Case 848
                Quote = "In the morning you beg to sleep more, in the afternoon you are dying to sleep, and at night you refuse to sleep."
                Exit Select
            Case 849
                Quote = "Doing nothing is hard, you never know when you're done."
                Exit Select
            Case 850
                Quote = "I always try to cheer myself up by singing when I get sad. Most of the time, it turns out that my voice is worse than my problems."
                Exit Select
            Case 851
                Quote = "God please give me patience, if you give me strength I will just punch them in the face."
                Exit Select
            Case 852
                Quote = "Life isn't measured by the number of breaths you take, but by the number of moments that take your breath away."
                Exit Select
            Case 853
                Quote = "When my boss asked me who is the stupid one, me or him? I told him everyone knows he doesn't hire stupid people."
                Exit Select
            Case 854
                Quote = "Everybody wants to go to heaven; but nobody wants to die."
                Exit Select
            Case 855
                Quote = "Some people walk into our lives and leave footprints on our hearts. Others walk into our lives and we want to leave footprints on their face!"
                Exit Select
            Case 856
                Quote = "Who says nothing is impossible? I've been doing nothing for years."
                Exit Select
            Case 857
                Quote = "They say that love is more important than money, but have you ever tried to pay your bills with a hug?"
                Exit Select
            Case 858
                Quote = "The road to success is always under construction. - Lily Tomlin"
                Exit Select
            Case 859
                Quote = "A good speech should be like a woman's skirt: long enough to cover the subject and short enough to create interest. Winston Churchill"
                Exit Select
            Case 860
                Quote = "Never take life seriously. Nobody gets out alive anyway."
                Exit Select
            Case 861
                Quote = "Smile today, tomorrow could be worse. "
                Exit Select
            Case 862
                Quote = "The ideal man doesn't smoke, doesn't drink, doesn't do drugs, doesn't swear, doesn't get angry, doesn't exist."
                Exit Select
            Case 863
                Quote = "Relationships these days start by pressing LIKE on her photo."
                Exit Select
            Case 864
                Quote = "Sometimes when I close my eyes, I can't see."
                Exit Select
            Case 865
                Quote = "Don't cry because it's over, smile because it happened. - Dr. Seuss"
                Exit Select
            Case 866
                Quote = "You know you're in love when you can't fall asleep because reality is finally better than your dreams. - Dr. Seuss"
                Exit Select
            Case 867
                Quote = "To be yourself in a world that is constantly trying to make you something else is the greatest accomplishment. - Ralph Waldo Emerson"
                Exit Select
            Case 868
                Quote = "Whenever you find yourself on the side of the majority, it is time to pause and reflect. -Mark Twain"
                Exit Select
            Case 869
                Quote = "When you're happy you enjoy the music. When you're sad you understand the lyrics."
                Exit Select
            Case 870
                Quote = "I never make the same mistake twice. I make it like five or six times, you know, just to be sure."
                Exit Select
            Case 871
                Quote = "Only dead fish go with the flow."
                Exit Select
            Case 872
                Quote = "I hate it when people text me: ""Call Me"". I'm gonna start calling people and when they answer, I'm gonna say: ""text Me"", and hang up."
                Exit Select
            Case 873
                Quote = "My life feels like a test I didn't study for."
                Exit Select
            Case 874
                Quote = """be strong,"" I whispered to my wifi signal."
                Exit Select
            Case 875
                Quote = "Happiness is… not having to set the alarm for the next day."
                Exit Select
            Case 876
                Quote = "Don't mistake silence for weakness. Smart people don't plan big moves out loud."
                Exit Select
            Case 877
                Quote = "My soulmate is out there somewhere, pushing a pull door… I just know it."
                Exit Select
            Case 878
                Quote = "Don't give up on your dreams. Keep sleeping."
                Exit Select
            Case 879
                Quote = "We are all a little broken. But the last time I checked, broken crayons still color the same."
                Exit Select
            Case 880
                Quote = "Life is short. Smile while you still have teeth."
                Exit Select
            Case 881
                Quote = "When you're stressed, you eat ice cream, cake, chocolate and sweets. Why? Because stressed spelled backwards is desserts."
                Exit Select
            Case 882
                Quote = "Defeat is not bitter unless you swallow it. - Joe Clark"
                Exit Select
            Case 883
                Quote = "Keep the ones that heard you when you never said a word."
                Exit Select
            Case 884
                Quote = "Ten years from now, make sure you can say that you chose your life. You didn't settle for it. - Mandy Hale"
                Exit Select
            Case 885
                Quote = "Before I used to be afraid of being alone. Now, I'm afraid of having the wrong people as company."
                Exit Select
            Case 886
                Quote = "Don't try to understand everything. Sometimes it is not meant to be understood, just accepted."
                Exit Select
            Case 887
                Quote = "Everyone has a chapter they don't read out loud."
                Exit Select
            Case 888
                Quote = "It's amazing how fast someone can become a stranger."
                Exit Select
            Case 889
                Quote = "It is not the length of life, but the depth."
                Exit Select
            Case 890
                Quote = "Tears are words the heart can't say."
                Exit Select
            Case 891
                Quote = "Though I saw it coming, it still hurts."
                Exit Select
            Case 892
                Quote = "what if I fall? oh, my darling, what if you fly?"
                Exit Select
            Case 893
                Quote = "Be comfortable being uncomfortable. Because one day, all this hard work will pay off."
                Exit Select
            Case 894
                Quote = "You don't have to see the whole staircase, just take the first step. - Martin Luther King"
                Exit Select
            Case 895
                Quote = "Most of you don't want success as much as you want to sleep! – Eric Thomas"
                Exit Select
            Case 896
                Quote = "I like to think that one day you'll be an old man like me talkin' a young man's ear off explainin' to him how you took the sourest lemon that life has to offer and turned it into something resembling lemonade. - This is us"
                Exit Select
            Case 897
                Quote = "I didn't come this far to only come this far."
                Exit Select
            Case 898
                Quote = "Don't be afraid to fail. Be afraid not to try."
                Exit Select
            Case 899
                Quote = "I'M NOT HERE TO BE AVERAGE. I'M HERE TO BE AWESOME."
                Exit Select
            Case 900
                Quote = "One day, I want to honestly say, ""I made it."""
                Exit Select
        End Select

        With appCell
            .Value = begValue & quote
            .Font.Name = "Comic Sans MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub ClearPen_Click(sender As Object, e As RibbonControlEventArgs) Handles ClearPen.Click
        Dim app As Object = Globals.ThisAddIn.Application
        Dim sht As Excel.Worksheet = Nothing
        Dim shp As Excel.Shape = Nothing

        sht = app.ActiveWorkbook.ActiveSheet
        For Each shp In sht.Shapes
            If shp.Type = Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight And (shp.Name = "RedPen" Or shp.Name = "BluePen" Or shp.Name = "BlackPen") Then
                shp.Select()
                shp.Delete()
            End If
        Next
    End Sub

    Private Sub Direction1_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction1.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "1"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If

    End Sub

    Private Sub Direction2_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction2.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "2"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub Direction3_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction3.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "3"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub Direction4_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction4.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "4"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub Direction5_Click(sender As Object, e As RibbonControlEventArgs) Handles Direction5.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        Dim connectorBeg As Object
        Dim connectorTarget As Object
        Dim shpBeg As Excel.Shape
        Dim shpTarget As Excel.Shape

        Dim Left As Double = appCell.Left
        Dim Top As Double = appCell.Top
        Dim Width As Double = appCell.Width
        Dim Height As Double = appCell.Height

        Dim location As Object
        location = app.inputbox("Select a cell", "Drawing connecting boxes", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If

        Dim TargetLeft As Double = location.Left
        Dim TargetTop As Double = location.Top
        Dim TargetWidth As Double = location.Width
        Dim TargetHeight As Double = location.Height



        If (Left > TargetLeft) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top + Height + 5, Left - 10, Top + Height + 12)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add target connector
            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop - 5.6, TargetLeft + TargetWidth + 11, TargetTop - 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False

            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()


            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.8, Top - 6.5, Left - 10, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.3, TargetTop + TargetHeight + 5.8, TargetLeft + TargetWidth + 11, TargetTop + TargetHeight + 13)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left > TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left - 5.9, Top + Height, Left - 15.1, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth + 6.2, TargetTop + TargetHeight, TargetLeft + TargetWidth + 15.3, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----

        ElseIf (Left < TargetLeft) And (Top = TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width + 6.2, Top + Height, Left + Width + 15.3, Top + Height)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.9, TargetTop + TargetHeight, TargetLeft - 15.1, TargetTop + TargetHeight)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             Left + Width + 6.3, Top + Height + 5.8, Left + Width + 11, Top + Height + 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft - 5.8, TargetTop - 6.5, TargetLeft - 10, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 7, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (TargetLeft > Left) And (Top <> TargetTop) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         Left + Width + 6.3, Top - 5.6, Left + Width + 11, Top - 13)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
        (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
         TargetLeft - 5.8, TargetTop + TargetHeight + 5, TargetLeft - 10, TargetTop + TargetHeight + 12)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft - 6, TargetTop + TargetHeight - 7, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top > TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top - 6, Left + Width, Top - 14)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop + TargetHeight + 6.2, TargetLeft + TargetWidth, TargetTop + TargetHeight + 14.2)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop + TargetHeight - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        ElseIf (Left = TargetLeft) And (Top < TargetTop) Then
            connectorBeg = app.ActiveSheet.Shapes.addconnector _
           (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
            Left + Width, Top + Height + 6.2, Left + Width, Top + Height + 14.2)
            With connectorBeg
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With

            connectorTarget = app.ActiveSheet.Shapes.addconnector _
            (Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
             TargetLeft + TargetWidth, TargetTop - 6, TargetLeft + TargetWidth, TargetTop - 14)
            With connectorTarget
                .line.weight = 1
                .line.ForeColor.RGB = RGB(0, 60, 60)
                .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
            End With
            'Add box
            shpBeg = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
    (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left + Width - 6, Top + Height - 6, 12.5, 12.5)
            With shpBeg.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpBeg.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpBeg.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpBeg.Fill.Visible = False
            'Add Target box
            shpTarget = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Addtextbox _
(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, TargetLeft + TargetWidth - 6, TargetTop - 6, 12.5, 12.5)
            With shpTarget.TextFrame2
                .TextRange.Text = "5"
                .MarginBottom = 0
                .MarginTop = 0
                .MarginRight = 0
                .MarginLeft = 3.3
            End With
            With shpTarget.Line
                .ForeColor.RGB = RGB(0, 60, 60)
            End With
            With shpTarget.TextFrame2.TextRange.Font
                .Size = 10
                .Name = "Arial"
                .Fill.ForeColor.RGB = RGB(0, 60, 60)
            End With
            shpTarget.Fill.Visible = False
            'Group the textbox and the connector1
            Dim MidObj As Object() = New Object() {shpBeg.Name, connectorBeg.name}
            app.activesheet.Shapes.Range(MidObj).Group()

            'Group the textbox and the connector2
            Dim MidObj2 As Object() = New Object() {shpTarget.Name, connectorTarget.name}
            app.activesheet.Shapes.Range(MidObj2).Group()
            '--End----
        End If
    End Sub

    Private Sub hyperlinkCell_Click(sender As Object, e As RibbonControlEventArgs) Handles hyperlinkCell.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim app As Object = Globals.ThisAddIn.Application
        'Origin Sheet
        Dim strBeginningSheetName As String = app.Activesheet.name
        'Store the Original sheet & cell location
        Dim OriSheetCellAddress As String = appCell.address


        Dim location As Object
        location = app.inputbox("Select a cell", "Creating a HyperLink", Type:=8)
        On Error Resume Next
        'The user clicks the  cancel button.
        If location.address = " " Then
            Exit Sub
        End If



        'Get the Target's sheet name
        Dim strDestinationSheetName As String = location.Parent.Name


        'Add the target address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=appCell,
             Address:="", SubAddress:=strDestinationSheetName & "!" & location.address,
             TextToDisplay:=appCell.Value.ToString)
        appCell.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
        appCell.font.color = RGB(0, 0, 0)

        'Add the origin address to textbox (Creating a hyperlink)
        app.Activesheet.Hyperlinks.Add(Anchor:=location,
             Address:="", SubAddress:=strBeginningSheetName & "!" & OriSheetCellAddress,
             TextToDisplay:=location.value.ToString)
        location.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
        location.font.color = RGB(0, 0, 0)
    End Sub

    Private Sub Calendar_ToggleButton_Click(sender As Object, e As RibbonControlEventArgs) Handles Calendar_ToggleButton.Click
        Globals.ThisAddIn.TaskPane2.Visible =
    TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
    End Sub

    Private Sub LeftBrace_Click(sender As Object, e As RibbonControlEventArgs) Handles LeftBrace.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        shp = app.ActiveSheet.Shapes.Addshape _
(Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftBrace, r.Left - 17 + r.Width, r.Top, 17, r.Height)
        shp.Select()
        app.Selection.ShapeRange.Adjustments.Item(1) = 0.4
        shp.Line.ForeColor.RGB = RGB(0, 0, 0)

    End Sub

    Private Sub DownApply_Click(sender As Object, e As RibbonControlEventArgs) Handles DownApply.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim arrow As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)
        shp = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 6, r.Top, r.Left + 6, r.Height + r.Top)
        shp.Line.ForeColor.RGB = RGB(255, 0, 0)

        arrow = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 6, r.Height + r.Top, r.Left + 13, r.Top + r.Height - 10)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {shp.Name, arrow.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()

    End Sub

    Private Sub RightApply_Click(sender As Object, e As RibbonControlEventArgs) Handles RightApply.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim arrow As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)
        shp = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left, r.Top + r.Height - 4.5, r.Left + r.Width, r.Top + r.Height - 4.5)
        shp.Line.ForeColor.RGB = RGB(255, 0, 0)

        arrow = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Height + r.Top - 4.5, r.Left + r.Width - 10, r.Top + r.Height - 10)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {shp.Name, arrow.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub PMingLiU_Click(sender As Object, e As RibbonControlEventArgs) Handles PMingLiU.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "新細明體"
    End Sub

    Private Sub Arial_Click(sender As Object, e As RibbonControlEventArgs) Handles Arial.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "Arial"
    End Sub

    Private Sub BookAntiqua_Click(sender As Object, e As RibbonControlEventArgs) Handles BookAntiqua.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "Book Antiqua"
    End Sub

    Private Sub Calibri_Click(sender As Object, e As RibbonControlEventArgs) Handles Calibri.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        app.Selection.Font.Name = "Calibri"
    End Sub

    Private Sub MultiIllu_Click(sender As Object, e As RibbonControlEventArgs) Handles MultiIllu.Click
        multi()
    End Sub
    Private Sub multi()
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim shp As Excel.Shape
        Dim sht As Excel.Worksheet = app.ActiveWorkbook.ActiveSheet
        Dim arrow As Excel.Shape
        Dim connectorref As Excel.Shape = Nothing
        Dim r As Excel.Range
        Dim i As Integer = Nothing

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)
        shp = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width + 8, r.Top + 8.25, r.Left + r.Width + 8, r.Height + r.Top + 8)
        shp.Line.ForeColor.RGB = RGB(255, 0, 0)


        arrow = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width + 8, r.Height + r.Top + 8, r.Left + r.Width + 22, r.Top + r.Height + 8)
        arrow.Line.ForeColor.RGB = RGB(255, 0, 0)
        arrow.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle

        Dim MidObj() As Object
        ReDim Preserve MidObj(2)
        MidObj(0) = shp.Name
        MidObj(1) = arrow.Name
        For i = 0 To (r.Height / 16.5 - 1)
            connectorref = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Top + 8.25 + i * 16.5, r.Left + r.Width + 8, r.Top + 8.25 + i * 16.5)
            connectorref.Line.ForeColor.RGB = RGB(255, 0, 0)
            'connectorref.Name = "connectorref" & i
            ReDim Preserve MidObj(i + 2)
            MidObj(i + 2) = connectorref.Name
        Next

        'Group shps all together
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub delta_Click(sender As Object, e As RibbonControlEventArgs) Handles delta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "δ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub epsilon_Click(sender As Object, e As RibbonControlEventArgs) Handles epsilon.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ε"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub zeta_Click(sender As Object, e As RibbonControlEventArgs) Handles zeta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ζ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub eta_Click(sender As Object, e As RibbonControlEventArgs) Handles eta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "η"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub theta_Click(sender As Object, e As RibbonControlEventArgs) Handles theta.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "θ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub mu_Click(sender As Object, e As RibbonControlEventArgs) Handles mu.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "μ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub pi_Click(sender As Object, e As RibbonControlEventArgs) Handles pi.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "π"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub rho_Click(sender As Object, e As RibbonControlEventArgs) Handles rho.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ρ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub phi_Click(sender As Object, e As RibbonControlEventArgs) Handles phi.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "φ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub psi_Click(sender As Object, e As RibbonControlEventArgs) Handles psi.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ψ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub omega_Click(sender As Object, e As RibbonControlEventArgs) Handles omega.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "ω"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 14
        End With
    End Sub

    Private Sub one_Click(sender As Object, e As RibbonControlEventArgs) Handles one.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅰ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub two_Click(sender As Object, e As RibbonControlEventArgs) Handles two.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅱ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub three_Click(sender As Object, e As RibbonControlEventArgs) Handles three.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅲ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub four_Click(sender As Object, e As RibbonControlEventArgs) Handles four.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅳ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub five_Click(sender As Object, e As RibbonControlEventArgs) Handles five.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅴ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub six_Click(sender As Object, e As RibbonControlEventArgs) Handles six.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅵ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub seven_Click(sender As Object, e As RibbonControlEventArgs) Handles seven.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅶ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub eight_Click(sender As Object, e As RibbonControlEventArgs) Handles eight.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅷ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub nine_Click(sender As Object, e As RibbonControlEventArgs) Handles nine.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅸ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub ten_Click(sender As Object, e As RibbonControlEventArgs) Handles ten.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Ⅹ"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Arial Unicode MS"
            .Font.Size = 12
        End With
    End Sub

    Private Sub ColorYellow_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorYellow.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(255, 255, 0)
    End Sub

    Private Sub ColorAqua_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorAqua.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(0, 255, 255)
    End Sub

    Private Sub ColorLime_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorLime.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(211, 112, 112)
    End Sub

    Private Sub ColorSilver_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorSilver.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(255, 255, 196)
    End Sub

    Private Sub ColorFuchsia_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorFuchsia.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.color = RGB(196, 196, 255)
    End Sub

    Private Sub ColorWhite_Click(sender As Object, e As RibbonControlEventArgs) Handles ColorWhite.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim appCell As Object = app.ActiveCell

        appCell.Interior.pattern = Microsoft.Office.Core.XlConstants.xlNone
    End Sub



    Private Sub RD_Click(sender As Object, e As RibbonControlEventArgs) Handles RD.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "RD"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub btnCB_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCB.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "CB"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub TB_Click(sender As Object, e As RibbonControlEventArgs) Handles TB.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "TB"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Tick2_Click(sender As Object, e As RibbonControlEventArgs) Handles Tick2.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape
        Dim connector2 As Excel.Shape
        Dim connector3 As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 4.8, r.Top + 4.5, r.Left + 8.8, r.Top + 9.5)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.2

        connector2 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 4.7, r.Top + r.Height - 1, r.Left + 7.7, r.Top + 3)
        connector2.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector2.Line.Weight = 1.2

        connector3 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 2.3, r.Top + r.Height - 6.5, r.Left + 4.8, r.Top + r.Height - 1)
        connector3.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector3.Line.Weight = 1.2

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, connector2.Name, connector3.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub Tick3_Click(sender As Object, e As RibbonControlEventArgs) Handles Tick3.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape
        Dim connector2 As Excel.Shape
        Dim connector3 As Excel.Shape
        Dim connector4 As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 5.8, r.Top + 2.7, r.Left + 8.8, r.Top + 8)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.15

        connector4 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 4.8, r.Top + 5.2, r.Left + 7.8, r.Top + 10.5)
        connector4.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector4.Line.Weight = 1.15

        connector2 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 4.7, r.Top + r.Height - 1, r.Left + 7.7, r.Top + 3)
        connector2.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector2.Line.Weight = 1.15

        connector3 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 2.3, r.Top + r.Height - 6.5, r.Left + 4.8, r.Top + r.Height - 1)
        connector3.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector3.Line.Weight = 1.15

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, connector2.Name, connector3.Name, connector4.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub FS_Click(sender As Object, e As RibbonControlEventArgs) Handles FS.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape
        Dim connector2 As Excel.Shape
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 5.8, r.Top + 2.7, r.Left + 10.8, r.Top + 14)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.3

        connector2 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 2.4, r.Top + 5, r.Left + 6.2, r.Top + 3)
        connector2.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector2.Line.Weight = 1.3


        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, connector2.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub NotMaterial_Click(sender As Object, e As RibbonControlEventArgs) Handles NotMaterial.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape = Nothing
        Dim shapeBeg As Excel.Shape = Nothing
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        shapeBeg = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, r.Left, r.Top, 50, r.Height)
        With shapeBeg.textframe2.textrange.Font
            .Size = 12
            .name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With
        With shapeBeg
            .Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            .Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End With
        With shapeBeg.TextFrame2
            .TextRange.Text = "M"
            .MarginBottom = 0
            .MarginTop = 0
            .MarginRight = 0
            .MarginLeft = 3
        End With
        shapeBeg.TextFrame.autosize = True

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + 3, r.Top + r.Height - 5, r.Left + 13, r.Top + 3)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.3

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, shapeBeg.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub

    Private Sub Minor_Pass_Click(sender As Object, e As RibbonControlEventArgs) Handles Minor_Pass.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "/Minor Pass/"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
        End With
    End Sub

    Private Sub Conclusion_Click(sender As Object, e As RibbonControlEventArgs) Handles Conclusion.Click
        Dim appCell As Object = Globals.ThisAddIn.Application.ActiveCell
        Dim begValue As String = Globals.ThisAddIn.Application.ActiveCell.Value

        With appCell
            .Value = begValue & "Conclusion:"
            .Font.Color = RGB(255, 0, 0)
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .font.bold = True
            .HorizontalAlignment = Excel.Constants.xlRight
        End With
    End Sub

    Private Sub ArrowBox_Click(sender As Object, e As RibbonControlEventArgs) Handles ArrowBox.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim connector1 As Excel.Shape = Nothing
        Dim connector2 As Excel.Shape = Nothing
        Dim shapeBeg As Excel.Shape = Nothing
        Dim r As Excel.Range

        Dim selection As String = app.Selection.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                                        RowAbsolute:=False, ColumnAbsolute:=False)

        r = app.Range(selection)

        shapeBeg = app.ActiveSheet.Shapes.AddTextbox _
            (Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, r.Left + r.Width + 27, r.Top + r.Height + 2.5, 29, 14)
        With shapeBeg.TextFrame2.TextRange.Font
            .Size = 10
            .Name = "Arial"
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With

        With shapeBeg.TextFrame2
            .TextRange.Text = ""
            .MarginBottom = 0
            .MarginTop = 1
            .MarginRight = 0
            .MarginLeft = 2
        End With
        shapeBeg.Line.ForeColor.RGB = RGB(0, 0, 0)

        connector1 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Top + r.Height, r.Left + r.Width, r.Top + r.Height + 10)
        connector1.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector1.Line.Weight = 1.3

        connector2 = app.ActiveSheet.Shapes.Addconnector _
(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, r.Left + r.Width, r.Top + r.Height + 10, r.Left + r.Width + 24, r.Top + r.Height + 10)
        connector2.Line.ForeColor.RGB = RGB(255, 0, 0)
        connector2.Line.Weight = 1.3
        connector2.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle

        'Group the shp and the arrow
        Dim MidObj As Object() = New Object() {connector1.Name, connector2.Name, shapeBeg.Name}
        app.ActiveSheet.Shapes.Range(MidObj).Group()
    End Sub
End Class
