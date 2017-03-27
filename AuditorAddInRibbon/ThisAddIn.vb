Public Class ThisAddIn
    Private taskPaneControl1 As CopyrightTaskPane
    Private taskPaneControl2 As CalendarTaskPaneControl
    Private WithEvents taskPaneValue As Microsoft.Office.Tools.CustomTaskPane
    Private WithEvents taskPaneValue2 As Microsoft.Office.Tools.CustomTaskPane
    Public ReadOnly Property TaskPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return taskPaneValue
        End Get
    End Property
    Public ReadOnly Property TaskPane2() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return taskPaneValue2
        End Get
    End Property
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        taskPaneControl1 = New CopyrightTaskPane()
        taskPaneValue = Me.CustomTaskPanes.Add(
            taskPaneControl1, "Copyright")
        taskPaneValue.Width = 420

        taskPaneControl2 = New CalendarTaskPaneControl()
        taskPaneValue2 = Me.CustomTaskPanes.Add(
            taskPaneControl2, "Calendar")
        taskPaneValue2.Width = 232

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Private Sub taskPaneValue_VisibleChanged(sender As Object, e As EventArgs) Handles taskPaneValue.VisibleChanged
        Globals.Ribbons.Ribbon1.Copyright_ToggleButton.Checked = taskPaneValue.Visible
    End Sub

    Private Sub taskPaneValue2_VisibleChanged(sender As Object, e As EventArgs) Handles taskPaneValue2.VisibleChanged
        Globals.Ribbons.Ribbon1.Calendar_ToggleButton.Checked = taskPaneValue2.Visible
    End Sub
End Class
