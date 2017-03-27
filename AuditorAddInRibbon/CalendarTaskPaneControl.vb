Public Class CalendarTaskPaneControl
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Globals.ThisAddIn.Application.ActiveCell.Value =
 MonthCalendar1.SelectionRange.Start.ToShortDateString()
    End Sub
End Class
