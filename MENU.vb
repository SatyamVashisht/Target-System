Public Class MENU
    Dim FRM As New Form
    Private Sub ENTRYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ENTRYToolStripMenuItem.Click
        FRM.Dispose()
        FRM = New PASSWORD2

        FRM.MdiParent = Me
        FRM.StartPosition = FormStartPosition.CenterScreen

        FRM.Show()
    End Sub

    Private Sub EMPLOYEEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EMPLOYEEToolStripMenuItem.Click
        FRM.Dispose()
        FRM = New PASSWORD

        FRM.MdiParent = Me
        FRM.StartPosition = FormStartPosition.CenterScreen

        FRM.Show()
    End Sub

    Private Sub TARGETToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TARGETToolStripMenuItem.Click
        FRM.Dispose()
        FRM = New TARGET

        FRM.MdiParent = Me
        FRM.StartPosition = FormStartPosition.CenterScreen

        FRM.Show()
    End Sub

    Private Sub VIEWToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VIEWToolStripMenuItem.Click
        FRM.Dispose()
        FRM = New VIEW

        FRM.MdiParent = Me
        FRM.StartPosition = FormStartPosition.CenterScreen

        FRM.Show()
    End Sub
End Class
