Public Class PASSWORD
    Dim frm As New Form
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "RAMSERSTA" Or TextBox1.Text = "ramsersta" Then
            frm.Dispose()
            frm = New EMPLOYEE


            frm.StartPosition = FormStartPosition.CenterScreen
            Me.Close()
            frm.Show()
        Else
            TextBox1.Text = ""
            MsgBox("INCORRECT PASSWORD")

        End If
    End Sub
End Class