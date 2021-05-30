Public Class PASSWORD2
    Dim FRM As New Form
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click



        If TextBox1.Text = "MANISHH" Or TextBox1.Text = "manishh" Then
            FRM.Dispose()
            FRM = New ENTRY


            FRM.StartPosition = FormStartPosition.CenterScreen
            Me.Close()
            FRM.Show()
        Else
            TextBox1.Text = ""
            MsgBox("INCORRECT PASSWORD")

        End If
    End Sub
End Class