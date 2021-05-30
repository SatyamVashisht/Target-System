Module Module1
    Dim inpForm As Windows.Forms.Form
    Dim inptxtbox As TextBox
    Dim lblDsc As Label
    Dim cmdOKBtn As Button
    Dim cmdCnclBtn As Button
    Dim lftPnl As Panel
    Dim stAlertIfEmpty As Boolean
    Public Function MyInputBox(ByVal Description As String, ByVal Title As String, ByVal DefaultValue As String, ByVal AlertIfEmpty As Boolean) As String
        stAlertIfEmpty = AlertIfEmpty
        inpForm = New Windows.Forms.Form
        inptxtbox = New TextBox
        lblDsc = New Label
        cmdOKBtn = New Button
        cmdCnclBtn = New Button
        lftPnl = New Panel
        'Left panel
        lftPnl.Parent = inpForm
        lftPnl.Dock = DockStyle.Right
        lftPnl.Width = 80
        lftPnl.TabIndex = 0
        'OK button
        cmdOKBtn.Parent = lftPnl
        cmdOKBtn.Text = "OK"
        cmdOKBtn.Top = 10
        cmdOKBtn.Visible = True
        inptxtbox.TabIndex = 0
        AddHandler cmdOKBtn.Click, AddressOf cmdOK_Click
        'Cancel Button
        cmdCnclBtn.Parent = lftPnl
        cmdCnclBtn.Text = "Cancel"
        cmdCnclBtn.Top = cmdOKBtn.Bottom + 5
        cmdCnclBtn.Visible = True
        cmdCnclBtn.TabIndex = 1
        AddHandler cmdCnclBtn.Click, AddressOf cmdCance_click
        'Description
        lblDsc.AutoSize = True
        lblDsc.Top = 5
        lblDsc.Left = 5
        lblDsc.Parent = inpForm
        lblDsc.Text = Description
        lblDsc.Visible = True
        lblDsc.BorderStyle = BorderStyle.Fixed3D
        lblDsc.BackColor = Color.LightGray

        'Input Textbox
        inptxtbox.Width = 450
        inptxtbox.Top = 40
        inptxtbox.Left = 5
        inptxtbox.Parent = inpForm
        inptxtbox.Text = DefaultValue
        inptxtbox.Visible = True
        inptxtbox.TabIndex = 0
        inptxtbox.PasswordChar = "*"


        'Form
        inpForm.Text = Title
        inpForm.StartPosition = FormStartPosition.CenterParent
        inpForm.CancelButton = cmdCnclBtn
        inpForm.AcceptButton = cmdOKBtn
        inpForm.ControlBox = False
        inpForm.Width = 550
        inpForm.Height = 150
        inpForm.ActiveControl = inptxtbox

        AddHandler inpForm.FormClosing, AddressOf inpForm_closing
        inpForm.ShowDialog()
        Return inptxtbox.Text
        inpForm.Dispose()
    End Function
    Sub cmdOK_Click(ByVal sender As Object, ByVal e As EventArgs)
        If inptxtbox.Text = "" And stAlertIfEmpty Then
            Dim vbanser As String = MsgBox("Submit empty field-?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Input Empty Field-?")
            If vbanser = 7 Then
                inpForm.ActiveControl = inptxtbox  'Promt the user if they leave the input box field blank and they click OK
                Exit Sub
            End If
            inptxtbox.Text = " "  'Return empty space so it will not be = Nothing
        End If
        inpForm.Close()
    End Sub
    Sub cmdCance_click(ByVal sender As Object, ByVal e As EventArgs)
        inptxtbox.Text = "!CANCELED!"  'This is what passes through when user presses Cancel button. You can change it to whatever you want.
        inpForm.Close()
    End Sub
    Sub inpForm_closing(ByVal sender As Object, ByVal e As EventArgs)
        'You can execute whatever you want here, this happens when the input box closes
    End Sub

End Module
