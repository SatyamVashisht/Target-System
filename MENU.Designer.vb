<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MENU
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ENTRYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VIEWToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EMPLOYEEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TARGETToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ENTRYToolStripMenuItem, Me.VIEWToolStripMenuItem, Me.EMPLOYEEToolStripMenuItem, Me.TARGETToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(8, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(1362, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ENTRYToolStripMenuItem
        '
        Me.ENTRYToolStripMenuItem.Name = "ENTRYToolStripMenuItem"
        Me.ENTRYToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ENTRYToolStripMenuItem.Text = "ENTRY"
        '
        'VIEWToolStripMenuItem
        '
        Me.VIEWToolStripMenuItem.Name = "VIEWToolStripMenuItem"
        Me.VIEWToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.VIEWToolStripMenuItem.Text = "VIEW"
        '
        'EMPLOYEEToolStripMenuItem
        '
        Me.EMPLOYEEToolStripMenuItem.Name = "EMPLOYEEToolStripMenuItem"
        Me.EMPLOYEEToolStripMenuItem.Size = New System.Drawing.Size(77, 20)
        Me.EMPLOYEEToolStripMenuItem.Text = "EMPLOYEE"
        '
        'TARGETToolStripMenuItem
        '
        Me.TARGETToolStripMenuItem.Name = "TARGETToolStripMenuItem"
        Me.TARGETToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.TARGETToolStripMenuItem.Text = "TARGET"
        '
        'MENU
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1362, 684)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "MENU"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MENU"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ENTRYToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VIEWToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EMPLOYEEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TARGETToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
