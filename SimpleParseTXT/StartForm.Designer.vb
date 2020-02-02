<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StartForm
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
        Me.PathTextBox = New System.Windows.Forms.TextBox()
        Me.ParseButton = New System.Windows.Forms.Button()
        Me.MakeButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TestButton = New System.Windows.Forms.Button()
        Me.HelpB = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'PathTextBox
        '
        Me.PathTextBox.Location = New System.Drawing.Point(20, 59)
        Me.PathTextBox.Name = "PathTextBox"
        Me.PathTextBox.Size = New System.Drawing.Size(522, 20)
        Me.PathTextBox.TabIndex = 2
        Me.PathTextBox.Text = "C:\folder\test.sg"
        '
        'ParseButton
        '
        Me.ParseButton.Location = New System.Drawing.Point(64, 112)
        Me.ParseButton.Name = "ParseButton"
        Me.ParseButton.Size = New System.Drawing.Size(147, 57)
        Me.ParseButton.TabIndex = 0
        Me.ParseButton.Text = "Parse .sg"
        Me.ParseButton.UseVisualStyleBackColor = True
        '
        'MakeButton
        '
        Me.MakeButton.Location = New System.Drawing.Point(64, 207)
        Me.MakeButton.Name = "MakeButton"
        Me.MakeButton.Size = New System.Drawing.Size(147, 57)
        Me.MakeButton.TabIndex = 1
        Me.MakeButton.Text = "Make .sg"
        Me.MakeButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "path to the .sg file"
        '
        'TestButton
        '
        Me.TestButton.Location = New System.Drawing.Point(263, 112)
        Me.TestButton.Name = "TestButton"
        Me.TestButton.Size = New System.Drawing.Size(147, 57)
        Me.TestButton.TabIndex = 4
        Me.TestButton.Text = "Test"
        Me.TestButton.UseVisualStyleBackColor = True
        '
        'HelpB
        '
        Me.HelpB.Location = New System.Drawing.Point(263, 207)
        Me.HelpB.Name = "HelpB"
        Me.HelpB.Size = New System.Drawing.Size(147, 57)
        Me.HelpB.TabIndex = 5
        Me.HelpB.Text = "Help"
        Me.HelpB.UseVisualStyleBackColor = True
        '
        'StartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 300)
        Me.Controls.Add(Me.HelpB)
        Me.Controls.Add(Me.TestButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MakeButton)
        Me.Controls.Add(Me.ParseButton)
        Me.Controls.Add(Me.PathTextBox)
        Me.Name = "StartForm"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PathTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ParseButton As System.Windows.Forms.Button
    Friend WithEvents MakeButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TestButton As System.Windows.Forms.Button
    Friend WithEvents HelpB As System.Windows.Forms.Button

End Class
