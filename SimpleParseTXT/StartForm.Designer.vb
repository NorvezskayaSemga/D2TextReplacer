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
        Me.TglobalTextBox1 = New System.Windows.Forms.TextBox()
        Me.TglobalTextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Tglobal1Button = New System.Windows.Forms.Button()
        Me.Tglobal2Button = New System.Windows.Forms.Button()
        Me.MapButton = New System.Windows.Forms.Button()
        Me.AutotranslateCheckBox = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'PathTextBox
        '
        Me.PathTextBox.Location = New System.Drawing.Point(15, 143)
        Me.PathTextBox.Name = "PathTextBox"
        Me.PathTextBox.Size = New System.Drawing.Size(462, 20)
        Me.PathTextBox.TabIndex = 30
        Me.PathTextBox.Text = "C:\folder\test.sg"
        '
        'ParseButton
        '
        Me.ParseButton.Location = New System.Drawing.Point(64, 178)
        Me.ParseButton.Name = "ParseButton"
        Me.ParseButton.Size = New System.Drawing.Size(147, 47)
        Me.ParseButton.TabIndex = 51
        Me.ParseButton.Text = "Parse .sg"
        Me.ParseButton.UseVisualStyleBackColor = True
        '
        'MakeButton
        '
        Me.MakeButton.Location = New System.Drawing.Point(64, 241)
        Me.MakeButton.Name = "MakeButton"
        Me.MakeButton.Size = New System.Drawing.Size(147, 47)
        Me.MakeButton.TabIndex = 53
        Me.MakeButton.Text = "Make .sg"
        Me.MakeButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 127)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "path to the .sg file"
        '
        'TestButton
        '
        Me.TestButton.Location = New System.Drawing.Point(263, 178)
        Me.TestButton.Name = "TestButton"
        Me.TestButton.Size = New System.Drawing.Size(147, 47)
        Me.TestButton.TabIndex = 52
        Me.TestButton.Text = "Test"
        Me.TestButton.UseVisualStyleBackColor = True
        '
        'HelpB
        '
        Me.HelpB.Location = New System.Drawing.Point(263, 241)
        Me.HelpB.Name = "HelpB"
        Me.HelpB.Size = New System.Drawing.Size(147, 47)
        Me.HelpB.TabIndex = 54
        Me.HelpB.Text = "Help"
        Me.HelpB.UseVisualStyleBackColor = True
        '
        'TglobalTextBox1
        '
        Me.TglobalTextBox1.Location = New System.Drawing.Point(12, 43)
        Me.TglobalTextBox1.Name = "TglobalTextBox1"
        Me.TglobalTextBox1.Size = New System.Drawing.Size(465, 20)
        Me.TglobalTextBox1.TabIndex = 10
        Me.TglobalTextBox1.Text = "C:\folder\TGlobal.dbf"
        '
        'TglobalTextBox2
        '
        Me.TglobalTextBox2.Location = New System.Drawing.Point(12, 68)
        Me.TglobalTextBox2.Name = "TglobalTextBox2"
        Me.TglobalTextBox2.Size = New System.Drawing.Size(465, 20)
        Me.TglobalTextBox2.TabIndex = 20
        Me.TglobalTextBox2.Text = "C:\folder\TGlobal.dbf"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(216, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "path to the dbf text files (ru/eng in any order)"
        '
        'Tglobal1Button
        '
        Me.Tglobal1Button.Location = New System.Drawing.Point(483, 43)
        Me.Tglobal1Button.Name = "Tglobal1Button"
        Me.Tglobal1Button.Size = New System.Drawing.Size(65, 20)
        Me.Tglobal1Button.TabIndex = 11
        Me.Tglobal1Button.Text = "Select"
        Me.Tglobal1Button.UseVisualStyleBackColor = True
        '
        'Tglobal2Button
        '
        Me.Tglobal2Button.Location = New System.Drawing.Point(483, 67)
        Me.Tglobal2Button.Name = "Tglobal2Button"
        Me.Tglobal2Button.Size = New System.Drawing.Size(65, 20)
        Me.Tglobal2Button.TabIndex = 21
        Me.Tglobal2Button.Text = "Select"
        Me.Tglobal2Button.UseVisualStyleBackColor = True
        '
        'MapButton
        '
        Me.MapButton.Location = New System.Drawing.Point(483, 143)
        Me.MapButton.Name = "MapButton"
        Me.MapButton.Size = New System.Drawing.Size(65, 20)
        Me.MapButton.TabIndex = 31
        Me.MapButton.Text = "Select"
        Me.MapButton.UseVisualStyleBackColor = True
        '
        'AutotranslateCheckBox
        '
        Me.AutotranslateCheckBox.AutoSize = True
        Me.AutotranslateCheckBox.Location = New System.Drawing.Point(12, 94)
        Me.AutotranslateCheckBox.Name = "AutotranslateCheckBox"
        Me.AutotranslateCheckBox.Size = New System.Drawing.Size(226, 17)
        Me.AutotranslateCheckBox.TabIndex = 55
        Me.AutotranslateCheckBox.Text = "Autotranslate by means of game resources"
        Me.AutotranslateCheckBox.UseVisualStyleBackColor = True
        '
        'StartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 300)
        Me.Controls.Add(Me.AutotranslateCheckBox)
        Me.Controls.Add(Me.MapButton)
        Me.Controls.Add(Me.Tglobal2Button)
        Me.Controls.Add(Me.Tglobal1Button)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TglobalTextBox2)
        Me.Controls.Add(Me.TglobalTextBox1)
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
    Friend WithEvents TglobalTextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TglobalTextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Tglobal1Button As System.Windows.Forms.Button
    Friend WithEvents Tglobal2Button As System.Windows.Forms.Button
    Friend WithEvents MapButton As System.Windows.Forms.Button
    Friend WithEvents AutotranslateCheckBox As System.Windows.Forms.CheckBox

End Class
