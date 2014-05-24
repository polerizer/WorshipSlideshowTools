<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class pptConvert
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
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.browse = New System.Windows.Forms.FolderBrowserDialog()
        Me.browse2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.txt2 = New System.Windows.Forms.TextBox()
        Me.lbl = New System.Windows.Forms.Label()
        Me.cbType = New System.Windows.Forms.ComboBox()
        Me.cb2 = New System.Windows.Forms.CheckBox()
        Me.cb3 = New System.Windows.Forms.CheckBox()
        Me.rb1 = New System.Windows.Forms.RadioButton()
        Me.rb2 = New System.Windows.Forms.RadioButton()
        Me.txtFiles = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.browse3 = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'btnConvert
        '
        Me.btnConvert.Location = New System.Drawing.Point(12, 12)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(269, 23)
        Me.btnConvert.TabIndex = 0
        Me.btnConvert.Text = "Convert to PPT"
        Me.btnConvert.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(330, 11)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Input Folder"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(330, 62)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Output Folder"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'txt1
        '
        Me.txt1.Location = New System.Drawing.Point(411, 14)
        Me.txt1.Name = "txt1"
        Me.txt1.ReadOnly = True
        Me.txt1.Size = New System.Drawing.Size(188, 20)
        Me.txt1.TabIndex = 3
        '
        'txt2
        '
        Me.txt2.Location = New System.Drawing.Point(411, 62)
        Me.txt2.Name = "txt2"
        Me.txt2.ReadOnly = True
        Me.txt2.Size = New System.Drawing.Size(188, 20)
        Me.txt2.TabIndex = 4
        '
        'lbl
        '
        Me.lbl.AutoSize = True
        Me.lbl.Location = New System.Drawing.Point(12, 91)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(87, 13)
        Me.lbl.TabIndex = 5
        Me.lbl.Text = "Not Converting..."
        '
        'cbType
        '
        Me.cbType.FormattingEnabled = True
        Me.cbType.Items.AddRange(New Object() {"PPT", "PPTX", "POT", "POTX", "ODP", "PDF", "HTML", "JPG", "BMP", "PNG", "TIF"})
        Me.cbType.Location = New System.Drawing.Point(12, 41)
        Me.cbType.Name = "cbType"
        Me.cbType.Size = New System.Drawing.Size(269, 21)
        Me.cbType.TabIndex = 6
        '
        'cb2
        '
        Me.cb2.AutoSize = True
        Me.cb2.Location = New System.Drawing.Point(12, 68)
        Me.cb2.Name = "cb2"
        Me.cb2.Size = New System.Drawing.Size(88, 17)
        Me.cb2.TabIndex = 8
        Me.cb2.Text = "Embed Fonts"
        Me.cb2.UseVisualStyleBackColor = True
        '
        'cb3
        '
        Me.cb3.AutoSize = True
        Me.cb3.Location = New System.Drawing.Point(106, 69)
        Me.cb3.Name = "cb3"
        Me.cb3.Size = New System.Drawing.Size(97, 17)
        Me.cb3.TabIndex = 9
        Me.cb3.Text = "Enable Macros"
        Me.cb3.UseVisualStyleBackColor = True
        '
        'rb1
        '
        Me.rb1.AutoSize = True
        Me.rb1.Checked = True
        Me.rb1.Location = New System.Drawing.Point(310, 17)
        Me.rb1.Name = "rb1"
        Me.rb1.Size = New System.Drawing.Size(14, 13)
        Me.rb1.TabIndex = 10
        Me.rb1.TabStop = True
        Me.rb1.UseVisualStyleBackColor = True
        '
        'rb2
        '
        Me.rb2.AutoSize = True
        Me.rb2.Location = New System.Drawing.Point(310, 41)
        Me.rb2.Name = "rb2"
        Me.rb2.Size = New System.Drawing.Size(14, 13)
        Me.rb2.TabIndex = 11
        Me.rb2.UseVisualStyleBackColor = True
        '
        'txtFiles
        '
        Me.txtFiles.Enabled = False
        Me.txtFiles.Location = New System.Drawing.Point(411, 39)
        Me.txtFiles.Name = "txtFiles"
        Me.txtFiles.ReadOnly = True
        Me.txtFiles.Size = New System.Drawing.Size(188, 20)
        Me.txtFiles.TabIndex = 13
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(330, 36)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Select Files"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'browse3
        '
        Me.browse3.Filter = "PowerPoint Supported Files|*.ppt;*.pptx;*.pptm;*.pot;*.potx;*.potm;*.odp"
        Me.browse3.Multiselect = True
        '
        'pptConvert
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(611, 134)
        Me.Controls.Add(Me.txtFiles)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.rb2)
        Me.Controls.Add(Me.rb1)
        Me.Controls.Add(Me.cb3)
        Me.Controls.Add(Me.cb2)
        Me.Controls.Add(Me.cbType)
        Me.Controls.Add(Me.lbl)
        Me.Controls.Add(Me.txt2)
        Me.Controls.Add(Me.txt1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnConvert)
        Me.Name = "pptConvert"
        Me.Text = "Convert"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents browse As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents browse2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents txt2 As System.Windows.Forms.TextBox
    Friend WithEvents lbl As System.Windows.Forms.Label
    Friend WithEvents cbType As System.Windows.Forms.ComboBox
    Friend WithEvents cb2 As System.Windows.Forms.CheckBox
    Friend WithEvents cb3 As System.Windows.Forms.CheckBox
    Friend WithEvents rb1 As System.Windows.Forms.RadioButton
    Friend WithEvents rb2 As System.Windows.Forms.RadioButton
    Friend WithEvents txtFiles As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents browse3 As System.Windows.Forms.OpenFileDialog

End Class
