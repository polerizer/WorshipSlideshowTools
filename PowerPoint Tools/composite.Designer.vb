<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class compose
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
        Me.browse = New System.Windows.Forms.OpenFileDialog()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.tree = New System.Windows.Forms.TreeView()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.browse2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.browse3 = New System.Windows.Forms.SaveFileDialog()
        Me.txtList = New System.Windows.Forms.ListBox()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.grpTitle = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSubtitle = New System.Windows.Forms.TextBox()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.btnTitle = New System.Windows.Forms.Button()
        Me.btnIndex = New System.Windows.Forms.Button()
        Me.grpIndex = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtComment = New System.Windows.Forms.TextBox()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.cbFonts = New System.Windows.Forms.CheckBox()
        Me.cbClose = New System.Windows.Forms.CheckBox()
        Me.grpTitle.SuspendLayout()
        Me.grpIndex.SuspendLayout()
        Me.SuspendLayout()
        '
        'browse
        '
        Me.browse.Filter = "PowerPoint Files|*.ppt;*.pptx;*.pptm;*.pot;*.potx;*.potm;*.odp"
        Me.browse.Title = "Select Presentation to Add"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(12, 59)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(97, 41)
        Me.btnAdd.TabIndex = 0
        Me.btnAdd.Text = "Add Slides"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(12, 106)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(97, 41)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save Slideshow"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'tree
        '
        Me.tree.Location = New System.Drawing.Point(115, 12)
        Me.tree.Name = "tree"
        Me.tree.Size = New System.Drawing.Size(399, 364)
        Me.tree.TabIndex = 2
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(12, 317)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(97, 41)
        Me.btnBrowse.TabIndex = 3
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'browse3
        '
        Me.browse3.DefaultExt = "ppt"
        Me.browse3.Filter = "PowerPoint Files|*.ppt;*.pptx;*.pptm;*.pot;*.potx;*.potm;*.odp"
        Me.browse3.Title = "Save As..."
        '
        'txtList
        '
        Me.txtList.FormattingEnabled = True
        Me.txtList.HorizontalScrollbar = True
        Me.txtList.Location = New System.Drawing.Point(6, 78)
        Me.txtList.Name = "txtList"
        Me.txtList.Size = New System.Drawing.Size(264, 199)
        Me.txtList.TabIndex = 4
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(12, 12)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(97, 41)
        Me.btnNew.TabIndex = 5
        Me.btnNew.Text = "New Slideshow"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'grpTitle
        '
        Me.grpTitle.Controls.Add(Me.Label2)
        Me.grpTitle.Controls.Add(Me.Label1)
        Me.grpTitle.Controls.Add(Me.txtSubtitle)
        Me.grpTitle.Controls.Add(Me.txtTitle)
        Me.grpTitle.Location = New System.Drawing.Point(521, 12)
        Me.grpTitle.Name = "grpTitle"
        Me.grpTitle.Size = New System.Drawing.Size(276, 75)
        Me.grpTitle.TabIndex = 9
        Me.grpTitle.TabStop = False
        Me.grpTitle.Text = "Title Slide"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Subtitle"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Title"
        '
        'txtSubtitle
        '
        Me.txtSubtitle.Location = New System.Drawing.Point(54, 47)
        Me.txtSubtitle.Name = "txtSubtitle"
        Me.txtSubtitle.Size = New System.Drawing.Size(216, 20)
        Me.txtSubtitle.TabIndex = 1
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(54, 19)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(216, 20)
        Me.txtTitle.TabIndex = 0
        '
        'btnTitle
        '
        Me.btnTitle.Location = New System.Drawing.Point(12, 176)
        Me.btnTitle.Name = "btnTitle"
        Me.btnTitle.Size = New System.Drawing.Size(97, 41)
        Me.btnTitle.TabIndex = 10
        Me.btnTitle.Text = "Insert Title Slide"
        Me.btnTitle.UseVisualStyleBackColor = True
        '
        'btnIndex
        '
        Me.btnIndex.Location = New System.Drawing.Point(12, 223)
        Me.btnIndex.Name = "btnIndex"
        Me.btnIndex.Size = New System.Drawing.Size(97, 41)
        Me.btnIndex.TabIndex = 11
        Me.btnIndex.Text = "Insert Index Slide"
        Me.btnIndex.UseVisualStyleBackColor = True
        '
        'grpIndex
        '
        Me.grpIndex.Controls.Add(Me.Label4)
        Me.grpIndex.Controls.Add(Me.Label3)
        Me.grpIndex.Controls.Add(Me.txtComment)
        Me.grpIndex.Controls.Add(Me.txtAuthor)
        Me.grpIndex.Controls.Add(Me.txtList)
        Me.grpIndex.Location = New System.Drawing.Point(521, 93)
        Me.grpIndex.Name = "grpIndex"
        Me.grpIndex.Size = New System.Drawing.Size(276, 283)
        Me.grpIndex.TabIndex = 12
        Me.grpIndex.TabStop = False
        Me.grpIndex.Text = "Index Slide"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Custom"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 26)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Author"
        '
        'txtComment
        '
        Me.txtComment.Location = New System.Drawing.Point(54, 45)
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(216, 20)
        Me.txtComment.TabIndex = 6
        '
        'txtAuthor
        '
        Me.txtAuthor.Location = New System.Drawing.Point(54, 19)
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.Size = New System.Drawing.Size(216, 20)
        Me.txtAuthor.TabIndex = 5
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(12, 270)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(97, 41)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close Slideshow"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'cbFonts
        '
        Me.cbFonts.AutoSize = True
        Me.cbFonts.Location = New System.Drawing.Point(12, 153)
        Me.cbFonts.Name = "cbFonts"
        Me.cbFonts.Size = New System.Drawing.Size(88, 17)
        Me.cbFonts.TabIndex = 14
        Me.cbFonts.Text = "Embed Fonts"
        Me.cbFonts.UseVisualStyleBackColor = True
        '
        'cbClose
        '
        Me.cbClose.AutoSize = True
        Me.cbClose.Location = New System.Drawing.Point(12, 364)
        Me.cbClose.Name = "cbClose"
        Me.cbClose.Size = New System.Drawing.Size(87, 17)
        Me.cbClose.TabIndex = 15
        Me.cbClose.Text = "Close on Exit"
        Me.cbClose.UseVisualStyleBackColor = True
        '
        'compose
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(804, 388)
        Me.Controls.Add(Me.cbClose)
        Me.Controls.Add(Me.cbFonts)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.grpIndex)
        Me.Controls.Add(Me.btnIndex)
        Me.Controls.Add(Me.btnTitle)
        Me.Controls.Add(Me.grpTitle)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.tree)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Name = "compose"
        Me.Text = "Compose"
        Me.grpTitle.ResumeLayout(False)
        Me.grpTitle.PerformLayout()
        Me.grpIndex.ResumeLayout(False)
        Me.grpIndex.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents browse As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents tree As System.Windows.Forms.TreeView
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents browse2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents browse3 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents txtList As System.Windows.Forms.ListBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents grpTitle As System.Windows.Forms.GroupBox
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSubtitle As System.Windows.Forms.TextBox
    Friend WithEvents btnTitle As System.Windows.Forms.Button
    Friend WithEvents btnIndex As System.Windows.Forms.Button
    Friend WithEvents grpIndex As System.Windows.Forms.GroupBox
    Friend WithEvents txtComment As System.Windows.Forms.TextBox
    Friend WithEvents txtAuthor As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cbFonts As System.Windows.Forms.CheckBox
    Friend WithEvents cbClose As System.Windows.Forms.CheckBox
End Class
