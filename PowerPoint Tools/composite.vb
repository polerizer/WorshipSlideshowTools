Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.PowerPoint.PpSlideLayout

Public Class compose
    Dim objPPT As PowerPoint.Application = New PowerPoint.Application
    Dim objPres As PowerPoint.Presentation
    Dim objCustomLayout As PowerPoint.CustomLayout
    Dim treeCount As Integer = 0
    Dim notANode As TreeNode = New TreeNode

    Private Sub compose_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If cbClose.Checked Then
                objPres.Close()
            End If
            If objPPT.Windows.Count = 0 Then
                objPPT.Quit()
            End If
        Catch
        End Try
    End Sub
    Private Sub compose_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objPPT.Visible = True
        Me.TopMost = True
        Me.TopMost = False
        toggle(False)
        notANode.Name = "Iizliektotallynaughtanodelolzlolololololhaithariluzginsuandalso!&%$^)@*$(#@!#*$)!@#$!12341873"
        txtSubtitle.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            If isPPTFile(tree.SelectedNode.Tag) And My.Computer.FileSystem.FileExists(tree.SelectedNode.Tag) Then
                objPres.Slides.InsertFromFile(tree.SelectedNode.Tag, objPres.Slides.Count)
                txtList.Items.Add(tree.SelectedNode.Tag.ToString)
            End If
        Catch
        End Try

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        browse3.ShowDialog()
        Dim fonts As MsoTriState = False
        If cbFonts.Checked = True Then
            fonts = MsoTriState.msoTrue
        End If
        If My.Computer.FileSystem.FileExists(browse3.FileName) = False Then
            Try
                objPres.SaveAs(browse3.FileName, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation, fonts)
            Catch
            End Try
        Else
            'If MsgBox("File Already Exists, Overwrite?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try
                objPres.SaveAs(browse3.FileName, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation, fonts)
            Catch ex As Exception

            End Try
            'End If
        End If
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        browse2.ShowDialog()
        addNodes(browse2.SelectedPath, notANode)
        tree.ExpandAll()
    End Sub

    Sub addNodes(ByVal path As String, ByVal parentNode As TreeNode)
        If My.Computer.FileSystem.DirectoryExists(path) = False Then
            Return
        End If
        Dim count As Integer = My.Computer.FileSystem.GetFiles(path).Count
        Dim fileName As String = ""

        'Insert file nodes
        If parentNode.Name = notANode.Name Then
            tree.Nodes.Add(path).Tag = path
            For i As Integer = 0 To (count - 1) Step 1
                fileName = My.Computer.FileSystem.GetFiles(path).Item(i)
                If fileName.EndsWith(".ppt") Or fileName.EndsWith(".pptx") Or fileName.EndsWith(".pptm") Or fileName.EndsWith(".pps") Or fileName.EndsWith(".ppsx") Or fileName.EndsWith(".pot") Or fileName.EndsWith(".potx") Or fileName.EndsWith(".potm") Or fileName.EndsWith(".odp") Then
                    tree.Nodes.Add(fileName.Substring(fileName.LastIndexOf("\") + 1, fileName.Length - fileName.LastIndexOf("\") - 1)).Tag = fileName
                End If
            Next
        Else
            For i As Integer = 0 To (count - 1) Step 1
                fileName = My.Computer.FileSystem.GetFiles(path).Item(i)
                If fileName.EndsWith(".ppt") Or fileName.EndsWith(".pptx") Or fileName.EndsWith(".pptm") Or fileName.EndsWith(".pps") Or fileName.EndsWith(".ppsx") Or fileName.EndsWith(".pot") Or fileName.EndsWith(".potx") Or fileName.EndsWith(".potm") Or fileName.EndsWith(".odp") Then
                    parentNode.Nodes.Add(fileName.Substring(fileName.LastIndexOf("\") + 1, fileName.Length - fileName.LastIndexOf("\") - 1)).Tag = fileName
                End If
            Next
        End If

        'Browse for subfolders, create nodes, etc.
        count = My.Computer.FileSystem.GetDirectories(path).Count
        If count >= 1 Then
            For i As Integer = 0 To (count - 1) Step 1
                Dim newNode As TreeNode = findNode(notANode, path).Nodes.Add(My.Computer.FileSystem.GetDirectories(path).Item(i).Substring(My.Computer.FileSystem.GetDirectories(path).Item(i).LastIndexOf("\") + 1, My.Computer.FileSystem.GetDirectories(path).Item(i).Length - My.Computer.FileSystem.GetDirectories(path).Item(i).LastIndexOf("\") - 1))
                newNode.Tag = My.Computer.FileSystem.GetDirectories(path).Item(i).ToString
                addNodes(My.Computer.FileSystem.GetDirectories(path).Item(i).ToString, newNode)
            Next
        End If

    End Sub

    Function findNode(ByVal startNode As TreeNode, ByVal tagLine As String) As TreeNode
        'checks initial node
        Dim node As TreeNode
        Dim i As Integer = 1

        If startNode.Name = notANode.Name Then
            startNode = tree.Nodes(0)
        End If

        If startNode.Tag = tagLine Then
            Return startNode
        Else
            'check for children
            If startNode.Nodes.Count > 0 Then
                node = findNode(startNode.Nodes(0), tagLine)
                If node.Name <> notANode.Name Then
                    Return node
                End If
            End If

            If startNode.Level >= 1 Then
                If startNode.Level >= 1 And startNode.Parent.Nodes.Count > 1 Then
                    i = startNode.Index + 1
                    Do Until i >= startNode.Parent.Nodes.Count
                        node = findNode(startNode.Parent.Nodes(i), tagLine)
                        If node.Name <> notANode.Name Then
                            Return node
                        Else
                            i += 1
                        End If
                    Loop
                End If
            ElseIf startNode.Level <= 0 And tree.Nodes.Count > 1 Then
                i = startNode.Index + 1
                Do Until i >= tree.Nodes.Count
                    node = findNode(tree.Nodes(i), tagLine)
                    If node.Name <> notANode.Name Then
                        Return node
                    Else
                        i += 1
                    End If
                Loop
            End If
        End If

        Return notANode
    End Function

    Function isPPTFile(ByVal str As String) As Boolean
        If str.EndsWith(".ppt") Or str.EndsWith(".pptx") Or str.EndsWith(".pptm") Or str.EndsWith(".pps") Or str.EndsWith(".ppsx") Or str.EndsWith(".pot") Or str.EndsWith(".potx") Or str.EndsWith(".potm") Or str.EndsWith(".odp") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub compose_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd

    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        If objPPT.Presentations.Count > 0 Then
            Dim answer As MsgBoxResult = MsgBox("Save previous slideshow before creating new one?", MsgBoxStyle.YesNoCancel)
            If answer = MsgBoxResult.Yes Then
                btnSave_Click(btnNew, e)
                Try
                    objPres.Close()
                Catch
                Finally
                    objPres = objPPT.Presentations.Add
                    toggle(True)
                    txtList.Items.Clear()
                    If txtComment.Text = "" Then
                        txtComment.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
                    End If
                    If txtSubtitle.Text = "" Then
                        txtSubtitle.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
                    End If
                End Try
            ElseIf answer = MsgBoxResult.No Then
                Try
                    objPres.Close()
                Catch
                Finally
                    objPres = objPPT.Presentations.Add
                    toggle(True)
                    txtList.Items.Clear()
                    If txtComment.Text = "" Then
                        txtComment.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
                    End If
                    If txtSubtitle.Text = "" Then
                        txtSubtitle.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
                    End If
                End Try
            End If
        Else
            Try
                objPres.Close()
            Catch
            Finally
                objPres = objPPT.Presentations.Add
                toggle(True)
                txtList.Items.Clear()
                If txtComment.Text = "" Then
                    txtComment.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
                End If
                If txtSubtitle.Text = "" Then
                    txtSubtitle.Text = My.Computer.Clock.LocalTime.DayOfWeek.ToString & ", " & My.Computer.Clock.LocalTime.Date.Month & "/" & My.Computer.Clock.LocalTime.Day & "/" & My.Computer.Clock.LocalTime.Year
                End If
            End Try
        End If
    End Sub

    Private Sub btnTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTitle.Click
        Dim objSlide As PowerPoint.Slide
        objCustomLayout = objPres.SlideMaster.CustomLayouts.Item(1)
        objSlide = objPres.Slides.AddSlide(objPres.Slides.Count + 1, objCustomLayout)
        objSlide.Layout = ppLayoutTitle
        objSlide.Shapes(1).TextFrame.TextRange.Text = txtTitle.Text
        objSlide.Shapes(2).TextFrame.TextRange.Text = txtSubtitle.Text
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            objPres.Close()
        Catch
        Finally
            toggle(False)
        End Try
    End Sub
    Sub toggle(ByVal state As Boolean)
        btnAdd.Enabled = state
        btnBrowse.Enabled = state
        btnClose.Enabled = state
        btnIndex.Enabled = state
        btnTitle.Enabled = state
        btnSave.Enabled = state
    End Sub

    Private Sub btnIndex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIndex.Click
        Dim objSlide As PowerPoint.Slide
        objCustomLayout = objPres.SlideMaster.CustomLayouts.Item(1)
        objSlide = objPres.Slides.AddSlide(objPres.Slides.Count + 1, objCustomLayout)
        objSlide.Layout = ppLayoutText
        objSlide.Shapes(1).TextFrame.TextRange.Text = objPres.Name & " by " & txtAuthor.Text
        objSlide.Shapes(2).TextFrame.TextRange.Text = txtComment.Text
        For i As Integer = 0 To txtList.Items.Count - 1 Step 1
            objSlide.Shapes(2).TextFrame.TextRange.Text = objSlide.Shapes(2).TextFrame.TextRange.Text & vbNewLine & txtList.Items.Item(i).ToString
        Next
    End Sub

End Class
