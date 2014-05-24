Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Public Class pptConvert
    Dim objPPT As PowerPoint.Application = New PowerPoint.Application

    Private Sub pptConvert_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        objPPT.Quit()
    End Sub

    Private Sub pptTools_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cbType.SelectedItem = "PPT"
    End Sub

    Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click

        Dim pptType As PowerPoint.PpSaveAsFileType = PowerPoint.PpSaveAsFileType.ppSaveAsPresentation
        Dim pptFonts As MsoTriState = MsoTriState.msoFalse
        objPPT.Visible = True

        'Set output settings
        Select Case cbType.SelectedItem
            Case "PPT"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsPresentation
            Case "PPTX"
                If cb3.Checked Then
                    pptType = PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled
                Else
                    pptType = PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation
                End If
            Case "POT"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsTemplate
            Case "POTX"
                If cb3.Checked Then
                    pptType = PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLTemplateMacroEnabled
                Else
                    pptType = PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLTemplate
                End If
            Case "ODP"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsOpenDocumentPresentation
            Case "HTML"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsHTML
            Case "PDF"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsPDF
            Case "JPG"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsJPG
            Case "BMP"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsBMP
            Case "TIF"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsTIF
            Case "PNG"
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsPNG
            Case Else
                pptType = PowerPoint.PpSaveAsFileType.ppSaveAsPresentation
        End Select

        If rb1.Checked And Not rb2.Checked Then
            If My.Computer.FileSystem.DirectoryExists(browse.SelectedPath) And My.Computer.FileSystem.DirectoryExists(browse2.SelectedPath) Then
                For i As Integer = 0 To (My.Computer.FileSystem.GetFiles(browse.SelectedPath).Count - 1) Step 1
                    If My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".ppt") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".pptx") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".pptm") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".pps") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".ppsx") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".pot") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".potx") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".potm") Or My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).EndsWith(".odp") Then
                        lbl.Text = "Converting '" & My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).Substring(My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).LastIndexOf("\") + 1, My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).Length - My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i).LastIndexOf("\") - 1) & "'"
                        Dim objPres As PowerPoint.Presentation = objPPT.Presentations.Open(My.Computer.FileSystem.GetFiles(browse.SelectedPath).Item(i))
                        objPres.SaveCopyAs(browse2.SelectedPath & "\" & objPres.Name.Substring(0, objPres.Name.IndexOf(".")), pptType, pptFonts)
                        lbl.Text = objPres.Name & " saved as " & browse2.SelectedPath & "\" & objPres.Name.Substring(0, objPres.Name.IndexOf("."))
                        objPres.Close()
                    End If
                Next
                lbl.Text = "Conversion complete."
            Else
                lbl.Text = "Conversion cannot begin: input or output folder does not exist."
            End If
        ElseIf rb2.Checked And Not rb1.Checked Then
            If My.Computer.FileSystem.DirectoryExists(browse2.SelectedPath) Then
                For i As Integer = 0 To (browse3.FileNames.Count - 1) Step 1
                    If My.Computer.FileSystem.FileExists(browse3.FileNames.ElementAt(i)) Then
                        lbl.Text = "Converting '" & browse3.FileNames.ElementAt(i).Substring(browse3.FileNames.ElementAt(i).LastIndexOf("\") + 1, browse3.FileNames.ElementAt(i).Length - browse3.FileNames.ElementAt(i).LastIndexOf("\") - 1) & "'"
                        Dim objPres As PowerPoint.Presentation = objPPT.Presentations.Open(browse3.FileNames.ElementAt(i))
                        objPres.SaveCopyAs(browse2.SelectedPath & "\" & objPres.Name.Substring(0, objPres.Name.IndexOf(".")), pptType, pptFonts)
                        lbl.Text = objPres.Name & " saved as " & browse2.SelectedPath & "\" & objPres.Name.Substring(0, objPres.Name.IndexOf("."))
                        objPres.Close()
                    End If
                Next
                lbl.Text = "Conversion complete."
            Else
                lbl.Text = "Conversion cannot begin: input or output folder does not exist."
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        browse.ShowDialog()
        txt1.Text = browse.SelectedPath
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        browse2.ShowDialog()
        txt2.Text = browse2.SelectedPath
    End Sub

    Private Sub cbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbType.SelectedIndexChanged
        btnConvert.Text = "Convert to " & cbType.SelectedItem.ToString
    End Sub

    Private Sub rb1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb1.CheckedChanged
        Button2.Enabled = rb1.Checked
        Button1.Enabled = rb2.Checked
        txt1.Enabled = rb1.Checked
        txtFiles.Enabled = rb2.Checked
    End Sub

    Private Sub rb2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb2.CheckedChanged
        Button2.Enabled = rb1.Checked
        Button1.Enabled = rb2.Checked
        txt1.Enabled = rb1.Checked
        txtFiles.Enabled = rb2.Checked
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        browse3.ShowDialog()
        txtFiles.Text = browse3.FileNames.Count.ToString & " objects selected."
    End Sub
End Class