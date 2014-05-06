Public Class MatrixScriptTool
   


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles NoRadio.CheckedChanged
        GroupBox1.Visible = False

        preview()

    End Sub

    Private Sub YesRadio_CheckedChanged(sender As Object, e As EventArgs) Handles YesRadio.CheckedChanged
        GroupBox1.Visible = True
        preview()


    End Sub
    Private Sub PSpreview()
        TextBox3.Text = premierSetupRec()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        preview()
        PSpreview()
        diagPreview()

    End Sub
    Private Sub reset()
        virRem.Value = 0
        proRem.Value = 0
        VRCustomRec.Text = ""
        NoRadio.Checked = True
        YesRadio.Checked = False
        VirusNoRB.Checked = False
        VirusYesRB.Checked = False
        SHDDCB.Checked = False
        LHDDCB.Checked = False
        StressCB.Checked = False
        QuickTestCB.Checked = False
        TestFailCB.Checked = False
        TestFailTB.Text = ""
        NoOSCorRB.Checked = False
        YesOSCorRB.Checked = False
        RecCB.Text = ""
        GroupBox1.Visible = False
        regMan.Checked = False
        regESP.Checked = False
        resPo.Checked = True
        softInst.Checked = False
        TextBox2.Visible = False
        CheckBox5.Checked = True
        CheckBox6.Checked = True
        CheckBox7.Checked = True
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        GroupBox2.Visible = False
        CheckBox10.Checked = False




    End Sub
    Private Sub preview()
        TextBox1.Text = VRRec()

    End Sub
    Private Function VRRec()
        Dim rec As String
        rec = "Virus Removal service is complete." & vbCrLf & "" + virRem.Value.ToString + " viruses have been removed from your system." & vbCrLf & "A  PC Tune Up was run and fixed all registry errors and security issues."
        rec = rec + vbCrLf + proRem.Value.ToString + " unwanted applications have been uninstalled." + vbCrLf + "Virus removal report saved to desktop for review with the customer. "
        If Not VRCustomRec.Text.Equals(Nothing) Then
            rec = rec + vbCrLf + VRCustomRec.Text
        End If
        Return rec

    End Function

    Private Sub virRem_ValueChanged(sender As Object, e As EventArgs) Handles virRem.ValueChanged
        preview()

    End Sub

    Private Sub proRem_ValueChanged(sender As Object, e As EventArgs) Handles proRem.ValueChanged
        preview()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim abc As String = "The WinPE virus removal scan has not been run." & vbCrLf & "Please boot to the ETTB flash drive and run the scan." & vbCrLf & " If you are having difficulty running the scan please call the Matrix at 1-855-55-MATRIX or refer to the Matrix virus removal ERG." & vbCrLf & "Thank you."
        Clipboard.SetText(abc)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim a As String = VRRec()
        Clipboard.SetText(a)

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            GroupBox2.Visible = True
        Else
            GroupBox2.Visible = False

        End If
    End Sub

    Private Sub softInst_CheckedChanged(sender As Object, e As EventArgs) Handles softInst.CheckedChanged
        If softInst.Checked = True Then
            TextBox2.Visible = True
        Else
            TextBox2.Visible = False

        End If
    End Sub
    Private Function isChecked(checkbox As CheckBox)

        If checkbox.Checked = True Then
            Return True
        Else
            Return False
        End If


    End Function

    Private Function premierSetupRec()
        Dim rec As String = "Premier Setup Complete."
        If isChecked(regMan) Then
            rec = rec + vbCrLf + "Manufacturer Warranty has been registered."
        End If
        If isChecked(regESP) Then
            rec = rec + vbCrLf + "Staples ESP has been registered."
        End If
        If isChecked(resPo) Then
            rec = rec + vbCrLf + "Restore Point created."
        End If
        If isChecked(softInst) Then
            rec = rec + vbCrLf + "The following titles have been installed on your PC: " & TextBox2.Text
        End If
        If isChecked(CheckBox5) Then
            rec = rec + vbCrLf + "Windows Updates are set to automatic."
        End If
        If isChecked(CheckBox6) Then
            rec = rec + vbCrLf + "Unwanted software has been uninstalled."
        End If
        If isChecked(CheckBox7) Then
            rec = rec + vbCrLf + "Unwanted Icons and Tiles have been removed."
        End If
        If isChecked(CheckBox8) Then
            rec = rec + vbCrLf + "Recovery Media created."
        End If
        If isChecked(CheckBox9) Then
            rec = rec + vbCrLf + "Data has been transferred from the old PC to the new PC."
            If isChecked(CheckBox10) Then
                rec = rec + vbCrLf + "The specific data you requested :" + RichTextBox1.Text + " has been transferred to your new PC"
            End If
        End If
        Return rec
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Clipboard.SetText(premierSetupRec())

    End Sub

    Private Sub regMan_CheckedChanged(sender As Object, e As EventArgs) Handles regMan.CheckedChanged
        PSpreview()

    End Sub

    Private Sub regESP_CheckedChanged(sender As Object, e As EventArgs) Handles regESP.CheckedChanged
        PSpreview()
    End Sub

    Private Sub resPo_CheckedChanged(sender As Object, e As EventArgs) Handles resPo.CheckedChanged
        PSpreview()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        PSpreview()
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        PSpreview()
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        PSpreview()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        PSpreview()
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        PSpreview()

    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        PSpreview()

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles VirusYesRB.CheckedChanged
        diagPreview()
    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles TestFailTB.TextChanged
        diagPreview()
    End Sub

    Private Sub TestFailCB_CheckedChanged(sender As Object, e As EventArgs) Handles TestFailCB.CheckedChanged
        If isChecked(TestFailCB) Then
            TestFailTB.Visible = True
        Else
            TestFailTB.Visible = False
        End If
        diagPreview()

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    Private Function isChecked(r As RadioButton)
        If r.Checked = True Then
            Return True
        Else : Return False

        End If
    End Function
    Private Function failTest()
        Return "PC Failed the following PC Doctor tests: " + TestFailTB.Text

    End Function

    Private Function issues()
        Dim issue As String = ""
        If isChecked(VirusYesRB) Then
            issue = "viruses."
            If isChecked(YesOSCorRB) Then
                issue = "viruses and OS corruption."
                If isChecked(TestFailCB) Then
                    issue = "viruses, OS corruption and " + failTest()
                    Return issue
                End If
                Return issue
            End If

        End If


        If isChecked(YesOSCorRB) Then
            issue = "OS corruption."
            If isChecked(TestFailCB) Then
                issue = "OS Corruption and " + failTest()
            End If
        End If
        If isChecked(TestFailCB) Then
            issue = failTest()
            If isChecked(VirusYesRB) Then
                issue = "viruses and " + failTest()

            End If
        End If
        If issue.Equals("") Then
            issue = "no issues."
        End If
        Return issue
    End Function
    Private Function performed()
        Dim perf As String = "Performed: Intake Analysis, Tune Up"
        If isChecked(QuickTestCB) Then
            perf = perf + ", PC Doctor Quick Test"
        End If
        If isChecked(SHDDCB) Then
            perf = perf + ", PC Doctor Short HDD Test"
        End If
        If isChecked(LHDDCB) Then
            perf = perf + ", PC Doctor Long HDD Test"
        End If
        If isChecked(StressCB) Then
            perf = perf + ", PC Doctor Stress Test"
        End If
        If isChecked(NoOSCorRB) Or isChecked(YesOSCorRB) Then
            perf = perf + ", System Files were verified for integrity."
        End If
        Return perf


    End Function
    Private Function diag()
        Dim rec As String
        rec = "The Diagnostic has been completed. A quick scan of all hardware and software was ran and we found " & issues()
        rec = rec + vbCrLf + "Recommendations: " + RecCB.Text
        rec = rec + vbCrLf + performed()
        rec = rec + vbCrLf + "Final Analysis was run and has been placed on the desktop for review with the customer."
        Return rec


    End Function


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Clipboard.SetText(diag())
    End Sub

    Private Sub StressCB_CheckedChanged(sender As Object, e As EventArgs) Handles StressCB.CheckedChanged
        diagPreview()
    End Sub

    Private Sub SHDDCB_CheckedChanged(sender As Object, e As EventArgs) Handles SHDDCB.CheckedChanged
        diagPreview()
    End Sub

    Private Sub LHDDCB_CheckedChanged(sender As Object, e As EventArgs) Handles LHDDCB.CheckedChanged
        diagPreview()
    End Sub

    Private Sub QuickTestCB_CheckedChanged(sender As Object, e As EventArgs) Handles QuickTestCB.CheckedChanged
        diagPreview()
    End Sub
    Private Sub diagPreview()
        DiagTB.Text = diag()
    End Sub

    Private Sub VirusNoRB_CheckedChanged(sender As Object, e As EventArgs) Handles VirusNoRB.CheckedChanged
        diagPreview()
    End Sub

    Private Sub YesOSCorRB_CheckedChanged(sender As Object, e As EventArgs) Handles YesOSCorRB.CheckedChanged
        diagPreview()

    End Sub

    Private Sub NoOSCorRB_CheckedChanged(sender As Object, e As EventArgs) Handles NoOSCorRB.CheckedChanged
        diagPreview()
    End Sub

    Private Sub RecCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RecCB.SelectedIndexChanged
        diagPreview()
    End Sub

    Private Sub VRCustomRec_TextChanged(sender As Object, e As EventArgs) Handles VRCustomRec.TextChanged
        preview()

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        PSpreview()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        PSpreview()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Clipboard.SetText("Rebooting PC, will reconnect when it comes back online.")

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Clipboard.SetText("Transferred work order to another Agent due to shift ending.")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim a As String
        a = "This PC currently has Sophos installed. Staples does not service virus removals for Sophos customers."
        a = a + vbCrLf + "Sophos includes remote/onsite removals. The customer will need to contact Sophos directly for service."
        a = a + vbCrLf + " We cannot run the Virus Removal service due to an issue with Sophos and the Tune Up utility."
        Clipboard.SetText(a)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim a As String
        a = "Store is about to close. Please secure the PC per policy and reconnect to the Matrix in the morning. Thank you."
        Clipboard.SetText(a)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim a As String
        a = "Going to lunch, I will be back in an hour."
        a = a + vbCrLf + "If you have an urgent issue, please call 1-855-55-Matrix."
        a = a + vbCrLf + "Otherwise, just leave a message in chat if you need anything and I'll answer it when I return. Thank you."
        Clipboard.SetText(a)

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub RecCB_TextChanged(sender As Object, e As EventArgs) Handles RecCB.TextChanged
        diagPreview()

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        reset()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        reset()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        reset()

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        reset()

    End Sub
End Class
