Public Class StartupForm 
        Sub New()
        InitSkins()
        InitializeComponent()

    End Sub
    Sub InitSkins()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
        'UserLookAndFeel.Default.SetSkinStyle("DevExpress Style")

    End Sub
Private Sub SimpleButton1_Click( sender As Object,  e As EventArgs) Handles SimpleButton1.Click
        DxErrorProvider1.ClearErrors
        If ComboBoxEdit1.EditValue = Nothing Then
            DxErrorProvider1.SetError(ComboBoxEdit1,"Missing Value"): Exit Sub
        End If
        If ComboBoxEdit2.EditValue = Nothing Then
            DxErrorProvider1.SetError(ComboBoxEdit2,"Missing Value"): Exit Sub
        End If
        If SpinEdit1.EditValue = Nothing Then
            DxErrorProvider1.SetError(SpinEdit1,"Missing Value"): Exit Sub
        End If
        If SpinEdit1.Enabled Then
            If My.Computer.FileSystem.FileExists(XtraForm1.DatabasePath) Then My.Computer.FileSystem.DeleteFile(XtraForm1.DatabasePath)
            For y As Integer = 0 to CInt(SpinEdit1.EditValue)
                Dim newRow As DataRow = XtraForm1.DataSet1.Bidders.NewRow
                newRow(1) = IIf(y<1,"Min. Bidder","Bidder" & y)
                XtraForm1.DataSet1.Tables(0).Rows.Add(newRow)

                Dim newRow2 As DataRow = XtraForm1.DataSet1.Price.NewRow
                newRow2(1) = IIf(y<1,"Min. Bidder","Bidder" & y)
                newRow2(3) = newRow(0)
                XtraForm1.DataSet1.Tables(1).Rows.Add(newRow2)
            Next
            XtraForm1.DataSet1.WriteXml(XtraForm1.DatabasePath)
        End If

        XtraForm1.EvaluationMethod = ComboBoxEdit2.SelectedIndex
        XtraForm1.ProjectRange = ComboBoxEdit1.SelectedIndex
        XtraForm1._CompanyName = TextEdit1.EditValue
        If ComboBoxEdit2.SelectedIndex <> 1 then XtraForm1.NavigationPane1.Pages(6).PageVisible = False
        xtraForm1.Show

        Me.Hide
End Sub


Private Sub SimpleButton2_Click( sender As Object,  e As EventArgs) Handles SimpleButton2.Click
        Setting.ShowDialog()
End Sub

Private Sub ComboBoxEdit1_SelectedIndexChanged( sender As Object,  e As EventArgs) Handles ComboBoxEdit1.SelectedIndexChanged
        If ComboBoxEdit1.EditValue = Nothing Then
            DxErrorProvider1.SetError(ComboBoxEdit1,"Missing Value")
        Else
            DxErrorProvider1.ClearErrors()
        End If
End Sub

Private Sub ComboBoxEdit2_SelectedIndexChanged( sender As Object,  e As EventArgs) Handles ComboBoxEdit2.SelectedIndexChanged
        If ComboBoxEdit2.EditValue = Nothing Then
            DxErrorProvider1.SetError(ComboBoxEdit2,"Missing Value")
        Else
            DxErrorProvider1.ClearErrors()
        End If
End Sub

    Private Sub StartupForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Are you sure you want to exit the application?", "Closing Application", MessageBoxButtons.YesNo,MessageBoxIcon.Asterisk) = Windows.Forms.DialogResult.Yes Then
            Me.Dispose
            Application.ExitThread
        Else
            e.Cancel = True
        End If
    End Sub

Private Sub SimpleButton3_Click( sender As Object,  e As EventArgs) Handles SimpleButton3.Click
        Me.Hide
        OpenNew.Show
End Sub
End Class