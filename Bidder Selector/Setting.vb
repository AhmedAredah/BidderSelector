
Public Class Setting

    Public Path As String = GetApplicationLocation(Application.ExecutablePath) & "\Database.xml"
    'Public Path As String = Application.ExecutablePath & "\Database.xml"

    Public Shared Function GetLnkTarget(lnkPath As String) As String
    Dim shl = New Shell32.Shell()
    ' Move this to class scope
    lnkPath = System.IO.Path.GetFullPath(lnkPath)
    Dim dir = shl.[NameSpace](System.IO.Path.GetDirectoryName(lnkPath))
    Dim itm = dir.Items().Item(System.IO.Path.GetFileName(lnkPath))
    Dim lnk = DirectCast(itm.GetLink, Shell32.ShellLinkObject)
    Return lnk.Target.Path
End Function

    Public Shared Function GetApplicationLocation(ApplicationLinkPath As String) As String
        If ApplicationLinkPath.EndsWith("lnk") Then
            If System.IO.File.Exists(ApplicationLinkPath) Then
                Return GetDirectoryFromFileName(GetLnkTarget(ApplicationLinkPath))
                'Dim shell As IWshRuntimeLibrary.WshShell = New IWshRuntimeLibrary.WshShell()
                'Dim link As IWshRuntimeLibrary.IWshShortcut = CType(shell.CreateShortcut(ApplicationLinkPath), IWshRuntimeLibrary.IWshShortcut)
                'Return GetDirectoryFromFileName(link.TargetPath)
            Else
                Return Application.StartupPath
            End If
        Else
            Return GetDirectoryFromFileName(ApplicationLinkPath)
        End If
    End Function

    Public Shared Function GetDirectoryFromFileName(ByVal FileFullPath As string) As String
        Do
            FileFullPath = FileFullPath.Remove(FileFullPath.Length - 1 , 1)
        Loop While Not FileFullPath.EndsWith("\")
        Return FileFullPath
    End Function

    Private Sub Setting_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadInformation()
        addDefaultRow()
    End Sub
    Sub LoadInformation()
        If Settings.Tables(0).Rows.Count > 0 Then Settings.Tables(0).Clear()
        If Settings.Tables(1).Rows.Count > 0 Then Settings.Tables(1).Clear()
        Try
            If My.Computer.FileSystem.FileExists(Path) Then Settings.ReadXml(Path)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub gridView1_RowUpdated(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.RowObjectEventArgs) Handles GridView1.RowUpdated
        If TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView).FocusedRowHandle = DevExpress.XtraGrid.GridControl.NewItemRowHandle Then Exit Sub
        Try
            Settings.AcceptChanges()
            GridControl1.Refresh()
            GridView1.RefreshData()
            Settings.WriteXml(Path)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub GridView1_CellValueChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles GridView1.CellValueChanged
        If TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView).FocusedRowHandle = DevExpress.XtraGrid.GridControl.NewItemRowHandle Then Exit Sub
        Try
            Settings.AcceptChanges()
            GridControl1.Refresh()
            GridView1.RefreshData()
            Settings.WriteXml(Path)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub GridControl3_KeyDown(sender As Object, e As KeyEventArgs) Handles GridControl1.KeyDown
        DeleteInsertRow(GridView1, e)
        Settings.AcceptChanges()
        GridControl1.Refresh()
        GridView1.RefreshData()
        Settings.WriteXml(Path)
    End Sub

    Public Shared Sub DeleteInsertRow(ByRef GridView As DevExpress.XtraGrid.Views.Grid.GridView, ByRef e As KeyEventArgs)
        If e.KeyData = Keys.Delete Then
            GridView.DeleteSelectedRows()
            GridView.RefreshData()
        End If
        If e.KeyData = Keys.Insert Then
            GridView.AddNewRow()
        End If
    End Sub

    Sub addDefaultRow()
        If Settings.Tables("weights").Rows.Count >= 1 Then Exit Sub
        Dim newRow As DataRow = Settings.Weights.NewRow
        newRow("name") = "Default"
        Settings.Tables("weights").Rows.Add(newRow)
    End Sub
    'Private Sub GridView1_RowCountChanged( sender As Object,  e As EventArgs) Handles GridView1.RowCountChanged
    '        Try
    '            Settings.AcceptChanges()
    '            GridControl1.Refresh
    '            GridView1.RefreshData
    '            Settings.WriteXml(Path)
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    'End Sub
    '    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles GridView1.CellEndEdit
    '    Settings.AcceptChanges()
    '    'GridView1.Refresh()
    '    Try
    '        Settings.WriteXml(Path)
    '    Catch ex As Exception
    '    End Try
    'End Sub
    'Private Sub GridView1_SelectionChanged(sender As Object, e As DevExpress.Data.SelectionChangedEventArgs) Handles GridView1.SelectionChanged
    '    If TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView).FocusedRowHandle = DevExpress.XtraGrid.GridControl.NewItemRowHandle Then Exit Sub
    '    Try
    '        Settings.AcceptChanges()
    '        GridControl1.Refresh()
    '        GridView1.RefreshData()
    '        Settings.WriteXml(Path)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub VGridControl1_CellValueChanged(sender As Object, e As DevExpress.XtraVerticalGrid.Events.CellValueChangedEventArgs) Handles VGridControl1.CellValueChanged, VGridControl2.CellValueChanged, VGridControl3.CellValueChanged, VGridControl4.CellValueChanged, VGridControl5.CellValueChanged, VGridControl6.CellValueChanged, VGridControl7.CellValueChanged
        Try
            Settings.AcceptChanges()
            TryCast(sender, DevExpress.XtraVerticalGrid.VGridControl).Refresh()
            Settings.WriteXml(Path)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Setting_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        XtraForm1.Settings.Clear()
        If My.Computer.FileSystem.FileExists(Path) Then XtraForm1.Settings.ReadXml(Path)
    End Sub
    Private Sub gridView2_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles GridView1.MouseDown
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            ContextMenuStrip1.Show(GridControl1, e.Location)
        End If
    End Sub
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        LoadDefaultSettings()
        GridControl1.RefreshDataSource()
        GridControl1.Refresh()
    End Sub
    Public Sub LoadDefaultSettingsWithnoCheck()

        Settings.Tables("DataTable1").Clear()

        Dim R_0 As DataRow = Settings.Tables("DataTable1").NewRow
        R_0("Rating Name") = "Not Submitted"
        R_0("Rating Value") = 0.000000001
        R_0("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R_0)

        Dim R0 As DataRow = Settings.Tables("DataTable1").NewRow
        R0("Rating Name") = "Very Poor"
        R0("Rating Value") = 1
        R0("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R0)

        Dim R1 As DataRow = Settings.Tables("DataTable1").NewRow
        R1("Rating Name") = "Poor"
        R1("Rating Value") = 2
        R1("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R1)

        Dim R2 As DataRow = Settings.Tables("DataTable1").NewRow
        R2("Rating Name") = "Moderately Poor"
        R2("Rating Value") = 3
        R2("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R2)

        Dim R3 As DataRow = Settings.Tables("DataTable1").NewRow
        R3("Rating Name") = "Fair"
        R3("Rating Value") = 4
        R3("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R3)

        Dim R04 As DataRow = Settings.Tables("DataTable1").NewRow
        R04("Rating Name") = "Moderately Good"
        R04("Rating Value") = 5
        R04("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R04)

        Dim R4 As DataRow = Settings.Tables("DataTable1").NewRow
        R4("Rating Name") = "Good"
        R4("Rating Value") = 6
        R4("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R4)

        Dim R5 As DataRow = Settings.Tables("DataTable1").NewRow
        R5("Rating Name") = "Very Good"
        R5("Rating Value") = 7
        R5("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R5)

        Dim R6 As DataRow = Settings.Tables("DataTable1").NewRow
        R6("Rating Name") = "Excellent"
        R6("Rating Value") = 8
        R6("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R6)

        Dim R06 As DataRow = Settings.Tables("DataTable1").NewRow
        R06("Rating Name") = "Outstanding"
        R06("Rating Value") = 9
        R06("GroupID") = 0
        Settings.Tables("DataTable1").Rows.Add(R06)

        'Penalties
        Dim R_1 As DataRow = Settings.Tables("DataTable1").NewRow
        R_1("Rating Name") = "Not Submitted"
        R_1("Rating Value") = 0.000000001
        R_1("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R_1)

        Dim R7 As DataRow = Settings.Tables("DataTable1").NewRow
        R7("Rating Name") = "Very low penalties"
        R7("Rating Value") = 9
        R7("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R7)

        Dim R71 As DataRow = Settings.Tables("DataTable1").NewRow
        R71("Rating Name") = "Low penalties"
        R71("Rating Value") = 8
        R71("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R71)


        Dim R8 As DataRow = Settings.Tables("DataTable1").NewRow
        R8("Rating Name") = "Moderately low penalties"
        R8("Rating Value") = 7
        R8("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R8)

        Dim R9 As DataRow = Settings.Tables("DataTable1").NewRow
        R9("Rating Name") = "Fair"
        R9("Rating Value") = 6
        R9("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R9)

        Dim R10 As DataRow = Settings.Tables("DataTable1").NewRow
        R10("Rating Name") = "Moderately high penalties"
        R10("Rating Value") = 5
        R10("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R10)

        Dim R11 As DataRow = Settings.Tables("DataTable1").NewRow
        R11("Rating Name") = "High penalties"
        R11("Rating Value") = 4
        R11("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R11)

        Dim R111 As DataRow = Settings.Tables("DataTable1").NewRow
        R111("Rating Name") = "Very high penalties "
        R111("Rating Value") = 3
        R111("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R111)

        Dim R0111 As DataRow = Settings.Tables("DataTable1").NewRow
        R0111("Rating Name") = "Severe penalties"
        R0111("Rating Value") = 2
        R0111("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R0111)

        Dim R00111 As DataRow = Settings.Tables("DataTable1").NewRow
        R00111("Rating Name") = "Not Accepted Penalties"
        R00111("Rating Value") = 1
        R00111("GroupID") = 1
        Settings.Tables("DataTable1").Rows.Add(R00111)


        'Contract not renewed
        Dim R_2 As DataRow = Settings.Tables("DataTable1").NewRow
        R_2("Rating Name") = "Not Submitted"
        R_2("Rating Value") = 0.000000001
        R_2("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R_2)

        Dim R12 As DataRow = Settings.Tables("DataTable1").NewRow
        R12("Rating Name") = "No Recorded Event"
        R12("Rating Value") = 9
        R12("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R12)

        Dim R012 As DataRow = Settings.Tables("DataTable1").NewRow
        R012("Rating Name") = "Not at all a problem"
        R012("Rating Value") = 8
        R012("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R012)

        Dim R13 As DataRow = Settings.Tables("DataTable1").NewRow
        R13("Rating Name") = "Minor Problem"
        R13("Rating Value") = 7
        R13("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R13)

        Dim R013 As DataRow = Settings.Tables("DataTable1").NewRow
        R013("Rating Name") = "Moderately Minor Problem"
        R013("Rating Value") = 6
        R013("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R013)

        Dim R14 As DataRow = Settings.Tables("DataTable1").NewRow
        R14("Rating Name") = "Fair"
        R14("Rating Value") = 5
        R14("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R14)

        Dim R15 As DataRow = Settings.Tables("DataTable1").NewRow
        R15("Rating Name") = "Moderately Major Problem"
        R15("Rating Value") = 4
        R15("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R15)

        Dim R16 As DataRow = Settings.Tables("DataTable1").NewRow
        R16("Rating Name") = "Major Problem"
        R16("Rating Value") = 3
        R16("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R16)

        Dim R161 As DataRow = Settings.Tables("DataTable1").NewRow
        R161("Rating Name") = "Sever Problem"
        R161("Rating Value") = 2
        R161("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R161)

        Dim R1610 As DataRow = Settings.Tables("DataTable1").NewRow
        R1610("Rating Name") = "Not Accepted Problem "
        R1610("Rating Value") = 1
        R1610("GroupID") = 2
        Settings.Tables("DataTable1").Rows.Add(R1610)

        'Past Failure
        Dim R_3 As DataRow = Settings.Tables("DataTable1").NewRow
        R_3("Rating Name") = "Not Submitted"
        R_3("Rating Value") = 0.000000001
        R_3("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R_3)

        Dim R17 As DataRow = Settings.Tables("DataTable1").NewRow
        R17("Rating Name") = "No Recorded Failure"
        R17("Rating Value") = 9
        R17("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R17)

        Dim R170 As DataRow = Settings.Tables("DataTable1").NewRow
        R170("Rating Name") = "Not at all a Failure"
        R170("Rating Value") = 8
        R170("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R170)

        Dim R18 As DataRow = Settings.Tables("DataTable1").NewRow
        R18("Rating Name") = "Minor Failure"
        R18("Rating Value") = 7
        R18("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R18)

        Dim R180 As DataRow = Settings.Tables("DataTable1").NewRow
        R180("Rating Name") = "Moderately Minor Failure"
        R180("Rating Value") = 6
        R180("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R180)

        Dim R19 As DataRow = Settings.Tables("DataTable1").NewRow
        R19("Rating Name") = "Fair"
        R19("Rating Value") = 5
        R19("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R19)

        Dim R20 As DataRow = Settings.Tables("DataTable1").NewRow
        R20("Rating Name") = "Moderately Major Failure"
        R20("Rating Value") = 4
        R20("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R20)

        Dim R21 As DataRow = Settings.Tables("DataTable1").NewRow
        R21("Rating Name") = "Major Failure"
        R21("Rating Value") = 3
        R21("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R21)

        Dim R211 As DataRow = Settings.Tables("DataTable1").NewRow
        R211("Rating Name") = "Severe Failure"
        R211("Rating Value") = 2
        R211("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R211)

        Dim R2111 As DataRow = Settings.Tables("DataTable1").NewRow
        R2111("Rating Name") = "Not Accepted Failure"
        R2111("Rating Value") = 1
        R2111("GroupID") = 3
        Settings.Tables("DataTable1").Rows.Add(R2111)

        Settings.WriteXml(Path)

    End Sub
    Public Sub LoadDefaultSettings()
        If MessageBox.Show("All the current data will be deleted", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Cancel Then Exit Sub
        LoadDefaultSettingsWithnoCheck()
    End Sub
End Class