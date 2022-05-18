Public Class RemoveBidders 
        Sub New()
        InitSkins()
        InitializeComponent()
    End Sub
    Sub InitSkins()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
    End Sub
Private Sub RemoveBidders_Load( sender As Object,  e As EventArgs) Handles MyBase.Load
        CheckEdit1.Checked=True
        LayoutControlItem1.Enabled=True
        LayoutControlItem4.Enabled=False
        SearchLookUpEdit1.Properties.DataSource = XtraForm1.BiddersBindingSource
        'colName.OptionsFilter.AutoFilterCondition = DevExpress.XtraGrid.Columns.AutoFilterCondition.Contains
End Sub

Private Sub CheckEdit1_CheckedChanged( sender As Object,  e As EventArgs)  Handles CheckEdit1.CheckedChanged
        If CheckEdit1.Checked Then CheckEdit2.Checked = False
        LayoutControlItem1.Enabled=True
        LayoutControlItem4.Enabled=False
End Sub

Private Sub CheckEdit2_CheckedChanged( sender As Object,  e As EventArgs) Handles CheckEdit2.CheckedChanged
        If CheckEdit2.Checked Then CheckEdit1.Checked = False
        LayoutControlItem1.Enabled=False
        LayoutControlItem4.Enabled=True
End Sub


Private Sub SearchLookUpEdit1_Popup( sender As Object,  e As EventArgs) Handles SearchLookUpEdit1.Popup
    SearchLookUpEdit1View.ActiveFilterCriteria = DevExpress.Data.Filtering.CriteriaOperator.Parse(String.Format("Index <> {0}",0))
End Sub

Private Sub SimpleButton1_Click( sender As Object,  e As EventArgs) Handles SimpleButton1.Click
        If CheckEdit1.Checked Then
            If SearchLookUpEdit1.EditValue = Nothing Then Exit Sub
            Dim result1() As DataRow = XtraForm1.DataSet1.Tables(0).Select("Index = " & SearchLookUpEdit1.EditValue)
            
            For Each row As DataRow In result1
                row.Delete
            Next

            If XtraForm1.DataSet1.Tables(1).Rows.Count > 0 Then
            Dim result2() As DataRow = XtraForm1.DataSet1.Tables(1).Select("BidderID = " & SearchLookUpEdit1.EditValue)
            For Each row As DataRow In result2
                row.Delete
            Next
            End If

            'XtraForm1.DataSet1.Tables(0).Rows.RemoveAt(XtraForm1.DataSet1.Tables(0).Rows.
            ''XtraForm1.DataSet1.Tables(0).Rows( CInt(SearchLookUpEdit1.EditValue)).Delete
        Else
            If SpinEdit1.Value > XtraForm1.DataSet1.Tables(0).Rows.Count-1 Then
                MsgBox(String.Format("Only {0} bidders are available{1}Deleting the Min. bidder is not allowed", XtraForm1.DataSet1.Tables(0).Rows.Count - 1, vbNewLine),MsgBoxStyle.Critical, "Warning")
                Exit Sub
            Else
                For y As Integer = XtraForm1.DataSet1.Tables(0).Rows.Count -1 to XtraForm1.DataSet1.Tables(0).Rows.Count - SpinEdit1.Value Step -1
                    XtraForm1.DataSet1.Tables(0).Rows(y).Delete
                Next
                For y As Integer = XtraForm1.DataSet1.Tables(1).Rows.Count -1 to XtraForm1.DataSet1.Tables(1).Rows.Count - SpinEdit1.Value Step -1
                    XtraForm1.DataSet1.Tables(1).Rows(y).Delete
                Next
            End If
        End If
        XtraForm1.VGridControl1.RefreshDataSource()
        XtraForm1.VGridControl2.RefreshDataSource()
        XtraForm1.VGridControl3.RefreshDataSource()
        XtraForm1.VGridControl4.RefreshDataSource()
        XtraForm1.VGridControl6.RefreshDataSource()
        XtraForm1.VGridControl7.RefreshDataSource()
        XtraForm1.VGridControl5.RefreshDataSource()
        Me.hide
End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown   
        If e.KeyCode = Keys.Escape Then Me.Hide
    End Sub

End Class