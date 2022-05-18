Public Class AddBidders 

    Sub New()
        InitSkins()
        InitializeComponent()
    End Sub
    Sub InitSkins()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
    End Sub

Private Sub AddBidders_Load( sender As Object,  e As EventArgs) Handles MyBase.Load

End Sub

Private Sub SimpleButton1_Click( sender As Object,  e As EventArgs) Handles SimpleButton1.Click
        
        For y As Integer = MaxInTheTable(XtraForm1.DataSet1.Tables(0)) + 1 to MaxInTheTable(XtraForm1.DataSet1.Tables(0)) + CInt(SpinEdit1.EditValue)
            'For y As Integer = XtraForm1.DataSet1.Tables(0).Rows.Count To XtraForm1.DataSet1.Tables(0).Rows.Count + CInt(SpinEdit1.EditValue) - 1
            Dim newRow As DataRow = XtraForm1.DataSet1.Bidders.NewRow
            newRow(1) = "Bidder" & y
            XtraForm1.DataSet1.Tables(0).Rows.Add(newRow)

            Dim NewRow2 As DataRow = XtraForm1.DataSet1.Price.NewRow
            newRow2(1) = "Bidder" & y
            XtraForm1.DataSet1.Tables(1).Rows.Add(NewRow2)

            XtraForm1.VGridControl1.RefreshDataSource
            XtraForm1.VGridControl2.RefreshDataSource()
            XtraForm1.VGridControl3.RefreshDataSource()
            XtraForm1.VGridControl4.RefreshDataSource()
            XtraForm1.VGridControl5.RefreshDataSource()
            XtraForm1.VGridControl6.RefreshDataSource()
            XtraForm1.VGridControl7.RefreshDataSource()
            Me.Close
            'XtraForm1.DataSet1.Tables(0).Rows.Add(XtraForm1.DataSet1.Bidders.NewRow)
        Next
End Sub
    Private Function MaxInTheTable(datatable As DataTable) As Integer
        Dim Max As Integer = 0
        For y As Integer = 0 to XtraForm1.DataSet1.Tables(0).Rows.Count-1
           If  Num(XtraForm1.DataSet1.Tables(0).Rows(y).Item("Name").ToString) > Max then
                Max = Num(XtraForm1.DataSet1.Tables(0).Rows(y).Item("Name").ToString)
           End If
        Next
        Return Max
    End Function
    Private shared Function Num(ByVal Text As String) As Integer
        If CheckForNumber(Text) Then
            Return Integer.Parse(System.Text.RegularExpressions.Regex.Replace(text, "[^\d]", ""))
        End If
End Function

    private shared Function CheckForNumber(strg As String) As Boolean
    Dim i As Integer
     
    For i = 1 To Len(strg)
        If IsNumeric(Mid(strg, i, 1)) Then
            CheckForNumber = True
            Exit Function
        End If
    Next i
     
End Function

        Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown   
            If e.KeyCode = Keys.Escape Then Me.Hide
        End Sub
        '    Dim myChars() As Char = Text.ToCharArray()
        '    For Each ch As Char In myChars
        '    If Char.IsDigit(ch) Then
        '        MessageBox.Show(ch)
        '    End If
        '    Next
            
        'Dim returnVal As String = String.Empty
        'Dim collection As system.Text.RegularExpressions.MatchCollection = system.Text.RegularExpressions.Regex.Matches(value, "\d+")
        'For Each m As system.Text.RegularExpressions.Match In collection
        '    returnVal += m.ToString()
        'Next
        'Return Convert.ToInt32(returnVal)
End Class