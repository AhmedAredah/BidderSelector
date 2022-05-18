Public Class XtraReport2

Private Sub XtraReport2_DataSourceDemanded( sender As Object,  e As EventArgs) Handles MyBase.DataSourceDemanded
        If DetailReport1.RowCount < 1 Then
            DetailReport1.Visible = False
            XrLabel7.Text = "No Bidder wins. " & vbNewLine & "All bidders are disqualified."
            XrLabel7.Font = New Font("Times New Roman", 11!, FontStyle.Bold, GraphicsUnit.Point, CType(0,Byte))
            XrLabel7.ForeColor = Color.Maroon
            XrLabel8.Visible = False
            GroupFooter2.Visible = False
        End If
        If DetailReport.RowCount < 1 Then
            DetailReport.Visible = False
        End If
End Sub
End Class