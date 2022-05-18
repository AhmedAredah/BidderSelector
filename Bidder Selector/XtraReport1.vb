Public Class XtraReport1


    Private Sub XtraReport1_DataSourceDemanded(sender As Object, e As EventArgs) Handles MyBase.DataSourceDemanded
        If DetailReport1.RowCount < 1 Then
            DetailReport1.Visible = False
            XrLabel7.Text = "No Bidder wins. " & vbNewLine & "All bidders are disqualified."
            XrLabel7.Font = New Font("Times New Roman", 11.0!, FontStyle.Bold, GraphicsUnit.Point, CType(0, Byte))
            XrLabel7.ForeColor = Color.Maroon
            XrLabel8.Visible = False
            'GroupFooter2.Visible = False
        End If
        If DetailReport.RowCount < 1 Then
            GroupHeader1.Visible = False
            Detail1.Visible = False
            'DetailReport.Visible = False
        End If

        If DetailReport2.RowCount < 1 Then
            GroupHeader3.Visible = False
            Detail3.Visible = False
            XrRichText2.Visible = False
            ReportFooter.HeightF = 160
            'DetailReport.Visible = False
        End If

        'If Parameters("Recalculate").Value = True Then
        '    Detail.Visible = True
        '    GroupHeader1.Visible = False
        '    Detail1.Visible = False
        '    'DetailReport.Visible = False
        '    XrLabel7.Visible = False
        '    XrLabel8.Visible = False
        'End If
    End Sub

    Private Sub XtraReport1_ParametersRequestBeforeShow(sender As Object, e As DevExpress.XtraReports.Parameters.ParametersRequestEventArgs) Handles MyBase.ParametersRequestBeforeShow
        If Parameters("Recalculate").Value = True Then
            Detail.Visible = True
            DetailReport.Visible = False
            XrLabel7.Visible = False
            XrLabel8.Visible = False
        End If
    End Sub


End Class