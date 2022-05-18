

Public Class XtraForm1

#Region "Declaration"

    Public DatabasePath As String = Setting.GetApplicationLocation(Application.ExecutablePath) & "\MainDatabase.BEV"
    'Public DatabasePath As String = Application.ExecutablePath & "\MainDatabase.BEV"
    Private SavedValue As Boolean = True
    Public SavedLocation As String = String.Empty
    Public EvaluationMethod As Nullable(Of Integer)
    Private _ProjectRange As Integer
    Public _CompanyName As String = String.Empty
    Private SplashScreenManager1 = New DevExpress.XtraSplashScreen.SplashScreenManager()

    Private Attributes_NormalizedArray() As NormalizationTable

    'Start and end of each Technical section.
    Private Attributes_StartIndex(,) As Double = New Double(5, 1) {{2, 4}, {7, 4}, {12, 4}, {17, 4}, {22, 4}, {27, 4}}

    'Factors for Each Technical Section
    Private Attribute_PeriorityArray() As Double = {0.189071388358477, 0.457628077251897, 0.160680803702146, 0.0751843582316964, 0.117435372455784, _
                                                     0.230487822196725, 0.278204159782441, 0.0476609374469192, 0.263519062312398, 0.180128018261516, _
                                                     0.139226427672306, 0.112734824589294, 0.211941790523574, 0.255539700453861, 0.280557256760966, _
                                                     0.160120858819289, 0.0676015279897491, 0.160037296200185, 0.302250613089002, 0.309989703901775, _
                                                     0.0663425430513966, 0.0726479137000823, 0.470916326808672, 0.174851850587607, 0.215241365852242, _
                                                     0.254334658647225, 0.3533947943569, 0.222213996478905, 0.100478829791923, 0.0695777207250471}
    Private Tech_OutputArray As New TechnicalOutputArray



    Private TechnicalEvaluationWeights(,) As Double = {{0.47, 0.14, 0.147, 0.1445, 0.048, 0.0505}, _
                                                      {0.3845, 0.1355, 0.161, 0.1775, 0.066, 0.0755}, _
                                                       {0.3395, 0.1525, 0.155, 0.199, 0.0745, 0.0795}, _
                                                       {0.297826086956522, 0.150434782608696, 0.147826086956522, 0.218260869565217, 0.09, 0.0956521739130435}, _
                                                      {0.220454545454545, 0.169090909090909, 0.16, 0.247272727272727, 0.103181818181818, 0.1}}


    Private ProjectEstimatedCost As Double
    Private FinancialEvaluation As FinancialOutputArray

    Public Property ProjectRange As Integer
        Set(value As Integer)
            _ProjectRange = value
        End Set
        Get
            Return _ProjectRange
        End Get
    End Property
#End Region


#Region "Helping Functions and subs"

    Public Property Saved() As Boolean
        Get
            Return SavedValue
        End Get

        Set(ByVal value As Boolean)
            If Not (value = SavedValue) Then
                SavedValue = value
                If Saved Then
                    Text = "Evaluation Criteria"
                Else
                    Text = "Evaluation Criteria *"
                End If
            End If
        End Set
    End Property

    Private Shared ReadOnly Property ClipboardData() As String
        Get
            Dim Idata As IDataObject = Clipboard.GetDataObject
            If Idata.ToString = "" Then Return ""
            If Idata.GetDataPresent(DataFormats.UnicodeText) Then
                Return DirectCast(Idata.GetData(DataFormats.UnicodeText), String)
            End If
            Return ""
        End Get
        'Set(value As String)
        '    Clipboard.SetDataObject(value)
        'End Set
    End Property

    Public Sub WaitForm(ByRef ParentForm As DevExpress.XtraEditors.XtraForm, Optional ToShow As Boolean = True)
        If ToShow Then
            SplashScreenManager1 = New DevExpress.XtraSplashScreen.SplashScreenManager _
                        (ParentForm, GetType(WaitForm1), True, True) With {.ClosingDelay = 500}
            SplashScreenManager1.ShowWaitForm()
        Else
            If SplashScreenManager1.IsSplashFormVisible Then SplashScreenManager1.CloseWaitForm()
        End If
    End Sub

    Sub LoadInformation()
        WaitForm(Me, True)
        If Settings.Tables(0).Rows.Count > 0 Then Settings.Tables(0).Clear()
        If Settings.Tables(1).Rows.Count > 0 Then Settings.Tables(1).Clear()
        If DataSet1.Tables(0).Rows.Count > 0 Then DataSet1.Tables(0).Clear()
        If DataSet1.Tables(1).Rows.Count > 0 Then DataSet1.Tables(1).Clear()
        Try
            If My.Computer.FileSystem.FileExists(DatabasePath) Then DataSet1.ReadXml(DatabasePath)
            If My.Computer.FileSystem.FileExists(Setting.Path) Then Settings.ReadXml(Setting.Path)
            'ApplyingFilters()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        WaitForm(Me, False)
    End Sub

    Private Shared Function validateDatatableHasData(DT As DataTable) As Boolean
        If DT.Rows.Count < 1 Then Return False
        For I As Integer = 0 To DT.Rows.Count - 1
            For II As Integer = 0 To DT.Columns.Count - 1
                If (DT.Rows(I)(II).GetType() Is GetType(Integer) Or DT.Rows(I)(II).GetType() Is GetType(Double)) Then
                    If DT.Rows(I)(II) = 0 Then Return False
                Else
                    If DT.Rows(I)(II).ToString = "" Then Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Function ValidateAnalysisOptions() As Boolean
        If Not BarEditItem1.EditValue = "" And Not BarEditItem2.EditValue = "" Then Return True
        Return False
    End Function


    Private Sub PasteTheData(Data As String)
        If Data = String.Empty Then Return
        Dim rowData As String() = Data.Split(New Char() {ControlChars.Cr, ControlChars.Tab})
        Dim RowDataCopy(rowData.GetUpperBound(0)) As Double
        For X As Integer = 0 To rowData.GetUpperBound(0)
            If ReturnScoreValue(rowData(X)).HasValue Then
                RowDataCopy(X) = ReturnScoreValue(rowData(X))
            Else
                MessageBox.Show(String.Format("Data is not formated correctly.{0}Data might have a typo.{0}Recheck your data again and repeat the process", vbNewLine), buttons:=MessageBoxButtons.OK, icon:=MessageBoxIcon.Error, caption:="Error")
                Exit Sub
            End If

        Next
        Dim newRow As DataRow = DataSet1.Tables(0).NewRow
        Dim I As Integer = 0
        While I < RowDataCopy.Length
            If I >= DataSet1.Tables(0).Columns.Count Then Exit While
            newRow(I) = RowDataCopy(I)
            Math.Max(Threading.Interlocked.Increment(I), I - 1)
        End While
    End Sub

    Private Sub ExitEditors()
        VGridControl1.CloseEditor()
        VGridControl2.CloseEditor()
        VGridControl3.CloseEditor()
        VGridControl4.CloseEditor()
        VGridControl5.CloseEditor()
        VGridControl6.CloseEditor()
        VGridControl7.CloseEditor()
    End Sub
#End Region

#Region "Form Actions"

    Private Sub XtraForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AddHandler NavigationPage1.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick
        AddHandler NavigationPage2.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick
        AddHandler NavigationPage3.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick
        AddHandler NavigationPage4.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick
        AddHandler NavigationPage5.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick
        AddHandler NavigationPage6.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick
        AddHandler NavigationPage7.CustomButtonClick, AddressOf NavigationPage1_CustomButtonClick

        If BarEditItem2.EditValue = "Technical and Financial Evaluation" Then
            NavigationPage6.CustomHeaderButtons.Item("Next").Properties.Visible = True
            NavigationPage6.CustomHeaderButtons.Item("|").Properties.Visible = True
        Else
            NavigationPage6.CustomHeaderButtons.Item("Next").Properties.Visible = False
            NavigationPage6.CustomHeaderButtons.Item("|").Properties.Visible = False
        End If
    End Sub

    Private Sub XtraForm1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        If Settings.Tables(0).Rows.Count < 23 Then
            Setting.LoadDefaultSettingsWithnoCheck()
            'ApplyingFilters()
            Settings.Clear()
            If My.Computer.FileSystem.FileExists(Setting.Path) Then Settings.ReadXml(Setting.Path)
        End If
        LoadInformation()

        BarEditItem1.EditValue = RepositoryItemComboBox1.Items(ProjectRange)
        BarEditItem4.EditValue = _CompanyName

        If EvaluationMethod.HasValue Then
            BarEditItem2.EditValue = RepositoryItemComboBox2.Items(EvaluationMethod)
        Else
            If DataSet1.Tables(1).Rows.Count > 0 Then
                If Not DataSet1.Tables(1).Rows(0)(2).ToString = "" Then
                    BarEditItem2.EditValue = RepositoryItemComboBox2.Items(1)
                Else
                    BarEditItem2.EditValue = RepositoryItemComboBox2.Items(0)
                End If
            Else
                BarEditItem2.EditValue = RepositoryItemComboBox2.Items(0)
            End If
        End If
        BarEditItem3.EditValue = CDbl(IIf(String.IsNullOrEmpty(DataSet1.Tables(1).Rows(0).Item(2).ToString), 0, DataSet1.Tables(1).Rows(0).Item(2)))
    End Sub

    'Private Sub ApplyingFilters
    '    RepositoryItemSearchLookUpEdit4View.ActiveFilterString = "GroupID = 3"
    '    RepositoryItemSearchLookUpEdit9View.ActiveFilterString = "GroupID = 2"
    '    RepositoryItemSearchLookUpEdit8View.ActiveFilterString = "GroupID = 1"
    'End Sub
    Private Sub XtraForm1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("Are you sure you want to exit the application?", "Closing Application", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) = Windows.Forms.DialogResult.Yes Then
            'Me.Dispose()
            'Application.Exit
            Application.ExitThread()
        Else
            e.Cancel = True
        End If
        'StartupForm.Show
    End Sub

    Private Sub VGridControl_CellValueChanged(sender As Object, e As DevExpress.XtraVerticalGrid.Events.CellValueChangedEventArgs) Handles VGridControl1.CellValueChanged, VGridControl2.CellValueChanged, VGridControl3.CellValueChanged, VGridControl4.CellValueChanged, VGridControl5.CellValueChanged, VGridControl6.CellValueChanged, VGridControl7.CellValueChanged
        Try
            DataSet1.AcceptChanges()
            TryCast(sender, DevExpress.XtraVerticalGrid.VGridControl).Refresh()
            DataSet1.WriteXml(DatabasePath)
            Saved = False
            If Not BarEditItem2.EditValue = "Technical and Financial Evaluation" Then
                If validateDatatableHasData(DataSet1.Tables(0)) And ValidateAnalysisOptions() And Not String.IsNullOrEmpty(CStr(BarEditItem3.EditValue)) Then
                    BarButtonItem4.Enabled = True
                Else
                    BarButtonItem4.Enabled = False
                End If
            Else
                If validateDatatableHasData(DataSet1.Tables(0)) And validateDatatableHasData(DataSet1.Tables(1)) And ValidateAnalysisOptions() Then
                    BarButtonItem4.Enabled = True
                Else
                    BarButtonItem4.Enabled = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub RibbonControl1_ApplicationButtonClick(sender As Object, e As EventArgs) Handles RibbonControl1.ApplicationButtonClick
        AboutBox1.Show()
    End Sub

    Private Sub BarButtonItem3_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem3.ItemClick
        Setting.ShowDialog()
    End Sub

    Private Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
        If SavedLocation.Length = 0 Then
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                SavedLocation = SaveFileDialog1.FileName
            End If
        End If

        If System.IO.File.Exists(String.Format("{0}\{1}", Application.StartupPath(), DatabasePath)) Then
            System.IO.File.Copy(String.Format("{0}\{1}", Application.StartupPath(), DatabasePath), SavedLocation, True)
            Saved = True
        Else
            Dim writer As Xml.XmlTextWriter = New Xml.XmlTextWriter(SavedLocation, Nothing)
            writer.Formatting = Xml.Formatting.Indented
            writer.WriteStartDocument()
            writer.WriteComment("This is the Bidder Evaluator Main database. This database is not to be changed manually!")
            DataSet1.WriteXml(writer)
            Saved = True
        End If
    End Sub

    Private Sub BarButtonItem2_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        SaveFileDialog1.Title = "Save as"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            SavedLocation = SaveFileDialog1.FileName
        End If

        If System.IO.File.Exists(String.Format("{0}\{1}", Application.StartupPath(), DatabasePath)) Then
            System.IO.File.Copy(String.Format("{0}\{1}", Application.StartupPath(), DatabasePath), SavedLocation, True)
            Saved = True
        Else
            Dim writer As Xml.XmlTextWriter = New Xml.XmlTextWriter(SavedLocation, Nothing)
            writer.Formatting = Xml.Formatting.Indented
            writer.WriteStartDocument()
            writer.WriteComment("This is the Bidder Evaluator Main database. This database is not to be changed manually!")
            DataSet1.WriteXml(writer)
            Saved = True
        End If
    End Sub

    Private Sub BarButtonItem5_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem5.ItemClick
        'Dim Bidders As New AddBidders
        AddBidders.ShowDialog()
    End Sub

    Private Sub BarButtonItem6_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem6.ItemClick
        RemoveBidders.ShowDialog()
    End Sub

    Private Sub VGridControl1_KeyDown(sender As Object, e As KeyEventArgs) Handles VGridControl1.KeyDown, VGridControl2.KeyDown, VGridControl3.KeyDown, VGridControl4.KeyDown, VGridControl6.KeyDown, VGridControl7.KeyDown
        Dim VView As DevExpress.XtraVerticalGrid.VGridControl = TryCast(sender, DevExpress.XtraVerticalGrid.VGridControl)
        If e.KeyCode = Keys.Control AndAlso e.KeyCode = Keys.C Then
            VView.CopyToClipboard()
        End If
    End Sub

    Private Sub BarButtonItem7_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem7.ItemClick
        Dim data As String() = ClipboardData.Split(ControlChars.Lf)
        If data.Length < 1 Then Return : MsgBox("Not compatible data", MsgBoxStyle.Critical, "Error")
        If MessageBox.Show(String.Format("Are you sure you want to import the data from clipboard and remove the current data.{0}All the currect data will be erased!", _
                vbNewLine), icon:=MessageBoxIcon.Information, buttons:=MessageBoxButtons.YesNo, caption:="Information") = Windows.Forms.DialogResult.No Then Exit Sub
        For Each row As String In data
            PasteTheData(row)
        Next
    End Sub

    Private Sub BarEditItem1_EditValueChanged(sender As Object, e As EventArgs) Handles BarEditItem1.EditValueChanged, BarEditItem2.EditValueChanged
        If TryCast(sender, DevExpress.XtraBars.BarEditItem).Name = "BarEditItem2" Then
            If BarEditItem2.EditValue = "Technical Evaluation Only" Then
                BarEditItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                If validateDatatableHasData(DataSet1.Tables(0)) And ValidateAnalysisOptions() Then
                    BarButtonItem4.Enabled = True
                Else
                    BarButtonItem4.Enabled = False
                End If
                NavigationPane1.Pages(6).PageVisible = False
                NavigationPage6.CustomHeaderButtons.Item("Next").Properties.Visible = False
                NavigationPage6.CustomHeaderButtons.Item("|").Properties.Visible = False
                If NavigationPane1.SelectedPageIndex = 6 Then NavigationPane1.SelectedPageIndex = 5
                'DataSet1.Tables(1).Clear()

            ElseIf BarEditItem2.EditValue = "Technical and Financial Evaluation" Then
                BarEditItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                If validateDatatableHasData(DataSet1.Tables(0)) And validateDatatableHasData(DataSet1.Tables(1)) And ValidateAnalysisOptions() Then
                    BarButtonItem4.Enabled = True
                Else
                    BarButtonItem4.Enabled = False
                End If
                NavigationPane1.Pages(6).PageVisible = True
                NavigationPage6.CustomHeaderButtons.Item("Next").Properties.Visible = True
                NavigationPage6.CustomHeaderButtons.Item("|").Properties.Visible = True
                If DataSet1.Tables(0).Rows.Count > 0 And DataSet1.Tables(1).Rows.Count < 1 Then
                    For I As Integer = 0 To DataSet1.Tables(0).Rows.Count - 1
                        Dim NewRow2 As DataRow = DataSet1.Price.NewRow
                        NewRow2(1) = DataSet1.Bidders.Rows(I).Item(1).ToString
                        NewRow2(3) = I
                        DataSet1.Tables(1).Rows.Add(NewRow2)
                    Next
                End If
            Else
                NavigationPane1.Pages(6).PageVisible = True
                NavigationPage6.CustomHeaderButtons.Item("Next").Properties.Visible = True
                NavigationPage6.CustomHeaderButtons.Item("|").Properties.Visible = True
            End If
        End If
    End Sub

    Private Sub BarButtonItem4_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem4.ItemClick
        WaitForm(Me, True)
        ExitEditors()
        'Technical Evaluation
        TechnicalEvaluation(BarEditItem1.EditValue.ToString)

        'Financial Evaluation
        If BarEditItem2.EditValue = "Technical and Financial Evaluation" Then
            Dim BidPrice(DataSet1.Tables(1).Rows.Count - 1) As Double
            For I As Integer = 0 To BidPrice.GetUpperBound(0)
                BidPrice(I) = DataSet1.Tables(1).Rows(I).Item(2)
            Next
            If DataSet1.Tables(1).Rows.Count > 0 Then
                ProjectEstimatedCost = CDbl(BarEditItem3.EditValue.ToString)
                DataSet1.Tables(1).Rows(0).Item(2) = ProjectEstimatedCost
                CalFinance(Tech_OutputArray.MinValue, CDbl(DataSet1.Tables(1).Rows(0).Item(2)), BidPrice, BarEditItem1.EditValue.ToString)
                'CDbl(DataSet1.Tables(1).Rows(0).Item(2))
            End If
            'CalFinance(CDbl(Tech_OutputArray.AttributeOutputArray(0)), CDbl(DataSet1.Tables(1).Rows(0).Item(2)), BidPrice , BarEditItem1.EditValue.ToString)
        End If
        BarButtonItem11.Enabled = True
        BarEditItem2.Enabled = False
        BarEditItem1.Enabled = False
        BarButtonItem4.Enabled = False
        'BarSubItem1.Enabled = True
        BarButtonItem12.Enabled = True

        NavigationPane1.Enabled = False
        BarEditItem3.Enabled = False
        BarEditItem4.Enabled = False
        WaitForm(Me, False)
    End Sub

    Private Sub BarButtonItem13_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem13.ItemClick
        'Dim Report As new XtraReport1
        'Report.DataSource = DataSet21
    End Sub

    Private Sub BarButtonItem11_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem11.ItemClick
        Dim X As New XtraForm2
        If BarEditItem2.EditValue = RepositoryItemComboBox2.Items(0) Then
            Dim Report As New XtraReport2()

            Report.Parameters("MinBidderTechnicalValue").Value = Math.Round(Tech_OutputArray.MinValue, 4)
            Dim row As DataRow = (From column In DataSet21.Tables(0).Rows Where column("TechnicalValue") = Math.Round(Tech_OutputArray.MaxValue, 4)).FirstOrDefault()
            If Not IsNothing(row) Then Report.Parameters("WinnerName").Value = row("Name").ToString

            Report.Parameters("ProjectScale").Value = BarEditItem1.EditValue
            Report.Parameters("CompanyName").Value = BarEditItem4.EditValue

            Report.DataSource = DataSet21
            X.DocumentViewer1.DocumentSource = Report

        ElseIf BarEditItem2.EditValue = RepositoryItemComboBox2.Items(1) Then
            Dim Report As New XtraReport1()

            Report.Parameters("MinBidderTechnicalValue").Value = Math.Round(Tech_OutputArray.MinValue, 4)
            Report.Parameters("ProjectScale").Value = BarEditItem1.EditValue

            If Not FinancialEvaluation.ReCalculateWithDifferentData Then
                Dim row As DataRow = (From column In DataSet21.Tables(0).Rows Where column("TechnicalValue") > Math.Round(Tech_OutputArray.MinValue, 4) _
                And Math.Round(column("FinancialValue"), 4) = Math.Round(FinancialEvaluation.MinFinancialValue, 4)).FirstOrDefault()
                If Not IsNothing(row) Then Report.Parameters("WinnerName").Value = row("Name").ToString Else Report.Parameters("WinnerName").Value = "NO Name"
            End If


            Report.Parameters("CompanyName").Value = BarEditItem4.EditValue
            Report.Parameters("ProjectEstimatedCost").Value = CDbl(BarEditItem3.EditValue)
            Report.Parameters("Recalculate").Value = CBool(FinancialEvaluation.ReCalculateWithDifferentData)


            Report.DataSource = DataSet21
            X.DocumentViewer1.DocumentSource = Report
        End If

        X.ShowDialog()

        'Dim printTool As New DevExpress.XtraReports.UI.ReportPrintTool(Report)

        'printTool.ShowRibbonPreview()
    End Sub


    Private Sub BarButtonItem12_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem12.ItemClick
        BarEditItem2.Enabled = True
        BarEditItem1.Enabled = True
        BarButtonItem11.Enabled = False
        'BarSubItem1.Enabled = False
        DataSet21.Clear()
        BarButtonItem4.Enabled = True
        BarButtonItem12.Enabled = False
        NavigationPane1.Enabled = True
        BarEditItem3.Enabled = True
        BarEditItem4.Enabled = True
    End Sub

    Private Sub NavigationPane1_Paint(sender As Object, e As PaintEventArgs) Handles NavigationPane1.Paint
        Try
            If NavigationPane1.Pages(6).PageVisible Then
                Dim aPen As Pen = New Pen(Color.SlateGray, 2.0F)
                Dim x As DevExpress.XtraEditors.ButtonPanel.BaseButtonsPanelViewInfo = NavigationPane1.ButtonsPanel.ViewInfo
                e.Graphics.DrawLine(aPen, x.Buttons(5).Bounds.X + 10, x.Buttons(5).Bounds.Y + x.Buttons(5).Bounds.Height, x.Buttons(5).Bounds.X + x.Buttons(5).Bounds.Width - 10, x.Buttons(5).Bounds.Y + x.Buttons(5).Bounds.Height)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub NavigationPage1_CustomButtonClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.Docking2010.ButtonEventArgs)
        If e.Button.Properties.Tag = "Next" AndAlso NavigationPane1.Pages(NavigationPane1.SelectedPageIndex + 1).PageVisible = True Then
            NavigationPane1.SelectNextPage()
        ElseIf e.Button.Properties.Tag = "Back" AndAlso NavigationPane1.Pages(NavigationPane1.SelectedPageIndex - 1).PageVisible = True Then
            NavigationPane1.SelectPrevPage()
        End If
    End Sub

    Private Sub BarButtonItem8_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem8.ItemClick
        If DataSet1.Tables(0).Rows.Count < 1 Then MsgBox("No data in the project", MsgBoxStyle.Critical, "Error") : Exit Sub
        For i As Integer = 0 To DataSet1.Tables(0).Columns.Count - 1
            If DataSet1.Tables(0).Rows(0).Item(i).ToString = "" Then MsgBox("Min. bidder is empty or a cell in Min. bidder is missing a value", MsgBoxStyle.Critical, "Error") : Exit Sub
        Next
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using DS As DataSet = DataSet1.Copy
                For i As Integer = 1 To DS.Tables(0).Rows.Count - 1
                    DS.Tables(0).Rows(i).Delete()
                Next
                DS.WriteXml(SaveFileDialog1.FileName)
            End Using
        End If
    End Sub

    Private Sub BarButtonItem9_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem9.ItemClick
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

            Using DS As DataSet = DataSet1.Clone
                'DT = DataSet1.Tables(0)
                DS.ReadXml(OpenFileDialog1.FileName)
                ExitEditors()

                If DS.Tables(0).Rows.Count < 1 Then Exit Sub
                For I As Integer = 0 To DS.Tables(0).Columns.Count - 1
                    DataSet1.Tables(0).Rows(0).Item(I) = DS.Tables(0).Rows(0).Item(I)
                Next
                DataSet1.WriteXml(DatabasePath)

                VGridControl1.RefreshDataSource()
                VGridControl2.RefreshDataSource()
                VGridControl3.RefreshDataSource()
                VGridControl4.RefreshDataSource()
                VGridControl5.RefreshDataSource()
                VGridControl6.RefreshDataSource()
                VGridControl7.RefreshDataSource()
            End Using
        End If
    End Sub

    Private Sub VGridControl1_Validated(sender As Object, e As EventArgs) Handles VGridControl1.Validated, VGridControl2.Validated, VGridControl3.Validated, VGridControl4.Validated, VGridControl6.Validated, VGridControl7.Validated
        Dim CurrentGrid As DevExpress.XtraVerticalGrid.VGridControl = DirectCast(sender, DevExpress.XtraVerticalGrid.VGridControl)
        If CurrentGrid.FocusedRow.Index = 0 Then
            For I As Integer = 0 To DataSet1.Bidders.Rows.Count - 1
                DataSet1.Price.Rows(I).Item(1) = DataSet1.Bidders.Rows(I).Item(1).ToString
            Next
            Exit Sub
        End If
    End Sub

    Private Sub VGridControl5_Validated(sender As Object, e As EventArgs) Handles VGridControl5.Validated
        Dim CurrentGrid As DevExpress.XtraVerticalGrid.VGridControl = DirectCast(sender, DevExpress.XtraVerticalGrid.VGridControl)
        If CurrentGrid.FocusedRow.Index = 0 Then
            For I As Integer = 0 To DataSet1.Price.Rows.Count - 1
                DataSet1.Bidders.Rows(I).Item(1) = DataSet1.Price.Rows(I).Item(1).ToString
            Next
            Exit Sub
        End If
    End Sub

    Private Sub GridView1_ValidatingEditor(ByVal sender As Object, ByVal e As DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs) Handles VGridControl1.ValidatingEditor, _
                                            VGridControl2.ValidatingEditor, VGridControl3.ValidatingEditor, VGridControl4.ValidatingEditor, VGridControl6.ValidatingEditor, VGridControl7.ValidatingEditor, VGridControl5.ValidatingEditor
        Dim CurrentGrid As DevExpress.XtraVerticalGrid.VGridControl = DirectCast(sender, DevExpress.XtraVerticalGrid.VGridControl)

        If e.Value.GetType Is GetType(String) And CurrentGrid.FocusedRow.Index = 0 Then
            Exit Sub
        End If
        If CDbl(e.Value) = 0 Then
            e.Valid = False
            e.ErrorText = "Zero and null values are not accepted"
        End If
    End Sub

    Private Sub BarButtonItem10_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem10.ItemClick
        Dim ClipText As String = String.Empty
        Dim ClipRow As String = String.Empty

        'Header row
        For Each Col As DataColumn In DataSet1.Tables(0).Columns
            If Not String.IsNullOrEmpty(ClipRow) Then
                ClipRow += ControlChars.Tab
            End If

            ClipRow += Col.ColumnName
        Next

        ClipText += ClipRow + ControlChars.NewLine

        'Data rows
        For Each row As DataRow In DataSet1.Tables(0).Rows
            ClipRow = String.Empty

            For Each col As DataColumn In DataSet1.Tables(0).Columns
                If Not String.IsNullOrEmpty(ClipRow) Then
                    ClipRow += ControlChars.Tab
                End If
                'If col.ColumnName = "FixedAndCurrentAssets" Or col.ColumnName = "Liquidity" Or col.ColumnName = "Previous financial penalties " Then
                '    ClipRow += row(col.ColumnName).ToString
                'Else
                If TypeOf (row(col.ColumnName)) Is Double Then
                    ClipRow += ReturnScoreName(row(col.ColumnName))
                Else
                    ClipRow += row(col.ColumnName).ToString
                End If

                'End If

            Next

            ClipText += ClipRow + ControlChars.NewLine
        Next

        Clipboard.SetText(ClipText)
    End Sub

#End Region

#Region "Analysis Subs and Functions"
    Private Shared Function ArrayComputations(ByRef datatable As DataTable, AttributeIndex As Integer) As Double(,)
        Dim BiddersNumber = datatable.Rows.Count - 1
        Dim ScoreArray(BiddersNumber, BiddersNumber) As Double

        For RowIndex As Integer = 0 To ScoreArray.GetUpperBound(0)
            For ColumnIndex As Integer = 0 To ScoreArray.GetUpperBound(1)
                ScoreArray(RowIndex, ColumnIndex) = datatable.Rows(RowIndex).Item(AttributeIndex) / datatable.Rows(ColumnIndex).Item(AttributeIndex)
            Next
        Next
        Return ScoreArray
    End Function

    'Generate the normalization table for each attribute
    ' Rows represents the attributes and Columns represents the Normalization table correspond to each attribute
    Private Function GenerateNormalizedTable(startIndex(,) As Double) As NormalizationTable() 'StartIndex is the start of each attribute, The first Column defines the start position and the second column defines how many attribute
        Dim _Array(startIndex(startIndex.GetUpperBound(0), 0) + startIndex(startIndex.GetUpperBound(0), 1) - startIndex(startIndex.GetLowerBound(0), 0)) As NormalizationTable
        _Array(0) = New NormalizationTable(ArrayComputations(DataSet1.Tables(0), 2), DataSet1.Tables(0).Columns(2).ColumnName)

        For I As Integer = 0 To _Array.GetUpperBound(0)
            _Array(I) = New NormalizationTable(ArrayComputations(DataSet1.Tables(0), I + 2), DataSet1.Tables(0).Columns(I + 2).ColumnName)
        Next
        Return _Array
    End Function

    Private Function GetAttributeGroup(AttributeIndex As Integer, StartIndex(,) As Double) As Integer  '0
        AttributeIndex = AttributeIndex + 2                                                            ' 2 7 12 17 22 25 30
        For I As Integer = 0 To StartIndex.GetUpperBound(0)
            If StartIndex(I, 0) > AttributeIndex Then
                Return I - 1
            ElseIf (StartIndex(I, 0) = AttributeIndex) Then
                Return I
            ElseIf (StartIndex(StartIndex.GetUpperBound(0), 0) < AttributeIndex) Then
                Return StartIndex.GetUpperBound(0)
            End If
        Next
        Return 0
    End Function

    Private Function GetBidderNormalizedValue(Attribute As NormalizationTable, bidderIndex As Integer) As Double
        Return Attribute.NormalizedColumn(bidderIndex)
    End Function

    Private Function GetBidderNormalizedArray(Attributes() As NormalizationTable, BidderIndex As Integer, _
                                              AttributeStartIndex As Integer, AttributeEndIndex As Integer) As Double()
        Dim Temp(AttributeEndIndex - AttributeStartIndex) As Double

        Dim ArrayIndex As Integer = 0


        For I As Integer = AttributeStartIndex - 2 To AttributeEndIndex - 2
            Temp(ArrayIndex) = Attributes(I).NormalizedColumn(BidderIndex)
            ArrayIndex = ArrayIndex + 1
        Next
        Return Temp
    End Function

    Private Function GetPeriorityArray(Periority As Double(), AttributeStartIndex As Integer, AttributeEndIndex As Integer) As Double()
        Dim Temp(AttributeEndIndex - AttributeStartIndex) As Double
        Dim ArrayIndex As Integer = 0
        For I As Integer = AttributeStartIndex - 2 To AttributeEndIndex - 2
            Temp(ArrayIndex) = Periority(I)
            ArrayIndex = ArrayIndex + 1
        Next
        Return Temp
    End Function

    Private Function GetSumofProductofPeriorityNormalized(Attributes() As NormalizationTable, Periority As Double(), _
                                                          BidderIndex As Integer, AttributeStartIndex As Integer, _
                                                          AttributeEndIndex As Integer) As Double
        Dim _Periority() As Double = GetPeriorityArray(Periority, AttributeStartIndex, AttributeEndIndex)

        Dim _BidderNormalizedArray() As Double = GetBidderNormalizedArray(Attributes, BidderIndex, AttributeStartIndex, AttributeEndIndex)

        Dim Result As Double = 0
        If _Periority.Count <> _BidderNormalizedArray.Count Then Exit Function
        For I As Integer = 0 To _Periority.GetUpperBound(0)
            Result = Result + (_Periority(I) * _BidderNormalizedArray(I))
        Next
        Return Result
    End Function

    Private Function GetDimensionOutPutArray(AttributesList() As NormalizationTable, Periority As Double(), StartIndex As Integer, EndIndex As Integer) As Double()
        Dim DimensionOutputArray(AttributesList(0).BiddersNumber) As Double
        For BidderIndex As Integer = 0 To DimensionOutputArray.GetUpperBound(0)
            DimensionOutputArray(BidderIndex) = GetSumofProductofPeriorityNormalized(AttributesList, Periority, BidderIndex, StartIndex, EndIndex)
        Next
        Return DimensionOutputArray 'Output for Bidders (Bidders as rows)
    End Function

    Private Function GetAllDimensionsOutputArray(AttributesList() As NormalizationTable, Periority As Double(), StartIndex(,) As Double) As Double()()
        Dim Temp()() As Double = New Double(CInt(StartIndex.GetUpperBound(0)))() {}
        For I As Integer = 0 To Temp.GetUpperBound(0)
            Temp(I) = GetDimensionOutPutArray(AttributesList, Periority, StartIndex(I, 0), StartIndex(I, 0) + StartIndex(I, 1))
        Next
        Return Temp 'Output for Bidders (Bidders as columns)
        '           Dimensions as rows
    End Function

    Public Enum ProjectSize As Integer
        LessThan5 = 0
        From5MTo100M = 1
        From100MTo250M = 2
        From250MTo500M = 3
        GreaterThan500M = 4
    End Enum

    Private Function GetTechnicalEvaluation(ProjectSize As ProjectSize, technicalEvaluationWeights As Double(,), _
                                        AttributesList() As NormalizationTable, Periority As Double(), _
                                        StartIndex(,) As Double) As Double()
        Dim temp(AttributesList(0).BiddersNumber) As Double
        For RowIndex_BidderIndex As Integer = 0 To temp.GetUpperBound(0)
            Dim AttributeGroup As Nullable(Of Integer) = Nothing
            For ColumnIndex_DimensionIndex As Integer = 0 To AttributesList.GetUpperBound(0)
                If AttributeGroup = GetAttributeGroup(ColumnIndex_DimensionIndex, StartIndex) Then Continue For

                temp(RowIndex_BidderIndex) = temp(RowIndex_BidderIndex) + (technicalEvaluationWeights(ProjectSize, GetAttributeGroup(ColumnIndex_DimensionIndex, StartIndex)) * _
                    GetAllDimensionsOutputArray(AttributesList, Periority, StartIndex)(GetAttributeGroup(ColumnIndex_DimensionIndex, StartIndex)).GetValue(RowIndex_BidderIndex))
                AttributeGroup = GetAttributeGroup(ColumnIndex_DimensionIndex, StartIndex)
            Next
            temp(RowIndex_BidderIndex) = temp(RowIndex_BidderIndex)
        Next
        Return temp
    End Function

    Private Sub TechnicalEvaluation(SelectedProject As String)
        'Dim NoBidders As Integer = DataSet1.Tables(0).Rows.Count
        'Create the Normalization Table

        Attributes_NormalizedArray = GenerateNormalizedTable(Attributes_StartIndex)


        Dim SelectedProjectSize As Integer
        'If String.Equals(SelectedProject, "< 5M") Then
        '    SelectedProjectSize = 0
        'ElseIf String.Equals(SelectedProject, "> 5M & < 100M") Then
        '    SelectedProjectSize = 1
        'ElseIf String.Equals(SelectedProject, "> 100M & < 250M") Then
        '    SelectedProjectSize = 2
        'ElseIf String.Equals(SelectedProject, "> 250M & < 500M") Then
        '    SelectedProjectSize = 3
        'ElseIf String.Equals(SelectedProject, "> 500M") Then
        '    SelectedProjectSize = 4
        'End If
        Select Case SelectedProject
            Case "< 5M"
                SelectedProjectSize = 0
            Case "> 5M & < 100M"
                SelectedProjectSize = 1
            Case "> 100M & < 250M"
                SelectedProjectSize = 2
            Case "> 250M & < 500M"
                SelectedProjectSize = 3
            Case "> 500M"
                SelectedProjectSize = 4
            Case Else
                Exit Select
        End Select
        Tech_OutputArray.AttributeOutputArray = GetTechnicalEvaluation(SelectedProjectSize, TechnicalEvaluationWeights, _
                                                        Attributes_NormalizedArray, Attribute_PeriorityArray, _
                                                        Attributes_StartIndex)

        DataSet21.Tables(0).Clear()
        For I As Integer = 1 To Tech_OutputArray.AttributeOutputArray.GetUpperBound(0) Step 1
            Dim newRow As DataRow = DataSet21.Tables(0).NewRow
            newRow(1) = DataSet1.Tables(0).Rows(I).Item(1)
            newRow(2) = Math.Round(Tech_OutputArray.AttributeOutputArray(I), 4)
            DataSet21.Tables(0).Rows.Add(newRow)
        Next
    End Sub

    Private Sub CalFinance(ByVal MinTechnical As Double, ByVal EstimatedCost As Double, ByVal BidPrice As Double(), ByVal SelectedProject As String)
        Dim SelectedProjectSize As Integer
        Select Case SelectedProject
            Case "< 5M"
                SelectedProjectSize = 0
            Case "> 5M & < 100M"
                SelectedProjectSize = 1
            Case "> 100M & < 250M"
                SelectedProjectSize = 2
            Case "> 250M & < 500M"
                SelectedProjectSize = 3
            Case "> 500M"
                SelectedProjectSize = 4
            Case Else
                Exit Select
        End Select

        FinancialEvaluation = New FinancialOutputArray(Tech_OutputArray.AttributeOutputArray, MinTechnical, EstimatedCost, SelectedProjectSize, BidPrice)
        If Not FinancialEvaluation.ReCalculateWithDifferentData Then
            For I As Integer = 0 To FinancialEvaluation.GetFinalFinancialProposal.GetUpperBound(0) - 1
                DataSet21.Tables(0).Rows(I)(3) = FinancialEvaluation.GetFinalFinancialProposal(I + 1)
                DataSet21.Tables(0).Rows(I)(4) = FinancialEvaluation.BidPrice(I + 1)
                DataSet21.Tables(0).Rows(I)(5) = FinancialEvaluation.GetTEP(I + 1)
                DataSet21.Tables(0).Rows(I)(6) = FinancialEvaluation.GetFinancialQualification(I + 1)
            Next
        End If

    End Sub

    Public Function ReturnScoreName(ByVal score As Double) As String
        If Settings.Tables("DataTable1").Rows.Count < 1 Then Setting.LoadDefaultSettingsWithnoCheck() ': ApplyingFilters()
        Dim ss As DataRow() = Settings.Tables("DataTable1").Select(String.Format("[Rating Value] = {0}", score))
        If ss.Count > 0 Then
            Return ss.First.Item("Rating Name").ToString
        Else
            Return score.ToString
        End If
    End Function

    Public Function ReturnScoreValue(ByVal score As String) As Double?
        If Settings.Tables("DataTable1").Rows.Count < 1 Then Setting.LoadDefaultSettingsWithnoCheck() ': ApplyingFilters()
        Dim ss As DataRow() = Settings.Tables("DataTable1").Select(String.Format("[Rating Name] = '{0}'", score))
        If ss.Count > 0 Then
            Return ss.First.Item("Rating Value").ToString
        Else
            Return Nothing
        End If
    End Function
#End Region

    Private Sub BarEditItem3_EditValueChanged(sender As Object, e As EventArgs) Handles BarEditItem3.EditValueChanged
        DataSet1.Tables(1).Rows(0).Item(2) = CDbl(BarEditItem3.EditValue)
    End Sub

    'Private Sub VGridControl5_ShowingEditor(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles VGridControl5.ShowingEditor
    '    If VGridControl5.FocusedRecord = 0 Then
    '        e.Cancel = True
    '    End If
    'End Sub

    'Private Sub VGridControl5_CustomDrawRowValueCell(sender As Object, e As DevExpress.XtraVerticalGrid.Events.CustomDrawRowValueCellEventArgs) Handles VGridControl5.CustomDrawRowValueCell

    '    'If e.RecordIndex = 0 Then
    '    '    e.CellText = "----"
    '    '    e.Handled = True
    '    'End If
    'End Sub
End Class