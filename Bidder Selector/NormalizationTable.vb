Public partial Class NormalizationTable

    private  _Score(,) As Double     'The Matrix
    Private _RowData() As Double     'The Sum of each row in the Matrix 
    Private _TotalColumnSumValue As Double     'The total Sum of the Rows in the matrix
    Private _Normalized() As Double 'The normalized column (Each row sum/_totalsum)  
    Private _BiddersNumber As Integer

    'Public Sub new()
    '    Name = ""
    '    _Score = Nothing
    '    _RowData = Nothing
    '    _TotalColumnSumValue = 0.0
    '    _Normalized = Nothing
    '    _BiddersNumber = 0
    'End Sub

        Public Sub new(ByVal BidderScore As Double(,), ByVal BidderName As String)
        Name = ""
        _Score = Nothing
        _RowData = Nothing
        _TotalColumnSumValue = 0.0
        _Normalized = Nothing
        _BiddersNumber = 0
        Name = BidderName
        Score = BidderScore
    End Sub
    Public ReadOnly Property TotalSum As Double
        Get
            Return _TotalColumnSumValue
        End Get
    end property

    Public Property Name As String

    Public WriteOnly Property Score As  Double(,)
        Set (ByVal Value As Double(,))
            _Score = value
            _RowData = RowAverageValues(_Score)
            _TotalColumnSumValue = Sum(_RowData)
            _Normalized = GetNormalizedColumn(_TotalColumnSumValue, _RowData)
            _BiddersNumber = GetBiddersNumber(_Score)
        End Set
    End property

    Public ReadOnly Property NormalizedColumn As  Double()
        Get
            Return _Normalized
        End Get
    End property

    Public ReadOnly Property BiddersNumber As  integer
        Get
            Return _BiddersNumber
        End Get
    End property

    Private Shared Function RowAverageValues(ByRef Score As Double(,)) As Double()
        If Score.GetUpperBound(0) = 0 Then Return Nothing
        Dim Total(Score.GetUpperBound(0)) As Double
        For I As Integer = 0 to Score.GetUpperBound(0)
            For II As Integer = 0 to Score.GetUpperBound(1)
                Total(I) = Total(I) + Score(I,II)
            Next
            Total(I) = Total(I) / (Score.GetUpperBound(0) + 1)
        Next
        Return Total
    End Function

    Private Shared Function GetBiddersNumber (ByRef Score As Double(,)) As Integer
        Return Score.GetUpperBound(0)
    End Function

    Private Shared Function Sum(Score As Double()) As Double
        If Score.Count = 0 Then Return Nothing
        Dim Total As Double
        For I As Integer = 0 to Score.GetUpperBound(0)
                Total = Total + Score(I)
        Next
        Return Total
    End Function

    Private Shared Function GetNormalizedColumn(ColumnsumValue As Double, RowData As Double()) As Double()
        If RowData.Count = 0 Or ColumnsumValue = 0 Or ColumnsumValue = Nothing Then Return Nothing
        Dim NormalCol(RowData.GetUpperBound(0)) As Double
        For I As Integer = 0 to RowData.GetUpperBound(0)
            NormalCol(I) = RowData(I)/ColumnsumValue
        Next
        Return NormalCol
    End Function

End Class
