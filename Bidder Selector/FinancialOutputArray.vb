Public Class FinancialOutputArray
    'Givens
    Private _TechnicalScore() As Double
    Private _MinTechnicalValue As Double
    Private _EstimatedCost As double
    Private _ProjectRange As XtraForm1.ProjectSize
    Private _BidPrices() As Double

    'For Calculations
    Private _FinalFinancialProposal() As Double
    Private _ApprovedMinBidPrice As Double
    Private Const _MinBidderPercent As Double = 0.82
    Private Const _MaxBidderPercent As Double = 1.18


    Public _AllBiddersAccepted As Boolean
    Public _OneBidderAccepted As Boolean
    Public _BidderEvaluation() As Integer
    Public _ReCalculateWithDifferentData As Boolean = False




    Public Sub New(ByVal TechnicalSchore As Double(), ByVal MinTechnical As Double, ByVal EstimatedCost As Double, _
                    ByVal ProjectRange As XtraForm1.ProjectSize, ByVal TheBidPrices As Double())
        _FinalFinancialProposal = Nothing

        _TechnicalScore = TechnicalSchore
        _MinTechnicalValue = MinTechnical
        _EstimatedCost = EstimatedCost
        _ProjectRange = ProjectRange
        _BidPrices = TheBidPrices
        SetTechnicalScore = TechnicalSchore
    End Sub

    Public ReadOnly Property BiddersCount() As Integer
        Get
            Return _TechnicalScore.Count
        End Get
    End Property


    Public WriteOnly Property SetTechnicalScore As Double()
        Set(value As Double())
            _TechnicalScore = value
            _AllBiddersAccepted = CheckAllinRange(_BidPrices, _EstimatedCost)
            If _AllBiddersAccepted = True Then
                _BidderEvaluation = CheckEachInTheRange(_BidPrices, _EstimatedCost)
                _ApprovedMinBidPrice = MinBidPrice(ApprovedBidPriceValue(ApprovedBidPrice(_BidPrices, CalcRange(_EstimatedCost)), _BidPrices))
                TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation)
                TEPMinBidPrice(TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation), _ApprovedMinBidPrice)

                _FinalFinancialProposal = FinalBidPrice(_BidPrices, TEP(TechnicalQualified(_TechnicalScore, _MinTechnicalValue), TEPMinBidPrice(TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation), MinBidPrice(_BidPrices)), MinBidPrice(_BidPrices)))
                _ReCalculateWithDifferentData = False
            Else
                _OneBidderAccepted = CheckAtLeastOneInRange(_BidPrices, _EstimatedCost)
                If _OneBidderAccepted = True Then
                    _BidderEvaluation = CheckEachInTheRange(_BidPrices, _EstimatedCost)
                    _ApprovedMinBidPrice = MinBidPrice(ApprovedBidPriceValue(ApprovedBidPrice(_BidPrices, CalcRange(_EstimatedCost)), _BidPrices))
                    TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation)
                    TEPMinBidPrice(TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation), _ApprovedMinBidPrice)

                    _FinalFinancialProposal = FinalBidPrice(_BidPrices, TEP(TechnicalQualified(_TechnicalScore, _MinTechnicalValue), TEPMinBidPrice(TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation), MinBidPrice(_BidPrices)), MinBidPrice(_BidPrices)))
                    _ReCalculateWithDifferentData = False
                Else
                    _ReCalculateWithDifferentData = True
                End If
            End If
        End Set
    End Property

    Public ReadOnly Property ReCalculateWithDifferentData() As Boolean
        Get
            Return _ReCalculateWithDifferentData
        End Get
    End Property

    Public ReadOnly Property GetFinalFinancialProposal() As Double()
        Get
            Return _FinalFinancialProposal
        End Get
    End Property

    Public ReadOnly Property GetTEP() As Double()
        Get
            Return TEP(TechnicalQualified(_TechnicalScore, _MinTechnicalValue), TEPMinBidPrice(TEPUnit(CalcCes(_ProjectRange), _TechnicalScore, _BidderEvaluation), MinBidPrice(_BidPrices)), MinBidPrice(_BidPrices))
        End Get
    End Property



    Public Property BidPrice As Double()
        Set(value As Double())
            _BidPrices = value
        End Set
        Get
            Return _BidPrices
        End Get
    End Property

    Private Function MinBidPrice(ApprovedBidPrices() As Double) As Double
        Dim Min As Double = ApprovedBidPrices(0)
        For I As Integer = 1 To ApprovedBidPrices.GetUpperBound(0)
            If (ApprovedBidPrices(I) < Min) AndAlso (ApprovedBidPrices(I) > 0) Then
                Min = ApprovedBidPrices(I)
            End If
        Next
        Return Min
    End Function
    Private Function CalcRange(ProjectEstimatedCost As Double) As Double()
        Dim Result(1) As Double
        Result(0) = _MinBidderPercent * ProjectEstimatedCost
        Result(1) = _MaxBidderPercent * ProjectEstimatedCost
        Return Result
    End Function

    Private Function CheckAllinRange(BidPrices() As Double, ProjectEstimatedCost As Double) As Boolean
        Dim Result As Integer = 1
        For I As Integer = 1 To BidPrices.Count - 1
            If BidPrices(I) >= CalcRange(ProjectEstimatedCost)(0) And BidPrices(I) <= CalcRange(ProjectEstimatedCost)(1) Then
                Result *= 0
            Else
                Result *= 1
            End If
        Next
        If Result = 1 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function CheckAtLeastOneInRange(BidPrices() As Double, ProjectEstimatedCost As Double) As Boolean
        For I As Integer = 1 To BidPrices.Count - 1
            If BidPrices(I) >= CalcRange(ProjectEstimatedCost)(0) And BidPrices(I) <= CalcRange(ProjectEstimatedCost)(1) Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Function CheckEachInTheRange(BidPrices() As Double, ProjectEstimatedCost As Double) As Integer()
        Dim Max As Double = _MaxBidderPercent * ProjectEstimatedCost
        Dim Min As Double = _MinBidderPercent * ProjectEstimatedCost
        Dim Result(BidPrices.GetUpperBound(0)) As Integer
        For I As Integer = 0 To BidPrices.GetUpperBound(0)
            'If (BidPrices(I) >= Min) And (BidPrices(I) <= Max) Then
            If (Min <= BidPrices(I)) AndAlso (BidPrices(I) <= Max) Then
                Result(I) = 0
            ElseIf BidPrices(I) < Min Then
                Result(I) = -1
            ElseIf BidPrices(I) > Max Then
                Result(I) = 1
            End If
        Next
        Return Result
    End Function

    Private Function CheckOneAtLeastInTheRange(List() As Boolean) As Boolean
        For I As Integer = 0 To List.GetUpperBound(0)
            If List(I) = True Then Return True
        Next
        Return False
    End Function


    Private Function CalcCes(ProjectRange As XtraForm1.ProjectSize) As Double()
        Dim Result(3) As Double
        If ProjectRange = XtraForm1.ProjectSize.LessThan5 Then
            Result(0) = (-20638 * Math.Pow(_MinTechnicalValue, 3)) + (13191 * Math.Pow(_MinTechnicalValue, 2)) - (2833.6 * _MinTechnicalValue) + 208.75
            Result(1) = (8384 * Math.Pow(_MinTechnicalValue, 3)) - (5542.2 * Math.Pow(_MinTechnicalValue, 2)) + (1259.5 * _MinTechnicalValue) - (103.42)
            Result(2) = (-771.07 * Math.Pow(_MinTechnicalValue, 3)) + (539.74 * Math.Pow(_MinTechnicalValue, 2)) - (136.86 * _MinTechnicalValue) + 14.843
            Result(3) = 0.63
        ElseIf ProjectRange = XtraForm1.ProjectSize.From5MTo100M Then
            Result(0) = (-28140 * Math.Pow(_MinTechnicalValue, 3)) + (17987 * Math.Pow(_MinTechnicalValue, 2)) - (3863.8 * _MinTechnicalValue) + 284.65
            Result(1) = (11104 * Math.Pow(_MinTechnicalValue, 3)) - (7343.2 * Math.Pow(_MinTechnicalValue, 2)) + (1669.5 * _MinTechnicalValue) - 137.12
            Result(2) = (-14287 * Math.Pow(_MinTechnicalValue, 3)) + (8667.3 * Math.Pow(_MinTechnicalValue, 2)) - (1735.9 * _MinTechnicalValue) + 118.35
            Result(3) = 0.56
        ElseIf ProjectRange = XtraForm1.ProjectSize.From100MTo250M Then
            Result(0) = (11256 * Math.Pow(_MinTechnicalValue, 3)) - (7195 * Math.Pow(_MinTechnicalValue, 2)) + (1545.6 * _MinTechnicalValue) - 113.86
            Result(1) = (-3422.4 * Math.Pow(_MinTechnicalValue, 3)) + (2262.3 * Math.Pow(_MinTechnicalValue, 2)) - (514.13 * _MinTechnicalValue) + 42.212
            Result(2) = (100 * Math.Pow(_MinTechnicalValue, 3)) - (70 * Math.Pow(_MinTechnicalValue, 2)) + (17.75 * _MinTechnicalValue) - 1.925
            Result(3) = 0.995
        ElseIf ProjectRange = XtraForm1.ProjectSize.From250MTo500M Then
            Result(0) = (3752 * Math.Pow(_MinTechnicalValue, 3)) - (2398.3 * Math.Pow(_MinTechnicalValue, 2)) + (515.18 * _MinTechnicalValue) - 37.954
            Result(1) = (-342.4 * Math.Pow(_MinTechnicalValue, 3)) + (226.32 * Math.Pow(_MinTechnicalValue, 2)) - (51.428 * _MinTechnicalValue) + 4.222
            Result(2) = (-142.13 * Math.Pow(_MinTechnicalValue, 3)) + (99.52 * Math.Pow(_MinTechnicalValue, 2)) - (25.241 * _MinTechnicalValue) + 2.7377
            Result(3) = 0.89
        ElseIf ProjectRange = XtraForm1.ProjectSize.GreaterThan500M Then
            Result(0) = (-15942 * Math.Pow(_MinTechnicalValue, 3)) + (10190 * Math.Pow(_MinTechnicalValue, 2)) - (2189 * _MinTechnicalValue) + 161.27
            Result(1) = (6416.8 * Math.Pow(_MinTechnicalValue, 3)) - (4241.7 * Math.Pow(_MinTechnicalValue, 2)) + (963.97 * _MinTechnicalValue) - 79.147
            Result(2) = (-605.33 * Math.Pow(_MinTechnicalValue, 3)) + (423.78 * Math.Pow(_MinTechnicalValue, 2)) - (107.47 * _MinTechnicalValue) + 11.656
            Result(3) = 0.705
        End If
        Return Result
    End Function

    Private Function TEPUnit(Ces() As Double, TechnicalScore() As Double, BidderEvaluation() As Integer) As Double()
        Dim Result(TechnicalScore.GetUpperBound(0)) As Double
        For I As Integer = 0 To TechnicalScore.GetUpperBound(0)
            '        Result(I) = (Ces(0) * Math.Pow(TechnicalScore(I), 3)) + (Ces(1) * Math.Pow(TechnicalScore(I), 2)) _
            '+ (Ces(2) * Math.Pow(TechnicalScore(I), 1)) + Ces(3)
            'Console.WriteLine(BidderEvaluation(I))
            If BidderEvaluation(I) = -1 Then
                Result(I) = 0
            Else
                Result(I) = (Ces(0) * Math.Pow(TechnicalScore(I), 3)) + (Ces(1) * Math.Pow(TechnicalScore(I), 2)) _
                    + (Ces(2) * Math.Pow(TechnicalScore(I), 1)) + Ces(3)
            End If
        Next
        Return Result
    End Function

    Private Function TechnicalQualified(TechnicalScore() As Double, MinTechnicalValue As Double) As Boolean()
        Dim Result(TechnicalScore.GetUpperBound(0)) As Boolean
        For I As Integer = 0 To TechnicalScore.GetUpperBound(0)
            If TechnicalScore(I) < MinTechnicalValue Then
                Result(I) = False
            Else
                Result(I) = True
            End If
        Next
        Return Result
    End Function

    Private Function ApprovedBidPrice(BidPrices() As Double, CalcRange() As Double) As Boolean()
        Dim Result(BidPrices.GetUpperBound(0)) As Boolean
        For I As Integer = 0 To BidPrices.GetUpperBound(0)
            If (BidPrices(I) >= CalcRange(0)) AndAlso (BidPrices(I) <= CalcRange(1)) Then
                Result(I) = True
            Else
                Result(I) = False
            End If
            Console.WriteLine(Result(I))
        Next
        Return Result
    End Function

    Private Function ApprovedBidPriceValue(ApprovedBidPrice() As Boolean, BidPrices() As Double) As Double()
        Dim Result(ApprovedBidPrice.GetUpperBound(0)) As Double
        For I As Integer = 0 To ApprovedBidPrice.GetUpperBound(0)
            If ApprovedBidPrice(I) = True Then
                Result(I) = BidPrices(I)
            Else
                Result(I) = 0
            End If
            Console.WriteLine(Result(I))
        Next
        Return Result
    End Function

    Private Function TEPMinBidPrice(TEPUnit() As Double, MinAcceptedBidPrice As Double) As Double()
        Dim Result(TEPUnit.GetUpperBound(0)) As Double
        For I As Integer = 0 To TEPUnit.GetUpperBound(0)
            Result(I) = TEPUnit(I) * MinAcceptedBidPrice

        Next
        Return Result
    End Function

    Private Function TEP(TechnicalQualified() As Boolean, TEPMinBidPrice() As Double, MinBidPrice As Double) As Double()
        Dim Result(TechnicalQualified.GetUpperBound(0)) As Double
        For I As Integer = 0 To TechnicalQualified.GetUpperBound(0)
            If TechnicalQualified(I) = True Then
                If MinBidPrice > TEPMinBidPrice(I) Then
                    Result(I) = TEPMinBidPrice(I)
                Else
                    Result(I) = TEPMinBidPrice(I) - MinBidPrice
                End If
            Else
                Result(I) = -10
            End If
        Next
        Return Result
    End Function

    Private Function FinalBidPrice(BidPrice() As Double, TEP() As Double) As Double()
        Dim Result(BidPrice.GetUpperBound(0)) As Double
        For I As Integer = 0 To TEP.GetUpperBound(0)
            If TEP(I) <> -10 Then
                Result(I) = BidPrice(I) - TEP(I)
            Else
                Result(I) = -10
            End If
        Next
        Return Result
    End Function

    Public Function MinFinancialValue() As Double
        '_FinalFinancialProposal.Min()
        If _FinalFinancialProposal.Count = 0 Then
            Return 0
        End If
        Dim Min As Double
        For I As Integer = 1 To _FinalFinancialProposal.GetUpperBound(0)
            If TechnicalQualified(_TechnicalScore, _MinTechnicalValue)(I) = True AndAlso
                CheckEachInTheRange(_BidPrices, _EstimatedCost)(I) >= 0 Then
                Min = _FinalFinancialProposal(I)
                Exit For
            End If
            'If Min > _FinalFinancialProposal(I) Then Min = _FinalFinancialProposal(I)
        Next
        For I As Integer = 1 To _FinalFinancialProposal.GetUpperBound(0)
            If TechnicalQualified(_TechnicalScore, _MinTechnicalValue)(I) = True AndAlso
                CheckEachInTheRange(_BidPrices, _EstimatedCost)(I) >= 0 Then
                If Min > _FinalFinancialProposal(I) Then Min = _FinalFinancialProposal(I)
            End If
            'If Min > _FinalFinancialProposal(I) Then Min = _FinalFinancialProposal(I)
        Next
        Return Min
    End Function

    Public Function GetFinancialQualification() As Boolean()
        Dim Result(_BidPrices.GetUpperBound(0)) As Boolean
        For I As Integer = 0 To _BidPrices.GetUpperBound(0)
            If CheckEachInTheRange(_BidPrices, _EstimatedCost)(I) >= 0 Then
                Result(I) = True
            Else
                Result(I) = False
            End If
        Next
        Return Result
    End Function
End Class