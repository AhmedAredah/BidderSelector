Public Class TechnicalOutputArray
    Private _AttributeOutputArray() As Double
    Private _MinVaue As Double
    Private _MaxValue As Double

    Public Sub new
        _AttributeOutputArray = Nothing
        _MinVaue = 0
        _MaxValue = 0
    End Sub

    Public Property AttributeOutputArray As Double ()
        Get
            Return _AttributeOutputArray
        End Get
        Set(value As Double())
            _AttributeOutputArray = value
            _MinVaue = _AttributeOutputArray(0)
            _MaxValue= CalcMax(_AttributeOutputArray)
            '_MinVaue = CalCMin(_AttributeOutputArray)
        End Set
    End Property

    Public ReadOnly Property MinValue As double
        Get
            Return _MinVaue
        End Get
    End Property

    Public ReadOnly Property MaxValue As Double
        Get
            Return _MAxValue
        End Get
    End Property

    Private Function CalCMin(Array As Double()) As Double
        Dim _Min As Double = 0
        For I As Integer = 0 to Array.GetUpperBound(0)
            If _Min > Array(I) then _Min = Array(I) Else Continue For
        Next
        Return _Min
    End Function

    Private Function CalcMax(Array As double()) As Double
        Dim _Max As Double = 0
        For I As Integer = 0 to Array.GetUpperBound(0)
            If _Max < Array(I) Then _Max = Array(I) Else Continue For
        Next
        Return _Max
    End Function
End Class
