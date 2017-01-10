Public Class Selection

    Public status As String

    Public selectionNumber As Integer

    ' Betfair Event details
    Public betfairEventId As String
    Public betfairEventName As String
    Public betfairEventDateTime As String
    Public betfairEventInplay As Boolean

    ' Betfair Under/Over1.5 market details
    Public betfairUnderOver15MarketId As String
    Public betfairUnderOver15MarketStatus As String
    Public betfairUnderOver15UnmathedBets As String

    Public betfairOver15SelectionId As String
    Public betfairOver15SelectionStatus As String
    Public betfairOver15BackOdds As String
    Public betfairOver15IfWinProfit As String

    Public betfairUnder15SelectionId As String
    Public betfairUnder15SelectionStatus As String
    Public betfairUnder15BackOdds As String
    Public betfairUnder15IfWinProfit As String


    ' Betfair Correct Score market details
    Public betfairCorrectScoreMarketId As String
    Public betfairCorrectScoreMarketStatus As String
    Public betfairCorrectScoreUnmathedBets As String

    Public betfairCorrectScore00SelectionId As String
    Public betfairCorrectScore00SelectionStatus As String
    Public betfairCorrectScore00BackOdds As String
    Public betfairCorrectScore00IfWinProfit As String

    Public betfairCorrectScore10SelectionId As String
    Public betfairCorrectScore10SelectionStatus As String
    Public betfairCorrectScore10BackOdds As String
    Public betfairCorrectScore10IfWinProfit As String

    Public betfairCorrectScore01SelectionId As String
    Public betfairCorrectScore01SelectionStatus As String
    Public betfairCorrectScore01BackOdds As String
    Public betfairCorrectScore01IfWinProfit As String

    Public Sub New(selectionNumber)

        status = "Not Selected"

    End Sub

    Public Sub getAllMarketData()

        Dim BetfairClass1 As New BetfairClass()

        ' Get the Correct Score and Under Over 1.5 prices
        BetfairClass1.PollBetFairUnderOver15Market(Me, 1, betfairEventId, My.Settings.NumberOfUkEvents)

        ' Get unmatched bets
        BetfairClass1.listCurrentOrder(Me)

        BetfairClass1 = Nothing

    End Sub

    Public Sub placeCorrectScore_00_Order()

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairCorrectScoreMarketId, betfairCorrectScore00SelectionId, CDbl(betfairCorrectScore00BackOdds), CDbl(2))
        BetfairClass1 = Nothing

    End Sub

End Class
