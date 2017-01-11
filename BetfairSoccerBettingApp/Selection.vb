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

    Public betfairOver15SelectionId As String
    Public betfairOver15SelectionStatus As String
    Public betfairOver15BackOdds As String
    Public betfairOver15IfWinProfit As String
    Public betfairOver15Orders As String


    Public betfairUnder15SelectionId As String
    Public betfairUnder15SelectionStatus As String
    Public betfairUnder15BackOdds As String
    Public betfairUnder15IfWinProfit As String
    Public betfairUnder15Orders As String


    ' Betfair Correct Score market details
    Public betfairCorrectScoreMarketId As String
    Public betfairCorrectScoreMarketStatus As String

    Public betfairCorrectScore00SelectionId As String
    Public betfairCorrectScore00SelectionStatus As String
    Public betfairCorrectScore00BackOdds As String
    Public betfairCorrectScore00IfWinProfit As String
    Public betfairCorrectScore00Orders As String


    Public betfairCorrectScore10SelectionId As String
    Public betfairCorrectScore10SelectionStatus As String
    Public betfairCorrectScore10BackOdds As String
    Public betfairCorrectScore10IfWinProfit As String
    Public betfairCorrectScore10Orders As String


    Public betfairCorrectScore01SelectionId As String
    Public betfairCorrectScore01SelectionStatus As String
    Public betfairCorrectScore01BackOdds As String
    Public betfairCorrectScore01IfWinProfit As String
    Public betfairCorrectScore01Orders As String


    Public Sub New(selectionNumber)

        status = "Selected"

    End Sub

    Public Sub getInitialBookDetails()

        ' Initialize main Id's
        betfairUnderOver15MarketId = Nothing
        betfairCorrectScoreMarketId = Nothing

        Dim BetfairClass1 As New BetfairClass()

        ' Get the Correct Score and Under Over 1.5 books using decsriptions
        BetfairClass1.PollBetFairInitialMarketDetails(Me, 1, betfairEventId, My.Settings.NumberOfUkEvents)

        BetfairClass1 = Nothing

    End Sub
    Public Sub getLatestMarketData()

        Dim BetfairClass1 As New BetfairClass()

        ' Get the market Status from listBook
        ' Call them individually as if one is OPEN and other CLOSED you wont get result back, also get all MATCHED and UNMATCHED data
        If String.IsNullOrEmpty(betfairCorrectScoreMarketId) Then
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Unable to refresh data as Market Id is null or empty for CORRECT_SCORE", EventLogEntryType.Error)
        Else
            BetfairClass1.listMarketBook(Me, betfairCorrectScoreMarketId)
            BetfairClass1.listMarketProfitAndLoss(Me, betfairCorrectScoreMarketId)
        End If
        If String.IsNullOrEmpty(betfairUnderOver15MarketId) Then
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Unable to refresh data as Market Id is null or empty for OVER_UNDER_15", EventLogEntryType.Error)
        Else
            BetfairClass1.listMarketBook(Me, betfairUnderOver15MarketId)
            BetfairClass1.listMarketProfitAndLoss(Me, betfairUnderOver15MarketId)
        End If

        BetfairClass1 = Nothing

    End Sub

    Public Sub placeCorrectScore_00_Order()

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairCorrectScoreMarketId, betfairCorrectScore00SelectionId, CDbl(betfairCorrectScore00BackOdds), CDbl(2))
        BetfairClass1 = Nothing

    End Sub

End Class
