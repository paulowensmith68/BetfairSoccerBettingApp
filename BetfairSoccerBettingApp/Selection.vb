Imports System.Threading

Public Class Selection

    Public status As String

    Public selectionNumber As Integer

    'Autobet flags
    Public autobetOver15BetMade As Boolean
    Public autobetUnder15BetMade As Boolean
    Public autobetCorrectScore00BetMade As Boolean
    Public autobetCorrectScore10BetMade As Boolean
    Public autobetCorrectScore01BetMade As Boolean
    Public autobetOver15TopUpBetMade As Boolean


    ' Betfair Event details
    Public betfairEventId As String
    Public betfairEventName As String
    Public betfairEventDateTime As String
    Public betfairEventInplay As Boolean
    Public betfairGoalsScored As String
    Public betfairGoal1DateTime As Date
    Public betfairGoal2DateTime As Date

    ' Betfair Under/Over0.5 market details
    Public betfairUnderOver05MarketId As String
    Public betfairUnderOver05MarketStatus As String

    ' Betfair Under/Over2.5 market details
    Public betfairUnderOver25MarketId As String
    Public betfairUnderOver25MarketStatus As String

    ' Betfair Under/Over3.5 market details
    Public betfairUnderOver35MarketId As String
    Public betfairUnderOver35MarketStatus As String

    ' Betfair Under/Over4.5 market details
    Public betfairUnderOver45MarketId As String
    Public betfairUnderOver45MarketStatus As String

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
            If betfairCorrectScoreMarketId = "Not Found" Then

                ' Market has been removed, update data
                betfairCorrectScoreMarketStatus = "CLOSED"
                betfairCorrectScoreMarketStatus = ""
                betfairCorrectScore00SelectionId = ""
                betfairCorrectScore00SelectionStatus = ""
                betfairCorrectScore00BackOdds = ""
                betfairCorrectScore00IfWinProfit = ""
                betfairCorrectScore00Orders = ""
                betfairCorrectScore10SelectionId = ""
                betfairCorrectScore10SelectionStatus = ""
                betfairCorrectScore10BackOdds = ""
                betfairCorrectScore10IfWinProfit = ""
                betfairCorrectScore10Orders = ""
                betfairCorrectScore01SelectionId = ""
                betfairCorrectScore01SelectionStatus = ""
                betfairCorrectScore01BackOdds = ""
                betfairCorrectScore01IfWinProfit = ""
                betfairCorrectScore01Orders = ""
            Else
                BetfairClass1.listMarketBook(Me, betfairCorrectScoreMarketId)
                BetfairClass1.listMarketProfitAndLoss(Me, betfairCorrectScoreMarketId)
            End If
        End If
        If String.IsNullOrEmpty(betfairUnderOver15MarketId) Then
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Unable to refresh data as Market Id is null or empty for OVER_UNDER_15", EventLogEntryType.Error)
        Else
            If betfairUnderOver15MarketId = "Not Found" Then

                ' Market has been removed, update data
                betfairUnderOver15MarketStatus = "CLOSED"
                betfairOver15SelectionId = ""
                betfairOver15SelectionStatus = ""
                betfairOver15BackOdds = ""
                betfairOver15IfWinProfit = ""
                betfairOver15Orders = ""
                betfairUnder15SelectionId = ""
                betfairUnder15SelectionStatus = ""
                betfairUnder15BackOdds = ""
                betfairUnder15IfWinProfit = ""
                betfairUnder15Orders = ""
            Else
                BetfairClass1.listMarketBook(Me, betfairUnderOver15MarketId)
                BetfairClass1.listMarketProfitAndLoss(Me, betfairUnderOver15MarketId)
            End If
        End If

        '
        ' Sometimes All markets are suspended during game whilst things are updated e.g. Yellow cards, Red cards, Injury Goals scored
        ' 
        ' Calulate Inplay time
        Dim eventDateTime As DateTime = DateTime.Parse(Me.betfairEventDateTime)
        Dim timeToStart As TimeSpan = DateTime.Now.Subtract(eventDateTime)
        Dim timeInplay As Double
        Dim formatTime As String = "####0.00"
        timeInplay = timeToStart.TotalMinutes

        If Me.betfairEventInplay = "True" And timeInplay < +105 And Me.betfairCorrectScoreMarketStatus = "SUSPENDED" Then

            ' Don't do anything - 
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : CORRECT_SCORE market SUSPENDED and still within playing time", EventLogEntryType.Error)

        Else

            ' Populate goals scored
            betfairGoalsScored = calculateGoalsScored()

        End If

        BetfairClass1 = Nothing

    End Sub
    Private Function calculateGoalsScored() As String

        ' Get the derived market Status from previous listMarketCatalogue
        Dim strGoalsScored As String = ""

        Try

            If Me.betfairCorrectScoreMarketId = "Not Found" Or Me.betfairCorrectScoreMarketStatus = "SUSPENDED" Or Me.betfairCorrectScoreMarketStatus = "CLOSED" Then
                Return "Match ended!"
            Else
                If Me.betfairUnderOver05MarketId = "Not Found" Or Me.betfairUnderOver05MarketStatus = "SUSPENDED" Or Me.betfairUnderOver05MarketStatus = "CLOSED" Then
                    If Me.betfairUnderOver15MarketId = "Not Found" Or Me.betfairUnderOver15MarketStatus = "SUSPENDED" Or Me.betfairUnderOver15MarketStatus = "CLOSED" Then
                        If Me.betfairUnderOver25MarketId = "Not Found" Or Me.betfairUnderOver25MarketStatus = "SUSPENDED" Or Me.betfairUnderOver25MarketStatus = "CLOSED" Then
                            If Me.betfairUnderOver35MarketId = "Not Found" Or Me.betfairUnderOver35MarketStatus = "SUSPENDED" Or Me.betfairUnderOver35MarketStatus = "CLOSED" Then
                                If Me.betfairUnderOver45MarketId = "Not Found" Or Me.betfairUnderOver45MarketStatus = "SUSPENDED" Or Me.betfairUnderOver45MarketStatus = "CLOSED" Then
                                    Return "Over 4.5"
                                Else
                                    Return "4 Goals scored"
                                End If
                            Else
                                Return "3 Goals scored"
                            End If
                        Else
                            Return "2 Goals scored"
                        End If
                    Else
                        Return "1 Goal scored"
                    End If
                Else
                    Return "0 - 0"
                End If
            End If

        Catch ex As Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : <calculateGoalsScored> Error calculating score: " + ex.Message, EventLogEntryType.Error)
            Return "Unknown Error"

        End Try

    End Function

    Public Sub placeOver15_Order(price As Double, stake As Double)

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairUnderOver15MarketId, betfairOver15SelectionId, price, stake)
        BetfairClass1 = Nothing

    End Sub

    Public Sub placeUnder15_Order(price As Double, stake As Double)

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairUnderOver15MarketId, betfairUnder15SelectionId, price, stake)
        BetfairClass1 = Nothing

    End Sub

    Public Sub placeCorrectScore00_Order(price As Double, stake As Double)

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairCorrectScoreMarketId, betfairCorrectScore00SelectionId, price, stake)
        BetfairClass1 = Nothing

    End Sub
    Public Sub placeCorrectScore10_Order(price As Double, stake As Double)

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairCorrectScoreMarketId, betfairCorrectScore10SelectionId, price, stake)
        BetfairClass1 = Nothing

    End Sub
    Public Sub placeCorrectScore01_Order(price As Double, stake As Double)

        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PlaceOrder(betfairCorrectScoreMarketId, betfairCorrectScore01SelectionId, price, stake)
        BetfairClass1 = Nothing

    End Sub
End Class
