﻿Imports MySql.Data.MySqlClient
Imports BetfairSoccerBettingApp.Api_ng_sample_code.TO
Imports BetfairSoccerBettingApp.Api_ng_sample_code
Public Class BetfairClass

    ' Holds the connection string to the database used.
    Public eventList As New List(Of BeffairEventClass)

    'Holds message received back from class
    Public returnMessage As String = ""

    Public Sub PollBetFairEvents(eventTypeId As Integer, maxResults As String, marketCountries As HashSet(Of String), Optional inplay As Boolean = False)

        Dim newEvent As BeffairEventClass

        Dim client As IClient = Nothing
        Dim clientType As String = Nothing
        client = New JsonRpcClient(globalBetFairUrl, globalBetFairAppKey, globalBetFairToken)
        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Starting to get list of Events for Event Id: " + eventTypeId.ToString + " Market Countries: " + DisplaySet(marketCountries), EventLogEntryType.Information)

        Try

            Dim marketFilter = New MarketFilter()
            Dim eventTypes = client.listEventTypes(marketFilter)
            Dim eventypeIds As ISet(Of String) = New HashSet(Of String)()

            ' Football is eventId 1
            eventypeIds.Add(eventTypeId)

            'ListMarketCatalogue parameters
            Dim time = New TimeRange()
            time.From = Date.Now()
            time.To = Date.Now.AddDays(globalBetFairDaysAhead)

            marketFilter = New MarketFilter()
            marketFilter.EventTypeIds = eventypeIds
            marketFilter.MarketStartTime = time

            ' Setup country codes required
            marketFilter.MarketCountries = marketCountries

            ' Set InPlayOnly : Restrict to markets that are currently in play if True or are not currently in play if false. If not specified, returns both.
            If inplay = True Then
                marketFilter.InPlayOnly = True
            End If


            Dim events = client.listEvents(marketFilter)
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Response from listEvents : " + events.Count.ToString, EventLogEntryType.Information)

            For Each footballEvent In events

                ' Processing event...
                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Processing event : " + footballEvent.Event.Name, EventLogEntryType.Information)

                ' Convert date to localtime
                Dim gmtOpenDate As DateTime
                gmtOpenDate = footballEvent.Event.OpenDate

                'GMT Standard Time
                Dim gmt As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time")
                gmtOpenDate = TimeZoneInfo.ConvertTimeFromUtc(gmtOpenDate, gmt)

                'Create instance of event class
                newEvent = New BeffairEventClass With {
                        .eventTypeId = eventTypeId,
                        .eventId = footballEvent.Event.Id.ToString,
                        .name = footballEvent.Event.Name,
                        .timezone = footballEvent.Event.Timezone,
                        .countryCode = footballEvent.Event.CountryCode,
                        .openDate = gmtOpenDate
                    }

                ' Add to list
                eventList.Add(newEvent)

            Next ' End of events

        Catch apiExcepion As APINGException
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, APINGExcepion msg : " + apiExcepion.Message, EventLogEntryType.Error)
            Exit Sub
        Catch ex As System.Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, system exception: " + ex.Message, EventLogEntryType.Error)
            Exit Sub

        Finally

        End Try

    End Sub

    Public Sub PollBetFairUnderOver15Market(ByRef selection As Selection, eventTypeId As Integer, eventId As String, maxResults As String)

        Dim client As IClient = Nothing
        Dim clientType As String = Nothing
        client = New JsonRpcClient(globalBetFairUrl, globalBetFairAppKey, globalBetFairToken)
        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Starting to get OVER_UNDER_15 and CORRECT_SCORE for Event Id: " + eventTypeId.ToString + " Event Id: " + eventId, EventLogEntryType.Information)

        Try

            Dim marketFilter = New MarketFilter()
            Dim eventTypes = client.listEventTypes(marketFilter)
            Dim eventypeIds As ISet(Of String) = New HashSet(Of String)()
            Dim eventIds As ISet(Of String) = New HashSet(Of String)()

            ' Football is eventId 1
            eventypeIds.Add(eventTypeId)

            ' Event Id
            eventIds.Add(eventId)

            ' Create new market filter
            marketFilter = New MarketFilter()
            marketFilter.EventTypeIds = eventypeIds

            ' Restrict to 1 event
            marketFilter.EventIds = eventIds

            ' Set-up market type codes e.g. WIN or MATCH ODDS
            marketFilter.MarketTypeCodes = New HashSet(Of String)({"CORRECT_SCORE", "OVER_UNDER_15"})

            ' Set-up order
            Dim marketSort = Api_ng_sample_code.TO.MarketSort.MAXIMUM_TRADED

            ' Set-up market projection
            Dim marketProjections As ISet(Of MarketProjection) = New HashSet(Of MarketProjection)()
            marketProjections.Add(MarketProjection.RUNNER_METADATA)
            marketProjections.Add(MarketProjection.EVENT)

            Dim marketCatalogues = client.listMarketCatalogue(marketFilter, marketProjections, marketSort, maxResults)
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Response from MarketCatalogue (event objects) : " + marketCatalogues.Count.ToString, EventLogEntryType.Information)

            For Each book In marketCatalogues
                Dim marketIdBook As String = book.MarketId
                Dim marketIdsBook As IList(Of String) = New List(Of String)()
                marketIdsBook.Add(marketIdBook)

                Dim priceData As ISet(Of PriceData) = New HashSet(Of PriceData)()
                'get all prices from the exchange
                priceData.Add(Api_ng_sample_code.TO.PriceData.EX_BEST_OFFERS)
                priceData.Add(Api_ng_sample_code.TO.PriceData.EX_TRADED)

                Dim priceProjection = New PriceProjection()
                priceProjection.PriceData = priceData

                Dim marketBook = client.listMarketBook(marketIdsBook, priceProjection)

                ' Look through the market books, there should only be 1
                For Each backBet In marketBook

                    If marketBook.Count = 1 Then

                        For i = 0 To backBet.Runners.Count - 1

                            If book.Runners(i).RunnerName = "Over 1.5 Goals" Then
                                selection.betfairUnderOver15MarketId = backBet.MarketId
                                selection.betfairUnderOver15MarketStatus = convertMarketStatus(backBet.Status)
                                selection.betfairEventInplay = backBet.IsInplay
                                selection.betfairOver15BackOdds = backBet.Runners(i).ExchangePrices.AvailableToBack(0).Price
                                selection.betfairOver15SelectionId = book.Runners(i).SelectionId
                                selection.betfairOver15SelectionStatus = convertRunnerStatus(backBet.Runners(i).Status)
                            ElseIf book.Runners(i).RunnerName = "Under 1.5 Goals" Then
                                selection.betfairUnderOver15MarketId = backBet.MarketId
                                selection.betfairUnderOver15MarketStatus = convertMarketStatus(backBet.Status)
                                selection.betfairEventInplay = backBet.IsInplay
                                selection.betfairUnder15BackOdds = backBet.Runners(i).ExchangePrices.AvailableToBack(0).Price
                                selection.betfairUnder15SelectionId = book.Runners(i).SelectionId
                                selection.betfairUnder15SelectionStatus = convertRunnerStatus(backBet.Runners(i).Status)
                            ElseIf book.Runners(i).RunnerName = "0 - 0" Then
                                selection.betfairCorrectScoreMarketId = backBet.MarketId
                                selection.betfairCorrectScoreMarketStatus = convertMarketStatus(backBet.Status)
                                selection.betfairEventInplay = backBet.IsInplay
                                selection.betfairCorrectScore00BackOdds = backBet.Runners(i).ExchangePrices.AvailableToBack(0).Price
                                selection.betfairCorrectScore00SelectionId = book.Runners(i).SelectionId
                                selection.betfairCorrectScore00SelectionStatus = convertRunnerStatus(backBet.Runners(i).Status)
                            ElseIf book.Runners(i).RunnerName = "1 - 0" Then
                                selection.betfairCorrectScoreMarketId = backBet.MarketId
                                selection.betfairCorrectScoreMarketStatus = convertMarketStatus(backBet.Status)
                                selection.betfairEventInplay = backBet.IsInplay
                                selection.betfairCorrectScore10BackOdds = backBet.Runners(i).ExchangePrices.AvailableToBack(0).Price
                                selection.betfairCorrectScore10SelectionId = book.Runners(i).SelectionId
                                selection.betfairCorrectScore10SelectionStatus = convertRunnerStatus(backBet.Runners(i).Status)
                            ElseIf book.Runners(i).RunnerName = "0 - 1" Then
                                selection.betfairCorrectScoreMarketId = backBet.MarketId
                                selection.betfairCorrectScoreMarketStatus = convertMarketStatus(backBet.Status)
                                selection.betfairEventInplay = backBet.IsInplay
                                selection.betfairCorrectScore01BackOdds = backBet.Runners(i).ExchangePrices.AvailableToBack(0).Price
                                selection.betfairCorrectScore01SelectionId = book.Runners(i).SelectionId
                                selection.betfairCorrectScore01SelectionStatus = convertRunnerStatus(backBet.Runners(i).Status)
                            End If

                        Next ' End of runners

                    End If

                Next ' End of backBet


                Dim marketProfitLoss = client.listMarketProfitAndLoss(marketIdsBook)

                ' Look through the market books, there should only be 1
                For Each profitLoss In marketProfitLoss

                    If marketProfitLoss.Count = 1 Then

                        For i = 0 To profitLoss.ProfitAndLosses.Count - 1

                            If profitLoss.ProfitAndLosses(i).SelectionId = selection.betfairOver15SelectionId Then
                                selection.betfairOver15IfWinProfit = profitLoss.ProfitAndLosses(i).IfWin

                            ElseIf profitLoss.ProfitAndLosses(i).SelectionId = selection.betfairUnder15SelectionId Then
                                selection.betfairUnder15IfWinProfit = profitLoss.ProfitAndLosses(i).IfWin

                            ElseIf profitLoss.ProfitAndLosses(i).SelectionId = selection.betfairCorrectScore00SelectionId Then
                                selection.betfairCorrectScore00IfWinProfit = profitLoss.ProfitAndLosses(i).IfWin

                            ElseIf profitLoss.ProfitAndLosses(i).SelectionId = selection.betfairCorrectScore10SelectionId Then
                                selection.betfairCorrectScore10IfWinProfit = profitLoss.ProfitAndLosses(i).IfWin

                            ElseIf profitLoss.ProfitAndLosses(i).SelectionId = selection.betfairCorrectScore01SelectionId Then
                                selection.betfairCorrectScore01IfWinProfit = profitLoss.ProfitAndLosses(i).IfWin

                            End If


                        Next ' End of runners

                    End If

                Next ' End of layBet

            Next

        Catch apiExcepion As APINGException
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, APINGExcepion msg : " + apiExcepion.Message, EventLogEntryType.Error)
            Exit Sub
        Catch ex As System.Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, system exception: " + ex.Message, EventLogEntryType.Error)
            Exit Sub

        Finally

        End Try

    End Sub

    Public Sub PlaceOrder(marketId As String, selectionId As String, price As Double, stake As Double)

        Dim client As IClient = Nothing
        Dim clientType As String = Nothing
        client = New JsonRpcClient(globalBetFairUrl, globalBetFairAppKey, globalBetFairToken)
        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Place Order for Market Id: " + marketId.ToString, EventLogEntryType.Information)


        Dim marketIds As IList(Of String) = New List(Of String)()
        marketIds.Add(marketId)

        ' place a back bet at rediculous odds so it doesn't get matched 
        ' Set-up Limit Order
        Dim LimitOrder = New LimitOrder()

        LimitOrder.Price = price
        LimitOrder.Size = stake

        ' placing a bet. set-up market projection
        Dim placeInstructions As IList(Of PlaceInstruction) = New List(Of PlaceInstruction)()
        Dim placeInstruction = New PlaceInstruction()

        placeInstruction.LimitOrder = LimitOrder
        placeInstruction.SelectionId = selectionId
        placeInstructions.Add(placeInstruction)

        Dim customerRef = "smith4p-autobet"
        Dim placeExecutionReport = client.placeOrders(marketId, customerRef, placeInstructions)

        Dim executionErrorcode As ExecutionReportErrorCode = placeExecutionReport.ErrorCode
        Dim instructionErrorCode As InstructionReportErrorCode = placeExecutionReport.InstructionReports(0).ErrorCode
        Console.WriteLine(vbLf & "PlaceExecutionReport error code is: " + executionErrorcode.ToString + vbLf & "InstructionReport error code is: " + instructionErrorCode.ToString)

        If executionErrorcode <> ExecutionReportErrorCode.BET_ACTION_ERROR AndAlso instructionErrorCode <> InstructionReportErrorCode.INVALID_BET_SIZE Then
            Environment.[Exit](0)
        End If

        Console.WriteLine(vbLf & "DONE!")

    End Sub

    Public Sub listCurrentOrder(ByRef selection As Selection)

        Dim client As IClient = Nothing
        Dim clientType As String = Nothing
        Dim unmatchedCSCount As Integer
        Dim unmatchedUO15Count As Integer

        client = New JsonRpcClient(globalBetFairUrl, globalBetFairAppKey, globalBetFairToken)
        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : List current Order for Unmatched Bets", EventLogEntryType.Information)

        Try

            ' Dim marketIds As ISet(Of String) = New HashSet(Of String)()
            ' marketIds.Add(marketId)

            Dim CurrentOrderSummaryReport = client.listCurrentOrders()

            Dim x As String
            x = CurrentOrderSummaryReport.CurrentOrders(0).MarketId

            For Each orderSummaryItem In CurrentOrderSummaryReport.CurrentOrders

                If orderSummaryItem.MarketId = selection.betfairCorrectScoreMarketId Then
                    If orderSummaryItem.SizeRemaining > 0 Then
                        unmatchedCSCount = unmatchedCSCount + 1
                    End If
                End If

                If orderSummaryItem.MarketId = selection.betfairUnderOver15MarketId Then
                    If orderSummaryItem.SizeRemaining > 0 Then
                        unmatchedUO15Count = unmatchedUO15Count + 1
                    End If
                End If

            Next

            ' Populate calling selection object
            selection.betfairUnderOver15UnmathedBets = unmatchedUO15Count.ToString
            selection.betfairCorrectScoreUnmathedBets = unmatchedCSCount.ToString

        Catch apiExcepion As APINGException
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, APINGExcepion msg : " + apiExcepion.Message, EventLogEntryType.Error)
            Exit Sub
        Catch ex As System.Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, system exception: " + ex.Message, EventLogEntryType.Error)
            Exit Sub

        Finally

        End Try

    End Sub
    Public Sub listMarketBook(ByRef selection As Selection, marketId As String)

        Dim client As IClient = Nothing
        Dim clientType As String = Nothing
        client = New JsonRpcClient(globalBetFairUrl, globalBetFairAppKey, globalBetFairToken)
        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Starting to get market book for Market Id: " + marketId.ToString, EventLogEntryType.Information)

        Try

            Dim marketIds As IList(Of String) = New List(Of String)()
            marketIds.Add(marketId)

            ' Set-up market projection
            Dim marketProjections As ISet(Of MarketProjection) = New HashSet(Of MarketProjection)()
            marketProjections.Add(MarketProjection.RUNNER_METADATA)
            marketProjections.Add(MarketProjection.EVENT)


            Dim markets = client.listMarketBook(marketIds, marketProjections)
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Response from listMarketBook : " + markets.Count.ToString, EventLogEntryType.Information)

            ' Store Market Details

        Catch apiExcepion As APINGException
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, APINGExcepion msg : " + apiExcepion.Message, EventLogEntryType.Error)
            Exit Sub
        Catch ex As System.Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Error getting Api data, system exception: " + ex.Message, EventLogEntryType.Error)
            Exit Sub

        Finally

        End Try

    End Sub

    Private Shared Function MarketIdNothing(ByVal s As BeffairEventClass) _
        As Boolean

        Return s.marketId Is Nothing

    End Function
    Private Shared Function DisplaySet(ByVal coll As HashSet(Of String)) As String
        Dim strReturn As String
        strReturn = "{"
        For Each i As String In coll
            strReturn = strReturn + " " + i
        Next i
        strReturn = strReturn + "}"
        Return strReturn
    End Function

    Private Function convertMarketStatus(statusEnum) As String

        If statusEnum = 0 Then
            Return "INACTIVE"
        ElseIf statusEnum = 1 Then
            Return "OPEN"
        ElseIf statusEnum = 2 Then
            Return "SUSPENDED"
        ElseIf statusEnum = 3 Then
            Return "CLOSED"
        Else
            Return "UNKNOWN"
        End If

    End Function
    Private Function convertRunnerStatus(statusEnum) As String

        If statusEnum = 0 Then
            Return "ACTIVE"
        ElseIf statusEnum = 1 Then
            Return "WINNER"
        ElseIf statusEnum = 2 Then
            Return "LOSER"
        ElseIf statusEnum = 3 Then
            Return "PLACED"   'The runner was placed, applies to EACH_WAY marketTypes only.
        ElseIf statusEnum = 4 Then
            Return "REMOVED_VACANT" ' applies to Greyhounds. Greyhound markets always return a fixed number of runners (traps). If a dog has been removed, the trap Is shown as vacant.
        ElseIf statusEnum = 5 Then
            Return "REMOVED"
        ElseIf statusEnum = 6 Then
            Return "HIDDEN"
        Else
            Return "UNKNOWN"
        End If

    End Function

End Class
