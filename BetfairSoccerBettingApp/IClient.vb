Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports BetfairSoccerBettingApp.Api_ng_sample_code.TO

Namespace Api_ng_sample_code
	Public Interface IClient
        '''        
        '''         * calls api-ng to get a list of event types
        '''         * 
        '''         * 
        Function listEventTypes(ByVal marketFilter As MarketFilter, Optional ByVal locale As String = Nothing) As IList(Of EventTypeResult)

        '''        
        '''         * calls api-ng to get a list of events
        '''         * 
        '''         * 
        Function listEvents(ByVal marketFilter As MarketFilter, Optional ByVal locale As String = Nothing) As IList(Of EventResult)

        '''        
        '''         * calls api-ng to get a list of market catalogues
        '''         * 
        Function listMarketCatalogue(ByVal marketFilter As MarketFilter, ByVal marketProjections As ISet(Of MarketProjection), ByVal marketSort As MarketSort, Optional ByVal maxResult As String = "1", Optional ByVal locale As String = Nothing) As IList(Of MarketCatalogue)

'''        
'''         * calls api-ng to get more detailed info about the specified markets
'''         * 
		Function listMarketBook(ByVal marketIds As IList(Of String), ByVal priceProjection As PriceProjection, Optional ByVal orderProjection? As OrderProjection = Nothing, Optional ByVal matchProjection? As MatchProjection = Nothing, Optional ByVal currencyCode As String = Nothing, Optional ByVal locale As String = Nothing) As IList(Of MarketBook)

        '''        
        '''         * places a bet
        '''         * 
        Function placeOrders(ByVal marketId As String, ByVal customerRef As String, ByVal placeInstructions As IList(Of PlaceInstruction), Optional ByVal locale As String = Nothing) As PlaceExecutionReport

        '''        
        '''         * calls api-ng to get profit and loss
        '''         * 
        Function listMarketProfitAndLoss(ByVal marketIds As IList(Of String), Optional ByVal includeSettledBets As Boolean = False, Optional ByVal includeBspBets As Boolean = False, Optional ByVal netOfCommission As Boolean = False) As IList(Of MarketProfitAndLoss)

        '*
        '         * Lists current orders
        '         * 

        Function listCurrentOrders(Optional ByVal betIds As ISet(Of [String]) = Nothing, Optional ByVal marketIds As ISet(Of [String]) = Nothing, Optional orderProjection As System.Nullable(Of OrderProjection) = Nothing, Optional placedDateRange As TimeRange = Nothing, Optional orderBy As System.Nullable(Of OrderBy) = Nothing, Optional sortDir As System.Nullable(Of SortDir) = Nothing,
    Optional fromRecord As System.Nullable(Of Integer) = Nothing, Optional recordCount As System.Nullable(Of Integer) = Nothing) As CurrentOrderSummaryReport


        '*
        '         * Cancels a bet, or decreases its size
        '         * 

        Function cancelOrders(marketId As String, instructions As IList(Of CancelInstruction), customerRef As String) As CancelExecutionReport

        '*
        '         * Replaces a bet: changes the price
        '         * 

        Function replaceOrders(marketId As [String], instructions As IList(Of ReplaceInstruction), customerRef As [String]) As ReplaceExecutionReport

        '*
        '         * updates a bet
        '         * 

        Function updateOrders(marketId As [String], instructions As IList(Of UpdateInstruction), customerRef As [String]) As UpdateExecutionReport


    End Interface
End Namespace
