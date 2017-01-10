Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Linq
Imports System.Text
Imports BetfairSoccerBettingApp.Api_ng_sample_code.TO
Imports System.Web.Services.Protocols
Imports System.Net
Imports System.IO
Imports BetfairSoccerBettingApp.Api_ng_sample_code.Json

Namespace Api_ng_sample_code
	Public Class JsonRpcClient
		Inherits HttpWebClientProtocol
		Implements IClient

		Private privateEndPoint As String
		Public Property EndPoint() As String
			Get
				Return privateEndPoint
			End Get
			Private Set(ByVal value As String)
				privateEndPoint = value
			End Set
		End Property
		Private Shared ReadOnly operationReturnTypeMap As IDictionary(Of String, Type) = New Dictionary(Of String, Type)()
		Public Const APPKEY_HEADER As String = "X-Application"
		Public Const SESSION_TOKEN_HEADER As String = "X-Authentication"
		Public Property CustomHeaders() As NameValueCollection
		Private Shared ReadOnly LIST_EVENT_TYPES_METHOD As String = "SportsAPING/v1.0/listEventTypes"
		Private Shared ReadOnly LIST_MARKET_CATALOGUE_METHOD As String = "SportsAPING/v1.0/listMarketCatalogue"
        Private Shared ReadOnly LIST_MARKET_BOOK_METHOD As String = "SportsAPING/v1.0/listMarketBook"
        Private Shared ReadOnly LIST_MARKET_PROFIT_AND_LOSS As String = "SportsAPING/v1.0/listMarketProfitAndLoss"
        Private Shared ReadOnly PLACE_ORDERS_METHOD As String = "SportsAPING/v1.0/placeOrders"
        Private Shared ReadOnly FILTER As String = "filter"
        Private Shared ReadOnly LOCALE As String = "locale"
		Private Shared ReadOnly CURRENCY_CODE As String = "currencyCode"
		Private Shared ReadOnly MARKET_PROJECTION As String = "marketProjection"
		Private Shared ReadOnly MATCH_PROJECTION As String = "matchProjection"
		Private Shared ReadOnly ORDER_PROJECTION As String = "orderProjection"
		Private Shared ReadOnly PRICE_PROJECTION As String = "priceProjection"
		Private Shared ReadOnly SORT As String = "sort"
		Private Shared ReadOnly MAX_RESULTS As String = "maxResults"
		Private Shared ReadOnly MARKET_IDS As String = "marketIds"
		Private Shared ReadOnly MARKET_ID As String = "marketId"
		Private Shared ReadOnly INSTRUCTIONS As String = "instructions"
        Private Shared ReadOnly CUSTOMER_REFERENCE As String = "customerRef"
        Private Shared ReadOnly INCLUDE_SETTLED_BETS As String = "includeSettledBets"
        Private Shared ReadOnly INCLUDE_BSP_BETS As String = "includeBspBets"
        Private Shared ReadOnly NET_OF_COMMISSION As String = "netOfCommission"

        '
        ' ----- c# structures added here https://github.com/betfair/API-NG-sample-code/blob/master/cSharp/Api-ng-sample-code/Api-ng-sample-code/TO/InstructionReportErrorCode.cs
        '
        Private Shared ReadOnly LIST_COMPETITIONS_METHOD As String = "SportsAPING/v1.0/listCompetitions"
        Private Shared ReadOnly LIST_COUNTRIES_METHOD As String = "SportsAPING/v1.0/listCountries"
        Private Shared ReadOnly LIST_CLEARED_ORDERS_METHOD As String = "SportsAPING/v1.0/listClearedOrders"
        Private Shared ReadOnly LIST_CURRENT_ORDERS_METHOD As String = "SportsAPING/v1.0/listCurrentOrders"
        Private Shared ReadOnly LIST_EVENTS_METHOD As String = "SportsAPING/v1.0/listEvents"
        Private Shared ReadOnly LIST_MARKET_TYPES As String = "SportsAPING/v1.0/listMarketTypes"
        Private Shared ReadOnly LIST_TIME_RANGES As String = "SportsAPING/v1.0/listTimeRanges"
        Private Shared ReadOnly LIST_VENUES As String = "SportsAPING/v1.0/listVenues"
        Private Shared ReadOnly CANCEL_ORDERS_METHOD As String = "SportsAPING/v1.0/cancelOrders"
        Private Shared ReadOnly REPLACE_ORDERS_METHOD As String = "SportsAPING/v1.0/replaceOrders"
        Private Shared ReadOnly UPDATE_ORDERS_METHOD As String = "SportsAPING/v1.0/updateOrders"

        Private Shared ReadOnly GET_ACCOUNT_DETAILS As String = "AccountAPING/v1.0/getAccountDetails"
        Private Shared ReadOnly GET_ACCOUNT_FUNDS As String = "AccountAPING/v1.0/getAccountFunds"
        Private Shared ReadOnly GET_ACCOUNT_STATEMENT As String = "AccountAPING/v1.0/getAccountStatement"
        Private Shared ReadOnly LIST_CURRENCY_RATES As String = "AccountAPING/v1.0/listCurrencyRates"
        Private Shared ReadOnly TRANSFER_FUNDS As String = "AccountAPING/v1.0/transferFunds"

        Private Shared ReadOnly LIST_RACE_DETAILS As String = "ScoresAPING/v1.0/listRaceDetails"

        Private Shared ReadOnly BET_IDS As String = "betIds"
        Private Shared ReadOnly RUNNER_IDS As String = "runnerIds"
        Private Shared ReadOnly SIDE As String = "side"
        Private Shared ReadOnly SETTLED_DATE_RANGE As String = "settledDateRange"
        Private Shared ReadOnly EVENT_TYPE_IDS As String = "eventTypeIds"
        Private Shared ReadOnly EVENT_IDS As String = "eventIds"
        Private Shared ReadOnly BET_STATUS As String = "betStatus"
        Private Shared ReadOnly PLACED_DATE_RANGE As String = "placedDateRange"
        Private Shared ReadOnly DATE_RANGE As String = "dateRange"
        Private Shared ReadOnly ORDER_BY As String = "orderBy"
        Private Shared ReadOnly GROUP_BY As String = "groupBy"
        Private Shared ReadOnly SORT_DIR As String = "sortDir"
        Private Shared ReadOnly FROM_RECORD As String = "fromRecord"
        Private Shared ReadOnly RECORD_COUNT As String = "recordCount"
        Private Shared ReadOnly GRANULARITY As String = "granularity"
        Private Shared ReadOnly INCLUDE_ITEM_DESCRIPTION As String = "includeItemDescription"
        Private Shared ReadOnly FROM_CURRENCY As String = "fromCurrency"
        Private Shared ReadOnly FROM As String = "from"
        Private Shared ReadOnly [TO] As String = "to"
        Private Shared ReadOnly AMOUNT As String = "amount"
        Private Shared ReadOnly WALLET As String = "wallet"
        Private Shared ReadOnly MARKET_VERSION As String = "marketVersion"
        Private Shared ReadOnly MEETINGIDS As String = "meetingIds"
        Private Shared ReadOnly RACEIDS As String = "raceIds"

        '-----

        Public Sub New(ByVal endPoint As String, ByVal appKey As String, ByVal sessionToken As String)
			Me.EndPoint = endPoint & "/json-rpc/v1"
			CustomHeaders = New NameValueCollection()
			If appKey IsNot Nothing Then
				CustomHeaders(APPKEY_HEADER) = appKey
			End If
			If sessionToken IsNot Nothing Then
				CustomHeaders(SESSION_TOKEN_HEADER) = sessionToken
			End If
		End Sub

        Public Function listEventTypes(ByVal marketFilter As MarketFilter, Optional ByVal locale As String = Nothing) As IList(Of EventTypeResult) Implements IClient.listEventTypes
            Dim args = New Dictionary(Of String, Object)()
            args(FILTER) = marketFilter
            args(JsonRpcClient.LOCALE) = locale
            Return Invoke(Of List(Of EventTypeResult))(LIST_EVENT_TYPES_METHOD, args)

        End Function

        Public Function listEvents(ByVal marketFilter As MarketFilter, Optional ByVal locale As String = Nothing) As IList(Of EventResult) Implements IClient.listEvents
            Dim args = New Dictionary(Of String, Object)()
            args(FILTER) = marketFilter
            args(JsonRpcClient.LOCALE) = locale
            Return Invoke(Of List(Of EventResult))(LIST_EVENTS_METHOD, args)

        End Function

        Public Function listMarketCatalogue(ByVal marketFilter As MarketFilter, ByVal marketProjections As ISet(Of MarketProjection), ByVal marketSort As MarketSort, Optional ByVal maxResult As String = "1", Optional ByVal locale As String = Nothing) As IList(Of MarketCatalogue) Implements IClient.listMarketCatalogue
			Dim args = New Dictionary(Of String, Object)()
			args(FILTER) = marketFilter
			args(MARKET_PROJECTION) = marketProjections
			args(SORT) = marketSort
			args(MAX_RESULTS) = maxResult
			args(JsonRpcClient.LOCALE) = locale
			Return Invoke(Of List(Of MarketCatalogue))(LIST_MARKET_CATALOGUE_METHOD, args)
		End Function

        Public Function listMarketProfitAndLoss(ByVal marketIds As IList(Of String), Optional ByVal includeSettledBets As Boolean = False, Optional ByVal includeBspBets As Boolean = False, Optional ByVal netOfCommission As Boolean = False) As IList(Of MarketProfitAndLoss) Implements IClient.listMarketProfitAndLoss
            Dim args = New Dictionary(Of String, Object)()
            args(MARKET_IDS) = marketIds
            args(INCLUDE_SETTLED_BETS) = includeSettledBets
            args(INCLUDE_BSP_BETS) = includeBspBets
            args(NET_OF_COMMISSION) = netOfCommission

            Return Invoke(Of List(Of MarketProfitAndLoss))(LIST_MARKET_PROFIT_AND_LOSS, args)
        End Function

        Public Function placeOrders(ByVal marketId As String, ByVal customerRef As String, ByVal placeInstructions As IList(Of PlaceInstruction), Optional ByVal locale As String = Nothing) As PlaceExecutionReport Implements IClient.placeOrders
            Dim args = New Dictionary(Of String, Object)()

            args(MARKET_ID) = marketId
            args(INSTRUCTIONS) = placeInstructions
            args(CUSTOMER_REFERENCE) = customerRef
            args(JsonRpcClient.LOCALE) = locale

            Return Invoke(Of PlaceExecutionReport)(PLACE_ORDERS_METHOD, args)
        End Function
        Public Function listMarketBook(ByVal marketIds As IList(Of String), ByVal priceProjection As PriceProjection, Optional ByVal orderProjection? As OrderProjection = Nothing, Optional ByVal matchProjection? As MatchProjection = Nothing, Optional ByVal currencyCode As String = Nothing, Optional ByVal locale As String = Nothing) As IList(Of MarketBook) Implements IClient.listMarketBook
            Dim args = New Dictionary(Of String, Object)()
            args(MARKET_IDS) = marketIds
            args(PRICE_PROJECTION) = priceProjection
            args(ORDER_PROJECTION) = orderProjection
            args(MATCH_PROJECTION) = matchProjection
            args(JsonRpcClient.LOCALE) = locale
            args(CURRENCY_CODE) = currencyCode
            Return Invoke(Of List(Of MarketBook))(LIST_MARKET_BOOK_METHOD, args)
        End Function
        Public Function listCurrentOrders(Optional ByVal betIds As ISet(Of [String]) = Nothing, Optional ByVal marketIds As ISet(Of [String]) = Nothing, Optional orderProjection As System.Nullable(Of OrderProjection) = Nothing, Optional placedDateRange As TimeRange = Nothing, Optional orderBy As System.Nullable(Of OrderBy) = Nothing, Optional sortDir As System.Nullable(Of SortDir) = Nothing,
    Optional fromRecord As System.Nullable(Of Integer) = Nothing, Optional recordCount As System.Nullable(Of Integer) = Nothing) As CurrentOrderSummaryReport Implements IClient.listCurrentOrders
            Dim args = New Dictionary(Of String, Object)()
            args(BET_IDS) = betIds
            args(MARKET_IDS) = marketIds
            args(ORDER_PROJECTION) = orderProjection
            args(PLACED_DATE_RANGE) = placedDateRange
            args(ORDER_BY) = orderBy
            args(SORT_DIR) = sortDir
            args(FROM_RECORD) = fromRecord
            args(RECORD_COUNT) = recordCount

            Return Invoke(Of CurrentOrderSummaryReport)(LIST_CURRENT_ORDERS_METHOD, args)
        End Function

        Public Function cancelOrders(marketId As String, instructions__1 As IList(Of CancelInstruction), customerRef As String) As CancelExecutionReport Implements IClient.cancelOrders
            Dim args = New Dictionary(Of String, Object)()
            args(MARKET_ID) = marketId
            args(INSTRUCTIONS) = instructions__1
            args(CUSTOMER_REFERENCE) = customerRef

            Return Invoke(Of CancelExecutionReport)(CANCEL_ORDERS_METHOD, args)
        End Function

        Public Function replaceOrders(marketId As [String], instructions__1 As IList(Of ReplaceInstruction), customerRef As [String]) As ReplaceExecutionReport Implements IClient.replaceOrders
            Dim args = New Dictionary(Of String, Object)()
            args(MARKET_ID) = marketId
            args(INSTRUCTIONS) = instructions__1
            args(CUSTOMER_REFERENCE) = customerRef

            Return Invoke(Of ReplaceExecutionReport)(REPLACE_ORDERS_METHOD, args)
        End Function

        Public Function updateOrders(marketId As [String], instructions__1 As IList(Of UpdateInstruction), customerRef As [String]) As UpdateExecutionReport Implements IClient.updateOrders
            Dim args = New Dictionary(Of String, Object)()
            args(MARKET_ID) = marketId
            args(INSTRUCTIONS) = instructions__1
            args(CUSTOMER_REFERENCE) = customerRef

            Return Invoke(Of UpdateExecutionReport)(UPDATE_ORDERS_METHOD, args)
        End Function

        Protected Function CreateWebRequest(ByVal uri As Uri) As WebRequest
            Dim request As WebRequest = WebRequest.Create(New Uri(EndPoint))
            request.Method = "POST"
            request.ContentType = "application/json-rpc"
            request.Headers.Add(HttpRequestHeader.AcceptCharset, "ISO-8859-1,utf-8")
            request.Headers.Add(CustomHeaders)
            Return request
        End Function

        Public Function Invoke(Of T)(ByVal method As String, Optional ByVal args As IDictionary(Of String, Object) = Nothing) As T
			If method Is Nothing Then
				Throw New ArgumentNullException("method")
			End If
			If method.Length = 0 Then
				Throw New ArgumentException(Nothing, "method")
			End If

			Dim request = CreateWebRequest(New Uri(EndPoint))

			Using stream As Stream = request.GetRequestStream()
			Using writer As New StreamWriter(stream, Encoding.UTF8)
				Dim [call] = New JsonRequest With {.Method = method, .Id = 1, .Params = args}
				JsonConvert.Export([call], writer)
			End Using
			End Using
			Console.WriteLine(vbLf & "Calling: " & method & " With args: " & JsonConvert.Serialize(Of IDictionary(Of String, Object))(args))

			Using response As WebResponse = GetWebResponse(request)
			Using stream As Stream = response.GetResponseStream()
			Using reader As New StreamReader(stream, Encoding.UTF8)
				Dim jsonResponse = JsonConvert.Import(Of T)(reader)
				Console.WriteLine(vbLf & "Got Response: " & JsonConvert.Serialize(Of JsonResponse(Of T))(jsonResponse))
				If jsonResponse.HasError Then
					Throw ReconstituteException(jsonResponse.Error)
				Else
					Return jsonResponse.Result
				End If
			End Using
			End Using
			End Using
		End Function


		Private Shared Function ReconstituteException(ByVal ex As Api_ng_sample_code.TO.Exception) As System.Exception
			Dim data = ex.Data

			' API-NG exception -- it must have "data" element to tell us which exception
			Dim exceptionName = data.Property("exceptionname").Value.ToString()
			Dim exceptionData = data.Property(exceptionName).Value.ToString()
			Return JsonConvert.Deserialize(Of APINGException)(exceptionData)
		End Function
	End Class
End Namespace
