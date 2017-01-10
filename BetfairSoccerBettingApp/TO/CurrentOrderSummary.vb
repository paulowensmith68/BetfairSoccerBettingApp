Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class CurrentOrderSummary
        <JsonProperty(PropertyName:="betId")>
        Public Property BetId() As String
            Get
                Return m_BetId
            End Get
            Set
                m_BetId = Value
            End Set
        End Property
        Private m_BetId As String

        <JsonProperty(PropertyName:="marketId")>
        Public Property MarketId() As String
            Get
                Return m_MarketId
            End Get
            Set
                m_MarketId = Value
            End Set
        End Property
        Private m_MarketId As String

        <JsonProperty(PropertyName:="selectionId")>
        Public Property SelectionId() As String
            Get
                Return m_SelectionId
            End Get
            Set
                m_SelectionId = Value
            End Set
        End Property
        Private m_SelectionId As String

        <JsonProperty(PropertyName:="handicap")>
        Public Property Handicap() As String
            Get
                Return m_Handicap
            End Get
            Set
                m_Handicap = Value
            End Set
        End Property
        Private m_Handicap As String

        <JsonProperty(PropertyName:="priceSize")>
        Public Property PriceSize() As PriceSize
            Get
                Return m_PriceSize
            End Get
            Set
                m_PriceSize = Value
            End Set
        End Property
        Private m_PriceSize As PriceSize

        <JsonProperty(PropertyName:="bspLiability")>
        Public Property BspLiability() As Double
            Get
                Return m_BspLiability
            End Get
            Set
                m_BspLiability = Value
            End Set
        End Property
        Private m_BspLiability As Double

        <JsonProperty(PropertyName:="side")>
        Public Property Side() As Side
            Get
                Return m_Side
            End Get
            Set
                m_Side = Value
            End Set
        End Property
        Private m_Side As Side

        <JsonProperty(PropertyName:="status")>
        Public Property Status() As OrderStatus
            Get
                Return m_Status
            End Get
            Set
                m_Status = Value
            End Set
        End Property
        Private m_Status As OrderStatus

        <JsonProperty(PropertyName:="persistenceType")>
        Public Property PersistenceType() As PersistenceType
            Get
                Return m_PersistenceType
            End Get
            Set
                m_PersistenceType = Value
            End Set
        End Property
        Private m_PersistenceType As PersistenceType

        <JsonProperty(PropertyName:="orderType")>
        Public Property OrderType() As OrderType
            Get
                Return m_OrderType
            End Get
            Set
                m_OrderType = Value
            End Set
        End Property
        Private m_OrderType As OrderType

        <JsonProperty(PropertyName:="placedDate")>
        Public Property PlacedDate() As DateTime
            Get
                Return m_PlacedDate
            End Get
            Set
                m_PlacedDate = Value
            End Set
        End Property
        Private m_PlacedDate As DateTime

        <JsonProperty(PropertyName:="matchedDate")>
        Public Property MatchedDate() As DateTime
            Get
                Return m_MatchedDate
            End Get
            Set
                m_MatchedDate = Value
            End Set
        End Property
        Private m_MatchedDate As DateTime

        <JsonProperty(PropertyName:="averagePriceMatched")>
        Public Property AveragePriceMatched() As Double
            Get
                Return m_AveragePriceMatched
            End Get
            Set
                m_AveragePriceMatched = Value
            End Set
        End Property
        Private m_AveragePriceMatched As Double

        <JsonProperty(PropertyName:="sizeMatched")>
        Public Property SizeMatched() As Double
            Get
                Return m_SizeMatched
            End Get
            Set
                m_SizeMatched = Value
            End Set
        End Property
        Private m_SizeMatched As Double

        <JsonProperty(PropertyName:="sizeRemaining")>
        Public Property SizeRemaining() As Double
            Get
                Return m_SizeRemaining
            End Get
            Set
                m_SizeRemaining = Value
            End Set
        End Property
        Private m_SizeRemaining As Double

        <JsonProperty(PropertyName:="sizeLapsed")>
        Public Property SizeLapsed() As Double
            Get
                Return m_SizeLapsed
            End Get
            Set
                m_SizeLapsed = Value
            End Set
        End Property
        Private m_SizeLapsed As Double

        <JsonProperty(PropertyName:="sizeCancelled")>
        Public Property SizeCancelled() As Double
            Get
                Return m_SizeCancelled
            End Get
            Set
                m_SizeCancelled = Value
            End Set
        End Property
        Private m_SizeCancelled As Double

        <JsonProperty(PropertyName:="sizeVoided")>
        Public Property SizeVoided() As Double
            Get
                Return m_SizeVoided
            End Get
            Set
                m_SizeVoided = Value
            End Set
        End Property
        Private m_SizeVoided As Double

        <JsonProperty(PropertyName:="regulatorAuthCode")>
        Public Property RegulatorAuthCode() As String
            Get
                Return m_RegulatorAuthCode
            End Get
            Set
                m_RegulatorAuthCode = Value
            End Set
        End Property
        Private m_RegulatorAuthCode As String

        <JsonProperty(PropertyName:="regulatorCode")>
        Public Property RegulatorCode() As String
            Get
                Return m_RegulatorCode
            End Get
            Set
                m_RegulatorCode = Value
            End Set
        End Property
        Private m_RegulatorCode As String

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "CurrentOrderSummary").AppendFormat(" : BetId={0}", BetId).AppendFormat(" : MarketId={0}", MarketId).AppendFormat(" : SelectionId={0}", SelectionId).AppendFormat(" : Handicap={0}", Handicap).AppendFormat(" : PriceSize={0}", PriceSize).AppendFormat(" : BspLiability={0}", BspLiability).AppendFormat(" : Side={0}", Side).AppendFormat(" : Status={0}", Status).AppendFormat(" : PersistenceType={0}", PersistenceType).AppendFormat(" : OrderType={0}", OrderType).AppendFormat(" : PlacedDate={0}", PlacedDate).AppendFormat(" : MatchedDate={0}", MatchedDate).AppendFormat(" : AveragePriceMatched={0}", AveragePriceMatched).AppendFormat(" : SizeMatched={0}", SizeMatched).AppendFormat(" : SizeRemaining={0}", SizeRemaining).AppendFormat(" : SizeLapsed={0}", SizeLapsed).AppendFormat(" : SizeCancelled={0}", SizeCancelled).AppendFormat(" : SizeVoided={0}", SizeVoided).AppendFormat(" : RegulatorAuthCode={0}", RegulatorAuthCode).AppendFormat(" : RegulatorCode={0}", RegulatorCode)

            Return sb.ToString()
        End Function
    End Class

End Namespace

