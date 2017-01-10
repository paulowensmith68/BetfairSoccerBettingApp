Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class EventResult
        <JsonProperty(PropertyName:="event")>
        Public Property [Event]() As [Event]
            Get
                Return m_Event
            End Get
            Set
                m_Event = Value
            End Set
        End Property
        Private m_Event As [Event]

        <JsonProperty(PropertyName:="marketCount")>
        Public Property MarketCount() As Integer
            Get
                Return m_MarketCount
            End Get
            Set
                m_MarketCount = Value
            End Set
        End Property
        Private m_MarketCount As Integer

        Public Overrides Function ToString() As String
            Return New StringBuilder().AppendFormat("{0}", "EventResult").AppendFormat(" : {0}", [Event]).AppendFormat(" : MarketCount={0}", MarketCount).ToString()
        End Function
    End Class
End Namespace
