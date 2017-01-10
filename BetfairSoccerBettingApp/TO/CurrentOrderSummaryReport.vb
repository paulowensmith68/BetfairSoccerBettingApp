
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class CurrentOrderSummaryReport
        <JsonProperty(PropertyName:="currentOrders")>
        Public Property CurrentOrders() As List(Of CurrentOrderSummary)
            Get
                Return m_CurrentOrders
            End Get
            Set
                m_CurrentOrders = Value
            End Set
        End Property
        Private m_CurrentOrders As List(Of CurrentOrderSummary)

        <JsonProperty(PropertyName:="moreAvailable")>
        Public Property MoreAvailable() As Boolean
            Get
                Return m_MoreAvailable
            End Get
            Set
                m_MoreAvailable = Value
            End Set
        End Property
        Private m_MoreAvailable As Boolean

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "CurrentOrderSummaryReport")

            If CurrentOrders IsNot Nothing AndAlso CurrentOrders.Count > 0 Then
                Dim idx As Integer = 0
                For Each currentOrder In CurrentOrders
                    sb.AppendFormat(" : CurrentOrder[{0}]={1}", System.Math.Max(System.Threading.Interlocked.Increment(idx), idx - 1), currentOrder)
                Next
            End If

            sb.AppendFormat(" : MoreAvailable={0}", MoreAvailable)

            Return sb.ToString()
        End Function
    End Class
End Namespace