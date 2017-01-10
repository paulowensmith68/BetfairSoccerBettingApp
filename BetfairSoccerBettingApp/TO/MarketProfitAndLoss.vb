Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class MarketProfitAndLoss
        <JsonProperty(PropertyName:="marketId")>
        Public Property MarketId() As String

        <JsonProperty(PropertyName:="commissionApplied")>
        Public Property CommissionApplied() As Double

        <JsonProperty(PropertyName:="profitAndLosses")>
        Public Property ProfitAndLosses() As List(Of RunnerProfitAndLoss)

        Public Overrides Function ToString() As String
            Dim sb = (New StringBuilder()).AppendFormat("{0}", "MarketProfitAndLoss").AppendFormat(" : MarketId={0}", MarketId).AppendFormat(" : CommissionApplied={0}", CommissionApplied)

            If ProfitAndLosses IsNot Nothing AndAlso ProfitAndLosses.Count > 0 Then
                Dim idx As Integer = 0
                For Each RunnerProfitAndLosses In ProfitAndLosses
                    sb.AppendFormat(" : RunnerProfitAndLosses[{0}]={1}", idx, RunnerProfitAndLosses)
                    idx += 1
                Next RunnerProfitAndLosses
            End If

            Return sb.ToString()
        End Function
    End Class
End Namespace
