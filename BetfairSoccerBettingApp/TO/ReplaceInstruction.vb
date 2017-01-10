Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class ReplaceInstruction
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

        <JsonProperty(PropertyName:="newPrice")>
        Public Property NewPrice() As Double
            Get
                Return m_NewPrice
            End Get
            Set
                m_NewPrice = Value
            End Set
        End Property
        Private m_NewPrice As Double

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "ReplaceInstruction").AppendFormat(" : BetId={0}", BetId).AppendFormat(" : NewPrice={0}", NewPrice)

            Return sb.ToString()
        End Function
    End Class
End Namespace