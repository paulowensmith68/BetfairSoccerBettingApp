
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class CancelInstruction
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

        <JsonProperty(PropertyName:="sizeReduction")>
        Public Property SizeReduction() As System.Nullable(Of Double)
            Get
                Return m_SizeReduction
            End Get
            Set
                m_SizeReduction = Value
            End Set
        End Property
        Private m_SizeReduction As System.Nullable(Of Double)

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "CancelInstruction").AppendFormat(" : BetId={0}", BetId).AppendFormat(" : SizeReduction={0}", SizeReduction)

            Return sb.ToString()
        End Function
    End Class
End Namespace

