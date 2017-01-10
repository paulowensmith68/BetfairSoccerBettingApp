
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class UpdateInstruction
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

        <JsonProperty(PropertyName:="newPersistenceType")>
        Public Property NewPersistenceType() As PersistenceType
            Get
                Return m_NewPersistenceType
            End Get
            Set
                m_NewPersistenceType = Value
            End Set
        End Property
        Private m_NewPersistenceType As PersistenceType

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "UpdateInstruction").AppendFormat(" : BetId={0}", BetId).AppendFormat(" : NewPersistenceType={0}", NewPersistenceType)

            Return sb.ToString()
        End Function
    End Class
End Namespace