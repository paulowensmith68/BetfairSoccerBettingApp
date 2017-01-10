
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class ReplaceInstructionReport
        <JsonProperty(PropertyName:="status")>
        Public Property Status() As InstructionReportStatus
            Get
                Return m_Status
            End Get
            Set
                m_Status = Value
            End Set
        End Property
        Private m_Status As InstructionReportStatus

        <JsonProperty(PropertyName:="placeInstructionReport")>
        Public Property PlaceInstructionReport() As PlaceInstructionReport
            Get
                Return m_PlaceInstructionReport
            End Get
            Set
                m_PlaceInstructionReport = Value
            End Set
        End Property
        Private m_PlaceInstructionReport As PlaceInstructionReport

        <JsonProperty(PropertyName:="errorCode")>
        Public Property ErrorCode() As InstructionReportErrorCode
            Get
                Return m_ErrorCode
            End Get
            Set
                m_ErrorCode = Value
            End Set
        End Property
        Private m_ErrorCode As InstructionReportErrorCode

        <JsonProperty(PropertyName:="cancelInstructionReport")>
        Public Property CancelInstructionReport() As CancelInstructionReport
            Get
                Return m_CancelInstructionReport
            End Get
            Set
                m_CancelInstructionReport = Value
            End Set
        End Property
        Private m_CancelInstructionReport As CancelInstructionReport

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "ReplaceInstructionReport").AppendFormat(" : Status={0}", Status).AppendFormat(" : PlaceInstructionReport={0}", PlaceInstructionReport).AppendFormat(" : ErrorCode={0}", ErrorCode).AppendFormat(" : CancelInstructionReport={0}", CancelInstructionReport)

            Return sb.ToString()
        End Function
    End Class
End Namespace