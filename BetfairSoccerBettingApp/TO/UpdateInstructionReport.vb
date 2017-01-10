
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class UpdateInstructionReport
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

        <JsonProperty(PropertyName:="instruction")>
        Public Property Instruction() As UpdateInstruction
            Get
                Return m_Instruction
            End Get
            Set
                m_Instruction = Value
            End Set
        End Property
        Private m_Instruction As UpdateInstruction

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "UpdateInstructionReport").AppendFormat(" : Status={0}", Status).AppendFormat(" : ErrorCode={0}", ErrorCode).AppendFormat(" : Instruction={0}", Instruction)

            Return sb.ToString()
        End Function
    End Class
End Namespace