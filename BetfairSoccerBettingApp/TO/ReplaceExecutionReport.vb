﻿
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class ReplaceExecutionReport
        <JsonProperty(PropertyName:="customerRef")>
        Public Property CustomerRef() As [String]
            Get
                Return m_CustomerRef
            End Get
            Set
                m_CustomerRef = Value
            End Set
        End Property
        Private m_CustomerRef As [String]

        <JsonProperty(PropertyName:="status")>
        Public Property Status() As ExecutionReportStatus
            Get
                Return m_Status
            End Get
            Set
                m_Status = Value
            End Set
        End Property
        Private m_Status As ExecutionReportStatus

        <JsonProperty(PropertyName:="errorCode")>
        Public Property ErrorCode() As ExecutionReportErrorCode
            Get
                Return m_ErrorCode
            End Get
            Set
                m_ErrorCode = Value
            End Set
        End Property
        Private m_ErrorCode As ExecutionReportErrorCode

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

        <JsonProperty(PropertyName:="instructionReports")>
        Public Property InstructionReports() As List(Of ReplaceInstructionReport)
            Get
                Return m_InstructionReports
            End Get
            Set
                m_InstructionReports = Value
            End Set
        End Property
        Private m_InstructionReports As List(Of ReplaceInstructionReport)

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()

            sb.AppendFormat("{0}", "ReplaceExecutionReport").AppendFormat(" : CustomerRef={0}", CustomerRef).AppendFormat(" : Status={0}", Status).AppendFormat(" : ErrorCode={0}", ErrorCode).AppendFormat(" : MarketId={0}", MarketId).AppendFormat(" : InstructionReports={0}", InstructionReports)

            Return sb.ToString()
        End Function
    End Class
End Namespace

