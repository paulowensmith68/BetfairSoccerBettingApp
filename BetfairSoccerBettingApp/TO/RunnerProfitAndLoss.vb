Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class RunnerProfitAndLoss
        <JsonProperty(PropertyName:="selectionId")>
        Public Property SelectionId() As Long

        <JsonProperty(PropertyName:="ifWin")>
        Public Property IfWin() As Double?

        <JsonProperty(PropertyName:="ifLose")>
        Public Property IfLose() As Double?

        <JsonProperty(PropertyName:="ifPlace")>
        Public Property IfPlace() As Double?

        Public Overrides Function ToString() As String
            Dim sb = (New StringBuilder()).AppendFormat("SelectionId={0}", SelectionId).AppendFormat(" : IfWin={0}", IfWin).AppendFormat(" : IfLose={0}", IfLose).AppendFormat(" : ifPlace={0}", IfPlace)

            Return sb.ToString()

        End Function
    End Class
End Namespace
