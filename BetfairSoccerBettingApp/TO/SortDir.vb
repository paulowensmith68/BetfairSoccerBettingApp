
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace Api_ng_sample_code.TO
    <JsonConverter(GetType(StringEnumConverter))>
    Public Enum SortDir
        EARLIEST_TO_LATEST
        LATEST_TO_EARLIEST
    End Enum
End Namespace
