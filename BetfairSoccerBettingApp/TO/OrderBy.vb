
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace Api_ng_sample_code.TO
    <JsonConverter(GetType(StringEnumConverter))>
    Public Enum OrderBy
        BY_BET
        BY_MARKET
        BY_MATCH_TIME
        BY_PLACE_TIME
        BY_SETTLED_TIME
        BY_VOID_TIME
    End Enum
End Namespace
