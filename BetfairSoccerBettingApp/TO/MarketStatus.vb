﻿Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace Api_ng_sample_code.TO
	<JsonConverter(GetType(StringEnumConverter))> _
	Public Enum MarketStatus
		INACTIVE
		OPEN
		SUSPENDED
		CLOSED
	End Enum
End Namespace
