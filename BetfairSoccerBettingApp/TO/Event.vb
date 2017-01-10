
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Api_ng_sample_code.TO
    Public Class [Event]
        <JsonProperty(PropertyName:="id")>
        Public Property Id() As String
            Get
                Return m_Id
            End Get
            Set
                m_Id = Value
            End Set
        End Property
        Private m_Id As String

        <JsonProperty(PropertyName:="name")>
        Public Property Name() As String
            Get
                Return m_Name
            End Get
            Set
                m_Name = Value
            End Set
        End Property
        Private m_Name As String

        <JsonProperty(PropertyName:="countryCode")>
        Public Property CountryCode() As String
            Get
                Return m_CountryCode
            End Get
            Set
                m_CountryCode = Value
            End Set
        End Property
        Private m_CountryCode As String

        <JsonProperty(PropertyName:="timezone")>
        Public Property Timezone() As String
            Get
                Return m_Timezone
            End Get
            Set
                m_Timezone = Value
            End Set
        End Property
        Private m_Timezone As String

        <JsonProperty(PropertyName:="venue")>
        Public Property Venue() As String
            Get
                Return m_Venue
            End Get
            Set
                m_Venue = Value
            End Set
        End Property
        Private m_Venue As String

        <JsonProperty(PropertyName:="openDate")>
        Public Property OpenDate() As System.Nullable(Of DateTime)
            Get
                Return m_OpenDate
            End Get
            Set
                m_OpenDate = Value
            End Set
        End Property
        Private m_OpenDate As System.Nullable(Of DateTime)

        Public Overrides Function ToString() As String
            Return New StringBuilder().AppendFormat("{0}", "Event").AppendFormat(" : Id={0}", Id).AppendFormat(" : Name={0}", Name).AppendFormat(" : CountryCode={0}", CountryCode).AppendFormat(" : Venue={0}", Venue).AppendFormat(" : Timezone={0}", Timezone).AppendFormat(" : OpenDate={0}", OpenDate).ToString()
        End Function
    End Class
End Namespace
