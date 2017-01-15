Imports System.ComponentModel
Imports System.IO
Imports System.Net.Mail

Public Class frmMain

    Public sel1 As New Selection(1)
    Public sel2 As New Selection(2)
    Public sel3 As New Selection(3)
    Public sel4 As New Selection(4)


    Private intFileNumber As Integer = FreeFile()


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        ' Uses standard 2 digit code https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2
        '
        Dim marketCountriesUkOnly As HashSet(Of String)
        marketCountriesUkOnly = New HashSet(Of String)({"GB"})
        Dim marketIndiaOnly As HashSet(Of String)
        marketIndiaOnly = New HashSet(Of String)({"IN"})
        Dim marketGreeceOnly As HashSet(Of String)
        marketGreeceOnly = New HashSet(Of String)({"GR"})

        Dim marketCountriesEurope As HashSet(Of String)
        marketCountriesEurope = New HashSet(Of String)({"GB", "FR", "DE", "IT", "ES", "PT", "NL", "GR"})

        ' Login
        Account.Login()

        ' Populate initial list of event data
        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PollBetFairEvents(1, My.Settings.NumberOfUkEvents, marketCountriesEurope)
        Me.dgvEvents.DataSource = BetfairClass1.eventList
        BetfairClass1 = Nothing

    End Sub

    Public Function WriteToEventLog(ByVal entry As String, Optional ByVal eventType As EventLogEntryType = EventLogEntryType.Information) As Boolean

        Dim objEventLog As New EventLog
        Dim strLogFile As String

        ' Write to Event Logs
        Try

            ' Always write to text log file in application directory
            strLogFile = My.Settings.ProcessLogPath & "BetFairFeedService_Stream" & globalStreamSportId.ToString & "_" & globalStreamName & "_Log_File_" & Format(Now, "_yyyy_MM_dd") & ".txt"
            FileOpen(intFileNumber, strLogFile, OpenMode.Append)
            Dim strDate As String = Format(Now, "yyyy-MM-dd")
            Dim strTimestamp As String = Format(Now, "HH.mm.ss.ffffff")
            Dim strEntryType As String = ""
            Select Case eventType
                Case EventLogEntryType.Information
                    strEntryType = "Information"
                Case EventLogEntryType.Error
                    strEntryType = "Error"
                Case EventLogEntryType.FailureAudit
                    strEntryType = "Failure Audit"
                Case EventLogEntryType.SuccessAudit
                    strEntryType = "Sucsess Audit"
                Case EventLogEntryType.Warning
                    strEntryType = "Warning"
                Case Else
                    strEntryType = "Unknown"
            End Select

            PrintLine(intFileNumber, strDate & "." & strTimestamp & ", " & strEntryType & ", " & entry)
            FileClose(intFileNumber)

            Return True

        Catch Ex As Exception

            Return False

        End Try

    End Function

    Private Sub btnExit_Click(sender As Object, e As EventArgs)

        ' Logout
        Account.Logout()

        Application.Exit()

    End Sub

    Private Sub timerRefreshSelections_Tick(sender As Object, e As EventArgs) Handles timerRefreshSelections.Tick

        '' Clean log rich textbox
        'If rtbLog.Lines.Count > 1000 Then
        '    rtbLog.Clear()
        'End If

        If tbxSel1EventName.Text <> "" Then
            Refreshsel1Info()
        End If

        If tbxSel2EventName.Text <> "" Then
            Refreshsel2Info()
        End If

        If tbxSel3EventName.Text <> "" Then
            Refreshsel3Info()
        End If

        If tbxSel4EventName.Text <> "" Then
            Refreshsel4Info()
        End If

    End Sub

    '
    ' Autobets
    '
    Private Sub btnSel1AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel1AutoBetOn.Click

        If btnSel1AutoBetOn.Text = "Autobet On" Then

            If tbxSel1EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel1.autobetOver15BetMade = False
                    sel1.autobetUnder15BetMade = False
                    sel1.autobetCorrectScore00BetMade = False
                    sel1.autobetCorrectScore10BetMade = False
                    sel1.autobetCorrectScore01BetMade = False
                    sel1.autobetOver15TopUpBetMade = False

                    ' Set the interval
                    timerSel1AutoBet.Interval = nudSettingsAutoBetRefresh.Value

                    ' Enable Autobet timer
                    timerSel1AutoBet.Enabled = True
                    btnSel1AutoBetOn.Text = "Autobet Off"
                    btnSel1AutoBetOn.BackColor = Color.LightSalmon

                    ' Disable the Select button
                    btnSel1.Enabled = False

                    ' Write to log
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 has been switched ON.", EventLogEntryType.Information)

                    ' Call tick
                    timerSel1AutoBet_Tick(sender, e)

                End If


            End If
        Else

            ' Disable Autobet timer
            timerSel1AutoBet.Enabled = False

            ' Switch off
            btnSel1AutoBetOn.Text = "Autobet On"
            btnSel1AutoBetOn.BackColor = Color.LightGreen

            ' Enable the Select button
            btnSel1.Enabled = True

            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 has been switched OFF.", EventLogEntryType.Information)

        End If

    End Sub

    Private Sub timerSel1AutoBet_Tick(sender As Object, e As EventArgs) Handles timerSel1AutoBet.Tick

        '
        ' Update status of each bet type
        '
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel1ProfitStatus00.BackColor = Color.LawnGreen
                btnSel1ProfitStatus00.Text = sel1.betfairCorrectScore00IfWinProfit
            Else
                btnSel1ProfitStatus00.BackColor = Color.White
                btnSel1ProfitStatus00.Text = ""
            End If
        Else
            btnSel1ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel1ProfitStatus10.BackColor = Color.LawnGreen
                btnSel1ProfitStatus10.Text = sel1.betfairCorrectScore10IfWinProfit
            Else
                btnSel1ProfitStatus10.BackColor = Color.White
                btnSel1ProfitStatus10.Text = ""
            End If
        Else
            btnSel1ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel1ProfitStatus01.BackColor = Color.LawnGreen
                btnSel1ProfitStatus01.Text = sel1.betfairCorrectScore01IfWinProfit
            Else
                btnSel1ProfitStatus01.BackColor = Color.White
                btnSel1ProfitStatus01.Text = ""
            End If
        Else
            btnSel1ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
            If CDbl(sel1.betfairUnder15IfWinProfit) > 0 Then
                btnSel1ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel1ProfitStatusUnder15.Text = sel1.betfairUnder15IfWinProfit
            Else
                btnSel1ProfitStatusUnder15.BackColor = Color.White
                btnSel1ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel1ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
            If CDbl(sel1.betfairOver15IfWinProfit) > 0 Then
                btnSel1ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel1ProfitStatusOver15.Text = sel1.betfairOver15IfWinProfit
            Else
                btnSel1ProfitStatusOver15.BackColor = Color.White
                btnSel1ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel1ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel1.betfairEventInplay = "True" Then
            ' Continue
        Else
            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
            Exit Sub
        End If

        ' Populate Unmatched Order counts
        If String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
            sel1.betfairOver15Orders = 0
        End If
        If String.IsNullOrEmpty(sel1.betfairUnder15Orders) Then
            sel1.betfairUnder15Orders = 0
        End If
        If String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
            sel1.betfairCorrectScore00Orders = 0
        End If
        If String.IsNullOrEmpty(sel1.betfairCorrectScore10Orders) Then
            sel1.betfairCorrectScore10Orders = 0
        End If
        If String.IsNullOrEmpty(sel1.betfairCorrectScore01Orders) Then
            sel1.betfairCorrectScore01Orders = 0
        End If


        ' 
        ' Look to Identifying starting position of the bet ..........
        '
        If sel1.autobetOver15BetMade = False Then

            If btnSel1ProfitStatusOver15.Text = "" And tbxSel1Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Identifying starting position", EventLogEntryType.Information)

                If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +40 Then

                        If Not String.IsNullOrEmpty(sel1.betfairOver15BackOdds) Then
                            If CDbl(sel1.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel1.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
                                    If CDbl(sel1.betfairOver15Orders) = 0 Then
                                        If sel1.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            ' Place back bet on Over1.5
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel1.betfairOver15BackOdds), "OUT")
                                            oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)
                                            stake = nudSettingsOver15Stake.Value.ToString
                                            sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel1.autobetOver15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel1.placeOver15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position - Inplay timer >40 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If

        End If


        ' 
        ' Look to ensure we have 0 - 0 covered
        '
        If sel1.autobetCorrectScore00BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetOver15BetMade = True And sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel1InplayTime.Text) > +25 And CDbl(tbxSel1InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then
                            If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                                    If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then
                                        If sel1.betfairCorrectScore00BackOdds <= nudSettingsCS00TargetPrice.Value Then

                                            ' calculate stake based on profit
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel1.betfairCorrectScore00BackOdds), "IN")
                                            oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel1.autobetCorrectScore00BetMade = True

                                            ' Place order on Correct Score 0-0 market
                                            sel1.placeCorrectScore00_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to ensure we have Under 1.5
        '
        If sel1.autobetUnder15BetMade = False Then

            ' Check the strategy has started and score only 1 goal
            If sel1.autobetOver15BetMade = True And tbxSel1Score.Text = "1 Goal scored" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +45 Then
                        If Not String.IsNullOrEmpty(sel1.betfairUnder15BackOdds) Then
                            If CDbl(sel1.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel1.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairUnder15Orders) Then
                                    If CDbl(sel1.betfairUnder15Orders) = 0 Then
                                        If sel1.betfairUnder15BackOdds >= nudSettingsUnder15TargetPrice.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsUnder15TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel1.betfairUnder15BackOdds), "IN")
                                            oddsMarket = CDbl(sel1.betfairUnder15BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel1.autobetUnder15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel1.placeUnder15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to cover 1 - 0, after 40 minutes play
        '
        If sel1.autobetCorrectScore10BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetOver15BetMade = True And sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel1InplayTime.Text) > +40 And CDbl(tbxSel1InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10BackOdds) Then
                            If CDbl(sel1.betfairCorrectScore10BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore10BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10Orders) Then
                                    If CDbl(sel1.betfairCorrectScore10Orders) = 0 Then
                                        If sel1.betfairCorrectScore10BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel1.betfairCorrectScore10BackOdds), "IN")
                                            oddsMarket = CDbl(sel1.betfairCorrectScore10BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 1 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel1.autobetCorrectScore10BetMade = True

                                            ' Place order on Correct Score 1-0 market
                                            sel1.placeCorrectScore10_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to cover 0 - 1, after 40 minutes play
        '
        If sel1.autobetCorrectScore01BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetOver15BetMade = True And sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel1InplayTime.Text) > +40 And CDbl(tbxSel1InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01BackOdds) Then
                            If CDbl(sel1.betfairCorrectScore01BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore01BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01Orders) Then
                                    If CDbl(sel1.betfairCorrectScore01Orders) = 0 Then
                                        If sel1.betfairCorrectScore01BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel1.betfairCorrectScore01BackOdds), "IN")
                                            oddsMarket = CDbl(sel1.betfairCorrectScore01BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 0 - 1 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel1.autobetCorrectScore01BetMade = True

                                            ' Place order on Correct Score 0-1 market
                                            sel1.placeCorrectScore01_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to boost the Over1.5 if we have taken 0-0 and either 1-0 or 0-1 after 40 minutes play
        '
        If sel1.autobetOver15TopUpBetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If (sel1.autobetOver15BetMade = True And sel1.autobetUnder15BetMade = False And sel1.autobetCorrectScore00BetMade = True And tbxSel1Score.Text = "0 - 0") And (sel1.autobetCorrectScore10BetMade = True Or sel1.autobetCorrectScore01BetMade = True) Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to boost Over 1.5", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel1InplayTime.Text) > +40 And CDbl(tbxSel1InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel1.betfairOver15BackOdds) Then
                            If CDbl(sel1.betfairOver15BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairOver15BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
                                    If CDbl(sel1.betfairOver15Orders) = 0 Then

                                        ' Place 2nd back bet on Over1.5
                                        Dim odds As Double
                                        Dim oddsMarket As Double
                                        Dim stake As Double
                                        odds = adjustOddsToMatch(CDbl(sel1.betfairOver15BackOdds), "OUT")
                                        oddsMarket = CDbl(sel1.betfairCorrectScore01BackOdds)
                                        stake = (nudSettingsOver15Stake.Value / 4)
                                        sendEmailToText("Match: " + sel1.betfairEventName + " Market: Boost to Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                        sel1.autobetOver15TopUpBetMade = True

                                        ' Place order on Over 1.5 market
                                        sel1.placeOver15_Order(odds, stake)

                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Boost Over 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Boost Over 1.5 position  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Boost Over 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Boost Over 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Boost Over 1.5 position - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Boost Over 1.5 position  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If

    End Sub

    Private Sub btnSel2AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel2AutoBetOn.Click

        If btnSel2AutoBetOn.Text = "Autobet On" Then

            If tbxSel2EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel2.autobetOver15BetMade = False
                    sel2.autobetUnder15BetMade = False
                    sel2.autobetCorrectScore00BetMade = False
                    sel2.autobetCorrectScore10BetMade = False
                    sel2.autobetCorrectScore01BetMade = False
                    sel2.autobetOver15TopUpBetMade = False

                    ' Set the interval
                    timerSel2AutoBet.Interval = nudSettingsAutoBetRefresh.Value

                    ' Enable Autobet timer
                    timerSel2AutoBet.Enabled = True
                    btnSel2AutoBetOn.Text = "Autobet Off"
                    btnSel2AutoBetOn.BackColor = Color.LightSalmon

                    ' Disable the Select button
                    btnSel2.Enabled = False

                    ' Write to log
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 has been switched ON.", EventLogEntryType.Information)

                    ' Call tick
                    timerSel2AutoBet_Tick(sender, e)

                End If


            End If
        Else

            ' Disable Autobet timer
            timerSel2AutoBet.Enabled = False

            ' Switch off
            btnSel2AutoBetOn.Text = "Autobet On"
            btnSel2AutoBetOn.BackColor = Color.LightGreen

            ' Enable the Select button
            btnSel2.Enabled = True

            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 has been switched OFF.", EventLogEntryType.Information)

        End If

    End Sub

    Private Sub timerSel2AutoBet_Tick(sender As Object, e As EventArgs) Handles timerSel2AutoBet.Tick

        '
        ' Update status of each bet type
        '
        If Not String.IsNullOrEmpty(sel2.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel2.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel2ProfitStatus00.BackColor = Color.LawnGreen
                btnSel2ProfitStatus00.Text = sel2.betfairCorrectScore00IfWinProfit
            Else
                btnSel2ProfitStatus00.BackColor = Color.White
                btnSel2ProfitStatus00.Text = ""
            End If
        Else
            btnSel2ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel2.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel2.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel2ProfitStatus10.BackColor = Color.LawnGreen
                btnSel2ProfitStatus10.Text = sel2.betfairCorrectScore10IfWinProfit
            Else
                btnSel2ProfitStatus10.BackColor = Color.White
                btnSel2ProfitStatus10.Text = ""
            End If
        Else
            btnSel2ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel2.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel2.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel2ProfitStatus01.BackColor = Color.LawnGreen
                btnSel2ProfitStatus01.Text = sel2.betfairCorrectScore01IfWinProfit
            Else
                btnSel2ProfitStatus01.BackColor = Color.White
                btnSel2ProfitStatus01.Text = ""
            End If
        Else
            btnSel2ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel2.betfairUnder15IfWinProfit) Then
            If CDbl(sel2.betfairUnder15IfWinProfit) > 0 Then
                btnSel2ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel2ProfitStatusUnder15.Text = sel2.betfairUnder15IfWinProfit
            Else
                btnSel2ProfitStatusUnder15.BackColor = Color.White
                btnSel2ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel2ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel2.betfairOver15IfWinProfit) Then
            If CDbl(sel2.betfairOver15IfWinProfit) > 0 Then
                btnSel2ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel2ProfitStatusOver15.Text = sel2.betfairOver15IfWinProfit
            Else
                btnSel2ProfitStatusOver15.BackColor = Color.White
                btnSel2ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel2ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel2.betfairEventInplay = "True" Then
            ' Continue
        Else
            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
            Exit Sub
        End If

        ' Populate Unmatched Order counts
        If String.IsNullOrEmpty(sel2.betfairOver15Orders) Then
            sel2.betfairOver15Orders = 0
        End If
        If String.IsNullOrEmpty(sel2.betfairUnder15Orders) Then
            sel2.betfairUnder15Orders = 0
        End If
        If String.IsNullOrEmpty(sel2.betfairCorrectScore00Orders) Then
            sel2.betfairCorrectScore00Orders = 0
        End If
        If String.IsNullOrEmpty(sel2.betfairCorrectScore10Orders) Then
            sel2.betfairCorrectScore10Orders = 0
        End If
        If String.IsNullOrEmpty(sel2.betfairCorrectScore01Orders) Then
            sel2.betfairCorrectScore01Orders = 0
        End If


        ' 
        ' Look to Identifying starting position of the bet ..........
        '
        If sel2.autobetOver15BetMade = False Then

            If btnSel2ProfitStatusOver15.Text = "" And tbxSel2Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Identifying starting position", EventLogEntryType.Information)

                If sel2.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +40 Then

                        If Not String.IsNullOrEmpty(sel2.betfairOver15BackOdds) Then
                            If CDbl(sel2.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel2.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel2.betfairOver15Orders) Then
                                    If CDbl(sel2.betfairOver15Orders) = 0 Then
                                        If sel2.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            ' Place back bet on Over1.5
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel2.betfairOver15BackOdds), "OUT")
                                            oddsMarket = CDbl(sel2.betfairCorrectScore00BackOdds)
                                            stake = nudSettingsOver15Stake.Value.ToString
                                            sendEmailToText("Match: " + sel2.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel2.autobetOver15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel2.placeOver15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position - Inplay timer >40 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If

        End If


        ' 
        ' Look to ensure we have 0 - 0 covered
        '
        If sel2.autobetCorrectScore00BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel2.autobetOver15BetMade = True And sel2.autobetUnder15BetMade = False And tbxSel2Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel2.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel2InplayTime.Text) > +25 And CDbl(tbxSel2InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel2.betfairCorrectScore00BackOdds) Then
                            If CDbl(sel2.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel2.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel2.betfairCorrectScore00Orders) Then
                                    If CDbl(sel2.betfairCorrectScore00Orders) = 0 Then
                                        If sel2.betfairCorrectScore00BackOdds <= nudSettingsCS00TargetPrice.Value Then

                                            ' calculate stake based on profit
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel2.betfairCorrectScore00BackOdds), "IN")
                                            oddsMarket = CDbl(sel2.betfairCorrectScore00BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel2.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel2.autobetCorrectScore00BetMade = True

                                            ' Place order on Correct Score 0-0 market
                                            sel2.placeCorrectScore00_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to ensure we have Under 1.5
        '
        If sel2.autobetUnder15BetMade = False Then

            ' Check the strategy has started and score only 1 goal
            If sel2.autobetOver15BetMade = True And tbxSel2Score.Text = "1 Goal scored" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                If sel2.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +45 Then
                        If Not String.IsNullOrEmpty(sel2.betfairUnder15BackOdds) Then
                            If CDbl(sel2.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel2.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel2.betfairUnder15Orders) Then
                                    If CDbl(sel2.betfairUnder15Orders) = 0 Then
                                        If sel2.betfairUnder15BackOdds >= nudSettingsUnder15TargetPrice.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsUnder15TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel2.betfairUnder15BackOdds), "IN")
                                            oddsMarket = CDbl(sel2.betfairUnder15BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel2.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel2.autobetUnder15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel2.placeUnder15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to cover 1 - 0, after 40 minutes play
        '
        If sel2.autobetCorrectScore10BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel2.autobetOver15BetMade = True And sel2.autobetUnder15BetMade = False And tbxSel2Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel2.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel2InplayTime.Text) > +40 And CDbl(tbxSel2InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel2.betfairCorrectScore10BackOdds) Then
                            If CDbl(sel2.betfairCorrectScore10BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel2.betfairCorrectScore10BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel2.betfairCorrectScore10Orders) Then
                                    If CDbl(sel2.betfairCorrectScore10Orders) = 0 Then
                                        If sel2.betfairCorrectScore10BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel2.betfairCorrectScore10BackOdds), "IN")
                                            oddsMarket = CDbl(sel2.betfairCorrectScore10BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel2.betfairEventName + " Market: Correct Score place back bet on 1 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel2.autobetCorrectScore10BetMade = True

                                            ' Place order on Correct Score 1-0 market
                                            sel2.placeCorrectScore10_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to cover 0 - 1, after 40 minutes play
        '
        If sel2.autobetCorrectScore01BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel2.autobetOver15BetMade = True And sel2.autobetUnder15BetMade = False And tbxSel2Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel2.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel2InplayTime.Text) > +40 And CDbl(tbxSel2InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel2.betfairCorrectScore01BackOdds) Then
                            If CDbl(sel2.betfairCorrectScore01BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel2.betfairCorrectScore01BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel2.betfairCorrectScore01Orders) Then
                                    If CDbl(sel2.betfairCorrectScore01Orders) = 0 Then
                                        If sel2.betfairCorrectScore01BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel2.betfairCorrectScore01BackOdds), "IN")
                                            oddsMarket = CDbl(sel2.betfairCorrectScore01BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel2.betfairEventName + " Market: Correct Score place back bet on 0 - 1 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel2.autobetCorrectScore01BetMade = True

                                            ' Place order on Correct Score 0-1 market
                                            sel2.placeCorrectScore01_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to boost the Over1.5 if we have taken 0-0 and either 1-0 or 0-1 after 40 minutes play
        '
        If sel2.autobetOver15TopUpBetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If (sel2.autobetOver15BetMade = True And sel2.autobetUnder15BetMade = False And sel2.autobetCorrectScore00BetMade = True And tbxSel2Score.Text = "0 - 0") And (sel2.autobetCorrectScore10BetMade = True Or sel2.autobetCorrectScore01BetMade = True) Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to boost Over 1.5", EventLogEntryType.Information)

                If sel2.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel2InplayTime.Text) > +40 And CDbl(tbxSel2InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel2.betfairOver15BackOdds) Then
                            If CDbl(sel2.betfairOver15BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel2.betfairOver15BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel2.betfairOver15Orders) Then
                                    If CDbl(sel2.betfairOver15Orders) = 0 Then

                                        ' Place 2nd back bet on Over1.5
                                        Dim odds As Double
                                        Dim oddsMarket As Double
                                        Dim stake As Double
                                        odds = adjustOddsToMatch(CDbl(sel2.betfairOver15BackOdds), "OUT")
                                        oddsMarket = CDbl(sel2.betfairCorrectScore01BackOdds)
                                        stake = (nudSettingsOver15Stake.Value / 4)
                                        sendEmailToText("Match: " + sel2.betfairEventName + " Market: Boost to Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                        sel2.autobetOver15TopUpBetMade = True

                                        ' Place order on Over 1.5 market
                                        sel2.placeOver15_Order(odds, stake)

                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Boost Over 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Boost Over 1.5 position  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Boost Over 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Boost Over 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Boost Over 1.5 position - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Boost Over 1.5 position  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If

    End Sub

    Private Sub btnSel3AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel3AutoBetOn.Click

        If btnSel3AutoBetOn.Text = "Autobet On" Then

            If tbxSel3EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel3.autobetOver15BetMade = False
                    sel3.autobetUnder15BetMade = False
                    sel3.autobetCorrectScore00BetMade = False
                    sel3.autobetCorrectScore10BetMade = False
                    sel3.autobetCorrectScore01BetMade = False
                    sel3.autobetOver15TopUpBetMade = False

                    ' Set the interval
                    timerSel3AutoBet.Interval = nudSettingsAutoBetRefresh.Value

                    ' Enable Autobet timer
                    timerSel3AutoBet.Enabled = True
                    btnSel3AutoBetOn.Text = "Autobet Off"
                    btnSel3AutoBetOn.BackColor = Color.LightSalmon

                    ' Disable the Select button
                    btnSel3.Enabled = False

                    ' Write to log
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 has been switched ON.", EventLogEntryType.Information)

                    ' Call tick
                    timerSel3AutoBet_Tick(sender, e)

                End If


            End If
        Else

            ' Disable Autobet timer
            timerSel3AutoBet.Enabled = False

            ' Switch off
            btnSel3AutoBetOn.Text = "Autobet On"
            btnSel3AutoBetOn.BackColor = Color.LightGreen

            ' Enable the Select button
            btnSel3.Enabled = True

            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 has been switched OFF.", EventLogEntryType.Information)

        End If

    End Sub

    Private Sub timerSel3AutoBet_Tick(sender As Object, e As EventArgs) Handles timerSel3AutoBet.Tick

        '
        ' Update status of each bet type
        '
        If Not String.IsNullOrEmpty(sel3.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel3.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel3ProfitStatus00.BackColor = Color.LawnGreen
                btnSel3ProfitStatus00.Text = sel3.betfairCorrectScore00IfWinProfit
            Else
                btnSel3ProfitStatus00.BackColor = Color.White
                btnSel3ProfitStatus00.Text = ""
            End If
        Else
            btnSel3ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel3.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel3.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel3ProfitStatus10.BackColor = Color.LawnGreen
                btnSel3ProfitStatus10.Text = sel3.betfairCorrectScore10IfWinProfit
            Else
                btnSel3ProfitStatus10.BackColor = Color.White
                btnSel3ProfitStatus10.Text = ""
            End If
        Else
            btnSel3ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel3.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel3.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel3ProfitStatus01.BackColor = Color.LawnGreen
                btnSel3ProfitStatus01.Text = sel3.betfairCorrectScore01IfWinProfit
            Else
                btnSel3ProfitStatus01.BackColor = Color.White
                btnSel3ProfitStatus01.Text = ""
            End If
        Else
            btnSel3ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel3.betfairUnder15IfWinProfit) Then
            If CDbl(sel3.betfairUnder15IfWinProfit) > 0 Then
                btnSel3ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel3ProfitStatusUnder15.Text = sel3.betfairUnder15IfWinProfit
            Else
                btnSel3ProfitStatusUnder15.BackColor = Color.White
                btnSel3ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel3ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel3.betfairOver15IfWinProfit) Then
            If CDbl(sel3.betfairOver15IfWinProfit) > 0 Then
                btnSel3ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel3ProfitStatusOver15.Text = sel3.betfairOver15IfWinProfit
            Else
                btnSel3ProfitStatusOver15.BackColor = Color.White
                btnSel3ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel3ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel3.betfairEventInplay = "True" Then
            ' Continue
        Else
            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
            Exit Sub
        End If

        ' Populate Unmatched Order counts
        If String.IsNullOrEmpty(sel3.betfairOver15Orders) Then
            sel3.betfairOver15Orders = 0
        End If
        If String.IsNullOrEmpty(sel3.betfairUnder15Orders) Then
            sel3.betfairUnder15Orders = 0
        End If
        If String.IsNullOrEmpty(sel3.betfairCorrectScore00Orders) Then
            sel3.betfairCorrectScore00Orders = 0
        End If
        If String.IsNullOrEmpty(sel3.betfairCorrectScore10Orders) Then
            sel3.betfairCorrectScore10Orders = 0
        End If
        If String.IsNullOrEmpty(sel3.betfairCorrectScore01Orders) Then
            sel3.betfairCorrectScore01Orders = 0
        End If


        ' 
        ' Look to Identifying starting position of the bet ..........
        '
        If sel3.autobetOver15BetMade = False Then

            If btnSel3ProfitStatusOver15.Text = "" And tbxSel3Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Identifying starting position", EventLogEntryType.Information)

                If sel3.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +40 Then

                        If Not String.IsNullOrEmpty(sel3.betfairOver15BackOdds) Then
                            If CDbl(sel3.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel3.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel3.betfairOver15Orders) Then
                                    If CDbl(sel3.betfairOver15Orders) = 0 Then
                                        If sel3.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            ' Place back bet on Over1.5
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel3.betfairOver15BackOdds), "OUT")
                                            oddsMarket = CDbl(sel3.betfairCorrectScore00BackOdds)
                                            stake = nudSettingsOver15Stake.Value.ToString
                                            sendEmailToText("Match: " + sel3.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel3.autobetOver15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel3.placeOver15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position - Inplay timer >40 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If

        End If


        ' 
        ' Look to ensure we have 0 - 0 covered
        '
        If sel3.autobetCorrectScore00BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel3.autobetOver15BetMade = True And sel3.autobetUnder15BetMade = False And tbxSel3Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel3.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel3InplayTime.Text) > +25 And CDbl(tbxSel3InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel3.betfairCorrectScore00BackOdds) Then
                            If CDbl(sel3.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel3.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel3.betfairCorrectScore00Orders) Then
                                    If CDbl(sel3.betfairCorrectScore00Orders) = 0 Then
                                        If sel3.betfairCorrectScore00BackOdds <= nudSettingsCS00TargetPrice.Value Then

                                            ' calculate stake based on profit
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel3.betfairCorrectScore00BackOdds), "IN")
                                            oddsMarket = CDbl(sel3.betfairCorrectScore00BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel3.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel3.autobetCorrectScore00BetMade = True

                                            ' Place order on Correct Score 0-0 market
                                            sel3.placeCorrectScore00_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to ensure we have Under 1.5
        '
        If sel3.autobetUnder15BetMade = False Then

            ' Check the strategy has started and score only 1 goal
            If sel3.autobetOver15BetMade = True And tbxSel3Score.Text = "1 Goal scored" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                If sel3.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +45 Then
                        If Not String.IsNullOrEmpty(sel3.betfairUnder15BackOdds) Then
                            If CDbl(sel3.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel3.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel3.betfairUnder15Orders) Then
                                    If CDbl(sel3.betfairUnder15Orders) = 0 Then
                                        If sel3.betfairUnder15BackOdds >= nudSettingsUnder15TargetPrice.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsUnder15TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel3.betfairUnder15BackOdds), "IN")
                                            oddsMarket = CDbl(sel3.betfairUnder15BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel3.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel3.autobetUnder15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel3.placeUnder15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to cover 1 - 0, after 40 minutes play
        '
        If sel3.autobetCorrectScore10BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel3.autobetOver15BetMade = True And sel3.autobetUnder15BetMade = False And tbxSel3Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel3.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel3InplayTime.Text) > +40 And CDbl(tbxSel3InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel3.betfairCorrectScore10BackOdds) Then
                            If CDbl(sel3.betfairCorrectScore10BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel3.betfairCorrectScore10BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel3.betfairCorrectScore10Orders) Then
                                    If CDbl(sel3.betfairCorrectScore10Orders) = 0 Then
                                        If sel3.betfairCorrectScore10BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel3.betfairCorrectScore10BackOdds), "IN")
                                            oddsMarket = CDbl(sel3.betfairCorrectScore10BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel3.betfairEventName + " Market: Correct Score place back bet on 1 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel3.autobetCorrectScore10BetMade = True

                                            ' Place order on Correct Score 1-0 market
                                            sel3.placeCorrectScore10_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to cover 0 - 1, after 40 minutes play
        '
        If sel3.autobetCorrectScore01BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel3.autobetOver15BetMade = True And sel3.autobetUnder15BetMade = False And tbxSel3Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel3.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel3InplayTime.Text) > +40 And CDbl(tbxSel3InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel3.betfairCorrectScore01BackOdds) Then
                            If CDbl(sel3.betfairCorrectScore01BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel3.betfairCorrectScore01BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel3.betfairCorrectScore01Orders) Then
                                    If CDbl(sel3.betfairCorrectScore01Orders) = 0 Then
                                        If sel3.betfairCorrectScore01BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel3.betfairCorrectScore01BackOdds), "IN")
                                            oddsMarket = CDbl(sel3.betfairCorrectScore01BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel3.betfairEventName + " Market: Correct Score place back bet on 0 - 1 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel3.autobetCorrectScore01BetMade = True

                                            ' Place order on Correct Score 0-1 market
                                            sel3.placeCorrectScore01_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to boost the Over1.5 if we have taken 0-0 and either 1-0 or 0-1 after 40 minutes play
        '
        If sel3.autobetOver15TopUpBetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If (sel3.autobetOver15BetMade = True And sel3.autobetUnder15BetMade = False And sel3.autobetCorrectScore00BetMade = True And tbxSel3Score.Text = "0 - 0") And (sel3.autobetCorrectScore10BetMade = True Or sel3.autobetCorrectScore01BetMade = True) Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to boost Over 1.5", EventLogEntryType.Information)

                If sel3.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel3InplayTime.Text) > +40 And CDbl(tbxSel3InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel3.betfairOver15BackOdds) Then
                            If CDbl(sel3.betfairOver15BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel3.betfairOver15BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel3.betfairOver15Orders) Then
                                    If CDbl(sel3.betfairOver15Orders) = 0 Then

                                        ' Place 2nd back bet on Over1.5
                                        Dim odds As Double
                                        Dim oddsMarket As Double
                                        Dim stake As Double
                                        odds = adjustOddsToMatch(CDbl(sel3.betfairOver15BackOdds), "OUT")
                                        oddsMarket = CDbl(sel3.betfairCorrectScore01BackOdds)
                                        stake = (nudSettingsOver15Stake.Value / 4)
                                        sendEmailToText("Match: " + sel3.betfairEventName + " Market: Boost to Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                        sel3.autobetOver15TopUpBetMade = True

                                        ' Place order on Over 1.5 market
                                        sel3.placeOver15_Order(odds, stake)

                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Boost Over 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Boost Over 1.5 position  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Boost Over 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Boost Over 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Boost Over 1.5 position - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Boost Over 1.5 position  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If

    End Sub

    Private Sub btnSel4AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel4AutoBetOn.Click

        If btnSel4AutoBetOn.Text = "Autobet On" Then

            If tbxSel4EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel4.autobetOver15BetMade = False
                    sel4.autobetUnder15BetMade = False
                    sel4.autobetCorrectScore00BetMade = False
                    sel4.autobetCorrectScore10BetMade = False
                    sel4.autobetCorrectScore01BetMade = False
                    sel4.autobetOver15TopUpBetMade = False

                    ' Set the interval
                    timerSel4AutoBet.Interval = nudSettingsAutoBetRefresh.Value

                    ' Enable Autobet timer
                    timerSel4AutoBet.Enabled = True
                    btnSel4AutoBetOn.Text = "Autobet Off"
                    btnSel4AutoBetOn.BackColor = Color.LightSalmon

                    ' Disable the Select button
                    btnSel4.Enabled = False

                    ' Write to log
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 has been switched ON.", EventLogEntryType.Information)

                    ' Call tick
                    timerSel4AutoBet_Tick(sender, e)

                End If


            End If
        Else

            ' Disable Autobet timer
            timerSel4AutoBet.Enabled = False

            ' Switch off
            btnSel4AutoBetOn.Text = "Autobet On"
            btnSel4AutoBetOn.BackColor = Color.LightGreen

            ' Enable the Select button
            btnSel4.Enabled = True

            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 has been switched OFF.", EventLogEntryType.Information)

        End If

    End Sub

    Private Sub timerSel4AutoBet_Tick(sender As Object, e As EventArgs) Handles timerSel4AutoBet.Tick

        '
        ' Update status of each bet type
        '
        If Not String.IsNullOrEmpty(sel4.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel4.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel4ProfitStatus00.BackColor = Color.LawnGreen
                btnSel4ProfitStatus00.Text = sel4.betfairCorrectScore00IfWinProfit
            Else
                btnSel4ProfitStatus00.BackColor = Color.White
                btnSel4ProfitStatus00.Text = ""
            End If
        Else
            btnSel4ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel4.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel4.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel4ProfitStatus10.BackColor = Color.LawnGreen
                btnSel4ProfitStatus10.Text = sel4.betfairCorrectScore10IfWinProfit
            Else
                btnSel4ProfitStatus10.BackColor = Color.White
                btnSel4ProfitStatus10.Text = ""
            End If
        Else
            btnSel4ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel4.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel4.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel4ProfitStatus01.BackColor = Color.LawnGreen
                btnSel4ProfitStatus01.Text = sel4.betfairCorrectScore01IfWinProfit
            Else
                btnSel4ProfitStatus01.BackColor = Color.White
                btnSel4ProfitStatus01.Text = ""
            End If
        Else
            btnSel4ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel4.betfairUnder15IfWinProfit) Then
            If CDbl(sel4.betfairUnder15IfWinProfit) > 0 Then
                btnSel4ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel4ProfitStatusUnder15.Text = sel4.betfairUnder15IfWinProfit
            Else
                btnSel4ProfitStatusUnder15.BackColor = Color.White
                btnSel4ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel4ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel4.betfairOver15IfWinProfit) Then
            If CDbl(sel4.betfairOver15IfWinProfit) > 0 Then
                btnSel4ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel4ProfitStatusOver15.Text = sel4.betfairOver15IfWinProfit
            Else
                btnSel4ProfitStatusOver15.BackColor = Color.White
                btnSel4ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel4ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel4.betfairEventInplay = "True" Then
            ' Continue
        Else
            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
            Exit Sub
        End If

        ' Populate Unmatched Order counts
        If String.IsNullOrEmpty(sel4.betfairOver15Orders) Then
            sel4.betfairOver15Orders = 0
        End If
        If String.IsNullOrEmpty(sel4.betfairUnder15Orders) Then
            sel4.betfairUnder15Orders = 0
        End If
        If String.IsNullOrEmpty(sel4.betfairCorrectScore00Orders) Then
            sel4.betfairCorrectScore00Orders = 0
        End If
        If String.IsNullOrEmpty(sel4.betfairCorrectScore10Orders) Then
            sel4.betfairCorrectScore10Orders = 0
        End If
        If String.IsNullOrEmpty(sel4.betfairCorrectScore01Orders) Then
            sel4.betfairCorrectScore01Orders = 0
        End If


        ' 
        ' Look to Identifying starting position of the bet ..........
        '
        If sel4.autobetOver15BetMade = False Then

            If btnSel4ProfitStatusOver15.Text = "" And tbxSel4Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Identifying starting position", EventLogEntryType.Information)

                If sel4.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +40 Then

                        If Not String.IsNullOrEmpty(sel4.betfairOver15BackOdds) Then
                            If CDbl(sel4.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel4.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel4.betfairOver15Orders) Then
                                    If CDbl(sel4.betfairOver15Orders) = 0 Then
                                        If sel4.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            ' Place back bet on Over1.5
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel4.betfairOver15BackOdds), "OUT")
                                            oddsMarket = CDbl(sel4.betfairCorrectScore00BackOdds)
                                            stake = nudSettingsOver15Stake.Value.ToString
                                            sendEmailToText("Match: " + sel4.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel4.autobetOver15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel4.placeOver15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position - Inplay timer >40 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If

        End If


        ' 
        ' Look to ensure we have 0 - 0 covered
        '
        If sel4.autobetCorrectScore00BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel4.autobetOver15BetMade = True And sel4.autobetUnder15BetMade = False And tbxSel4Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel4.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel4InplayTime.Text) > +25 And CDbl(tbxSel4InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel4.betfairCorrectScore00BackOdds) Then
                            If CDbl(sel4.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel4.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel4.betfairCorrectScore00Orders) Then
                                    If CDbl(sel4.betfairCorrectScore00Orders) = 0 Then
                                        If sel4.betfairCorrectScore00BackOdds <= nudSettingsCS00TargetPrice.Value Then

                                            ' calculate stake based on profit
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel4.betfairCorrectScore00BackOdds), "IN")
                                            oddsMarket = CDbl(sel4.betfairCorrectScore00BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel4.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel4.autobetCorrectScore00BetMade = True

                                            ' Place order on Correct Score 0-0 market
                                            sel4.placeCorrectScore00_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to ensure we have Under 1.5
        '
        If sel4.autobetUnder15BetMade = False Then

            ' Check the strategy has started and score only 1 goal
            If sel4.autobetOver15BetMade = True And tbxSel4Score.Text = "1 Goal scored" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                If sel4.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +45 Then
                        If Not String.IsNullOrEmpty(sel4.betfairUnder15BackOdds) Then
                            If CDbl(sel4.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel4.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel4.betfairUnder15Orders) Then
                                    If CDbl(sel4.betfairUnder15Orders) = 0 Then
                                        If sel4.betfairUnder15BackOdds >= nudSettingsUnder15TargetPrice.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsUnder15TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel4.betfairUnder15BackOdds), "IN")
                                            oddsMarket = CDbl(sel4.betfairUnder15BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel4.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel4.autobetUnder15BetMade = True

                                            ' Place order on Over 1.5 market
                                            sel4.placeUnder15_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If



        ' 
        ' Look to cover 1 - 0, after 40 minutes play
        '
        If sel4.autobetCorrectScore10BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel4.autobetOver15BetMade = True And sel4.autobetUnder15BetMade = False And tbxSel4Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel4.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel4InplayTime.Text) > +40 And CDbl(tbxSel4InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel4.betfairCorrectScore10BackOdds) Then
                            If CDbl(sel4.betfairCorrectScore10BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel4.betfairCorrectScore10BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel4.betfairCorrectScore10Orders) Then
                                    If CDbl(sel4.betfairCorrectScore10Orders) = 0 Then
                                        If sel4.betfairCorrectScore10BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel4.betfairCorrectScore10BackOdds), "IN")
                                            oddsMarket = CDbl(sel4.betfairCorrectScore10BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel4.betfairEventName + " Market: Correct Score place back bet on 1 - 0 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel4.autobetCorrectScore10BetMade = True

                                            ' Place order on Correct Score 1-0 market
                                            sel4.placeCorrectScore10_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to cover 0 - 1, after 40 minutes play
        '
        If sel4.autobetCorrectScore01BetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If sel4.autobetOver15BetMade = True And sel4.autobetUnder15BetMade = False And tbxSel4Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel4.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel4InplayTime.Text) > +40 And CDbl(tbxSel4InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel4.betfairCorrectScore01BackOdds) Then
                            If CDbl(sel4.betfairCorrectScore01BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel4.betfairCorrectScore01BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel4.betfairCorrectScore01Orders) Then
                                    If CDbl(sel4.betfairCorrectScore01Orders) = 0 Then
                                        If sel4.betfairCorrectScore01BackOdds <= nudSettingsCS10and01TargetGross.Value Then

                                            ' calculate profit
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            odds = adjustOddsToMatch(CDbl(sel4.betfairCorrectScore01BackOdds), "IN")
                                            oddsMarket = CDbl(sel4.betfairCorrectScore01BackOdds)
                                            stake = grossPerMarket / (oddsMarket - 1)
                                            sendEmailToText("Match: " + sel4.betfairEventName + " Market: Correct Score place back bet on 0 - 1 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                            sel4.autobetCorrectScore01BetMade = True

                                            ' Place order on Correct Score 0-1 market
                                            sel4.placeCorrectScore01_Order(odds, stake)

                                        End If
                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1  - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' 
        ' Look to boost the Over1.5 if we have taken 0-0 and either 1-0 or 0-1 after 40 minutes play
        '
        If sel4.autobetOver15TopUpBetMade = False Then

            ' Check the strategy has started, no Under 1.5 bet made and score still 0 - 0
            If (sel4.autobetOver15BetMade = True And sel4.autobetUnder15BetMade = False And sel4.autobetCorrectScore00BetMade = True And tbxSel4Score.Text = "0 - 0") And (sel4.autobetCorrectScore10BetMade = True Or sel4.autobetCorrectScore01BetMade = True) Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to boost Over 1.5", EventLogEntryType.Information)

                If sel4.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check time is after 40 minutes

                    If CDbl(tbxSel4InplayTime.Text) > +40 And CDbl(tbxSel4InplayTime.Text) < +60 Then
                        If Not String.IsNullOrEmpty(sel4.betfairOver15BackOdds) Then
                            If CDbl(sel4.betfairOver15BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel4.betfairOver15BackOdds) < nudSettingsCS00UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel4.betfairOver15Orders) Then
                                    If CDbl(sel4.betfairOver15Orders) = 0 Then

                                        ' Place 2nd back bet on Over1.5
                                        Dim odds As Double
                                        Dim oddsMarket As Double
                                        Dim stake As Double
                                        odds = adjustOddsToMatch(CDbl(sel4.betfairOver15BackOdds), "OUT")
                                        oddsMarket = CDbl(sel4.betfairCorrectScore01BackOdds)
                                        stake = (nudSettingsOver15Stake.Value / 4)
                                        sendEmailToText("Match: " + sel4.betfairEventName + " Market: Boost to Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: £" + FormatNumber(CDbl(stake), 2).ToString)
                                        sel4.autobetOver15TopUpBetMade = True

                                        ' Place order on Over 1.5 market
                                        sel4.placeOver15_Order(odds, stake)

                                    Else
                                        ' Unmatched orders
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Boost Over 1.5 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                                    End If
                                Else
                                    ' Unmatched orders are either NULL or EMPTY
                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Boost Over 1.5 position  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                                End If
                            Else
                                ' Odds are either Odds not within limits
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Boost Over 1.5 position - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                            End If
                        Else
                            ' Odds are either NULL or EMPTY
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Boost Over 1.5 position - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Not first half of match
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Boost Over 1.5 position - Inplay timer >60 mins, no further action taken", EventLogEntryType.Information)
                    End If
                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Boost Over 1.5 position  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If

    End Sub


    '
    ' Selection Button
    '
    Private Sub btnsel1_Click(sender As Object, e As EventArgs) Handles btnSel1.Click

        Dim selectedRowCount As Integer =
        dgvEvents.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount > 0 Then

            'Initialize some fields
            tbxSel1InplayStatus.Text = ""
            tbxSel1Score.Text = ""
            tbxSel1CorrectScore00Orders.Text = ""
            tbxSel1CorrectScore10Orders.Text = ""
            tbxSel1CorrectScore01Orders.Text = ""
            tbxSel1RefreshLight.Text = ""
            tbxSel1InplayTime.Text = ""
            tbxSel1EventName.Text = ""
            tbxSel1EventDateTime.Text = ""
            tbxSel1Goal1.Text = ""
            tbxSel1Goal2.Text = ""

            tbxSel1CorrectScoreStatus.Text = ""
            tbxSel1CorrectScore00IfWin.Text = ""
            tbxSel1CorrectScore00Odds.Text = ""
            tbxSel1CorrectScore00Status.Text = ""
            tbxSel1CorrectScore00Orders.Text = ""

            tbxSel1CorrectScore01IfWin.Text = ""
            tbxSel1CorrectScore01Odds.Text = ""
            tbxSel1CorrectScore01Status.Text = ""
            tbxSel1CorrectScore01Orders.Text = ""

            tbxSel1CorrectScore10IfWin.Text = ""
            tbxSel1CorrectScore10Odds.Text = ""
            tbxSel1CorrectScore10Status.Text = ""
            tbxSel1CorrectScore10Orders.Text = ""

            tbxSel1UnderOver15MarketStatus.Text = ""

            tbxSel1Under15Odds.Text = ""
            tbxSel1IUnder15fWinProfit.Text = ""
            tbxSel1IUnder15Status.Text = ""
            tbxSel1IUnder15Orders.Text = ""

            tbxSel1Over15Odds.Text = ""
            tbxSel1IOver15fWinProfit.Text = ""
            tbxSel1IOver15Status.Text = ""
            tbxSel1IOver15Orders.Text = ""

            ' Reset colored buttons
            tbxSel1RefreshLight.BackColor = Color.White
            tbxSel1InplayStatus.BackColor = Color.White
            tbxSel1CorrectScore00Status.BackColor = Color.White
            tbxSel1CorrectScore10Status.BackColor = Color.White
            tbxSel1CorrectScore01Status.BackColor = Color.White
            tbxSel1IUnder15Status.BackColor = Color.White
            tbxSel1IOver15Status.BackColor = Color.White

            ' Refresh screen
            Application.DoEvents()

            ' Copy data from dgv
            tbxSel1EventName.Text = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            grpSel1.Text = "Selection 1 - " + dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel1.betfairEventName = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel1.betfairEventDateTime = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            tbxSel1EventDateTime.Text = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            sel1.betfairEventId = dgvEvents.SelectedRows(0).Cells(1).Value.ToString()

            ' Refresh 
            Refreshsel1Info()

            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True

            ' Enable Autobet Button
            If tbxSel1Score.Text = "0 - 0" Or tbxSel1Score.Text = "1 Goal scored" Then
                btnSel1AutoBetOn.Enabled = True
            End If

        Else

            grpSel1.Text = "Selection 1"
            tbxSel1EventName.Text = ""
            btnSel1AutoBetOn.Enabled = False

        End If


    End Sub

    Private Sub btnSel2_Click(sender As Object, e As EventArgs) Handles btnSel2.Click

        Dim selectedRowCount As Integer =
        dgvEvents.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount > 0 Then

            'Initialize some fields
            tbxSel2InplayStatus.Text = ""
            tbxSel2Score.Text = ""
            tbxSel2CorrectScore00Orders.Text = ""
            tbxSel2CorrectScore10Orders.Text = ""
            tbxSel2CorrectScore01Orders.Text = ""
            tbxSel2RefreshLight.Text = ""
            tbxSel2InplayTime.Text = ""
            tbxSel2EventName.Text = ""
            tbxSel2EventDateTime.Text = ""
            tbxSel2Goal1.Text = ""
            tbxSel2Goal2.Text = ""

            tbxSel2CorrectScoreStatus.Text = ""
            tbxSel2CorrectScore00IfWin.Text = ""
            tbxSel2CorrectScore00Odds.Text = ""
            tbxSel2CorrectScore00Status.Text = ""
            tbxSel2CorrectScore00Orders.Text = ""

            tbxSel2CorrectScore01IfWin.Text = ""
            tbxSel2CorrectScore01Odds.Text = ""
            tbxSel2CorrectScore01Status.Text = ""
            tbxSel2CorrectScore01Orders.Text = ""

            tbxSel2CorrectScore10IfWin.Text = ""
            tbxSel2CorrectScore10Odds.Text = ""
            tbxSel2CorrectScore10Status.Text = ""
            tbxSel2CorrectScore10Orders.Text = ""

            tbxSel2UnderOver15MarketStatus.Text = ""

            tbxSel2Under15Odds.Text = ""
            tbxSel2IUnder15fWinProfit.Text = ""
            tbxSel2IUnder15Status.Text = ""
            tbxSel2IUnder15Orders.Text = ""

            tbxSel2Over15Odds.Text = ""
            tbxSel2IOver15fWinProfit.Text = ""
            tbxSel2IOver15Status.Text = ""
            tbxSel2IOver15Orders.Text = ""

            ' Reset colored buttons
            tbxSel2RefreshLight.BackColor = Color.White
            tbxSel2InplayStatus.BackColor = Color.White
            tbxSel2CorrectScore00Status.BackColor = Color.White
            tbxSel2CorrectScore10Status.BackColor = Color.White
            tbxSel2CorrectScore01Status.BackColor = Color.White
            tbxSel2IUnder15Status.BackColor = Color.White
            tbxSel2IOver15Status.BackColor = Color.White

            ' Refresh screen
            Application.DoEvents()

            ' Copy data from dgv
            tbxSel2EventName.Text = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            grpSel2.Text = "Selection 2 - " + dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel2.betfairEventName = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel2.betfairEventDateTime = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            tbxSel2EventDateTime.Text = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            sel2.betfairEventId = dgvEvents.SelectedRows(0).Cells(1).Value.ToString()


            ' Refresh 
            Refreshsel2Info()

            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True

            ' Enable Autobet Button
            If tbxSel2Score.Text = "0 - 0" Or tbxSel2Score.Text = "1 Goal scored" Then
                btnSel2AutoBetOn.Enabled = True
            End If

        Else

            grpSel2.Text = "Selection 2"
            tbxSel2EventName.Text = ""
            btnSel2AutoBetOn.Enabled = False

        End If
    End Sub

    Private Sub btnSel3_Click(sender As Object, e As EventArgs) Handles btnSel3.Click

        Dim selectedRowCount As Integer = dgvEvents.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount > 0 Then

            'Initialize some fields
            tbxSel3InplayStatus.Text = ""
            tbxSel3Score.Text = ""
            tbxSel3CorrectScore00Orders.Text = ""
            tbxSel3CorrectScore10Orders.Text = ""
            tbxSel3CorrectScore01Orders.Text = ""
            tbxSel3RefreshLight.Text = ""
            tbxSel3InplayTime.Text = ""
            tbxSel3EventName.Text = ""
            tbxSel3EventDateTime.Text = ""
            tbxSel3Goal1.Text = ""
            tbxSel3Goal2.Text = ""

            tbxSel3CorrectScoreStatus.Text = ""
            tbxSel3CorrectScore00IfWin.Text = ""
            tbxSel3CorrectScore00Odds.Text = ""
            tbxSel3CorrectScore00Status.Text = ""
            tbxSel3CorrectScore00Orders.Text = ""

            tbxSel3CorrectScore01IfWin.Text = ""
            tbxSel3CorrectScore01Odds.Text = ""
            tbxSel3CorrectScore01Status.Text = ""
            tbxSel3CorrectScore01Orders.Text = ""

            tbxSel3CorrectScore10IfWin.Text = ""
            tbxSel3CorrectScore10Odds.Text = ""
            tbxSel3CorrectScore10Status.Text = ""
            tbxSel3CorrectScore10Orders.Text = ""

            tbxSel3UnderOver15MarketStatus.Text = ""

            tbxSel3Under15Odds.Text = ""
            tbxSel3IUnder15fWinProfit.Text = ""
            tbxSel3IUnder15Status.Text = ""
            tbxSel3IUnder15Orders.Text = ""

            tbxSel3Over15Odds.Text = ""
            tbxSel3IOver15fWinProfit.Text = ""
            tbxSel3IOver15Status.Text = ""
            tbxSel3IOver15Orders.Text = ""

            ' Reset colored buttons
            tbxSel3RefreshLight.BackColor = Color.White
            tbxSel3InplayStatus.BackColor = Color.White
            tbxSel3CorrectScore00Status.BackColor = Color.White
            tbxSel3CorrectScore10Status.BackColor = Color.White
            tbxSel3CorrectScore01Status.BackColor = Color.White
            tbxSel3IUnder15Status.BackColor = Color.White
            tbxSel3IOver15Status.BackColor = Color.White

            ' Refresh screen
            Application.DoEvents()

            ' Copy data from dgv
            tbxSel3EventName.Text = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            grpSel3.Text = "Selection 3 - " + dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel3.betfairEventName = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel3.betfairEventDateTime = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            tbxSel3EventDateTime.Text = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            sel3.betfairEventId = dgvEvents.SelectedRows(0).Cells(1).Value.ToString()


            ' Refresh 
            Refreshsel3Info()

            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True

            ' Enable Autobet Button
            If tbxSel3Score.Text = "0 - 0" Or tbxSel3Score.Text = "1 Goal scored" Then
                btnSel3AutoBetOn.Enabled = True
            End If

        Else

            grpSel3.Text = "Selection 3"
            tbxSel3EventName.Text = ""
            btnSel3AutoBetOn.Enabled = False

        End If
    End Sub

    Private Sub btnSel4_Click(sender As Object, e As EventArgs) Handles btnSel4.Click

        Dim selectedRowCount As Integer = dgvEvents.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount > 0 Then

            'Initialize some fields
            tbxSel4InplayStatus.Text = ""
            tbxSel4Score.Text = ""
            tbxSel4CorrectScore00Orders.Text = ""
            tbxSel4CorrectScore10Orders.Text = ""
            tbxSel4CorrectScore01Orders.Text = ""
            tbxSel4RefreshLight.Text = ""
            tbxSel4InplayTime.Text = ""
            tbxSel4EventName.Text = ""
            tbxSel4EventDateTime.Text = ""
            tbxSel4Goal1.Text = ""
            tbxSel4Goal2.Text = ""

            tbxSel4CorrectScoreStatus.Text = ""
            tbxSel4CorrectScore00IfWin.Text = ""
            tbxSel4CorrectScore00Odds.Text = ""
            tbxSel4CorrectScore00Status.Text = ""
            tbxSel4CorrectScore00Orders.Text = ""

            tbxSel4CorrectScore01IfWin.Text = ""
            tbxSel4CorrectScore01Odds.Text = ""
            tbxSel4CorrectScore01Status.Text = ""
            tbxSel4CorrectScore01Orders.Text = ""

            tbxSel4CorrectScore10IfWin.Text = ""
            tbxSel4CorrectScore10Odds.Text = ""
            tbxSel4CorrectScore10Status.Text = ""
            tbxSel4CorrectScore10Orders.Text = ""

            tbxSel4UnderOver15MarketStatus.Text = ""

            tbxSel4Under15Odds.Text = ""
            tbxSel4IUnder15fWinProfit.Text = ""
            tbxSel4IUnder15Status.Text = ""
            tbxSel4IUnder15Orders.Text = ""

            tbxSel4Over15Odds.Text = ""
            tbxSel4IOver15fWinProfit.Text = ""
            tbxSel4IOver15Status.Text = ""
            tbxSel4IOver15Orders.Text = ""

            ' Reset colored buttons
            tbxSel4RefreshLight.BackColor = Color.White
            tbxSel4InplayStatus.BackColor = Color.White
            tbxSel4CorrectScore00Status.BackColor = Color.White
            tbxSel4CorrectScore10Status.BackColor = Color.White
            tbxSel4CorrectScore01Status.BackColor = Color.White
            tbxSel4IUnder15Status.BackColor = Color.White
            tbxSel4IOver15Status.BackColor = Color.White

            ' Refresh screen
            Application.DoEvents()

            ' Copy data from dgv
            tbxSel4EventName.Text = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            grpSel4.Text = "Selection 4 - " + dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel4.betfairEventName = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel4.betfairEventDateTime = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            tbxSel4EventDateTime.Text = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            sel4.betfairEventId = dgvEvents.SelectedRows(0).Cells(1).Value.ToString()

            ' Refresh 
            Refreshsel4Info()

            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True

            ' Enable Autobet Button
            If tbxSel4Score.Text = "0 - 0" Or tbxSel4Score.Text = "1 Goal scored" Then
                btnSel4AutoBetOn.Enabled = True
            End If

        Else

            grpSel4.Text = "Selection 4"
            tbxSel4EventName.Text = ""
            btnSel4AutoBetOn.Enabled = False

        End If

    End Sub

    Private Sub Refreshsel1Info()

        ' Get Initial book details, like marketId's and selectionId's
        sel1.getInitialBookDetails()

        ' Get latest data from Betfair
        sel1.getLatestMarketData()

        ' Update Inplay status
        If sel1.betfairEventInplay = "False" Then
            tbxSel1InplayStatus.BackColor = Color.OrangeRed
        Else
            tbxSel1InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel1Score.Text

        ' Get latest score
        tbxSel1Score.Text = sel1.betfairGoalsScored


        ' Update form
        Application.DoEvents()


        ' Detect score change
        If strPreviousScore = tbxSel1Score.Text Then
            ' Same score
        Else

            ' If first time through...ignore
            If strPreviousScore <> "" Then

                ' If match ended...ignore
                If tbxSel1Score.Text <> "Match ended!" Then

                    ' 1st Goal scored since last tick
                    If tbxSel1Score.Text = "1 Goal scored" Then
                        tbxSel1Goal1.Text = tbxSel1InplayTime.Text.ToString
                        sel1.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel1.Text + ", Goal 1 scored at: " + tbxSel1InplayTime.Text.ToString, EventLogEntryType.Information)

                        ' Send text
                        sendEmailToText("Goal 1 scored in match: " + sel1.betfairEventName + " at Inplay timer time: " + tbxSel1InplayTime.Text.ToString)

                    Else
                        If tbxSel1Score.Text = "2 Goals scored" Then
                            tbxSel1Goal2.Text = tbxSel1InplayTime.Text.ToString
                            sel1.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel1.Text + ", Goal 2 scored at: " + tbxSel1InplayTime.Text.ToString, EventLogEntryType.Information)

                            ' Send text
                            sendEmailToText("Goal 2 scored in match: " + sel1.betfairEventName + " at Inplay timer time: " + tbxSel1InplayTime.Text.ToString)

                        End If
                    End If
                End If
            End If
        End If


        ' Market Status
        tbxSel1CorrectScoreStatus.Text = sel1.betfairCorrectScoreMarketStatus
        tbxSel1UnderOver15MarketStatus.Text = sel1.betfairUnderOver15MarketStatus

        tbxSel1Over15Odds.Text = sel1.betfairOver15BackOdds
        tbxSel1Under15Odds.Text = sel1.betfairUnder15BackOdds
        tbxSel1CorrectScore00Odds.Text = sel1.betfairCorrectScore00BackOdds
        tbxSel1CorrectScore10Odds.Text = sel1.betfairCorrectScore10BackOdds
        tbxSel1CorrectScore01Odds.Text = sel1.betfairCorrectScore01BackOdds

        If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
            If Double.Parse(sel1.betfairOver15IfWinProfit) >= 0 Then
                tbxSel1IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel1IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
            If Double.Parse(sel1.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel1IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel1IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
            If Double.Parse(sel1.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel1CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
            If Double.Parse(sel1.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel1CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
            If Double.Parse(sel1.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel1CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel1IOver15fWinProfit.Text = sel1.betfairOver15IfWinProfit
        tbxSel1IUnder15fWinProfit.Text = sel1.betfairUnder15IfWinProfit
        tbxSel1CorrectScore00IfWin.Text = sel1.betfairCorrectScore00IfWinProfit
        tbxSel1CorrectScore10IfWin.Text = sel1.betfairCorrectScore10IfWinProfit
        tbxSel1CorrectScore01IfWin.Text = sel1.betfairCorrectScore01IfWinProfit

        If sel1.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel1IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel1IUnder15Status.BackColor = Color.OrangeRed
        End If
        If sel1.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel1IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel1IOver15Status.BackColor = Color.OrangeRed
        End If
        If sel1.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel1CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel1CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If sel1.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel1CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel1CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If sel1.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel1CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel1CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel1IUnder15Status.Text = sel1.betfairUnder15SelectionStatus
        tbxSel1IOver15Status.Text = sel1.betfairOver15SelectionStatus
        tbxSel1CorrectScore00Status.Text = sel1.betfairCorrectScore00SelectionStatus
        tbxSel1CorrectScore10Status.Text = sel1.betfairCorrectScore10SelectionStatus
        tbxSel1CorrectScore01Status.Text = sel1.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel1IOver15Orders.Text = sel1.betfairOver15Orders
        tbxSel1IUnder15Orders.Text = sel1.betfairUnder15Orders

        tbxSel1CorrectScore00Orders.Text = sel1.betfairCorrectScore00Orders
        tbxSel1CorrectScore10Orders.Text = sel1.betfairCorrectScore10Orders
        tbxSel1CorrectScore01Orders.Text = sel1.betfairCorrectScore01Orders

        ' Update refresh date/time
        tbxSel1RefreshLight.BackColor = Color.DarkGreen
        tbxSel1RefreshLight.ForeColor = Color.White
        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        tbxSel1RefreshLight.Text = time.ToString(format)

        ' Update the Inplay datetime
        Dim eventDateTime As DateTime = DateTime.Parse(tbxSel1EventDateTime.Text)
        Dim timeToStart As TimeSpan = DateTime.Now.Subtract(eventDateTime)
        Dim formatTime As String = "####0.00"
        tbxSel1InplayTime.Text = timeToStart.TotalMinutes.ToString(formatTime)

    End Sub

    Private Sub Refreshsel2Info()

        ' Get Initial book details, like marketId's and selectionId's
        sel2.getInitialBookDetails()

        ' Get latest data from Betfair
        sel2.getLatestMarketData()

        ' Update Inplay status
        If sel2.betfairEventInplay = "False" Then
            tbxSel2InplayStatus.BackColor = Color.OrangeRed
        Else
            tbxSel2InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel2Score.Text

        ' Get latest score
        tbxSel2Score.Text = sel2.betfairGoalsScored

        ' Update form
        Application.DoEvents()


        ' Detect score change
        If strPreviousScore = tbxSel2Score.Text Then
            ' Same score
        Else

            ' If first time through...ignore
            If strPreviousScore <> "" Then

                ' If match ended...ignore
                If tbxSel2Score.Text <> "Match ended!" Then

                    ' 1st Goal scored since last tick
                    If tbxSel2Score.Text = "1 Goal scored" Then
                        tbxSel2Goal1.Text = tbxSel2InplayTime.Text.ToString
                        sel2.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 1 scored at: " + tbxSel2InplayTime.Text.ToString, EventLogEntryType.Information)

                        ' Send text
                        sendEmailToText("Goal 1 scored in match: " + sel2.betfairEventName + " at Inplay timer time: " + tbxSel2InplayTime.Text.ToString)

                    Else
                        If tbxSel2Score.Text = "2 Goals scored" Then
                            tbxSel2Goal2.Text = tbxSel2InplayTime.Text.ToString
                            sel2.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 2 scored at: " + tbxSel2InplayTime.Text.ToString, EventLogEntryType.Information)

                            ' Send text
                            sendEmailToText("Goal 2 scored in match: " + sel2.betfairEventName + " at Inplay timer time: " + tbxSel2InplayTime.Text.ToString)

                        End If
                    End If
                End If
            End If
        End If


        ' Market Status
        tbxSel2CorrectScoreStatus.Text = sel2.betfairCorrectScoreMarketStatus
        tbxSel2UnderOver15MarketStatus.Text = sel2.betfairUnderOver15MarketStatus

        tbxSel2Over15Odds.Text = sel2.betfairOver15BackOdds
        tbxSel2Under15Odds.Text = sel2.betfairUnder15BackOdds
        tbxSel2CorrectScore00Odds.Text = sel2.betfairCorrectScore00BackOdds
        tbxSel2CorrectScore10Odds.Text = sel2.betfairCorrectScore10BackOdds
        tbxSel2CorrectScore01Odds.Text = sel2.betfairCorrectScore01BackOdds

        If sel2.betfairOver15IfWinProfit IsNot Nothing Then
            If Double.Parse(sel2.betfairOver15IfWinProfit) >= 0 Then
                tbxSel2IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel2IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel2.betfairUnder15IfWinProfit IsNot Nothing Then
            If Double.Parse(sel2.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel2IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel2IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel2.betfairCorrectScore00IfWinProfit IsNot Nothing Then
            If Double.Parse(sel2.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel2CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel2CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel2.betfairCorrectScore10IfWinProfit IsNot Nothing Then
            If Double.Parse(sel2.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel2CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel2CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel2.betfairCorrectScore01IfWinProfit IsNot Nothing Then
            If Double.Parse(sel2.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel2CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel2CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel2IOver15fWinProfit.Text = sel2.betfairOver15IfWinProfit
        tbxSel2IUnder15fWinProfit.Text = sel2.betfairUnder15IfWinProfit
        tbxSel2CorrectScore00IfWin.Text = sel2.betfairCorrectScore00IfWinProfit
        tbxSel2CorrectScore10IfWin.Text = sel2.betfairCorrectScore10IfWinProfit
        tbxSel2CorrectScore01IfWin.Text = sel2.betfairCorrectScore01IfWinProfit

        If sel2.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel2IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel2IUnder15Status.BackColor = Color.OrangeRed
        End If
        If sel2.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel2IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel2IOver15Status.BackColor = Color.OrangeRed
        End If
        If sel2.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel2CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel2CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If sel2.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel2CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel2CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If sel2.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel2CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel2CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel2IUnder15Status.Text = sel2.betfairUnder15SelectionStatus
        tbxSel2IOver15Status.Text = sel2.betfairOver15SelectionStatus
        tbxSel2CorrectScore00Status.Text = sel2.betfairCorrectScore00SelectionStatus
        tbxSel2CorrectScore10Status.Text = sel2.betfairCorrectScore10SelectionStatus
        tbxSel2CorrectScore01Status.Text = sel2.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel2IOver15Orders.Text = sel2.betfairOver15Orders
        tbxSel2IUnder15Orders.Text = sel2.betfairUnder15Orders

        tbxSel2CorrectScore00Orders.Text = sel2.betfairCorrectScore00Orders
        tbxSel2CorrectScore10Orders.Text = sel2.betfairCorrectScore10Orders
        tbxSel2CorrectScore01Orders.Text = sel2.betfairCorrectScore01Orders

        ' Update refresh date/time
        tbxSel2RefreshLight.BackColor = Color.DarkGreen
        tbxSel2RefreshLight.ForeColor = Color.White
        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        tbxSel2RefreshLight.Text = time.ToString(format)

        ' Update the Inplay datetime
        Dim eventDateTime As DateTime = DateTime.Parse(tbxSel2EventDateTime.Text)
        Dim timeToStart As TimeSpan = DateTime.Now.Subtract(eventDateTime)
        Dim formatTime As String = "####0.00"
        tbxSel2InplayTime.Text = timeToStart.TotalMinutes.ToString(formatTime)

    End Sub

    Private Sub Refreshsel3Info()

        ' Get Initial book details, like marketId's and selectionId's
        sel3.getInitialBookDetails()

        ' Get latest data from Betfair
        sel3.getLatestMarketData()

        ' Update Inplay status
        If sel3.betfairEventInplay = "False" Then
            tbxSel3InplayStatus.BackColor = Color.OrangeRed
        Else
            tbxSel3InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel3Score.Text

        ' Get latest score
        tbxSel3Score.Text = sel3.betfairGoalsScored

        ' Update form
        Application.DoEvents()

        ' Detect score change
        If strPreviousScore = tbxSel3Score.Text Then
            ' Same score
        Else

            ' If first time through...ignore
            If strPreviousScore <> "" Then

                ' If match ended...ignore
                If tbxSel3Score.Text <> "Match ended!" Then

                    ' 1st Goal scored since last tick
                    If tbxSel3Score.Text = "1 Goal scored" Then
                        tbxSel3Goal1.Text = tbxSel3InplayTime.Text.ToString
                        sel3.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel3.Text + ", Goal 1 scored at: " + tbxSel3InplayTime.Text.ToString, EventLogEntryType.Information)

                        ' Send text
                        sendEmailToText("Goal 1 scored in match: " + sel3.betfairEventName + " at Inplay timer time: " + tbxSel3InplayTime.Text.ToString)

                    Else
                        If tbxSel3Score.Text = "2 Goals scored" Then
                            tbxSel3Goal2.Text = tbxSel3InplayTime.Text.ToString
                            sel3.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel3.Text + ", Goal 2 scored at: " + tbxSel3InplayTime.Text.ToString, EventLogEntryType.Information)

                            ' Send text
                            sendEmailToText("Goal 2 scored in match: " + sel3.betfairEventName + " at Inplay timer time: " + tbxSel3InplayTime.Text.ToString)

                        End If
                    End If
                End If
            End If
        End If

        ' Market Status
        tbxSel3CorrectScoreStatus.Text = sel3.betfairCorrectScoreMarketStatus
        tbxSel3UnderOver15MarketStatus.Text = sel3.betfairUnderOver15MarketStatus

        tbxSel3Over15Odds.Text = sel3.betfairOver15BackOdds
        tbxSel3Under15Odds.Text = sel3.betfairUnder15BackOdds
        tbxSel3CorrectScore00Odds.Text = sel3.betfairCorrectScore00BackOdds
        tbxSel3CorrectScore10Odds.Text = sel3.betfairCorrectScore10BackOdds
        tbxSel3CorrectScore01Odds.Text = sel3.betfairCorrectScore01BackOdds

        If sel3.betfairOver15IfWinProfit IsNot Nothing Then
            If Double.Parse(sel3.betfairOver15IfWinProfit) >= 0 Then
                tbxSel3IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel3IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel3.betfairUnder15IfWinProfit IsNot Nothing Then
            If Double.Parse(sel3.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel3IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel3IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel3.betfairCorrectScore00IfWinProfit IsNot Nothing Then
            If Double.Parse(sel3.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel3CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel3CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel3.betfairCorrectScore10IfWinProfit IsNot Nothing Then
            If Double.Parse(sel3.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel3CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel3CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel3.betfairCorrectScore01IfWinProfit IsNot Nothing Then
            If Double.Parse(sel3.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel3CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel3CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel3IOver15fWinProfit.Text = sel3.betfairOver15IfWinProfit
        tbxSel3IUnder15fWinProfit.Text = sel3.betfairUnder15IfWinProfit
        tbxSel3CorrectScore00IfWin.Text = sel3.betfairCorrectScore00IfWinProfit
        tbxSel3CorrectScore10IfWin.Text = sel3.betfairCorrectScore10IfWinProfit
        tbxSel3CorrectScore01IfWin.Text = sel3.betfairCorrectScore01IfWinProfit

        If sel3.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel3IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel3IUnder15Status.BackColor = Color.OrangeRed
        End If
        If sel3.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel3IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel3IOver15Status.BackColor = Color.OrangeRed
        End If
        If sel3.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel3CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel3CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If sel3.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel3CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel3CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If sel3.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel3CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel3CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel3IUnder15Status.Text = sel3.betfairUnder15SelectionStatus
        tbxSel3IOver15Status.Text = sel3.betfairOver15SelectionStatus
        tbxSel3CorrectScore00Status.Text = sel3.betfairCorrectScore00SelectionStatus
        tbxSel3CorrectScore10Status.Text = sel3.betfairCorrectScore10SelectionStatus
        tbxSel3CorrectScore01Status.Text = sel3.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel3IOver15Orders.Text = sel3.betfairOver15Orders
        tbxSel3IUnder15Orders.Text = sel3.betfairUnder15Orders

        tbxSel3CorrectScore00Orders.Text = sel3.betfairCorrectScore00Orders
        tbxSel3CorrectScore10Orders.Text = sel3.betfairCorrectScore10Orders
        tbxSel3CorrectScore01Orders.Text = sel3.betfairCorrectScore01Orders

        ' Update refresh date/time
        tbxSel3RefreshLight.BackColor = Color.DarkGreen
        tbxSel3RefreshLight.ForeColor = Color.White
        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        tbxSel3RefreshLight.Text = time.ToString(format)

        ' Update the Inplay datetime
        Dim eventDateTime As DateTime = DateTime.Parse(tbxSel3EventDateTime.Text)
        Dim timeToStart As TimeSpan = DateTime.Now.Subtract(eventDateTime)
        Dim formatTime As String = "####0.00"
        tbxSel3InplayTime.Text = timeToStart.TotalMinutes.ToString(formatTime)

    End Sub

    Private Sub Refreshsel4Info()

        ' Get Initial book details, like marketId's and selectionId's
        sel4.getInitialBookDetails()

        ' Get latest data from Betfair
        sel4.getLatestMarketData()

        ' Update Inplay status
        If sel4.betfairEventInplay = "False" Then
            tbxSel4InplayStatus.BackColor = Color.OrangeRed
        Else
            tbxSel4InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel4Score.Text

        ' Get latest score
        tbxSel4Score.Text = sel4.betfairGoalsScored

        ' Update form
        Application.DoEvents()


        ' Detect score change
        If strPreviousScore = tbxSel4Score.Text Then
            ' Same score
        Else

            ' If first time through...ignore
            If strPreviousScore <> "" Then

                ' If match ended...ignore
                If tbxSel4Score.Text <> "Match ended!" Then

                    ' 1st Goal scored since last tick
                    If tbxSel4Score.Text = "1 Goal scored" Then
                        tbxSel4Goal1.Text = tbxSel4InplayTime.Text.ToString
                        sel4.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel4.Text + ", Goal 1 scored at: " + tbxSel4InplayTime.Text.ToString, EventLogEntryType.Information)

                        ' Send text
                        sendEmailToText("Goal 1 scored in match: " + sel4.betfairEventName + " at Inplay timer time: " + tbxSel4InplayTime.Text.ToString)

                    Else
                        If tbxSel4Score.Text = "2 Goals scored" Then
                            tbxSel4Goal2.Text = tbxSel4InplayTime.Text.ToString
                            sel4.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel4.Text + ", Goal 2 scored at: " + tbxSel4InplayTime.Text.ToString, EventLogEntryType.Information)

                            ' Send text
                            sendEmailToText("Goal 2 scored in match: " + sel4.betfairEventName + " at Inplay timer time: " + tbxSel4InplayTime.Text.ToString)

                        End If
                    End If
                End If
            End If
        End If

        ' Market Status
        tbxSel4CorrectScoreStatus.Text = sel4.betfairCorrectScoreMarketStatus
        tbxSel4UnderOver15MarketStatus.Text = sel4.betfairUnderOver15MarketStatus

        tbxSel4Over15Odds.Text = sel4.betfairOver15BackOdds
        tbxSel4Under15Odds.Text = sel4.betfairUnder15BackOdds
        tbxSel4CorrectScore00Odds.Text = sel4.betfairCorrectScore00BackOdds
        tbxSel4CorrectScore10Odds.Text = sel4.betfairCorrectScore10BackOdds
        tbxSel4CorrectScore01Odds.Text = sel4.betfairCorrectScore01BackOdds

        If sel4.betfairOver15IfWinProfit IsNot Nothing Then
            If Double.Parse(sel4.betfairOver15IfWinProfit) >= 0 Then
                tbxSel4IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel4IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel4.betfairUnder15IfWinProfit IsNot Nothing Then
            If Double.Parse(sel4.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel4IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel4IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel4.betfairCorrectScore00IfWinProfit IsNot Nothing Then
            If Double.Parse(sel4.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel4CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel4CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel4.betfairCorrectScore10IfWinProfit IsNot Nothing Then
            If Double.Parse(sel4.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel4CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel4CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel4.betfairCorrectScore01IfWinProfit IsNot Nothing Then
            If Double.Parse(sel4.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel4CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel4CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel4IOver15fWinProfit.Text = sel4.betfairOver15IfWinProfit
        tbxSel4IUnder15fWinProfit.Text = sel4.betfairUnder15IfWinProfit
        tbxSel4CorrectScore00IfWin.Text = sel4.betfairCorrectScore00IfWinProfit
        tbxSel4CorrectScore10IfWin.Text = sel4.betfairCorrectScore10IfWinProfit
        tbxSel4CorrectScore01IfWin.Text = sel4.betfairCorrectScore01IfWinProfit

        If sel4.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel4IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel4IUnder15Status.BackColor = Color.OrangeRed
        End If
        If sel4.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel4IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel4IOver15Status.BackColor = Color.OrangeRed
        End If
        If sel4.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel4CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel4CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If sel4.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel4CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel4CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If sel4.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel4CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel4CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel4IUnder15Status.Text = sel4.betfairUnder15SelectionStatus
        tbxSel4IOver15Status.Text = sel4.betfairOver15SelectionStatus
        tbxSel4CorrectScore00Status.Text = sel4.betfairCorrectScore00SelectionStatus
        tbxSel4CorrectScore10Status.Text = sel4.betfairCorrectScore10SelectionStatus
        tbxSel4CorrectScore01Status.Text = sel4.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel4IOver15Orders.Text = sel4.betfairOver15Orders
        tbxSel4IUnder15Orders.Text = sel4.betfairUnder15Orders

        tbxSel4CorrectScore00Orders.Text = sel4.betfairCorrectScore00Orders
        tbxSel4CorrectScore10Orders.Text = sel4.betfairCorrectScore10Orders
        tbxSel4CorrectScore01Orders.Text = sel4.betfairCorrectScore01Orders

        ' Update refresh date/time
        tbxSel4RefreshLight.BackColor = Color.DarkGreen
        tbxSel4RefreshLight.ForeColor = Color.White
        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        tbxSel4RefreshLight.Text = time.ToString(format)

        ' Update the Inplay datetime
        Dim eventDateTime As DateTime = DateTime.Parse(tbxSel4EventDateTime.Text)
        Dim timeToStart As TimeSpan = DateTime.Now.Subtract(eventDateTime)
        Dim formatTime As String = "####0.00"
        tbxSel4InplayTime.Text = timeToStart.TotalMinutes.ToString(formatTime)

    End Sub

    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If MsgBox("Are you sure you want to Exit ?", vbYesNo) = vbNo Then
            e.Cancel = True
        Else

            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Logging out of Betfair", EventLogEntryType.Information)

            ' Login
            Account.Logout()

        End If

    End Sub

    Private Sub btnEmailTest_Click(sender As Object, e As EventArgs)

        Dim drResult As DialogResult = frmEmail.ShowDialog()

    End Sub

    Public Sub sendEmailToText(message As String)
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("paulowensmith68@gmail.com", "rdbosmtupcwjltcx")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("paulowensmith68@gmail.com")
            e_mail.To.Add("tlgrp1144839@txtlocal.co.uk")
            e_mail.Subject = "Betfair App"
            e_mail.IsBodyHtml = False
            e_mail.Body = message + "##"
            Smtp_Server.Send(e_mail)
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Text sent successfully. Message: " + message, EventLogEntryType.Information)

        Catch ex As Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Text sending error: " + ex.Message, EventLogEntryType.Error)
        End Try

    End Sub

    Private Function adjustOddsToMatch(odds As Double, directionOdds As String) As Double

        ' Adjustable multiplier
        Dim tickMultiplier As Integer = 1

        ' Direction Odds - either OUT or IN (OUT e.g. is going from 10 to 15, IN is going 3 to 2)

        ' Odds between 1-5 increments 0.1
        ' Odds between 6-9 increments 0.2
        ' Odds over 10 increments 0.5
        ' odds over 20 increments 1
        ' odds over 30 increment 2
        If odds > 1 And odds < 6 Then
            If directionOdds = "OUT" Then
                'Subtract
                odds = odds - 0.01
                Return odds
            Else
                ' Add
                odds = odds + 0.01
                Return odds
            End If

        ElseIf odds >= 6 And odds < 10 Then

            If directionOdds = "OUT" Then
                'Subtract
                odds = odds - 0.02
                Return odds
            Else
                ' Add
                odds = odds + 0.02
                Return odds
            End If

        ElseIf odds >= 10 And odds < 20 Then

            If directionOdds = "OUT" Then
                'Subtract
                odds = odds - 0.5
                Return odds
            Else
                ' Add
                odds = odds + 0.5
                Return odds
            End If


        ElseIf odds >= 20 And odds < 30 Then

            If directionOdds = "OUT" Then
                'Subtract
                odds = odds - 1
                Return odds
            Else
                ' Add
                odds = odds + 1
                Return odds
            End If
        Else
            ' Return odds
            Return odds
        End If

    End Function

End Class
