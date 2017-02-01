Imports System.ComponentModel
Imports System.IO
Imports System.Net.Mail

Public Class frmMain

    Public Sel1 As New Selection(1)
    Public Sel2 As New Selection(2)
    Public Sel3 As New Selection(3)
    Public Sel4 As New Selection(4)


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
        marketCountriesEurope = New HashSet(Of String)({"GB", "FR", "DE", "IT", "ES", "PT", "NL", "GR", "TR"})

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
                    sel1.autobetOver15StartegyStarted = False
                    sel1.autobetUnder15BetMade = False
                    sel1.autobetCorrectScore00BetMade = False
                    sel1.autobetCorrectScore00BTLBetMade = False
                    sel1.autobetCorrectScore10BTLBetMade = False
                    sel1.autobetCorrectScore01BTLBetMade = False
                    sel1.autobetOver15TopUpBetMade = False
                    sel1.autobetCashOutNoGoalsAtHalfTime = False

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
        ' Look to back OVER 1.5
        '
        If cbxOver15On.Checked = True Then

            If tbxSel1Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 position", EventLogEntryType.Information)

                If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +30 Then

                        If Not String.IsNullOrEmpty(sel1.betfairOver15BackOdds) Then
                            If CDbl(sel1.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel1.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
                                    If CDbl(sel1.betfairOver15Orders) = 0 Then
                                        If sel1.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 All pre-conditions met......", EventLogEntryType.Information)

                                            sel1.autobetOver15StartegyStarted = True

                                            If CDbl(tbxSel1InplayTime.Text) > +0 And sel1.autobetOver15Back1 = False Then

                                                ' Place back bet on Over1.5
                                                Dim odds As Double
                                                Dim oddsMarket As Double
                                                Dim stake As Double
                                                odds = adjustOddsDownLadder(CDbl(sel1.betfairOver15BackOdds), 2)
                                                oddsMarket = CDbl(sel1.betfairOver15BackOdds)
                                                stake = nudSettingsOver15Stake.Value
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeOver15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetOver15Back1 = True

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 BACK BET 1 with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - OVER 1.5 Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            End If

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
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS00BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore00BTLBetMade = True Then
            If CDbl(tbxSel1InplayTime.Text) >= CDbl(sel1.autobetCorrectScore00BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore00BTLBetMade = False
            End If
        End If


        ' 
        ' Look to cover 0-0
        '
        If sel1.autobetCorrectScore00BetMade = False And sel1.autobetCorrectScore00BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL00On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                                        If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct score 0-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)

                                                If CDbl(tbxSel1InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore00_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore00_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore00Orders = "1"
                                                sel1.autobetCorrectScore00BTLTimeLastBet = tbxSel1InplayTime.Text
                                                sel1.autobetCorrectScore00BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct score 0-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If

                        '' Check time band, this is the final bet to cover difference
                        'If CDbl(tbxSel1InplayTime.Text) > +30 And CDbl(tbxSel1InplayTime.Text) < +45 Then
                        '    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then
                        '        If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                        '            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                        '                If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                        '                    ' calculate stake based on profit, minus any balance
                        '                    Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                        '                    Dim currentIfWin As Double = 0
                        '                    Dim liability As Double = 0
                        '                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                        '                        currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                        '                    End If
                        '                    Dim odds As Double
                        '                    Dim oddsMarket As Double
                        '                    Dim stake As Double
                        '                    odds = adjustOddsDownLadder(CDbl(sel1.betfairCorrectScore00BackOdds), 3)
                        '                    oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)
                        '                    liability = grossPerMarket - currentIfWin
                        '                    stake = liability / (oddsMarket - 1)
                        '                    If stake > 0 Then
                        '                        If stake < +2 Then
                        '                            stake = +2
                        '                        End If
                        '                        sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                        '                        sel1.autobetCorrectScore00BetMade = True

                        '                        ' Place order on Correct Score 0-0 market
                        '                        Dim orderStatus As String
                        '                        orderStatus = sel1.placeCorrectScore00_Order(odds, stake, "Back")
                        '                        checkOrderStatus(sel1, orderStatus)

                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0 CLOSE Strategy BACK BET with the following.....", EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Liability: " + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-0  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                        '                    Else
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 position - Stake calculated as < 0, no further action taken", EventLogEntryType.Information)
                        '                    End If
                        '                Else
                        '                    ' Unmatched orders
                        '                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                        '                End If
                        '            Else
                        '                ' Unmatched orders are either NULL or EMPTY
                        '                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '            End If
                        '        Else
                        '            ' Odds are either Odds not within limits
                        '            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                        '        End If
                        '    Else
                        '        ' Odds are either NULL or EMPTY
                        '        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '    End If
                        'Else
                        '    ' Not first half of match
                        '    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        'End If

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
        If cbxUnder15On.Checked = True Then

            If sel1.autobetUnder15BetMade = False Then

                ' Check the strategy has started and score only 1 goal
                If sel1.autobetOver15StartegyStarted = True Then

                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                    If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                        ' Check in first half
                        If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +45 Then
                            If Not String.IsNullOrEmpty(sel1.betfairUnder15BackOdds) Then
                                If CDbl(sel1.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel1.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                    If Not String.IsNullOrEmpty(sel1.betfairUnder15Orders) Then
                                        If CDbl(sel1.betfairUnder15Orders) = 0 Then

                                            ' calculate liability
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            Dim liability As Double = 0
                                            Dim currentLiabilityUnder15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
                                                currentLiabilityUnder15 = CDbl(sel1.betfairUnder15IfWinProfit)
                                            End If
                                            Dim currentLiabilityOver15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
                                                currentLiabilityOver15 = CDbl(sel1.betfairCorrectScore00IfloseProfit)
                                            End If

                                            ' Set odds
                                            odds = adjustOddsDownLadder(CDbl(sel1.betfairUnder15BackOdds), 2)
                                            oddsMarket = CDbl(sel1.betfairUnder15BackOdds)

                                            If currentLiabilityUnder15 < 0 Then
                                                If currentLiabilityOver15 > 0 Then
                                                    ' Expected path
                                                    liability = ((currentLiabilityUnder15 * -1) + (currentLiabilityOver15))
                                                    stake = liability / (oddsMarket)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 position - Negative liability on both markets, Stake: " + FormatNumber(CDbl(stake), 2).ToString + " Price at market: " + oddsMarket.ToString, EventLogEntryType.Information)
                                                Else
                                                    stake = 0
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 position - Liability didn't look right, stake set to 0", EventLogEntryType.Information)
                                                End If
                                            End If

                                            If stake > 0 Then
                                                If stake < +2 Then
                                                    stake = +2
                                                End If
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetUnder15BetMade = True

                                                ' Place order on Under 1.5 market
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeUnder15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  Liability (includes £10 profit): £" + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            Else
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 position - Stake calculated as <= 0, no further action taken", EventLogEntryType.Information)
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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5 - Inplay timer not between +0 and 45 mins, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Market not open
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                    End If
                End If
            End If
        End If

        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS10BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore10BTLBetMade = True Then
            If CDbl(tbxSel1InplayTime.Text) >= CDbl(sel1.autobetCorrectScore10BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore10BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore10BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 1 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "1 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL10On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore10BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore10BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10Orders) Then
                                        If CDbl(sel1.betfairCorrectScore10Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct score 1-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore10IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore10BackOdds)

                                                If CDbl(tbxSel1InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 1 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 1-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore10_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 1-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore10_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore10Orders = "1"
                                                sel1.autobetCorrectScore10BTLTimeLastBet = tbxSel1InplayTime.Text
                                                sel1.autobetCorrectScore10BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct score 1-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS01BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore01BTLBetMade = True Then
            If CDbl(tbxSel1InplayTime.Text) >= CDbl(sel1.autobetCorrectScore01BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore01BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore01BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 1
            If sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "0 - 1" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL01On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore01BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore01BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01Orders) Then
                                        If CDbl(sel1.betfairCorrectScore01Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct score 0-1 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore01IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore01BackOdds)

                                                If CDbl(tbxSel1InplayTime.Text) <= 1 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 1 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 1 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 1 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-1 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore01_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-1 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore01_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore01Orders = "1"
                                                sel1.autobetCorrectScore01BTLTimeLastBet = tbxSel1InplayTime.Text
                                                sel1.autobetCorrectScore01BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct score 0-1 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel1 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' Check the strategy has got to half time and no Goal
        If sel1.autobetCashOutNoGoalsAtHalfTime = False Then
            If sel1.autobetOver15StartegyStarted = True And sel1.autobetUnder15BetMade = False And tbxSel1Score.Text = "0 - 0" Then
                If CDbl(tbxSel1InplayTime.Text) > +50 And CDbl(tbxSel1InplayTime.Text) < +65 Then
                    sel1.autobetCashOutNoGoalsAtHalfTime = True
                    sendEmailToText("Match: " + sel1.betfairEventName + " reached 1/2 time and no goals. Check CASH OUT option")
                End If
            End If
        End If


    End Sub

    Private Sub btnSel2AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel2AutoBetOn.Click

        If btnSel2AutoBetOn.Text = "Autobet On" Then

            If tbxSel2EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel1.autobetOver15StartegyStarted = False
                    sel1.autobetUnder15BetMade = False
                    sel1.autobetCorrectScore00BetMade = False
                    sel1.autobetCorrectScore00BTLBetMade = False
                    sel1.autobetCorrectScore10BTLBetMade = False
                    sel1.autobetCorrectScore01BTLBetMade = False
                    sel1.autobetOver15TopUpBetMade = False
                    sel1.autobetCashOutNoGoalsAtHalfTime = False

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
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel2ProfitStatus00.BackColor = Color.LawnGreen
                btnSel2ProfitStatus00.Text = sel1.betfairCorrectScore00IfWinProfit
            Else
                btnSel2ProfitStatus00.BackColor = Color.White
                btnSel2ProfitStatus00.Text = ""
            End If
        Else
            btnSel2ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel2ProfitStatus10.BackColor = Color.LawnGreen
                btnSel2ProfitStatus10.Text = sel1.betfairCorrectScore10IfWinProfit
            Else
                btnSel2ProfitStatus10.BackColor = Color.White
                btnSel2ProfitStatus10.Text = ""
            End If
        Else
            btnSel2ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel2ProfitStatus01.BackColor = Color.LawnGreen
                btnSel2ProfitStatus01.Text = sel1.betfairCorrectScore01IfWinProfit
            Else
                btnSel2ProfitStatus01.BackColor = Color.White
                btnSel2ProfitStatus01.Text = ""
            End If
        Else
            btnSel2ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
            If CDbl(sel1.betfairUnder15IfWinProfit) > 0 Then
                btnSel2ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel2ProfitStatusUnder15.Text = sel1.betfairUnder15IfWinProfit
            Else
                btnSel2ProfitStatusUnder15.BackColor = Color.White
                btnSel2ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel2ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
            If CDbl(sel1.betfairOver15IfWinProfit) > 0 Then
                btnSel2ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel2ProfitStatusOver15.Text = sel1.betfairOver15IfWinProfit
            Else
                btnSel2ProfitStatusOver15.BackColor = Color.White
                btnSel2ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel2ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel1.betfairEventInplay = "True" Then
            ' Continue
        Else
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
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
        ' Look to back OVER 1.5
        '
        If cbxOver15On.Checked = True Then

            If tbxSel2Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 position", EventLogEntryType.Information)

                If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +30 Then

                        If Not String.IsNullOrEmpty(sel1.betfairOver15BackOdds) Then
                            If CDbl(sel1.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel1.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
                                    If CDbl(sel1.betfairOver15Orders) = 0 Then
                                        If sel1.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 All pre-conditions met......", EventLogEntryType.Information)

                                            sel1.autobetOver15StartegyStarted = True

                                            If CDbl(tbxSel2InplayTime.Text) > +0 And sel1.autobetOver15Back1 = False Then

                                                ' Place back bet on Over1.5
                                                Dim odds As Double
                                                Dim oddsMarket As Double
                                                Dim stake As Double
                                                odds = adjustOddsDownLadder(CDbl(sel1.betfairOver15BackOdds), 2)
                                                oddsMarket = CDbl(sel1.betfairOver15BackOdds)
                                                stake = nudSettingsOver15Stake.Value
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeOver15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetOver15Back1 = True

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 BACK BET 1 with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - OVER 1.5 Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            End If

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
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS00BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore00BTLBetMade = True Then
            If CDbl(tbxSel2InplayTime.Text) >= CDbl(sel1.autobetCorrectScore00BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore00BTLBetMade = False
            End If
        End If


        ' 
        ' Look to cover 0-0
        '
        If sel1.autobetCorrectScore00BetMade = False And sel1.autobetCorrectScore00BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel2Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL00On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                                        If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct score 0-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)

                                                If CDbl(tbxSel2InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore00_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore00_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore00Orders = "1"
                                                sel1.autobetCorrectScore00BTLTimeLastBet = tbxSel2InplayTime.Text
                                                sel1.autobetCorrectScore00BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct score 0-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If

                        '' Check time band, this is the final bet to cover difference
                        'If CDbl(tbxSel2InplayTime.Text) > +30 And CDbl(tbxSel2InplayTime.Text) < +45 Then
                        '    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then
                        '        If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                        '            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                        '                If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                        '                    ' calculate stake based on profit, minus any balance
                        '                    Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                        '                    Dim currentIfWin As Double = 0
                        '                    Dim liability As Double = 0
                        '                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                        '                        currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                        '                    End If
                        '                    Dim odds As Double
                        '                    Dim oddsMarket As Double
                        '                    Dim stake As Double
                        '                    odds = adjustOddsDownLadder(CDbl(sel1.betfairCorrectScore00BackOdds), 3)
                        '                    oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)
                        '                    liability = grossPerMarket - currentIfWin
                        '                    stake = liability / (oddsMarket - 1)
                        '                    If stake > 0 Then
                        '                        If stake < +2 Then
                        '                            stake = +2
                        '                        End If
                        '                        sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                        '                        sel1.autobetCorrectScore00BetMade = True

                        '                        ' Place order on Correct Score 0-0 market
                        '                        Dim orderStatus As String
                        '                        orderStatus = sel1.placeCorrectScore00_Order(odds, stake, "Back")
                        '                        checkOrderStatus(sel1, orderStatus)

                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0 CLOSE Strategy BACK BET with the following.....", EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Liability: " + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-0  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                        '                    Else
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 position - Stake calculated as < 0, no further action taken", EventLogEntryType.Information)
                        '                    End If
                        '                Else
                        '                    ' Unmatched orders
                        '                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                        '                End If
                        '            Else
                        '                ' Unmatched orders are either NULL or EMPTY
                        '                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '            End If
                        '        Else
                        '            ' Odds are either Odds not within limits
                        '            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                        '        End If
                        '    Else
                        '        ' Odds are either NULL or EMPTY
                        '        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '    End If
                        'Else
                        '    ' Not first half of match
                        '    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        'End If

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
        If cbxUnder15On.Checked = True Then

            If sel1.autobetUnder15BetMade = False Then

                ' Check the strategy has started and score only 1 goal
                If sel1.autobetOver15StartegyStarted = True Then

                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                    If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                        ' Check in first half
                        If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +45 Then
                            If Not String.IsNullOrEmpty(sel1.betfairUnder15BackOdds) Then
                                If CDbl(sel1.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel1.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                    If Not String.IsNullOrEmpty(sel1.betfairUnder15Orders) Then
                                        If CDbl(sel1.betfairUnder15Orders) = 0 Then

                                            ' calculate liability
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            Dim liability As Double = 0
                                            Dim currentLiabilityUnder15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
                                                currentLiabilityUnder15 = CDbl(sel1.betfairUnder15IfWinProfit)
                                            End If
                                            Dim currentLiabilityOver15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
                                                currentLiabilityOver15 = CDbl(sel1.betfairCorrectScore00IfloseProfit)
                                            End If

                                            ' Set odds
                                            odds = adjustOddsDownLadder(CDbl(sel1.betfairUnder15BackOdds), 2)
                                            oddsMarket = CDbl(sel1.betfairUnder15BackOdds)

                                            If currentLiabilityUnder15 < 0 Then
                                                If currentLiabilityOver15 > 0 Then
                                                    ' Expected path
                                                    liability = ((currentLiabilityUnder15 * -1) + (currentLiabilityOver15))
                                                    stake = liability / (oddsMarket)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 position - Negative liability on both markets, Stake: " + FormatNumber(CDbl(stake), 2).ToString + " Price at market: " + oddsMarket.ToString, EventLogEntryType.Information)
                                                Else
                                                    stake = 0
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 position - Liability didn't look right, stake set to 0", EventLogEntryType.Information)
                                                End If
                                            End If

                                            If stake > 0 Then
                                                If stake < +2 Then
                                                    stake = +2
                                                End If
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetUnder15BetMade = True

                                                ' Place order on Under 1.5 market
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeUnder15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  Liability (includes £10 profit): £" + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            Else
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 position - Stake calculated as <= 0, no further action taken", EventLogEntryType.Information)
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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5 - Inplay timer not between +0 and 45 mins, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Market not open
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                    End If
                End If
            End If
        End If

        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS10BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore10BTLBetMade = True Then
            If CDbl(tbxSel2InplayTime.Text) >= CDbl(sel1.autobetCorrectScore10BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore10BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore10BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 1 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel2Score.Text = "1 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL10On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore10BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore10BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10Orders) Then
                                        If CDbl(sel1.betfairCorrectScore10Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct score 1-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore10IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore10BackOdds)

                                                If CDbl(tbxSel2InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 1 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 1-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore10_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 1-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore10_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore10Orders = "1"
                                                sel1.autobetCorrectScore10BTLTimeLastBet = tbxSel2InplayTime.Text
                                                sel1.autobetCorrectScore10BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct score 1-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS01BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore01BTLBetMade = True Then
            If CDbl(tbxSel2InplayTime.Text) >= CDbl(sel1.autobetCorrectScore01BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore01BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore01BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 1
            If sel1.autobetUnder15BetMade = False And tbxSel2Score.Text = "0 - 1" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL01On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel2InplayTime.Text) > +0 And CDbl(tbxSel2InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore01BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore01BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01Orders) Then
                                        If CDbl(sel1.betfairCorrectScore01Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct score 0-1 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore01IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore01BackOdds)

                                                If CDbl(tbxSel2InplayTime.Text) <= 1 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 1 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 1 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 1 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-1 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore01_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-1 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore01_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore01Orders = "1"
                                                sel1.autobetCorrectScore01BTLTimeLastBet = tbxSel2InplayTime.Text
                                                sel1.autobetCorrectScore01BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct score 0-1 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel2 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' Check the strategy has got to half time and no Goal
        If sel1.autobetCashOutNoGoalsAtHalfTime = False Then
            If sel1.autobetOver15StartegyStarted = True And sel1.autobetUnder15BetMade = False And tbxSel2Score.Text = "0 - 0" Then
                If CDbl(tbxSel2InplayTime.Text) > +50 And CDbl(tbxSel2InplayTime.Text) < +65 Then
                    sel1.autobetCashOutNoGoalsAtHalfTime = True
                    sendEmailToText("Match: " + sel1.betfairEventName + " reached 1/2 time and no goals. Check CASH OUT option")
                End If
            End If
        End If


    End Sub

    Private Sub btnSel3AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel3AutoBetOn.Click

        If btnSel3AutoBetOn.Text = "Autobet On" Then

            If tbxSel3EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel1.autobetOver15StartegyStarted = False
                    sel1.autobetUnder15BetMade = False
                    sel1.autobetCorrectScore00BetMade = False
                    sel1.autobetCorrectScore00BTLBetMade = False
                    sel1.autobetCorrectScore10BTLBetMade = False
                    sel1.autobetCorrectScore01BTLBetMade = False
                    sel1.autobetOver15TopUpBetMade = False
                    sel1.autobetCashOutNoGoalsAtHalfTime = False

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
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel3ProfitStatus00.BackColor = Color.LawnGreen
                btnSel3ProfitStatus00.Text = sel1.betfairCorrectScore00IfWinProfit
            Else
                btnSel3ProfitStatus00.BackColor = Color.White
                btnSel3ProfitStatus00.Text = ""
            End If
        Else
            btnSel3ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel3ProfitStatus10.BackColor = Color.LawnGreen
                btnSel3ProfitStatus10.Text = sel1.betfairCorrectScore10IfWinProfit
            Else
                btnSel3ProfitStatus10.BackColor = Color.White
                btnSel3ProfitStatus10.Text = ""
            End If
        Else
            btnSel3ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel3ProfitStatus01.BackColor = Color.LawnGreen
                btnSel3ProfitStatus01.Text = sel1.betfairCorrectScore01IfWinProfit
            Else
                btnSel3ProfitStatus01.BackColor = Color.White
                btnSel3ProfitStatus01.Text = ""
            End If
        Else
            btnSel3ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
            If CDbl(sel1.betfairUnder15IfWinProfit) > 0 Then
                btnSel3ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel3ProfitStatusUnder15.Text = sel1.betfairUnder15IfWinProfit
            Else
                btnSel3ProfitStatusUnder15.BackColor = Color.White
                btnSel3ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel3ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
            If CDbl(sel1.betfairOver15IfWinProfit) > 0 Then
                btnSel3ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel3ProfitStatusOver15.Text = sel1.betfairOver15IfWinProfit
            Else
                btnSel3ProfitStatusOver15.BackColor = Color.White
                btnSel3ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel3ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel1.betfairEventInplay = "True" Then
            ' Continue
        Else
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
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
        ' Look to back OVER 1.5
        '
        If cbxOver15On.Checked = True Then

            If tbxSel3Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 position", EventLogEntryType.Information)

                If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +30 Then

                        If Not String.IsNullOrEmpty(sel1.betfairOver15BackOdds) Then
                            If CDbl(sel1.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel1.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
                                    If CDbl(sel1.betfairOver15Orders) = 0 Then
                                        If sel1.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 All pre-conditions met......", EventLogEntryType.Information)

                                            sel1.autobetOver15StartegyStarted = True

                                            If CDbl(tbxSel3InplayTime.Text) > +0 And sel1.autobetOver15Back1 = False Then

                                                ' Place back bet on Over1.5
                                                Dim odds As Double
                                                Dim oddsMarket As Double
                                                Dim stake As Double
                                                odds = adjustOddsDownLadder(CDbl(sel1.betfairOver15BackOdds), 2)
                                                oddsMarket = CDbl(sel1.betfairOver15BackOdds)
                                                stake = nudSettingsOver15Stake.Value
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeOver15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetOver15Back1 = True

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 BACK BET 1 with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - OVER 1.5 Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            End If

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
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS00BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore00BTLBetMade = True Then
            If CDbl(tbxSel3InplayTime.Text) >= CDbl(sel1.autobetCorrectScore00BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore00BTLBetMade = False
            End If
        End If


        ' 
        ' Look to cover 0-0
        '
        If sel1.autobetCorrectScore00BetMade = False And sel1.autobetCorrectScore00BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel3Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL00On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                                        If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct score 0-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)

                                                If CDbl(tbxSel3InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore00_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore00_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore00Orders = "1"
                                                sel1.autobetCorrectScore00BTLTimeLastBet = tbxSel3InplayTime.Text
                                                sel1.autobetCorrectScore00BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct score 0-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If

                        '' Check time band, this is the final bet to cover difference
                        'If CDbl(tbxSel3InplayTime.Text) > +30 And CDbl(tbxSel3InplayTime.Text) < +45 Then
                        '    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then
                        '        If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                        '            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                        '                If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                        '                    ' calculate stake based on profit, minus any balance
                        '                    Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                        '                    Dim currentIfWin As Double = 0
                        '                    Dim liability As Double = 0
                        '                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                        '                        currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                        '                    End If
                        '                    Dim odds As Double
                        '                    Dim oddsMarket As Double
                        '                    Dim stake As Double
                        '                    odds = adjustOddsDownLadder(CDbl(sel1.betfairCorrectScore00BackOdds), 3)
                        '                    oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)
                        '                    liability = grossPerMarket - currentIfWin
                        '                    stake = liability / (oddsMarket - 1)
                        '                    If stake > 0 Then
                        '                        If stake < +2 Then
                        '                            stake = +2
                        '                        End If
                        '                        sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                        '                        sel1.autobetCorrectScore00BetMade = True

                        '                        ' Place order on Correct Score 0-0 market
                        '                        Dim orderStatus As String
                        '                        orderStatus = sel1.placeCorrectScore00_Order(odds, stake, "Back")
                        '                        checkOrderStatus(sel1, orderStatus)

                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0 CLOSE Strategy BACK BET with the following.....", EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Liability: " + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-0  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                        '                    Else
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 position - Stake calculated as < 0, no further action taken", EventLogEntryType.Information)
                        '                    End If
                        '                Else
                        '                    ' Unmatched orders
                        '                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                        '                End If
                        '            Else
                        '                ' Unmatched orders are either NULL or EMPTY
                        '                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '            End If
                        '        Else
                        '            ' Odds are either Odds not within limits
                        '            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                        '        End If
                        '    Else
                        '        ' Odds are either NULL or EMPTY
                        '        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '    End If
                        'Else
                        '    ' Not first half of match
                        '    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        'End If

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
        If cbxUnder15On.Checked = True Then

            If sel1.autobetUnder15BetMade = False Then

                ' Check the strategy has started and score only 1 goal
                If sel1.autobetOver15StartegyStarted = True Then

                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                    If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                        ' Check in first half
                        If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +45 Then
                            If Not String.IsNullOrEmpty(sel1.betfairUnder15BackOdds) Then
                                If CDbl(sel1.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel1.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                    If Not String.IsNullOrEmpty(sel1.betfairUnder15Orders) Then
                                        If CDbl(sel1.betfairUnder15Orders) = 0 Then

                                            ' calculate liability
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            Dim liability As Double = 0
                                            Dim currentLiabilityUnder15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
                                                currentLiabilityUnder15 = CDbl(sel1.betfairUnder15IfWinProfit)
                                            End If
                                            Dim currentLiabilityOver15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
                                                currentLiabilityOver15 = CDbl(sel1.betfairCorrectScore00IfloseProfit)
                                            End If

                                            ' Set odds
                                            odds = adjustOddsDownLadder(CDbl(sel1.betfairUnder15BackOdds), 2)
                                            oddsMarket = CDbl(sel1.betfairUnder15BackOdds)

                                            If currentLiabilityUnder15 < 0 Then
                                                If currentLiabilityOver15 > 0 Then
                                                    ' Expected path
                                                    liability = ((currentLiabilityUnder15 * -1) + (currentLiabilityOver15))
                                                    stake = liability / (oddsMarket)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 position - Negative liability on both markets, Stake: " + FormatNumber(CDbl(stake), 2).ToString + " Price at market: " + oddsMarket.ToString, EventLogEntryType.Information)
                                                Else
                                                    stake = 0
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 position - Liability didn't look right, stake set to 0", EventLogEntryType.Information)
                                                End If
                                            End If

                                            If stake > 0 Then
                                                If stake < +2 Then
                                                    stake = +2
                                                End If
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetUnder15BetMade = True

                                                ' Place order on Under 1.5 market
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeUnder15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  Liability (includes £10 profit): £" + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            Else
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 position - Stake calculated as <= 0, no further action taken", EventLogEntryType.Information)
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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5 - Inplay timer not between +0 and 45 mins, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Market not open
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                    End If
                End If
            End If
        End If

        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS10BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore10BTLBetMade = True Then
            If CDbl(tbxSel3InplayTime.Text) >= CDbl(sel1.autobetCorrectScore10BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore10BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore10BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 1 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel3Score.Text = "1 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL10On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore10BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore10BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10Orders) Then
                                        If CDbl(sel1.betfairCorrectScore10Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct score 1-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore10IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore10BackOdds)

                                                If CDbl(tbxSel3InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 1 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 1-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore10_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 1-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore10_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore10Orders = "1"
                                                sel1.autobetCorrectScore10BTLTimeLastBet = tbxSel3InplayTime.Text
                                                sel1.autobetCorrectScore10BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct score 1-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS01BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore01BTLBetMade = True Then
            If CDbl(tbxSel3InplayTime.Text) >= CDbl(sel1.autobetCorrectScore01BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore01BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore01BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 1
            If sel1.autobetUnder15BetMade = False And tbxSel3Score.Text = "0 - 1" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL01On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel3InplayTime.Text) > +0 And CDbl(tbxSel3InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore01BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore01BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01Orders) Then
                                        If CDbl(sel1.betfairCorrectScore01Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct score 0-1 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore01IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore01BackOdds)

                                                If CDbl(tbxSel3InplayTime.Text) <= 1 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 1 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 1 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 1 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-1 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore01_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-1 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore01_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore01Orders = "1"
                                                sel1.autobetCorrectScore01BTLTimeLastBet = tbxSel3InplayTime.Text
                                                sel1.autobetCorrectScore01BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct score 0-1 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel3 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' Check the strategy has got to half time and no Goal
        If sel1.autobetCashOutNoGoalsAtHalfTime = False Then
            If sel1.autobetOver15StartegyStarted = True And sel1.autobetUnder15BetMade = False And tbxSel3Score.Text = "0 - 0" Then
                If CDbl(tbxSel3InplayTime.Text) > +50 And CDbl(tbxSel3InplayTime.Text) < +65 Then
                    sel1.autobetCashOutNoGoalsAtHalfTime = True
                    sendEmailToText("Match: " + sel1.betfairEventName + " reached 1/2 time and no goals. Check CASH OUT option")
                End If
            End If
        End If


    End Sub

    Private Sub btnSel4AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel4AutoBetOn.Click

        If btnSel4AutoBetOn.Text = "Autobet On" Then

            If tbxSel4EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Initialize flags
                    sel1.autobetOver15StartegyStarted = False
                    sel1.autobetUnder15BetMade = False
                    sel1.autobetCorrectScore00BetMade = False
                    sel1.autobetCorrectScore00BTLBetMade = False
                    sel1.autobetCorrectScore10BTLBetMade = False
                    sel1.autobetCorrectScore01BTLBetMade = False
                    sel1.autobetOver15TopUpBetMade = False
                    sel1.autobetCashOutNoGoalsAtHalfTime = False

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
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore00IfWinProfit) > 0 Then
                btnSel4ProfitStatus00.BackColor = Color.LawnGreen
                btnSel4ProfitStatus00.Text = sel1.betfairCorrectScore00IfWinProfit
            Else
                btnSel4ProfitStatus00.BackColor = Color.White
                btnSel4ProfitStatus00.Text = ""
            End If
        Else
            btnSel4ProfitStatus00.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore10IfWinProfit) > 0 Then
                btnSel4ProfitStatus10.BackColor = Color.LawnGreen
                btnSel4ProfitStatus10.Text = sel1.betfairCorrectScore10IfWinProfit
            Else
                btnSel4ProfitStatus10.BackColor = Color.White
                btnSel4ProfitStatus10.Text = ""
            End If
        Else
            btnSel4ProfitStatus10.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
            If CDbl(sel1.betfairCorrectScore01IfWinProfit) > 0 Then
                btnSel4ProfitStatus01.BackColor = Color.LawnGreen
                btnSel4ProfitStatus01.Text = sel1.betfairCorrectScore01IfWinProfit
            Else
                btnSel4ProfitStatus01.BackColor = Color.White
                btnSel4ProfitStatus01.Text = ""
            End If
        Else
            btnSel4ProfitStatus01.Text = "NULL"

        End If
        If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
            If CDbl(sel1.betfairUnder15IfWinProfit) > 0 Then
                btnSel4ProfitStatusUnder15.BackColor = Color.LawnGreen
                btnSel4ProfitStatusUnder15.Text = sel1.betfairUnder15IfWinProfit
            Else
                btnSel4ProfitStatusUnder15.BackColor = Color.White
                btnSel4ProfitStatusUnder15.Text = ""
            End If
        Else
            btnSel4ProfitStatusUnder15.Text = "NULL"
        End If
        If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
            If CDbl(sel1.betfairOver15IfWinProfit) > 0 Then
                btnSel4ProfitStatusOver15.BackColor = Color.LawnGreen
                btnSel4ProfitStatusOver15.Text = sel1.betfairOver15IfWinProfit
            Else
                btnSel4ProfitStatusOver15.BackColor = Color.White
                btnSel4ProfitStatusOver15.Text = ""
            End If
        Else
            btnSel4ProfitStatusOver15.Text = "NULL"
        End If


        ' Check the status of the Event, must be Inplay
        '
        If sel1.betfairEventInplay = "True" Then
            ' Continue
        Else
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Event not in play, exiting Auto bet loop", EventLogEntryType.Information)
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
        ' Look to back OVER 1.5
        '
        If cbxOver15On.Checked = True Then

            If tbxSel4Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 position", EventLogEntryType.Information)

                If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                    ' Check in first half
                    If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +30 Then

                        If Not String.IsNullOrEmpty(sel1.betfairOver15BackOdds) Then
                            If CDbl(sel1.betfairOver15BackOdds) > nudSettingsOver15LowerPrice.Value And CDbl(sel1.betfairOver15BackOdds) < nudSettingsOver15UpperPrice.Value Then
                                If Not String.IsNullOrEmpty(sel1.betfairOver15Orders) Then
                                    If CDbl(sel1.betfairOver15Orders) = 0 Then
                                        If sel1.betfairOver15BackOdds >= nudSettingsOver15TargetPrice.Value Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 All pre-conditions met......", EventLogEntryType.Information)

                                            sel1.autobetOver15StartegyStarted = True

                                            If CDbl(tbxSel4InplayTime.Text) > +0 And sel1.autobetOver15Back1 = False Then

                                                ' Place back bet on Over1.5
                                                Dim odds As Double
                                                Dim oddsMarket As Double
                                                Dim stake As Double
                                                odds = adjustOddsDownLadder(CDbl(sel1.betfairOver15BackOdds), 2)
                                                oddsMarket = CDbl(sel1.betfairOver15BackOdds)
                                                stake = nudSettingsOver15Stake.Value
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeOver15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on OVER1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetOver15Back1 = True

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 BACK BET 1 with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - OVER 1.5 Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            End If

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
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS00BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore00BTLBetMade = True Then
            If CDbl(tbxSel4InplayTime.Text) >= CDbl(sel1.autobetCorrectScore00BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore00BTLBetMade = False
            End If
        End If


        ' 
        ' Look to cover 0-0
        '
        If sel1.autobetCorrectScore00BetMade = False And sel1.autobetCorrectScore00BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel4Score.Text = "0 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover 0 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL00On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                                        If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct score 0-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)

                                                If CDbl(tbxSel4InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore00_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore00_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore00Orders = "1"
                                                sel1.autobetCorrectScore00BTLTimeLastBet = tbxSel4InplayTime.Text
                                                sel1.autobetCorrectScore00BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct score 0-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If

                        '' Check time band, this is the final bet to cover difference
                        'If CDbl(tbxSel4InplayTime.Text) > +30 And CDbl(tbxSel4InplayTime.Text) < +45 Then
                        '    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then
                        '        If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then
                        '            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                        '                If CDbl(sel1.betfairCorrectScore00Orders) = 0 Then

                        '                    ' calculate stake based on profit, minus any balance
                        '                    Dim grossPerMarket As Double = nudSettingsCS00TargetGross.Value
                        '                    Dim currentIfWin As Double = 0
                        '                    Dim liability As Double = 0
                        '                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00IfWinProfit) Then
                        '                        currentIfWin = CDbl(sel1.betfairCorrectScore00IfWinProfit)
                        '                    End If
                        '                    Dim odds As Double
                        '                    Dim oddsMarket As Double
                        '                    Dim stake As Double
                        '                    odds = adjustOddsDownLadder(CDbl(sel1.betfairCorrectScore00BackOdds), 3)
                        '                    oddsMarket = CDbl(sel1.betfairCorrectScore00BackOdds)
                        '                    liability = grossPerMarket - currentIfWin
                        '                    stake = liability / (oddsMarket - 1)
                        '                    If stake > 0 Then
                        '                        If stake < +2 Then
                        '                            stake = +2
                        '                        End If
                        '                        sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score place back bet on 0 - 0 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                        '                        sel1.autobetCorrectScore00BetMade = True

                        '                        ' Place order on Correct Score 0-0 market
                        '                        Dim orderStatus As String
                        '                        orderStatus = sel1.placeCorrectScore00_Order(odds, stake, "Back")
                        '                        checkOrderStatus(sel1, orderStatus)

                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0 CLOSE Strategy BACK BET with the following.....", EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Liability: " + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-0  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                        '                    Else
                        '                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 position - Stake calculated as < 0, no further action taken", EventLogEntryType.Information)
                        '                    End If
                        '                Else
                        '                    ' Unmatched orders
                        '                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 position - Unmatched orders, no further action taken", EventLogEntryType.Information)
                        '                End If
                        '            Else
                        '                ' Unmatched orders are either NULL or EMPTY
                        '                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0  - Unmatched orders are NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '            End If
                        '        Else
                        '            ' Odds are either Odds not within limits
                        '            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 - Odds not within correct Upper/Lower limits, no further action taken", EventLogEntryType.Information)
                        '        End If
                        '    Else
                        '        ' Odds are either NULL or EMPTY
                        '        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0 - Odds NULL or EMPTY, no further action taken", EventLogEntryType.Information)
                        '    End If
                        'Else
                        '    ' Not first half of match
                        '    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        'End If

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
        If cbxUnder15On.Checked = True Then

            If sel1.autobetUnder15BetMade = False Then

                ' Check the strategy has started and score only 1 goal
                If sel1.autobetOver15StartegyStarted = True Then

                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover UNDER 1.5", EventLogEntryType.Information)

                    If sel1.betfairUnderOver15MarketStatus = "OPEN" Then

                        ' Check in first half
                        If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +45 Then
                            If Not String.IsNullOrEmpty(sel1.betfairUnder15BackOdds) Then
                                If CDbl(sel1.betfairUnder15BackOdds) > nudSettingsUnder15LowerPrice.Value And CDbl(sel1.betfairUnder15BackOdds) < nudSettingsUnder15UpperPrice.Value Then
                                    If Not String.IsNullOrEmpty(sel1.betfairUnder15Orders) Then
                                        If CDbl(sel1.betfairUnder15Orders) = 0 Then

                                            ' calculate liability
                                            Dim odds As Double
                                            Dim oddsMarket As Double
                                            Dim stake As Double
                                            Dim liability As Double = 0
                                            Dim currentLiabilityUnder15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairUnder15IfWinProfit) Then
                                                currentLiabilityUnder15 = CDbl(sel1.betfairUnder15IfWinProfit)
                                            End If
                                            Dim currentLiabilityOver15 As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairOver15IfWinProfit) Then
                                                currentLiabilityOver15 = CDbl(sel1.betfairCorrectScore00IfloseProfit)
                                            End If

                                            ' Set odds
                                            odds = adjustOddsDownLadder(CDbl(sel1.betfairUnder15BackOdds), 2)
                                            oddsMarket = CDbl(sel1.betfairUnder15BackOdds)

                                            If currentLiabilityUnder15 < 0 Then
                                                If currentLiabilityOver15 > 0 Then
                                                    ' Expected path
                                                    liability = ((currentLiabilityUnder15 * -1) + (currentLiabilityOver15))
                                                    stake = liability / (oddsMarket)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 position - Negative liability on both markets, Stake: " + FormatNumber(CDbl(stake), 2).ToString + " Price at market: " + oddsMarket.ToString, EventLogEntryType.Information)
                                                Else
                                                    stake = 0
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 position - Liability didn't look right, stake set to 0", EventLogEntryType.Information)
                                                End If
                                            End If

                                            If stake > 0 Then
                                                If stake < +2 Then
                                                    stake = +2
                                                End If
                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Over/Under1.5 place back bet on UNDER 1.5 Price: " + oddsMarket.ToString + " Stake: " + FormatNumber(CDbl(stake), 2).ToString)
                                                sel1.autobetUnder15BetMade = True

                                                ' Place order on Under 1.5 market
                                                Dim orderStatus As String
                                                orderStatus = sel1.placeUnder15_Order(odds, stake, "Back")
                                                checkOrderStatus(sel1, orderStatus)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + odds.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  Liability (includes £10 profit): £" + FormatNumber(CDbl(liability), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  Order Return Status : " + orderStatus, EventLogEntryType.Information)

                                            Else
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 position - Stake calculated as <= 0, no further action taken", EventLogEntryType.Information)
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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5 - Inplay timer not between +0 and 45 mins, no further action taken", EventLogEntryType.Information)
                        End If
                    Else
                        ' Market not open
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - UNDER 1.5  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                    End If
                End If
            End If
        End If

        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS10BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore10BTLBetMade = True Then
            If CDbl(tbxSel4InplayTime.Text) >= CDbl(sel1.autobetCorrectScore10BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore10BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore10BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 1 - 0
            If sel1.autobetUnder15BetMade = False And tbxSel4Score.Text = "1 - 0" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover 1 - 0", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL10On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore10BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore10BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10Orders) Then
                                        If CDbl(sel1.betfairCorrectScore10Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct score 1-0 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore10IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore10IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore10BackOdds)

                                                If CDbl(tbxSel4InplayTime.Text) <= 10 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 10 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 10 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 1 - 0 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 1-0 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore10_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 1-0 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore10_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 1-0  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore10Orders = "1"
                                                sel1.autobetCorrectScore10BTLTimeLastBet = tbxSel4InplayTime.Text
                                                sel1.autobetCorrectScore10BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct score 1-0 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 1-0  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        '  
        ' Trying to stop frequent BTL, sometimes in quiet markets the LAY bet disappears. Reset flag after nudSettingsCS01BTLMinsGap minutes
        '
        If sel1.autobetCorrectScore01BTLBetMade = True Then
            If CDbl(tbxSel4InplayTime.Text) >= CDbl(sel1.autobetCorrectScore01BTLTimeLastBet) + nudSettingsBTLMinsGap.Value Then
                sel1.autobetCorrectScore01BTLBetMade = False
            End If
        End If


        If sel1.autobetCorrectScore01BTLBetMade = False Then

            ' Check no Under 1.5 bet made and score still 0 - 1
            If sel1.autobetUnder15BetMade = False And tbxSel4Score.Text = "0 - 1" Then

                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Looking to cover 0 - 1", EventLogEntryType.Information)

                If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                    ' Check strategy selected
                    If cbxBTL01On.Checked = True Then

                        ' Check time band, this is the initial strategy
                        If CDbl(tbxSel4InplayTime.Text) > +0 And CDbl(tbxSel4InplayTime.Text) < +30 Then

                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01BackOdds) Then

                                If CDbl(sel1.betfairCorrectScore01BackOdds) > nudSettingsCS10_CS01LowerPrice.Value And CDbl(sel1.betfairCorrectScore01BackOdds) < nudSettingsCS10_CS01UpperPrice.Value Then

                                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01Orders) Then
                                        If CDbl(sel1.betfairCorrectScore01Orders) = 0 Then

                                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct score 0-1 position start, BTL strategy", EventLogEntryType.Information)

                                            ' calculate stake based on profit, minus any balance
                                            Dim grossPerMarket As Double = nudSettingsCS10and01TargetGross.Value
                                            Dim currentIfWin As Double = 0
                                            Dim liability As Double = 0
                                            If Not String.IsNullOrEmpty(sel1.betfairCorrectScore01IfWinProfit) Then
                                                currentIfWin = CDbl(sel1.betfairCorrectScore01IfWinProfit)
                                            End If
                                            liability = grossPerMarket - currentIfWin

                                            ' Check liability, if already reached don't bet any more
                                            If liability > 0 Then

                                                ' Calculate Odds 
                                                Dim oddsMarket As Double
                                                Dim oddsBack As Double
                                                Dim oddsLay As Double
                                                Dim stake As Double

                                                oddsMarket = CDbl(sel1.betfairCorrectScore01BackOdds)

                                                If CDbl(tbxSel4InplayTime.Text) <= 1 Then
                                                    stake = nudSettingsBTLStakeLT10.Value
                                                Else
                                                    stake = nudSettingsBTLStakeGT10.Value
                                                End If
                                                If oddsMarket > 1 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 8)
                                                ElseIf oddsMarket > 4 And oddsMarket <= 4 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 6)
                                                ElseIf oddsMarket > 6 And oddsMarket <= 1 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 4)
                                                ElseIf oddsMarket > 1 And oddsMarket <= 20 Then
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                Else
                                                    oddsLay = adjustOddsDownLadder(oddsMarket, 2)
                                                End If

                                                ' Bit odd (excuse the pun), but this enures BACK bet is matched, LAY bet just sits there until matched.
                                                oddsBack = oddsLay

                                                sendEmailToText("Match: " + sel1.betfairEventName + " Market: Correct Score BTL strategy, placing BACK bet and LAY bet 0 - 1 Market BACK price: " + oddsMarket.ToString + ", Order BACK Price:" + oddsBack.ToString + " Order LAY Price: " + oddsLay.ToString)

                                                ' Place BACK order on Correct Score 0-1 market
                                                Dim orderStatusBack As String
                                                orderStatusBack = sel1.placeCorrectScore01_Order(oddsBack, stake + nudSettingsBTLExtraBACK.Value, "Back")
                                                checkOrderStatus(sel1, orderStatusBack)

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1 BTL Strategy BACK BET with the following.....", EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsBack.ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusBack, EventLogEntryType.Information)


                                                ' Check if any LAY stake should be made
                                                If stake > 0 Then
                                                    ' Place LAY order on Correct Score 0-1 market
                                                    Dim orderStatusLay As String
                                                    orderStatusLay = sel1.placeCorrectScore01_Order(oddsLay, stake, "Lay")
                                                    checkOrderStatus(sel1, orderStatusLay)

                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1 BTL Strategy LAY BET with the following.....", EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1  Market Odds: " + oddsMarket.ToString + " Adjusted Odds: " + oddsLay.ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1  Stake: " + FormatNumber(CDbl(stake), 2).ToString, EventLogEntryType.Information)
                                                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - CORRECT SCORE 0-1  Order Return Status : " + orderStatusLay, EventLogEntryType.Information)

                                                End If

                                                ' Set orders to 1 so that duplicate orders not fired off
                                                sel1.betfairCorrectScore01Orders = "1"
                                                sel1.autobetCorrectScore01BTLTimeLastBet = tbxSel4InplayTime.Text
                                                sel1.autobetCorrectScore01BTLBetMade = True

                                            Else

                                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct score 0-1 position - Liability already reached.", EventLogEntryType.Information)

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
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1  - Inplay timer not within time window, no further action taken", EventLogEntryType.Information)
                        End If
                    End If

                Else
                    ' Market not open
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Autobet for Sel4 - Correct Score 0-1  - Market not OPEN, no further action taken", EventLogEntryType.Information)
                End If
            End If
        End If


        ' Check the strategy has got to half time and no Goal
        If sel1.autobetCashOutNoGoalsAtHalfTime = False Then
            If sel1.autobetOver15StartegyStarted = True And sel1.autobetUnder15BetMade = False And tbxSel4Score.Text = "0 - 0" Then
                If CDbl(tbxSel4InplayTime.Text) > +50 And CDbl(tbxSel4InplayTime.Text) < +65 Then
                    sel1.autobetCashOutNoGoalsAtHalfTime = True
                    sendEmailToText("Match: " + sel1.betfairEventName + " reached 1/2 time and no goals. Check CASH OUT option")
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
            RefreshSel1Info()
            RefreshSel1Goals()

            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True
            timerRefreshGoals.Enabled = True


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
            RefreshSel2Goals()


            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True
            timerRefreshGoals.Enabled = True


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
            RefreshSel3Goals()


            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True
            timerRefreshGoals.Enabled = True


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
            RefreshSel4Goals()


            ' Enable Refresh Timer
            timerRefreshSelections.Enabled = True
            timerRefreshGoals.Enabled = True


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

    Private Sub RefreshSel1Info()

        ' Get Initial book details, like marketId's and selectionId's
        Sel1.getInitialBookDetails()

        ' Get latest data from Betfair
        Sel1.getLatestMarketData()

        ' Update Inplay status
        If Sel1.betfairEventInplay = False Then
            tbxSel1InplayStatus.BackColor = Color.Red
            Sel1.betfairEventInPlayLastCycle = False
        Else

            tbxSel1InplayStatus.BackColor = Color.GreenYellow

            If Sel1.betfairEventInPlayLastCycle = False Then

                ' Ensure we don't keep resetting Event Date/Time
                Sel1.betfairEventInPlayLastCycle = True

                Dim dateTime As DateTime = DateTime.Now
                Dim formatEventDateTime As String = "M/d/yyyy h:mm tt"
                tbxSel1EventDateTime.Text = dateTime.ToString(formatEventDateTime)
                Sel1.betfairEventDateTime = dateTime.ToString(formatEventDateTime)

            End If

        End If

        ' Market Status
        tbxSel1CorrectScoreStatus.Text = Sel1.betfairCorrectScoreMarketStatus
        tbxSel1UnderOver15MarketStatus.Text = Sel1.betfairUnderOver15MarketStatus

        tbxSel1Over15Odds.Text = Sel1.betfairOver15BackOdds
        tbxSel1Under15Odds.Text = Sel1.betfairUnder15BackOdds
        tbxSel1CorrectScore00Odds.Text = Sel1.betfairCorrectScore00BackOdds
        tbxSel1CorrectScore10Odds.Text = Sel1.betfairCorrectScore10BackOdds
        tbxSel1CorrectScore01Odds.Text = Sel1.betfairCorrectScore01BackOdds

        If Not String.IsNullOrEmpty(Sel1.betfairOver15IfWinProfit) Then
            If Double.Parse(Sel1.betfairOver15IfWinProfit) >= 0 Then
                tbxSel1IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel1IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel1.betfairUnder15IfWinProfit) Then
            If Double.Parse(Sel1.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel1IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel1IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel1.betfairCorrectScore00IfWinProfit) Then
            If Double.Parse(Sel1.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel1CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel1.betfairCorrectScore10IfWinProfit) Then
            If Double.Parse(Sel1.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel1CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel1.betfairCorrectScore01IfWinProfit) Then
            If Double.Parse(Sel1.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel1CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel1IOver15fWinProfit.Text = Sel1.betfairOver15IfWinProfit
        tbxSel1IUnder15fWinProfit.Text = Sel1.betfairUnder15IfWinProfit
        tbxSel1CorrectScore00IfWin.Text = Sel1.betfairCorrectScore00IfWinProfit
        tbxSel1CorrectScore10IfWin.Text = Sel1.betfairCorrectScore10IfWinProfit
        tbxSel1CorrectScore01IfWin.Text = Sel1.betfairCorrectScore01IfWinProfit

        If Sel1.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel1IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel1IUnder15Status.BackColor = Color.OrangeRed
        End If
        If Sel1.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel1IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel1IOver15Status.BackColor = Color.OrangeRed
        End If
        If Sel1.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel1CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel1CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If Sel1.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel1CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel1CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If Sel1.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel1CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel1CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel1IUnder15Status.Text = Sel1.betfairUnder15SelectionStatus
        tbxSel1IOver15Status.Text = Sel1.betfairOver15SelectionStatus
        tbxSel1CorrectScore00Status.Text = Sel1.betfairCorrectScore00SelectionStatus
        tbxSel1CorrectScore10Status.Text = Sel1.betfairCorrectScore10SelectionStatus
        tbxSel1CorrectScore01Status.Text = Sel1.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel1IOver15Orders.Text = Sel1.betfairOver15Orders
        tbxSel1IUnder15Orders.Text = Sel1.betfairUnder15Orders

        tbxSel1CorrectScore00Orders.Text = Sel1.betfairCorrectScore00Orders
        tbxSel1CorrectScore10Orders.Text = Sel1.betfairCorrectScore10Orders
        tbxSel1CorrectScore01Orders.Text = Sel1.betfairCorrectScore01Orders

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

    Private Sub RefreshSel1Goals()

        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel1Score.Text

        ' Get latest score
        tbxSel1Score.Text = Sel1.betfairGoalsScored

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
                        Sel1.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel1.Text + ", Goal 1 scored at: " + tbxSel1InplayTime.Text.ToString, EventLogEntryType.Information)

                        If Sel1.textGoal1Sent = False Then
                            sendEmailToText("Goal 1 scored in match: " + Sel1.betfairEventName + " at Inplay timer time: " + tbxSel1InplayTime.Text.ToString)
                            Sel1.textGoal1Sent = True
                        End If
                    Else
                        If tbxSel1Score.Text = "2 Goals scored" Then
                            tbxSel1Goal2.Text = tbxSel1InplayTime.Text.ToString
                            Sel1.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel1.Text + ", Goal 2 scored at: " + tbxSel1InplayTime.Text.ToString, EventLogEntryType.Information)

                            If Sel1.textGoal1Sent = False Then
                                sendEmailToText("Goal 2 scored in match: " + Sel1.betfairEventName + " at Inplay timer time: " + tbxSel1InplayTime.Text.ToString)
                                Sel1.textGoal2Sent = True
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub RefreshSel2Info()

        ' Get Initial book details, like marketId's and selectionId's
        Sel2.getInitialBookDetails()

        ' Get latest data from Betfair
        Sel2.getLatestMarketData()

        ' Update Inplay status
        If Sel2.betfairEventInplay = False Then
            tbxSel2InplayStatus.BackColor = Color.Red
            Sel2.betfairEventInPlayLastCycle = False
        Else

            tbxSel2InplayStatus.BackColor = Color.GreenYellow

            If Sel2.betfairEventInPlayLastCycle = False Then

                ' Ensure we don't keep resetting Event Date/Time
                Sel2.betfairEventInPlayLastCycle = True

                Dim dateTime As DateTime = DateTime.Now
                Dim formatEventDateTime As String = "M/d/yyyy h:mm tt"
                tbxSel2EventDateTime.Text = dateTime.ToString(formatEventDateTime)
                Sel2.betfairEventDateTime = dateTime.ToString(formatEventDateTime)

            End If

        End If

        ' Market Status
        tbxSel2CorrectScoreStatus.Text = Sel2.betfairCorrectScoreMarketStatus
        tbxSel2UnderOver15MarketStatus.Text = Sel2.betfairUnderOver15MarketStatus

        tbxSel2Over15Odds.Text = Sel2.betfairOver15BackOdds
        tbxSel2Under15Odds.Text = Sel2.betfairUnder15BackOdds
        tbxSel2CorrectScore00Odds.Text = Sel2.betfairCorrectScore00BackOdds
        tbxSel2CorrectScore10Odds.Text = Sel2.betfairCorrectScore10BackOdds
        tbxSel2CorrectScore01Odds.Text = Sel2.betfairCorrectScore01BackOdds

        If Not String.IsNullOrEmpty(Sel2.betfairOver15IfWinProfit) Then
            If Double.Parse(Sel2.betfairOver15IfWinProfit) >= 0 Then
                tbxSel2IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel2IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel2.betfairUnder15IfWinProfit) Then
            If Double.Parse(Sel2.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel2IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel2IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel2.betfairCorrectScore00IfWinProfit) Then
            If Double.Parse(Sel2.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel2CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel2CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel2.betfairCorrectScore10IfWinProfit) Then
            If Double.Parse(Sel2.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel2CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel2CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel2.betfairCorrectScore01IfWinProfit) Then
            If Double.Parse(Sel2.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel2CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel2CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel2IOver15fWinProfit.Text = Sel2.betfairOver15IfWinProfit
        tbxSel2IUnder15fWinProfit.Text = Sel2.betfairUnder15IfWinProfit
        tbxSel2CorrectScore00IfWin.Text = Sel2.betfairCorrectScore00IfWinProfit
        tbxSel2CorrectScore10IfWin.Text = Sel2.betfairCorrectScore10IfWinProfit
        tbxSel2CorrectScore01IfWin.Text = Sel2.betfairCorrectScore01IfWinProfit

        If Sel2.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel2IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel2IUnder15Status.BackColor = Color.OrangeRed
        End If
        If Sel2.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel2IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel2IOver15Status.BackColor = Color.OrangeRed
        End If
        If Sel2.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel2CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel2CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If Sel2.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel2CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel2CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If Sel2.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel2CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel2CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel2IUnder15Status.Text = Sel2.betfairUnder15SelectionStatus
        tbxSel2IOver15Status.Text = Sel2.betfairOver15SelectionStatus
        tbxSel2CorrectScore00Status.Text = Sel2.betfairCorrectScore00SelectionStatus
        tbxSel2CorrectScore10Status.Text = Sel2.betfairCorrectScore10SelectionStatus
        tbxSel2CorrectScore01Status.Text = Sel2.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel2IOver15Orders.Text = Sel2.betfairOver15Orders
        tbxSel2IUnder15Orders.Text = Sel2.betfairUnder15Orders

        tbxSel2CorrectScore00Orders.Text = Sel2.betfairCorrectScore00Orders
        tbxSel2CorrectScore10Orders.Text = Sel2.betfairCorrectScore10Orders
        tbxSel2CorrectScore01Orders.Text = Sel2.betfairCorrectScore01Orders

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

    Private Sub RefreshSel2Goals()

        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel2Score.Text

        ' Get latest score
        tbxSel2Score.Text = Sel2.betfairGoalsScored

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
                        Sel2.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 1 scored at: " + tbxSel2InplayTime.Text.ToString, EventLogEntryType.Information)

                        If Sel2.textGoal1Sent = False Then
                            sendEmailToText("Goal 1 scored in match: " + Sel2.betfairEventName + " at Inplay timer time: " + tbxSel2InplayTime.Text.ToString)
                            Sel2.textGoal1Sent = True
                        End If
                    Else
                        If tbxSel2Score.Text = "2 Goals scored" Then
                            tbxSel2Goal2.Text = tbxSel2InplayTime.Text.ToString
                            Sel2.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 2 scored at: " + tbxSel2InplayTime.Text.ToString, EventLogEntryType.Information)

                            If Sel2.textGoal1Sent = False Then
                                sendEmailToText("Goal 2 scored in match: " + Sel2.betfairEventName + " at Inplay timer time: " + tbxSel2InplayTime.Text.ToString)
                                Sel2.textGoal2Sent = True
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub RefreshSel3Info()

        ' Get Initial book details, like marketId's and selectionId's
        Sel3.getInitialBookDetails()

        ' Get latest data from Betfair
        Sel3.getLatestMarketData()

        ' Update Inplay status
        If Sel3.betfairEventInplay = False Then
            tbxSel3InplayStatus.BackColor = Color.Red
            Sel3.betfairEventInPlayLastCycle = False
        Else

            tbxSel3InplayStatus.BackColor = Color.GreenYellow

            If Sel3.betfairEventInPlayLastCycle = False Then

                ' Ensure we don't keep resetting Event Date/Time
                Sel3.betfairEventInPlayLastCycle = True

                Dim dateTime As DateTime = DateTime.Now
                Dim formatEventDateTime As String = "M/d/yyyy h:mm tt"
                tbxSel3EventDateTime.Text = dateTime.ToString(formatEventDateTime)
                Sel3.betfairEventDateTime = dateTime.ToString(formatEventDateTime)

            End If

        End If

        ' Market Status
        tbxSel3CorrectScoreStatus.Text = Sel3.betfairCorrectScoreMarketStatus
        tbxSel3UnderOver15MarketStatus.Text = Sel3.betfairUnderOver15MarketStatus

        tbxSel3Over15Odds.Text = Sel3.betfairOver15BackOdds
        tbxSel3Under15Odds.Text = Sel3.betfairUnder15BackOdds
        tbxSel3CorrectScore00Odds.Text = Sel3.betfairCorrectScore00BackOdds
        tbxSel3CorrectScore10Odds.Text = Sel3.betfairCorrectScore10BackOdds
        tbxSel3CorrectScore01Odds.Text = Sel3.betfairCorrectScore01BackOdds

        If Not String.IsNullOrEmpty(Sel3.betfairOver15IfWinProfit) Then
            If Double.Parse(Sel3.betfairOver15IfWinProfit) >= 0 Then
                tbxSel3IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel3IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel3.betfairUnder15IfWinProfit) Then
            If Double.Parse(Sel3.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel3IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel3IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel3.betfairCorrectScore00IfWinProfit) Then
            If Double.Parse(Sel3.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel3CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel3CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel3.betfairCorrectScore10IfWinProfit) Then
            If Double.Parse(Sel3.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel3CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel3CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel3.betfairCorrectScore01IfWinProfit) Then
            If Double.Parse(Sel3.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel3CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel3CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel3IOver15fWinProfit.Text = Sel3.betfairOver15IfWinProfit
        tbxSel3IUnder15fWinProfit.Text = Sel3.betfairUnder15IfWinProfit
        tbxSel3CorrectScore00IfWin.Text = Sel3.betfairCorrectScore00IfWinProfit
        tbxSel3CorrectScore10IfWin.Text = Sel3.betfairCorrectScore10IfWinProfit
        tbxSel3CorrectScore01IfWin.Text = Sel3.betfairCorrectScore01IfWinProfit

        If Sel3.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel3IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel3IUnder15Status.BackColor = Color.OrangeRed
        End If
        If Sel3.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel3IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel3IOver15Status.BackColor = Color.OrangeRed
        End If
        If Sel3.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel3CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel3CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If Sel3.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel3CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel3CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If Sel3.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel3CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel3CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel3IUnder15Status.Text = Sel3.betfairUnder15SelectionStatus
        tbxSel3IOver15Status.Text = Sel3.betfairOver15SelectionStatus
        tbxSel3CorrectScore00Status.Text = Sel3.betfairCorrectScore00SelectionStatus
        tbxSel3CorrectScore10Status.Text = Sel3.betfairCorrectScore10SelectionStatus
        tbxSel3CorrectScore01Status.Text = Sel3.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel3IOver15Orders.Text = Sel3.betfairOver15Orders
        tbxSel3IUnder15Orders.Text = Sel3.betfairUnder15Orders

        tbxSel3CorrectScore00Orders.Text = Sel3.betfairCorrectScore00Orders
        tbxSel3CorrectScore10Orders.Text = Sel3.betfairCorrectScore10Orders
        tbxSel3CorrectScore01Orders.Text = Sel3.betfairCorrectScore01Orders

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

    Private Sub RefreshSel3Goals()

        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel3Score.Text

        ' Get latest score
        tbxSel3Score.Text = Sel3.betfairGoalsScored

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
                        Sel3.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel3.Text + ", Goal 1 scored at: " + tbxSel3InplayTime.Text.ToString, EventLogEntryType.Information)

                        If Sel3.textGoal1Sent = False Then
                            sendEmailToText("Goal 1 scored in match: " + Sel3.betfairEventName + " at Inplay timer time: " + tbxSel3InplayTime.Text.ToString)
                            Sel3.textGoal1Sent = True
                        End If
                    Else
                        If tbxSel3Score.Text = "2 Goals scored" Then
                            tbxSel3Goal2.Text = tbxSel3InplayTime.Text.ToString
                            Sel3.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel3.Text + ", Goal 2 scored at: " + tbxSel3InplayTime.Text.ToString, EventLogEntryType.Information)

                            If Sel3.textGoal1Sent = False Then
                                sendEmailToText("Goal 2 scored in match: " + Sel3.betfairEventName + " at Inplay timer time: " + tbxSel3InplayTime.Text.ToString)
                                Sel3.textGoal2Sent = True
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub RefreshSel4Info()

        ' Get Initial book details, like marketId's and selectionId's
        Sel4.getInitialBookDetails()

        ' Get latest data from Betfair
        Sel4.getLatestMarketData()

        ' Update Inplay status
        If Sel4.betfairEventInplay = False Then
            tbxSel4InplayStatus.BackColor = Color.Red
            Sel4.betfairEventInPlayLastCycle = False
        Else

            tbxSel4InplayStatus.BackColor = Color.GreenYellow

            If Sel4.betfairEventInPlayLastCycle = False Then

                ' Ensure we don't keep resetting Event Date/Time
                Sel4.betfairEventInPlayLastCycle = True

                Dim dateTime As DateTime = DateTime.Now
                Dim formatEventDateTime As String = "M/d/yyyy h:mm tt"
                tbxSel4EventDateTime.Text = dateTime.ToString(formatEventDateTime)
                Sel4.betfairEventDateTime = dateTime.ToString(formatEventDateTime)

            End If

        End If

        ' Market Status
        tbxSel4CorrectScoreStatus.Text = Sel4.betfairCorrectScoreMarketStatus
        tbxSel4UnderOver15MarketStatus.Text = Sel4.betfairUnderOver15MarketStatus

        tbxSel4Over15Odds.Text = Sel4.betfairOver15BackOdds
        tbxSel4Under15Odds.Text = Sel4.betfairUnder15BackOdds
        tbxSel4CorrectScore00Odds.Text = Sel4.betfairCorrectScore00BackOdds
        tbxSel4CorrectScore10Odds.Text = Sel4.betfairCorrectScore10BackOdds
        tbxSel4CorrectScore01Odds.Text = Sel4.betfairCorrectScore01BackOdds

        If Not String.IsNullOrEmpty(Sel4.betfairOver15IfWinProfit) Then
            If Double.Parse(Sel4.betfairOver15IfWinProfit) >= 0 Then
                tbxSel4IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel4IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel4.betfairUnder15IfWinProfit) Then
            If Double.Parse(Sel4.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel4IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel4IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel4.betfairCorrectScore00IfWinProfit) Then
            If Double.Parse(Sel4.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel4CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel4CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel4.betfairCorrectScore10IfWinProfit) Then
            If Double.Parse(Sel4.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel4CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel4CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If Not String.IsNullOrEmpty(Sel4.betfairCorrectScore01IfWinProfit) Then
            If Double.Parse(Sel4.betfairCorrectScore01IfWinProfit) >= 0 Then
                tbxSel4CorrectScore01IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel4CorrectScore01IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        tbxSel4IOver15fWinProfit.Text = Sel4.betfairOver15IfWinProfit
        tbxSel4IUnder15fWinProfit.Text = Sel4.betfairUnder15IfWinProfit
        tbxSel4CorrectScore00IfWin.Text = Sel4.betfairCorrectScore00IfWinProfit
        tbxSel4CorrectScore10IfWin.Text = Sel4.betfairCorrectScore10IfWinProfit
        tbxSel4CorrectScore01IfWin.Text = Sel4.betfairCorrectScore01IfWinProfit

        If Sel4.betfairUnder15SelectionStatus = "ACTIVE" Then
            tbxSel4IUnder15Status.BackColor = Color.LawnGreen
        Else
            tbxSel4IUnder15Status.BackColor = Color.OrangeRed
        End If
        If Sel4.betfairOver15SelectionStatus = "ACTIVE" Then
            tbxSel4IOver15Status.BackColor = Color.LawnGreen
        Else
            tbxSel4IOver15Status.BackColor = Color.OrangeRed
        End If
        If Sel4.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
            tbxSel4CorrectScore00Status.BackColor = Color.LawnGreen
        Else
            tbxSel4CorrectScore00Status.BackColor = Color.OrangeRed
        End If
        If Sel4.betfairCorrectScore10SelectionStatus = "ACTIVE" Then
            tbxSel4CorrectScore10Status.BackColor = Color.LawnGreen
        Else
            tbxSel4CorrectScore10Status.BackColor = Color.OrangeRed
        End If
        If Sel4.betfairCorrectScore01SelectionStatus = "ACTIVE" Then
            tbxSel4CorrectScore01Status.BackColor = Color.LawnGreen
        Else
            tbxSel4CorrectScore01Status.BackColor = Color.OrangeRed
        End If

        tbxSel4IUnder15Status.Text = Sel4.betfairUnder15SelectionStatus
        tbxSel4IOver15Status.Text = Sel4.betfairOver15SelectionStatus
        tbxSel4CorrectScore00Status.Text = Sel4.betfairCorrectScore00SelectionStatus
        tbxSel4CorrectScore10Status.Text = Sel4.betfairCorrectScore10SelectionStatus
        tbxSel4CorrectScore01Status.Text = Sel4.betfairCorrectScore01SelectionStatus

        ' Populate unmatched bets
        tbxSel4IOver15Orders.Text = Sel4.betfairOver15Orders
        tbxSel4IUnder15Orders.Text = Sel4.betfairUnder15Orders

        tbxSel4CorrectScore00Orders.Text = Sel4.betfairCorrectScore00Orders
        tbxSel4CorrectScore10Orders.Text = Sel4.betfairCorrectScore10Orders
        tbxSel4CorrectScore01Orders.Text = Sel4.betfairCorrectScore01Orders

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

    Private Sub RefreshSel4Goals()

        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel4Score.Text

        ' Get latest score
        tbxSel4Score.Text = Sel4.betfairGoalsScored

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
                        Sel4.betfairGoal1DateTime = Now()
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel4.Text + ", Goal 1 scored at: " + tbxSel4InplayTime.Text.ToString, EventLogEntryType.Information)

                        If Sel4.textGoal1Sent = False Then
                            sendEmailToText("Goal 1 scored in match: " + Sel4.betfairEventName + " at Inplay timer time: " + tbxSel4InplayTime.Text.ToString)
                            Sel4.textGoal1Sent = True
                        End If
                    Else
                        If tbxSel4Score.Text = "2 Goals scored" Then
                            tbxSel4Goal2.Text = tbxSel4InplayTime.Text.ToString
                            Sel4.betfairGoal2DateTime = Now()
                            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel4.Text + ", Goal 2 scored at: " + tbxSel4InplayTime.Text.ToString, EventLogEntryType.Information)

                            If Sel4.textGoal1Sent = False Then
                                sendEmailToText("Goal 2 scored in match: " + Sel4.betfairEventName + " at Inplay timer time: " + tbxSel4InplayTime.Text.ToString)
                                Sel4.textGoal2Sent = True
                            End If
                        End If
                    End If
                End If
            End If
        End If

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

    Private Function adjustOddsDownLadder(odds As Double, ticks As Integer) As Double

        ' Betfair table
        '1.01 → 2	0.01
        '2→ 3	    0.02
        '3 → 4	    0.05
        '4 → 6	    0.1
        '6 → 10	    0.2
        '10 → 20	0.5
        '20 → 30	  1
        '30 → 50	  2
        '50 → 100	  5
        '100 → 1000	 10

        ' Going down the tick sizes, so need to check when going from bottom end to top end of another
        Try

            If odds > 1.05 And odds <= 2 Then
                odds = odds - (0.01 * ticks)
                Return odds
            ElseIf odds > 2 And odds <= 3 Then
                If odds >= (2 + (0.02 * ticks)) Then
                    odds = odds - (0.02 * ticks)
                Else
                    Dim ticksAbove As Integer
                    Dim ticksBelow As Integer
                    ticksAbove = (odds - 2) / 0.02
                    ticksBelow = ticks - ticksAbove
                    odds = 2 - (0.01 * ticksBelow)
                End If
                Return odds
            ElseIf odds > 3 And odds <= 4 Then
                If odds >= (3 + (0.05 * ticks)) Then
                    odds = odds - (0.05 * ticks)
                Else
                    Dim ticksAbove As Integer
                    Dim ticksBelow As Integer
                    ticksAbove = (odds - 3) / 0.05
                    ticksBelow = ticks - ticksAbove
                    odds = 3 - (0.02 * ticksBelow)
                End If
                Return odds
            ElseIf odds > 4 And odds <= 6 Then
                If odds >= (4 + (0.1 * ticks)) Then
                    odds = odds - (0.1 * ticks)
                Else
                    Dim ticksAbove As Integer
                    Dim ticksBelow As Integer
                    ticksAbove = (odds - 4) / 0.1
                    ticksBelow = ticks - ticksAbove
                    odds = 4 - (0.05 * ticksBelow)
                End If
                Return odds
            ElseIf odds > 6 And odds <= 10 Then
                If odds >= (6 + (0.2 * ticks)) Then
                    odds = odds - (0.2 * ticks)
                Else
                    Dim ticksAbove As Integer
                    Dim ticksBelow As Integer
                    ticksAbove = (odds - 6) / 0.2
                    ticksBelow = ticks - ticksAbove
                    odds = 6 - (0.1 * ticksBelow)
                End If
                Return odds
            ElseIf odds > 10 And odds <= 20 Then
                If odds >= (10 + (0.5 * ticks)) Then
                    odds = odds - (0.5 * ticks)
                Else
                    Dim ticksAbove As Integer
                    Dim ticksBelow As Integer
                    ticksAbove = (odds - 10) / 0.5
                    ticksBelow = ticks - ticksAbove
                    odds = 10 - (0.2 * ticksBelow)
                End If
                Return odds
            ElseIf odds > 20 And odds <= 30 Then
                If odds >= (20 + (1 * ticks)) Then
                    odds = odds - (1 * ticks)
                Else
                    Dim ticksAbove As Integer
                    Dim ticksBelow As Integer
                    ticksAbove = (odds - 20) / 1
                    ticksBelow = ticks - ticksAbove
                    odds = 20 - (0.5 * ticksBelow)
                End If
                Return odds
            ElseIf odds > 30 And odds <= 50 Then
                odds = odds - (2 * ticks)
                Return odds
            ElseIf odds > 50 And odds <= 100 Then
                odds = odds - (5 * ticks)
                Return odds
            ElseIf odds > 100 And odds <= 1000 Then
                odds = odds - 10
                Return odds
            Else
                Return odds
            End If

        Catch ex As Exception
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : adjustOddsDownLadder error, Odds passed were: " + odds.ToString + " exception: " + ex.Message, EventLogEntryType.Error)
            Return odds
        End Try


    End Function

    Public Sub checkOrderStatus(sel As Selection, status As String)

        ' ExecutionReportStatus
        ' ======================
        ' SUCCESS               Order processed successfully
        ' FAILURE               Order failed.
        ' PROCESSED_WITH_ERRORS The order itself has been accepted, but at least one (possibly all) actions have generated errors. This error only occurs for replaceOrders, cancelOrders And updateOrders 
        '                       operations.The placeOrders operation will Not return PROCESSED_WITH_ERRORS status as it Is an atomic operation.
        ' TIMEOUT               Order timed out.
        '
        ' ExecutionReportErrorCode
        ' ========================
        ' ERROR_IN_MATCHER        The matcher Is Not healthy
        ' PROCESSED_WITH_ERRORS   The order itself has been accepted, but at least one (possibly all) actions have generated errors
        ' BET_ACTION_ERROR        There Is an error with an action that has caused the entire order to be rejected. Check the instructionReports errorCode for the reason for the rejection of the order.
        ' INVALID_ACCOUNT_STATE   Order rejected due to the account's status (suspended, inactive, dup cards)
        ' INVALID_WALLET_STATUS   Order rejected due to the account's wallet's status
        ' INSUFFICIENT_FUNDS      Account has exceeded its exposure limit Or available to bet limit
        ' LOSS_LIMIT_EXCEEDED     The Account has exceed the self imposed loss limit
        ' MARKET_SUSPENDED        Market Is suspended
        ' MARKET_NOT_OPEN_FOR_BETTING   Market Is Not open For betting. It Is either Not yet active, suspended Or Closed awaiting settlement.
        ' DUPLICATE_TRANSACTION   Duplicate customer reference data submitted - Please note: There Is a time window associated with the de-duplication of duplicate submissions which Is 60 second
        ' INVALID_ORDER           Order cannot be accepted by the matcher due To the combination Of actions. For example, bets being edited are Not On the same market, Or order includes both edits And placement
        ' INVALID_MARKET_ID       Market doesn't exist
        ' PERMISSION_DENIED       Business rules do Not allow order to be placed. You are either attempting to place the order using a Delayed Application Key Or from a restricted jurisdiction (i.e. USA)
        ' DUPLICATE_BETIDS        duplicate bet ids found
        ' NO_ACTION_REQUIRED      Order hasn't been passed to matcher as system detected there will be no state change
        ' SERVICE_UNAVAILABLE     The requested service Is unavailable
        ' REJECTED_BY_REGULATOR   The regulator rejected the order. On the Italian Exchange this error will occur if more than 50 bets are sent in a single placeOrders request.
        ' NO_CHASING              A specific error code that relates to Spanish Exchange markets only which indicates that the bet placed contravenes the Spanish regulatory rules relating to loss chasing.
        ' REGULATOR_IS_NOT_AVAILABLE  The underlying regulator service Is Not available.
        ' TOO_MANY_INSTRUCTIONS   The amount of orders exceeded the maximum amount allowed to be executed

        If status = "SUCCESS" Then
            ' Continue
        Else
            sendEmailToText("Match: " + sel.betfairEventName + " placeOrder has failed....please look at logs.")
        End If
    End Sub

    Private Sub nudSettingsSelectionRefresh_ValueChanged(sender As Object, e As EventArgs) Handles nudSettingsSelectionRefresh.ValueChanged
        timerRefreshSelections.Interval = nudSettingsSelectionRefresh.Value
    End Sub

    Private Sub btnCancelOrders_Click(sender As Object, e As EventArgs) Handles btnCancelOrders.Click
        sel1.cancelAllOrdersOverUnder15_Order()

    End Sub

    Private Sub btnCancelOrdersCS_Click(sender As Object, e As EventArgs) Handles btnCancelOrdersCS.Click
        sel1.cancelAllOrdersCorrectScore_Order()
    End Sub
End Class
