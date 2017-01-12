Imports System.IO

Public Class frmMain

    Public sel1 As New Selection(1)
    Public sel2 As New Selection(2)
    Public sel3 As New Selection(3)
    Public sel4 As New Selection(4)


    Private intFileNumber As Integer = FreeFile()


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Set global log to this forms rtb
        grtbLog = Me.rtbLog

        '
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

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' Logout
        Account.Logout()

        Application.Exit()

    End Sub

    Private Sub timerRefreshSelections_Tick(sender As Object, e As EventArgs) Handles timerRefreshSelections.Tick

        ' Clean log rich textbox
        If rtbLog.Lines.Count > 1000 Then
            rtbLog.Clear()
        End If

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

    Private Sub btnSel1AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel1AutoBetOn.Click

        If btnSel1AutoBetOn.Text = "Auto Bet On" Then

            If tbxSel1EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Set the interval
                    timerSel1AutoBet.Interval = nudSettingsAutoBetRefresh.Value

                    ' Enable Auto Bet timer
                    timerSel1AutoBet.Enabled = True

                    btnSel1AutoBetOn.Text = "Auto Bet Off"
                    btnSel1AutoBetOn.BackColor = Color.LightSalmon

                    ' Write to log
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Auto Bet for Sel1 has been switched on.", EventLogEntryType.Information)

                    ' Call tick
                    timerSel1AutoBet_Tick(sender, e)

                End If


            End If
        Else

            ' Disable Auto Bet timer
            timerSel1AutoBet.Enabled = False

            ' Switch off
            btnSel1AutoBetOn.Text = "Auto Bet On"
            btnSel1AutoBetOn.BackColor = Color.LightGreen


        End If

    End Sub

    Private Sub timerSel1AutoBet_Tick(sender As Object, e As EventArgs) Handles timerSel1AutoBet.Tick

        '
        ' Do we need any stakes, check the status of bets on this Event 
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


        ' If any of the bets are missing then continue
        If btnSel1ProfitStatus00.Text = "" Or btnSel1ProfitStatus10.Text = "" Or btnSel1ProfitStatus01.Text = "" Or btnSel1ProfitStatusUnder15.Text = "" Or btnSel1ProfitStatusOver15.Text = "" Then
            ' Continue
        Else
            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Auto Bet for Sel1 - No bets required (or some of the profit fields are null)", EventLogEntryType.Information)

        End If

        '
        ' Check the status of the Event, must be Inplay
        '
        If sel1.betfairEventInplay = "True" Then
            ' Continue
        Else
            ' Write to log
            gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Auto Bet for Sel1 - Event not in play", EventLogEntryType.Information)
        End If


        ' 
        ' Look for Correct Score 0-0 bet
        '
        If btnSel1ProfitStatus00.Text = "" Then

            If sel1.betfairCorrectScoreMarketStatus = "OPEN" Then

                ' Check in first half
                If CDbl(tbxSel1InplayTime.Text) > +0 And CDbl(tbxSel1InplayTime.Text) < +45 Then

                    If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00BackOdds) Then
                        If CDbl(sel1.betfairCorrectScore00BackOdds) > nudSettingsCS00LowerPrice.Value And CDbl(sel1.betfairCorrectScore00BackOdds) < nudSettingsCS00UpperPrice.Value Then

                            If sel1.betfairCorrectScore00BackOdds <= nudSettingsCS00TargetPrice.Value Then

                                If Not String.IsNullOrEmpty(sel1.betfairCorrectScore00Orders) Then
                                    If CDbl(sel1.betfairCorrectScore00Orders) > 1 Then

                                        'Unmatched Orders
                                        ' Write to log
                                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Auto Bet for Sel1 - I WOULD CANCEL UNMATCHED ORDERS", EventLogEntryType.Information)
                                    End If
                                End If

                                ' Write to log
                                gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Auto Bet for Sel1 - I WOULD PLACE A BET NOW", EventLogEntryType.Information)

                            End If
                        End If
                    End If
                End If

            End If
        End If

        ' Place order on 0-0 market
        'sel1.placeCorrectScore_00_Order()

    End Sub

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

            ' Enable Auto Bet Button
            btnSel1AutoBetOn.Enabled = True

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

            ' Enable Auto Bet Button
            btnSel2AutoBetOn.Enabled = True

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

            ' Enable Auto Bet Button
            btnSel3AutoBetOn.Enabled = True

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

            ' Enable Auto Bet Button
            btnSel4AutoBetOn.Enabled = True

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
            tbxSel1InplayStatus.BackColor = Color.Yellow
        Else
            tbxSel1InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel1Score.Text

        ' Get latest score
        tbxSel1Score.Text = sel1.betfairGoalsScored

        ' Detect score change
        If strPreviousScore = tbxSel1Score.Text Then
            ' Same score
        Else
            ' If first time through...ignore
            If tbxSel1Score.Text <> "" Then
                ' Goal scored since last tick
                If tbxSel1Goal1.Text = "" Then
                    tbxSel1Goal1.Text = tbxSel1InplayTime.Text.ToString
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel1.Text + ", Goal 1 scored at: " + tbxSel1InplayTime.Text.ToString, EventLogEntryType.Information)

                Else
                    If tbxSel1Goal2.Text = "" Then
                        tbxSel1Goal2.Text = tbxSel1InplayTime.Text.ToString
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel1.Text + ", Goal 2 scored at: " + tbxSel1InplayTime.Text.ToString, EventLogEntryType.Information)
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

        If sel1.betfairOver15IfWinProfit IsNot Nothing Then
            If Integer.Parse(sel1.betfairOver15IfWinProfit) >= 0 Then
                tbxSel1IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel1IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel1.betfairUnder15IfWinProfit IsNot Nothing Then
            If Integer.Parse(sel1.betfairUnder15IfWinProfit) >= 0 Then
                tbxSel1IUnder15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel1IUnder15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel1.betfairCorrectScore00IfWinProfit IsNot Nothing Then
            If Double.Parse(sel1.betfairCorrectScore00IfWinProfit) >= 0 Then
                tbxSel1CorrectScore00IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore00IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel1.betfairCorrectScore10IfWinProfit IsNot Nothing Then
            If Double.Parse(sel1.betfairCorrectScore10IfWinProfit) >= 0 Then
                tbxSel1CorrectScore10IfWin.ForeColor = Color.DarkGreen
            Else
                tbxSel1CorrectScore10IfWin.ForeColor = Color.OrangeRed
            End If
        End If
        If sel1.betfairCorrectScore01IfWinProfit IsNot Nothing Then
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
            tbxSel2InplayStatus.BackColor = Color.Yellow
        Else
            tbxSel2InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel2Score.Text

        ' Get latest score
        tbxSel2Score.Text = sel2.betfairGoalsScored

        ' Detect score change
        If strPreviousScore = tbxSel2Score.Text Then
            ' Same score
        Else
            ' If first time through...ignore
            If tbxSel2Score.Text <> "" Then
                ' Goal scored since last tick
                If tbxSel2Goal1.Text = "" Then
                    tbxSel2Goal1.Text = tbxSel2InplayTime.Text.ToString
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 1 scored at: " + tbxSel2InplayTime.Text.ToString, EventLogEntryType.Information)
                Else
                    If tbxSel2Goal2.Text = "" Then
                        tbxSel2Goal2.Text = tbxSel2InplayTime.Text.ToString
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 2 scored at: " + tbxSel2InplayTime.Text.ToString, EventLogEntryType.Information)
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
            If Integer.Parse(sel2.betfairOver15IfWinProfit) >= 0 Then
                tbxSel2IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel2IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel2.betfairUnder15IfWinProfit IsNot Nothing Then
            If Integer.Parse(sel2.betfairUnder15IfWinProfit) >= 0 Then
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
            tbxSel3InplayStatus.BackColor = Color.Yellow
        Else
            tbxSel3InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel3Score.Text

        ' Get latest score
        tbxSel3Score.Text = sel3.betfairGoalsScored

        ' Detect score change
        If strPreviousScore = tbxSel3Score.Text Then
            ' Same score
        Else
            ' If first time through...ignore
            If tbxSel3Score.Text <> "" Then
                ' Goal scored since last tick
                If tbxSel3Goal1.Text = "" Then
                    tbxSel3Goal1.Text = tbxSel3InplayTime.Text.ToString
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel3.Text + ", Goal 1 scored at: " + tbxSel3InplayTime.Text.ToString, EventLogEntryType.Information)
                Else
                    If tbxSel3Goal2.Text = "" Then
                        tbxSel3Goal2.Text = tbxSel3InplayTime.Text.ToString
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel3.Text + ", Goal 2 scored at: " + tbxSel3InplayTime.Text.ToString, EventLogEntryType.Information)
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
            If Integer.Parse(sel3.betfairOver15IfWinProfit) >= 0 Then
                tbxSel3IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel3IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel3.betfairUnder15IfWinProfit IsNot Nothing Then
            If Integer.Parse(sel3.betfairUnder15IfWinProfit) >= 0 Then
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
            tbxSel4InplayStatus.BackColor = Color.Yellow
        Else
            tbxSel4InplayStatus.BackColor = Color.Green
        End If


        ' Determine change of goals
        Dim strPreviousScore As String
        strPreviousScore = tbxSel4Score.Text

        ' Get latest score
        tbxSel4Score.Text = sel4.betfairGoalsScored

        ' Detect score change
        If strPreviousScore = tbxSel4Score.Text Then
            ' Same score
        Else
            ' If first time through...ignore
            If tbxSel4Score.Text <> "" Then
                ' Goal scored since last tick
                If tbxSel4Goal1.Text = "" Then
                    tbxSel4Goal1.Text = tbxSel4InplayTime.Text.ToString
                    gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 1 scored at: " + tbxSel4InplayTime.Text.ToString, EventLogEntryType.Information)
                Else
                    If tbxSel4Goal2.Text = "" Then
                        tbxSel4Goal2.Text = tbxSel4InplayTime.Text.ToString
                        gobjEvent.WriteToEventLog("BetfairSoccerBettingApp : Selection: " + grpSel2.Text + ", Goal 2 scored at: " + tbxSel4InplayTime.Text.ToString, EventLogEntryType.Information)
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
            If Integer.Parse(sel4.betfairOver15IfWinProfit) >= 0 Then
                tbxSel4IOver15fWinProfit.ForeColor = Color.DarkGreen
            Else
                tbxSel4IOver15fWinProfit.ForeColor = Color.OrangeRed
            End If
        End If
        If sel4.betfairUnder15IfWinProfit IsNot Nothing Then
            If Integer.Parse(sel4.betfairUnder15IfWinProfit) >= 0 Then
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

End Class
