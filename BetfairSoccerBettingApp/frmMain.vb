Imports System.IO

Public Class frmMain

    Public sel1 As New Selection(1)
    Public sel2 As New Selection(2)
    Public sel3 As New Selection(3)
    Public sel4 As New Selection(4)


    Private intFileNumber As Integer = FreeFile()


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        marketCountriesEurope = New HashSet(Of String)({"FR", "DE", "IT", "ES", "PT", "NL", "GR"})

        ' Login
        Account.Login()

        ' Populate initial list of event data
        Dim BetfairClass1 As New BetfairClass()
        BetfairClass1.PollBetFairEvents(1, My.Settings.NumberOfUkEvents, marketCountriesUkOnly)
        Me.dgvEvents.DataSource = BetfairClass1.eventList
        BetfairClass1 = Nothing


        ' Refresh log on screen
        Dim SR As StreamReader
        SR = File.OpenText(gobjLogFileName)
        rtbLog.AppendText(SR.ReadToEnd)
        SR.Close()


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

    Private Sub btnsel1_Click(sender As Object, e As EventArgs) Handles btnSel1.Click

        Dim selectedRowCount As Integer =
        dgvEvents.Rows.GetRowCount(DataGridViewElementStates.Selected)

        If selectedRowCount > 0 Then

            sel1.selectionNumber = 1
            tbxSel1EventName.Text = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            grpSel1.Text = "Selection 1 - " + dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel1.betfairEventName = dgvEvents.SelectedRows(0).Cells(2).Value.ToString()
            sel1.betfairEventDateTime = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            tbxSel1EventDateTime.Text = dgvEvents.SelectedRows(0).Cells(5).Value.ToString()
            sel1.betfairEventId = dgvEvents.SelectedRows(0).Cells(1).Value.ToString()

            sel1.getAllMarketData()

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

    Private Sub Refreshsel1Info()

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

        ' Determine Score
        If sel1.betfairEventInplay = "0" Then
            tbxSel1Score.Text = "Not Started"
        Else
            If sel1.betfairCorrectScore00SelectionStatus = "ACTIVE" Then
                tbxSel1Score.Text = "0 - 0"
            ElseIf sel1.betfairCorrectScore01SelectionStatus = "ACTIVE" And sel1.betfairCorrectScore10SelectionStatus <> "ACTIVE" Then
                tbxSel1Score.Text = "0 - 1"
            ElseIf sel1.betfairCorrectScore10SelectionStatus = "ACTIVE" And sel1.betfairCorrectScore01SelectionStatus <> "ACTIVE" Then
                tbxSel1Score.Text = "1 - 0"
            Else
                tbxSel1Score.Text = "Score Over 1.5"
            End If
        End If


        ' Populate unmatched bets
        tbxSel1UnmatchedCorrectScore.Text = sel1.betfairCorrectScoreUnmathedBets
        tbxSel1UnmatchedUnderOver15.Text = sel1.betfairUnderOver15UnmathedBets


        ' Update refresh date/time
        tbxSel1RefreshLight.BackColor = Color.DarkGreen
        tbxSel1RefreshLight.ForeColor = Color.White
        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        tbxSel1RefreshLight.Text = time.ToString(format)

        ' Update the Inplay datetime
        Dim eventDateTime As DateTime = DateTime.Parse(tbxSel1EventDateTime.Text)
        Dim timeToStart As TimeSpan = DateTime.Now.Subtract(eventDateTime)
        Dim formatTime As String = "##,##0.00"
        tbxSel1InplayTime.Text = timeToStart.TotalMinutes.ToString(formatTime)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' Logout
        Account.Logout()

        Application.Exit()

    End Sub

    Private Sub timerRefreshSelections_Tick(sender As Object, e As EventArgs) Handles timerRefreshSelections.Tick

        If tbxSel1EventName.Text <> "" Then

            Refreshsel1Info()

        Else

            tbxSel1RefreshLight.BackColor = Color.White
            tbxSel1RefreshLight.ForeColor = Color.Black
            tbxSel1RefreshLight.Text = ""

        End If

    End Sub

    Private Sub btnSel1AutoBetOn_Click(sender As Object, e As EventArgs) Handles btnSel1AutoBetOn.Click

        If btnSel1AutoBetOn.Text = "Auto Bet On" Then

            If tbxSel1EventName.Text <> "" Then

                If MsgBox("Please confirm you want to switch Automatic Betting on?", MsgBoxStyle.YesNo, "Automatic Betting Confirmation") = MsgBoxResult.Yes Then

                    ' Enable Auto Bet timer
                    timerSel1AutoBet.Enabled = True

                    btnSel1AutoBetOn.Text = "Auto Bet Off"
                    btnSel1AutoBetOn.BackColor = Color.LightSalmon

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

        ' Place order on 0-0 market
        'sel1.placeCorrectScore_00_Order()

    End Sub


End Class
