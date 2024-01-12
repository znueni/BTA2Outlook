Imports System.Runtime.InteropServices
Imports iTextSharp.text.pdf, System.Text.RegularExpressions, Microsoft.Office.Interop, Org.BouncyCastle.Crypto.Engines

Public Class BTA2OutlookApplication
    Public Shared Sub Main(args() As String)
        Application.EnableVisualStyles()
        If args.Length = 0 Then
            MsgBox("Click OK and then browse and select the ITN travel iterny PDF document.")
            Dim f As New OpenFileDialog() With {.Filter = "PDF|*.pdf|*.*|*.*", .FileName = "ITN_*.pdf"}
            If f.ShowDialog() = DialogResult.OK Then
                args = New String() {f.FileName}
            Else
                Application.Run(New BTA2OutlookApplication)
                Exit Sub
            End If
        End If
        Dim text As String = ExtractAllPDFText(args(0))
        Application.Run(New BTA2OutlookApplication(text, args(0)))
    End Sub

    Const PAGESEPERATOR As String = "**PAGE**"
    Private Shared Function ExtractAllPDFText(pdffilename As String) As String
        Dim reader As New iTextSharp.text.pdf.PdfReader(pdffilename)
        Dim text As New System.Text.StringBuilder()
        For i As Integer = 1 To reader.NumberOfPages
            text.AppendLine(iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, i))
            text.AppendLine(PAGESEPERATOR)
        Next
        Return text.ToString
    End Function

    Public Sub New(PDFText As String, filename As String)
        InitializeComponent()
        ITNEvents = Parse(PDFText)
        RefreshListview()
        Me.Text &= " - " & IO.Path.GetFileName(filename)
        originalPDFfilename = filename
    End Sub
    Public Sub New()
        Application.EnableVisualStyles()
        InitializeComponent()
    End Sub

    Dim ITNEvents As List(Of ITNEvent), originalPDFfilename As String

    Public Function Parse(Text As String) As List(Of ITNEvent)
        Dim list As New List(Of ITNEvent)
        Dim lines As String() = Text.Split(New Char() {vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)

        'TU 20.09.2022 WestCord Hotel Delft, LM DELFT
        'WE 21.09.2022 1 standard single room

        'FR 23.09.2022 17:15 Recommended time of arrival at the airport Bruxelles
        'FR 23.09.2022 20:15 Departure : BRU Bruxelles

        'TH 30.03.2023 18:00 Recommended time of arrival at the airport Copenhagen
        'TH 30.03.2023 20:00 Departure : CPH Copenhagen

        'TU 01.11.2022 Holiday Inn Exp Plaza Saldanha, LISBON
        'FR 04.11.2022 1 standard single room
        Dim rxPhone As New Regex("[^0-9()+]*")

        Dim rxDateline As New Regex("\A\w{2}\s(\d{2}\.[01]\d\.20\d\d)\s")
        Dim rxNonDepatureDateline As New Regex("\A\w{2}\s\d{2}\.[01]\d\.20\d\d\s\D")
        Dim rxDepature As New Regex("\A\w{2}\s(\d{2}\.[01]\d\.20\d\d\s\d\d:\d\d)\sDeparture\s*:\s*(.*)$")

        Dim rxCarPickup As New Regex("\A\w{2}\s(\d{2}\.[01]\d\.20\d\d\s\d\d:\d\d)\sPick up\s*:\s*(.*)$")
        Dim rxCarDropOff As New Regex("\A\w{2}\s(\d{2}\.[01]\d\.20\d\d\s\d\d:\d\d)\sDrop off\s*:\s*(.*)$")

        '\A\w{2}\s\d{2}\.[01]\d\.20\d\d\s
        '(\A\w{2}\s\d{2}\.[01]\d\.20\d\d\s)(.*)$

        Dim rxFlightArrival As New Regex("(\d\d:\d\d)\sArrival\s*:\s*(.*)$")
        Dim rxFlightArrivalD As New Regex("(\d\d.\d\d.20\d\d)\s*(\d\d:\d\d)\sArrival\s*:\s*(.*)$")

        Dim CaptureFlight As Boolean = False, CaptureHotel As Boolean = False, f As ITNFlight = Nothing, h As ITNHotel = Nothing, c As ITNCar
        For i As Integer = 0 To lines.Length - 1
            If rxDepature.IsMatch(lines(i)) Then
                f = New ITNFlight() With {.Rx = rxDepature.Match(lines(i))}
                f.DateStart = Date.Parse(f.Rx.Groups(1).Value)
                f.LocFrom = f.Rx.Groups(2).Value
                list.Add(f)
                i += 1
                If lines(i).Contains("Arrival") Then
                    If rxFlightArrival.IsMatch(lines(i)) Then
                        f.Rx = rxFlightArrival.Match(lines(i))
                        f.LocTo = f.Rx.Groups(2).Value
                        f.DateEnd = Date.Parse(f.DateStart.Date & " " & f.Rx.Groups(1).Value)
                        CaptureFlight = True
                    ElseIf rxFlightArrivalD.IsMatch(lines(i)) Then
                        f.Rx = rxFlightArrivalD.Match(lines(i))
                        f.LocTo = f.Rx.Groups(3).Value
                        f.DateEnd = Date.Parse(f.Rx.Groups(1).Value & " " & f.Rx.Groups(2).Value)
                        CaptureFlight = True
                    End If
                End If

            ElseIf lines(i).TrimEnd().EndsWith("Car rental") Then
                c = New ITNCar() With {.Name = lines(i)}
                list.Add(c)
                Do
                    i += 1
                    If rxCarPickup.IsMatch(lines(i)) Then
                        c.Rx = rxCarPickup.Match(lines(i))
                        c.DateStart = Date.Parse(c.Rx.Groups(1).Value)
                        c.Location = c.Rx.Groups(2).Value
                        c.EventType = ITNCar.CarEventType.Pickup
                        Exit Do
                    ElseIf rxCarDropOff.IsMatch(lines(i)) Then
                        c.Rx = rxCarDropOff.Match(lines(i))
                        c.DateStart = Date.Parse(c.Rx.Groups(1).Value)
                        c.Location = c.Rx.Groups(2).Value
                        c.EventType = ITNCar.CarEventType.Dropoff
                        Exit Do
                    End If
                Loop


            ElseIf rxDateline.IsMatch(lines(i)) Then
                    CaptureFlight = False
                    CaptureHotel = False
                    If rxDateline.IsMatch(lines(i + 1)) AndAlso lines(i + 1).Contains("room") Then
                        h = New ITNHotel With {.Rx = rxDateline.Match(lines(i))}
                        list.Add(h)
                        CaptureHotel = True
                        h.Name = lines(i).Substring(h.Rx.Length).Trim()
                        h.DateStart = Date.Parse(h.Rx.Groups(1).Value)
                        i += 1
                        h.Rx = rxDateline.Match(lines(i))
                        h.DateEnd = Date.Parse(h.Rx.Groups(1).Value)
                    End If

                ElseIf CaptureHotel AndAlso lines(i).Contains(":") Then
                    If lines(i).StartsWith("rate per night/room") Then
                        h.Rate = lines(i).Split(":"c)(1).Trim
                    ElseIf lines(i).StartsWith("Conf. N°") Then
                        h.Confirmation = lines(i).Split(":"c)(1).Trim
                    ElseIf lines(i).StartsWith("Address") Then
                        h.Address = lines(i).Split(":"c)(1).Trim & ", " & lines(i + 1).Trim()
                        i += 2
                        If Not lines(i).StartsWith("Tel") Then
                            h.Address &= vbCrLf & lines(i)
                        Else
                            h.Phone = rxPhone.Replace(lines(i).Trim(), String.Empty)
                            If h.Phone.StartsWith("+") AndAlso h.Phone.Contains("(0)") Then h.Phone = h.Phone.Replace("(0)", "")
                        End If
                        'If lines(i).StartsWith("Tel") Then
                        '    h.Address &= vbCrLf & lines(i)
                        'ElseIf lines(i + 1).StartsWith("Tel") Then
                        '    h.Address &= vbCrLf & lines(i) & vbCrLf & lines(i + 1)
                        'End If
                    End If

                ElseIf CaptureFlight AndAlso lines(i).Contains(":") Then
                    If lines(i).StartsWith("Flight N°") Then
                    f.Flightno = lines(i).Split(":"c)(1).Trim
                ElseIf lines(i).StartsWith("Reservation Number") Then
                    f.Reservation = lines(i).Split(":"c)(1).Trim
                    If lines(i + 1).StartsWith("(Your reference ") Then f.Reservation &= " " & lines(i + 1).Trim()
                ElseIf lines(i).StartsWith("Seat") Then
                    f.Seat = lines(i).Split(":"c)(1).Trim
                ElseIf lines(i).StartsWith("Baggage") Then
                    f.Baggage = lines(i).Split(":"c)(1).Trim
                    'ElseIf lines(i).StartsWith("") Then
                End If




            End If
        Next
        Return list
    End Function


    Private Sub RefreshListview()
        lst.Items.Clear()
        For Each e As ITNEvent In ITNEvents
            Dim li As ListViewItem = e.Listviewitem
            lst.Items.Add(li)
            li.Tag = e
            If TypeOf e Is ITNFlight Then
                Dim f As ITNFlight = e
                li.ImageIndex = 0
                li.Group = lst.Groups(0)
            ElseIf TypeOf e Is ITNHotel Then
                Dim h As ITNHotel = e
                li.ImageIndex = 1
                li.Group = lst.Groups(1)
            ElseIf TypeOf e Is ITNCar Then
                Dim c As ITNCar = e
                li.ImageIndex = 2
                li.Group = lst.Groups(2)
            End If
            li.Checked = True
        Next
        lst.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) 'Handles Button1.Click
        ITNEvents = Parse(Clipboard.GetText())
        RefreshListview()
    End Sub

    Private Sub btnCreateEventsInOutlook_Click(sender As Object, e As EventArgs) Handles btnCreateEventsInOutlook.Click
        lblWait.Visible = True
        Try
            Dim App As New Outlook.Application
            For Each li As ListViewItem In lst.CheckedItems
                Try
                    Dim Cal As Outlook.AppointmentItem = App.CreateItem(Outlook.OlItemType.olAppointmentItem)
                    If TypeOf li.Tag Is ITNFlight Then
                        Dim f As ITNFlight = li.Tag
                        Cal.Subject = $"✈ {f.Flightno} to {f.LocTo}"
                        Cal.Location = f.LocFrom
                        Cal.Body = $"Flight {f.Flightno}{vbCrLf}{vbCrLf}Reservation {f.Reservation}"
                        Cal.Start = f.DateStart
                        Cal.End = f.DateEnd
                        Cal.BusyStatus = Outlook.OlBusyStatus.olOutOfOffice
                        'Wir adden zudem noch einen reminder 24h vor dem flug, damit man einchecken kann:
                        Cal.ReminderOverrideDefault = True
                        Cal.ReminderMinutesBeforeStart = 60 * 24
                        Cal.ReminderPlaySound = True

                        If chkAddBoardingblocker.Checked Then
                            Cal.Save()
                            Marshal.ReleaseComObject(Cal)
                            Cal = App.CreateItem(Outlook.OlItemType.olAppointmentItem)
                            Cal.Subject = $"✈ Boarding/Security"
                            Cal.Body = String.Empty
                            Cal.Location = f.LocFrom
                            Cal.Start = f.DateStart.AddHours(-1)
                            Cal.End = f.DateStart
                            Cal.BusyStatus = Outlook.OlBusyStatus.olOutOfOffice
                            Cal.ReminderSet = False
                        End If

                    ElseIf TypeOf li.Tag Is ITNHotel Then
                        Dim h As ITNHotel = li.Tag
                        Cal.Subject = $"Hotel {h.Name}"
                        Cal.Location = h.Address
                        Cal.Start = h.DateStart.Date.AddHours(20)
                        Cal.End = h.DateEnd.Date.AddHours(11)
                        'Dim nights As Integer = Math.Ceiling(Cal.End.Subtract(Cal.Start).TotalDays)
                        Cal.Body = $"{h.Name}{vbCrLf}{h.Address}{vbCrLf}📞 {h.Phone}{vbCrLf}{vbCrLf}Daily rate {h.Rate}{vbCrLf}Conf # {h.Confirmation}"
                        Cal.AllDayEvent = True
                        Cal.BusyStatus = Outlook.OlBusyStatus.olWorkingElsewhere

                    ElseIf TypeOf li.Tag Is ITNCar Then
                        Dim c As ITNCar = li.Tag
                        Cal.Subject = c.EventType.ToString() & " - " & c.Name
                        Cal.Location = c.Location
                        Cal.BusyStatus = Outlook.OlBusyStatus.olOutOfOffice
                        Cal.Start = c.DateStart
                        Cal.End = c.DateStart.AddMinutes(30)
                        Cal.AllDayEvent = False

                    End If
                    Cal.Save()
                    Marshal.ReleaseComObject(Cal)
                Catch ex As Exception
                    MsgBox($"{ex.GetType().Name}: {ex.Message} when creating the appointment on {li.Text}", MsgBoxStyle.Critical)
                End Try

            Next
            Marshal.ReleaseComObject(App)
            MsgBox("Calendar items created. Go to your Outlook calendar to view and ensure.", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox($"{ex.GetType().Name}: {ex.Message} during communication with Outlook", MsgBoxStyle.Critical)
        End Try
        lblWait.Visible = False
    End Sub

    Private Sub btnSaveICSs_Click(sender As Object, e As EventArgs) Handles btnSaveICSs.Click
        UseWaitCursor = False : Enabled = True : Application.DoEvents()
        Dim icses As New List(Of ICSEvent)
        For Each li As ListViewItem In lst.CheckedItems
            Try
                If TypeOf li.Tag Is ITNFlight Then
                    Dim f As ITNFlight = li.Tag
                    icses.Add(New ICSEvent() With {
                        .subject = $"✈ {f.Flightno} to {f.LocTo}",
                        .location = f.LocFrom,
                        .description = $"Flight {f.Flightno}{vbCrLf}{vbCrLf}Reservation {f.Reservation}",
                        .dtStart = f.DateStart,
                        .dtEnd = f.DateEnd,
                        .Reminder = True, .ReminderMinutesBeforeEvent = 60 * 24,
                        .Busy = ICSEvent.BusyStatusTypes.OOF,
                        .IsAllDay = False, .IsTransparent = False
                    })

                    If chkAddBoardingblocker.Checked Then
                        icses.Add(New ICSEvent() With {
                        .subject = $"✈ Boarding/Security",
                        .location = f.LocFrom,
                        .description = "",
                        .dtStart = f.DateStart.AddHours(-1),
                        .dtEnd = f.DateStart,
                        .Reminder = False,
                        .Busy = ICSEvent.BusyStatusTypes.OOF,
                        .IsAllDay = False, .IsTransparent = False
                    })

                    End If

                ElseIf TypeOf li.Tag Is ITNHotel Then
                    Dim h As ITNHotel = li.Tag
                    icses.Add(New ICSEvent() With {
                        .subject = $"Hotel {h.Name}",
                        .location = h.Address,
                        .description = $"{h.Name}{vbCrLf}{h.Address}{vbCrLf}📞 {h.Phone}{vbCrLf}{vbCrLf}Daily rate {h.Rate}{vbCrLf}Conf # {h.Confirmation}",
                        .dtStart = h.DateStart.Date,
                        .dtEnd = h.DateEnd.Date.AddDays(1),
                        .Reminder = False,
                        .Busy = ICSEvent.BusyStatusTypes.TENTATIVE,
                        .IsAllDay = True, .IsTransparent = True
                    })

                ElseIf TypeOf li.Tag Is ITNCar Then
                    Dim c As ITNCar = li.Tag
                    icses.Add(New ICSEvent() With {
                       .subject = c.EventType.ToString() & " - " & c.Name,
                       .location = c.Location,
                       .dtStart = c.DateStart,
                       .dtEnd = c.DateStart.AddMinutes(30),
                       .Reminder = True, .ReminderMinutesBeforeEvent = 30,
                       .Busy = ICSEvent.BusyStatusTypes.OOF,
                       .IsAllDay = False, .IsTransparent = False
                   })

                End If
            Catch ex As Exception
                MsgBox($"{ex.GetType().Name}: {ex.Message} when creating the appointment on {li.Text}", MsgBoxStyle.Critical)
                UseWaitCursor = False : Enabled = True
            End Try
        Next

        Dim d As New SaveFileDialog() With {.Filter = "ics|*.ics|all|*.*", .DefaultExt = "*.ics", .FileName = IO.Path.GetFileNameWithoutExtension(originalPDFfilename)}
        If d.ShowDialog() = DialogResult.OK Then
            Dim fs As New IO.FileStream(d.FileName, IO.FileMode.Create)
            Dim sw As New IO.StreamWriter(fs, System.Text.Encoding.UTF8)
            sw.Write(ICSEvent.ToICSCalendar(icses.ToArray()))
            sw.Close()
        End If

        UseWaitCursor = False : Enabled = True
    End Sub
End Class

Public MustInherit Class ITNEvent
    Public Rx As Match
    Public DateStart, DateEnd As Date
    Public MustOverride Function Listviewitem() As ListViewItem
    Public Overridable ReadOnly Property Duration As TimeSpan
        Get
            Return DateEnd.Subtract(DateStart)
        End Get
    End Property
    'Public Overridable Function Label() As TableLayoutPanel
    '    'Return New Label With {
    '    '    .ImageAlign = ContentAlignment.MiddleLeft,
    '    '    .TextAlign = ContentAlignment.MiddleLeft,
    '    '    .AutoEllipsis = True,
    '    '    .AutoSize = False,
    '    '    .Height = 200}
    '    Dim t As New TableLayoutPanel() With {
    '        .RowCount = 1,
    '        .ColumnCount = 4,
    '        .Tag = Me,
    '        .BackColor = Color.White,
    '        .Width = 500,
    '        .AutoSize = True
    '       }
    '    Dim img As New PictureBox
    '    t.Controls.Add(img)
    '    t.SetCellPosition(img, New TableLayoutPanelCellPosition(0, 0))
    '    Dim lbl As New Label
    '    lbl.Text = DateStart.ToString
    '    t.Controls.Add(lbl)
    '    t.SetCellPosition(lbl, New TableLayoutPanelCellPosition(1, 0))
    '    Dim lbl2 As New Label
    '    lbl2.Text = DateEnd.ToString
    '    t.Controls.Add(lbl2)
    '    t.SetCellPosition(lbl2, New TableLayoutPanelCellPosition(2, 0))
    '    Return t
    'End Function

End Class
Public Class ITNFlight
    Inherits ITNEvent
    Public LocFrom, LocTo As String, Flightno As String, Reservation As String, Seat As String, Baggage As String
    'Public Overrides Function Label() As TableLayoutPanel
    '    Dim t As TableLayoutPanel = MyBase.Label()
    '    DirectCast(t.Controls(0), PictureBox).Image = My.Resources.plane
    '    Dim l1 As New Label With {.Text = $"Flight from {LocFrom} to {LocTo}"}
    '    t.Controls.Add(l1)
    '    t.SetCellPosition(l1, New TableLayoutPanelCellPosition(3, 0))
    '    Return t
    'End Function

    Public Overrides Function Listviewitem() As ListViewItem
        Return New ListViewItem(New String() {DateStart.ToString, IIf(DateStart.Date = DateEnd.Date, DateEnd.ToShortTimeString(), DateEnd.ToString()), $"Flight {Flightno} from {LocFrom} to {LocTo} {vbCrLf}Seat {Seat}, Conf N° {Reservation}", Duration.ToString})
    End Function
End Class
Public Class ITNHotel
    Inherits ITNEvent
    Public Name, Confirmation, Rate, Address, Phone As String
    'Public Overrides Function Label() As TableLayoutPanel
    '    Dim t As TableLayoutPanel = MyBase.Label()
    '    DirectCast(t.Controls(0), PictureBox).Image = My.Resources.hotel
    '    Dim l1 As New Label With {.Text = $"Hotel in {Address} for {Rate}"}
    '    t.Controls.Add(l1)
    '    t.SetCellPosition(l1, New TableLayoutPanelCellPosition(3, 0))
    '    Return t
    'End Function

    Public Overrides Function Listviewitem() As ListViewItem
        Return New ListViewItem(New String() {DateStart.ToLongDateString, DateEnd.ToLongDateString, $"Hotel {Name} {vbCrLf}in {Address} for {Rate}/night {vbCrLf}Conf N°: {Confirmation}", Duration.TotalDays & " days"})
    End Function

    Public Overrides ReadOnly Property Duration As TimeSpan
        Get
            Return DateEnd.Date.Subtract(DateStart.Date)
        End Get
    End Property
End Class

Public Class ITNCar
    Inherits ITNEvent
    Public Location As String, EventType As CarEventType, Name As String
    Public Enum CarEventType
        Pickup
        Dropoff
    End Enum
    Public Overrides Function Listviewitem() As ListViewItem
        Return New ListViewItem(New String() {DateStart.ToString, DateEnd.ToString, $"{Name} {EventType.ToString()} in {Location}", ""})
    End Function
End Class

Public Class ICSEvent
    Public description As String, subject As String, location As String, dtStart As Date, dtEnd As Date, ReminderMinutesBeforeEvent As Integer, Busy As BusyStatusTypes, IsTransparent As Boolean, Reminder As Boolean
    Public IsAllDay As Boolean

    Public Shared Function ToICSCalendar(events As ICSEvent()) As String
        Const ENDE As String = "END:VCALENDAR"
        Const HEAD As String = "BEGIN:VCALENDAR
PRODID:-//Znueni//BTA2Outlook//EN
VERSION:2.0
METHOD:PUBLISH
"
        Dim ics As New System.Text.StringBuilder(HEAD)
        For Each e As ICSEvent In events
            ics.AppendLine(e.ToString())
        Next
        ics.Append(ENDE)
        Return ics.ToString()
    End Function

    Public Shared Function ToICSString(s As String) As String
        Return s?.Replace("\", "\\").Replace(",", "\,").Replace(vbCrLf, "\n")
    End Function
    Public Shared Function ToICSDate(d As Date, AllDay As Boolean) As String 'including the : because it can have a ;VALUE=DATE:
        If AllDay Then Return ";VALUE=DATE:" & d.ToString("yyyyMMdd") Else Return ":" & d.ToString("yyyyMMdd") & "T" & d.ToString("HHmm") & "00"
    End Function

    Public Overrides Function ToString() As String
        Dim r As String =
            $"BEGIN:VALARM
TRIGGER:-PT{ReminderMinutesBeforeEvent}M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM
".Trim()

        Dim ics As String = $"
BEGIN:VEVENT
CLASS:PUBLIC
DESCRIPTION:{ToICSString(description)}
DTEND{ToICSDate(dtEnd, IsAllDay)}
DTSTAMP{ToICSDate(Date.Now, False)}
DTSTART{ToICSDate(dtStart, IsAllDay)}
LOCATION:{ToICSString(location)}
SEQUENCE:0
SUMMARY:{ToICSString(subject)}
TRANSP:{If(IsTransparent, "TRANSPARENT", "OPAQUE")}
X-MICROSOFT-CDO-BUSYSTATUS:{Busy.ToString()}
X-MICROSOFT-CDO-IMPORTANCE:1
X-MICROSOFT-DISALLOW-COUNTER:FALSE
X-MS-OLK-AUTOFILLLOCATION:FALSE
X-MS-OLK-CONFTYPE:0
".Trim()

        Const ENDE As String = "END:VEVENT"

        If Reminder Then Return ics & vbCrLf & r & vbCrLf & ENDE Else Return ics & vbCrLf & ENDE

    End Function

    Public Enum BusyStatusTypes
        BUSY
        FREE
        OOF
        TENTATIVE
    End Enum
End Class
