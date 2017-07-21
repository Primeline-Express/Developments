Option Compare Text
Option Explicit On
Imports System.Net.Mail
Imports System.Configuration

Public Class Form1

    Const strDefaultFolder As String = "\\APOLLO1\Azyra\"
    Const strPODsFolder As String = "Derry Morgan EDI\"
    Const strEDIFolder As String = "EDI Messages\Live EDI\"
    Const strMainEDIsFolder As String = "Express XML EDI Files\"
    Const strTestFolder As String = "TEST\"
    Const strAttachmentLoc As String = "\\APOLLO1\Azyra\EDI Messages\Live EDI\Express XML EDI Files\XML Details\XML Details.docx"
    Const strHead1 = "Customer Name"
    Const strHead2 = "Time"
    Const strHead3 = "Received"
    Const strHead4 = "Unsuccessful"
    Const strHead5 = "Successful"
    Public strEDIs As String = strDefaultFolder & strEDIFolder & strMainEDIsFolder
    Public strTimco As String = strDefaultFolder & strEDIFolder & "TI Midwood\"
    Public strPODs As String = strDefaultFolder & strPODsFolder
    Public strEagle As String = strDefaultFolder & strEDIFolder & "Primeline Eagle\"
    Public strBulkEvents As String = strEDIs & "PLE Events\"
    Public listDetails As New List(Of String)()
    Public listRecipients As New List(Of String)()
    Const fldCustName As Long = 1
    Const fldFileName As Long = 2
    Const fldFolder As Long = 3
    Const fldRecCnt As Long = 4
    Const fldUnSuccCnt As Long = 5
    Const fldSuccCnt As Long = 6 'Muppet
    Const fldTime As Long = 7
    Const lngFields As Long = 7
    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Date.Now.Hour > 13 Then
            Call ArraySetup()
        Else
            If Date.Now.Hour < 9 Then
                Call ArchiveFiles()
                Application.Exit()
            End If
        End If
    End Sub

    Public Sub LiveButton_Click(sender As Object, e As EventArgs) Handles LiveButton.Click
        Call ArraySetup()
    End Sub

    Public Sub GetUnSuccessfulFiles()
        Dim Mail As New MailMessage
        Dim SMTP As New SmtpClient("smtp.gmail.com")
        Dim lngCount As Long
        Dim lngCount3 As Long
        Dim lngTotal As Long
        Dim i As Long
        Dim strEmailAddress As String = ConfigurationManager.AppSettings("strEmailAddress")
        Dim strEmailPassword As String = ConfigurationManager.AppSettings("strEmailPassword")
        Dim MyFiles1() As String = IO.Directory.GetFiles(strEDIs & "UnSuccessful\")
        lngCount = MyFiles1.Length
        Dim MyFiles3() As String = IO.Directory.GetFiles(strEDIs & strTestFolder & "UnSuccessful\")
        lngCount3 = MyFiles3.Length
        lngTotal = lngCount + lngCount3

        If lngTotal > 0 Then
            Mail.Subject = "XML File Monitor Unsuccessful Files"
            Mail.From = New MailAddress("reporting@primelineexpress.co.uk", "Primeline Express Reporting", System.Text.Encoding.UTF8)
            SMTP.Credentials = New System.Net.NetworkCredential(strEmailAddress, strEmailPassword) '<-- Password Here
            Mail.To.Add("reporting@primelineexpress.co.uk")
            If lngCount > 0 Then
                For i = 0 To lngCount - 1
                    Dim attachment As System.Net.Mail.Attachment
                    attachment = New System.Net.Mail.Attachment(MyFiles1(i))
                    Mail.Attachments.Add(attachment)
                Next i
            End If
            If lngCount3 > 0 Then
                For i = 0 To lngCount3 - 1
                    Dim attachment As System.Net.Mail.Attachment
                    attachment = New System.Net.Mail.Attachment(MyFiles3(i))
                    Mail.Attachments.Add(attachment)
                Next i
            End If
            Mail.ReplyToList.Add("reporting@primelineexpress.co.uk")
            'Mail.ReplyTo = New MailAddress("reporting@primelineexpress.co.uk")
            SMTP.EnableSsl = True
            SMTP.Port = "587"
            SMTP.Send(Mail)
        End If
    End Sub
    Public Sub AddToList(ByRef listDetails As List(Of String), ByVal strCustName As String, ByVal strFileName As String, ByVal strFolder As String, ByVal strTime As String)
        listDetails.Add(strCustName)
        listDetails.Add(strFileName)
        listDetails.Add(strFolder)
        listDetails.Add(0)
        listDetails.Add(0)
        listDetails.Add(0)
        listDetails.Add(strTime)
    End Sub
    Public Sub ArraySetup(Optional ByVal booTest As Boolean = False)
        Dim lngCSVCount As Long
        Dim lngCustCount As Long
        'Have below in order you want them on report
        AddToList(listDetails, "GBS", "GBS", strEDIs & strTestFolder, "12:00")
        AddToList(listDetails, "Kurt Geiger", "Kurt Geiger", strEDIs, "12:00")
        AddToList(listDetails, "TBS", "TBS", strEDIs & strTestFolder, "12:00")
        AddToList(listDetails, "Henkel", "Henkel", strEDIs & strTestFolder, "14:00")
        AddToList(listDetails, "Sealeys", "Sealeys", strEDIs, "14:00")
        AddToList(listDetails, "Dune", "Dune", strEDIs, "14:05")
        AddToList(listDetails, "Eagle", "XCR", strEagle, "16:00")
        AddToList(listDetails, "Automint", "Automint", strEDIs, "17:15")
        AddToList(listDetails, "Febi", "febi", strEDIs, "18:00")
        AddToList(listDetails, "Sony", "Sony", strEDIs & strTestFolder, "18:30")
        AddToList(listDetails, "Timco", "primeline_express", strTimco, "19:30")
        AddToList(listDetails, "Volvo", "Volvo", strEDIs, "19:30")
        AddToList(listDetails, "Technicolor", "Technicolor", strEDIs, "20:15")
        AddToList(listDetails, "Cinram", "Cinram", strEDIs & strTestFolder, "20:30")
        AddToList(listDetails, "Decora", "decora", strEDIs & strTestFolder, "21:00")
        AddToList(listDetails, "Ralawise", "ralawise", strEDIs, "21:05")


        'Add customers above here, last 2 (Bulk Events, PODs) are coded to appear after split line, so only add customers above those

        lngCSVCount = 2 'Set for how many are counted as CSVs instead of XML, and will appear at bottom of report
        AddToList(listDetails, "Bulk Events", "Bulk", strBulkEvents, "")
        AddToList(listDetails, "Westcoast/Clarity PODs", "Westcoast", strPODs, "")

        lngCustCount = ((listDetails.Count()) / lngFields)
        listDetails.Insert(0, lngCustCount)
        UpdateXMLStatus(listDetails, lngCSVCount, lngCustCount)
        TextEmail(listDetails, lngCustCount, lngCSVCount, booTest)
        GetUnSuccessfulFiles()
        Application.Exit()
    End Sub
    Public Function UpdateFileTypeCounts(ByRef listDetails As List(Of String), ByVal strFolder As String, ByVal strFileName As String, ByVal strFileType As String, ByVal lngCustCount As Long) As String
        Dim j As Long
        Dim lngFileCount As Long
        Dim lngCounter As Long = 0
        Dim strFileNames() As String = IO.Directory.GetFiles(strFolder, strFileType)
        lngFileCount = strFileNames.Count()
        If lngFileCount > 0 Then
            Select Case strFileType
                Case "*.xml"
                    For j = 0 To lngFileCount - 1
                        If InStr(1, strFileNames(j), strFileName) Then
                            lngCounter = lngCounter + 1
                        End If
                    Next j
                    If lngCounter > 0 Then
                        UpdateFileTypeCounts = lngCounter
                    Else
                        UpdateFileTypeCounts = ""
                    End If
                Case "*.csv"
                    UpdateFileTypeCounts = lngFileCount
                Case Else
                    UpdateFileTypeCounts = ""
            End Select
        Else
            UpdateFileTypeCounts = ""
        End If
    End Function
    Public Sub UpdateXMLStatus(ByRef listDetails As List(Of String), ByVal lngCSVCount As Long, ByVal lngCustCount As Long)
        Dim i As Long
        Dim strSearchFold As String
        Dim strSearchSuccFold As String
        Dim strSearchUnSuccFold As String
        Dim strSuccSuff As String = "Successful\"
        Dim strUnSuccSuff As String = "UnSuccessful\"
        Dim strFileName As String
        Dim strFileType As String

        For i = 1 To lngCustCount
            Select Case i
                Case <= (lngCustCount - lngCSVCount)
                    strFileType = "*.xml"
                Case Else
                    strFileType = "*.csv"
            End Select
            strSearchFold = listDetails((((i - 1) * lngFields) + fldFolder))
            strSearchSuccFold = strSearchFold & strSuccSuff
            strSearchUnSuccFold = strSearchFold & strUnSuccSuff
            strFileName = listDetails(((i - 1) * lngFields) + fldFileName)

            listDetails(((i - 1) * lngFields) + fldRecCnt) = UpdateFileTypeCounts(listDetails, strSearchFold, strFileName, strFileType, lngCustCount)
            listDetails(((i - 1) * lngFields) + fldUnSuccCnt) = UpdateFileTypeCounts(listDetails, strSearchUnSuccFold, strFileName, strFileType, lngCustCount)
            listDetails(((i - 1) * lngFields) + fldSuccCnt) = UpdateFileTypeCounts(listDetails, strSearchSuccFold, strFileName, strFileType, lngCustCount)
        Next
    End Sub

    Public Sub TextEmail(ByRef listDetails As List(Of String), ByVal lngCustCount As Long, ByVal lngCSVCount As Long, Optional ByVal booTest As Boolean = False)

        Dim i As Long
        Dim lngTableWidth As Long
        Dim lngColWidth As Long
        Dim lngCustWidth As Long
        Dim strHtmlBody As String
        Dim strHeader As String = ""
        Dim strXMLBody As String = ""
        Dim strCSVBody As String = ""
        Dim strSplitRow As String = ""
        Dim strFooter As String = ""
        Dim lngLastrow As Long

        lngColWidth = 1
        lngCustWidth = (32 * 4)
        lngTableWidth = 750
        lngLastrow = lngCustCount - lngCSVCount
        strHeader = ConstructHeader(lngTableWidth, lngColWidth)

        For i = 1 To lngLastrow
            strXMLBody = strXMLBody & AddTableRows(listDetails, lngColWidth, i)
        Next i

        strSplitRow = AddSplitRow(lngColWidth)

        For i = lngLastrow + 1 To lngCustCount
            strCSVBody = strCSVBody & AddTableRows(listDetails, lngColWidth, i)
        Next i

        strFooter = AddFooter(lngColWidth)

        strHtmlBody = strHeader & strXMLBody & strSplitRow & strCSVBody & strFooter
        CreateRecipientsList(booTest)
        SendEmail(strHtmlBody, listRecipients)
    End Sub

    Public Shared Sub SendEmail(ByVal strHtmlBody As String, ByVal listRecipients As List(Of String))
        Dim Mail As New MailMessage
        Dim i As Long

        For i = 0 To listRecipients.Count() - 1
            Mail.To.Add(listRecipients(i))
        Next
        Dim strEmailAddress = ConfigurationManager.AppSettings("strEmailAddress")
        dim strEmailPassword = ConfigurationManager.AppSettings("strEmailPassword")
        Mail.Subject = "XML File Monitor"
        Mail.From = New MailAddress("reporting@primelineexpress.co.uk", "Primeline Express Reporting", System.Text.Encoding.UTF8)
        Dim attachment As System.Net.Mail.Attachment
        attachment = New System.Net.Mail.Attachment(strAttachmentLoc)
        Mail.Attachments.Add(attachment)
        Mail.ReplyToList.Add("reporting@primelineexpress.co.uk")
        Mail.IsBodyHtml = True
        Mail.Body = strHtmlBody 'Message Here

        Dim SMTP As New SmtpClient("smtp.gmail.com") With {
            .Credentials = New System.Net.NetworkCredential(strEmailAddress, strEmailPassword), '<-- Password Here
            .EnableSsl = True,
            .Port = "587"
        }

        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        Try
            SMTP.Send(Mail)
            Exit Try
        Catch exc As Exception
            SMTP.Send(Mail)
        End Try
    End Sub

    Private Sub ArchiveButton_Click(sender As Object, e As EventArgs) Handles ArchiveButton.Click
        Call ArchiveFiles()
        MessageBox.Show("Files archived")
        Application.Exit()
    End Sub
    Public Function ConstructHeader(ByVal lngTableWidth As Long, ByVal lngColWidth As Long) As String
        Dim strOutput As String
        strOutput = ""
        strOutput = strOutput & "<html>" & vbCrLf
        strOutput = strOutput & "<body>" & vbCrLf
        strOutput = strOutput & "<table id=""t01""style=""width:" & lngTableWidth & "px; color: black; border-collapse: collapse; padding-top: 1px 1px 1px 1px; border-style: hidden; margin: 1px 1px 1px 1px;"">" & vbCrLf
        strOutput = strOutput & "  <tr>" & vbCrLf
        strOutput = strOutput & "    <th colspan=""" & lngColWidth + 4 & """ style=""background-color: #009933; color: white; text-align: left; padding: 1px 1px 1px 1px; text-align: left; width:100%; border-collapse: collapse;""><h1 style=""margin: 1px 0px 1px 0px; font-size: 22px; text-align: center;"">Primeline Express XML File Monitor</h1> " & vbCrLf
        strOutput = strOutput & "  </tr>" & vbCrLf
        strOutput = strOutput & InsertRow(lngColWidth, strHead1, strHead2, strHead3, strHead4, strHead5, "#009933", "#009933",,,, "white", "white", "white", "white", "white", "18")
        ConstructHeader = strOutput
    End Function
    Public Function AddTableRows(ByRef listDetails As List(Of String), ByVal lngColWidth As Long, ByVal i As Long) As String
        Dim strOutput As String
        Dim strCol1 As String = listDetails((i - 1) * lngFields + fldCustName)
        Dim strCol2 As String = listDetails(((i - 1) * lngFields) + fldTime)
        Dim strCol3 As String = listDetails(((i - 1) * lngFields) + fldRecCnt)
        Dim strCol4 As String = listDetails(((i - 1) * lngFields) + fldUnSuccCnt)
        Dim strCol5 As String = listDetails(((i - 1) * lngFields) + fldSuccCnt)
        strOutput = InsertRow(lngColWidth, strCol1, strCol2, strCol3, strCol4, strCol5)
        AddTableRows = strOutput
    End Function
    Public Function AddSplitRow(ByVal lngColWidth As Long) As String
        Dim strOutput As String
        strOutput = InsertRow(lngColWidth,,,,,,,, "#D9D9D9", "#D9D9D9", "#D9D9D9")
        AddSplitRow = strOutput
    End Function
    Public Function AddFooter(ByVal lngColWidth As Long) As String
        Dim strOutput As String = ""
        strOutput = strOutput & "</table>" & vbCrLf
        strOutput = strOutput & "</body>" & vbCrLf
        strOutput = strOutput & "</html>"
        AddFooter = strOutput
    End Function
    Public Function InsertRow(ByVal lngColWidth As Long, Optional ByVal strCol1 As String = "", Optional ByVal strCol2 As String = "", Optional ByVal strCol3 As String = "", Optional ByVal strCol4 As String = "", Optional ByVal strCol5 As String = "", Optional ByVal strBk1 As String = "#D9D9D9", Optional ByVal strBk2 As String = "#D9D9D9", Optional ByVal strBk3 As String = "#7030A0", Optional ByVal strBk4 As String = "#003399", Optional ByVal strBk5 As String = "#009933", Optional ByVal strFr1 As String = "black", Optional ByVal strFr2 As String = "black", Optional ByVal strFr3 As String = "white", Optional ByVal strFr4 As String = "white", Optional ByVal strFr5 As String = "white", Optional ByVal strFntSize As String = "14") As String
        Dim strOutput As String
        strOutput = "<tr>" & vbCrLf
        strOutput = strOutput & " <th colspan=""" & lngColWidth & """ style=""background-color: " & strBk1 & ";  color: " & strFr1 & "; text-align: left; padding: 1px 1px 1px 1px; text-align: left; border-collapse: collapse;border-bottom:1px solid white;""><h3 style=""margin: 1px 0px 1px 0px; font-size: " & strFntSize & "px; direction: rtl; text-align: center;"">" & strCol1 & "</h3></th>" & vbCrLf
        strOutput = strOutput & " <th colspan=""" & lngColWidth & """ style=""background-color: " & strBk2 & ";  color: " & strFr2 & "; text-align: left; padding: 1px 1px 1px 1px; text-align: left; border-collapse: collapse;border-bottom:1px solid white;""><h3 style=""margin: 1px 0px 1px 0px; font-size: " & strFntSize & "px; direction: rtl; text-align: center;"">" & strCol2 & "</h3></th>" & vbCrLf
        strOutput = strOutput & " <th colspan=""" & lngColWidth & """ style=""background-color: " & strBk3 & ";  color: " & strFr3 & "; text-align: left; padding: 1px 1px 1px 1px; text-align: left; border-collapse: collapse;border-bottom:1px solid white;""><h3 style=""margin: 1px 0px 1px 0px; font-size: " & strFntSize & "px; direction: rtl; text-align: center;"">" & strCol3 & "</h3></th>" & vbCrLf
        strOutput = strOutput & " <th colspan=""" & lngColWidth & """ style=""background-color: " & strBk4 & ";  color: " & strFr4 & "; text-align: left; padding: 1px 1px 1px 1px; text-align: left; border-collapse: collapse;border-bottom:1px solid white;""><h3 style=""margin: 1px 0px 1px 0px; font-size: " & strFntSize & "px; direction: rtl; text-align: center;"">" & strCol4 & "</h3></th>" & vbCrLf
        strOutput = strOutput & " <th colspan=""" & lngColWidth & """ style=""background-color: " & strBk5 & ";  color: " & strFr5 & "; text-align: left; padding: 1px 1px 1px 1px; text-align: left; border-collapse: collapse;border-bottom:1px solid white;""><h3 style=""margin: 1px 0px 1px 0px; font-size: " & strFntSize & "px; direction: rtl; text-align: center;"">" & strCol5 & "</h3></th>" & vbCrLf
        strOutput = strOutput & "  </tr>" & vbCrLf
        InsertRow = strOutput
    End Function
    Public Sub CreateRecipientsList(Optional ByVal booTest As Boolean = False)
        listRecipients.Add("jason.casey@primelineexpress.co.uk")
        listRecipients.Add("colin.white@primelineexpress.co.uk")
        If booTest = False Then
            listRecipients.Add("dom.cliffe@primelineexpress.co.uk")
            listRecipients.Add("robert.eadie@primelineexpress.co.uk")
            listRecipients.Add("jason.oriordan@primelineexpress.ie")
            listRecipients.Add("scott.beales@primelineexpress.co.uk")
            listRecipients.Add("andrew.orr@primelineexpress.co.uk")
            listRecipients.Add("ian.kavanagh@primelineexpress.ie")
            listRecipients.Add("clientservices@primelineexpress.co.uk")
            listRecipients.Add("john.ashall@primelineexpress.co.uk")
            listRecipients.Add("nick.spence@primelineexpress.co.uk")
            listRecipients.Add("warwicktransport@primelineexpress.co.uk")
            listRecipients.Add("stuart.topping@primelineexpress.co.uk")
            listRecipients.Add("warwickwarehouse@primelineexpress.co.uk")
        End If
    End Sub
    Private Sub TestButton_Click(sender As Object, e As EventArgs) Handles TestButton.Click
        Call ArraySetup(True)
    End Sub
    Public Sub ArchiveFiles()
        Const strArc As String = "Archive\"
        Dim strSuccFold1 = strEDIs & "Successful\"
        Dim strSuccFold2 = strEDIs & strTestFolder & "Successful\"
        Dim strSuccFold3 = strTimco & "Successful\"
        Dim strSuccFold4 = strEagle & "Successful\"
        Dim strSuccFold5 = strPODs & "Successful\"
        Dim strSuccFold6 = strBulkEvents & "Successful\"
        Dim i As Long
        'Define folders above and add to list here to archive by Year/Month/Day
        Dim listFolders As New List(Of String) From {
            strSuccFold1,
            strSuccFold2,
            strSuccFold3,
            strSuccFold4,
            strSuccFold5,
            strSuccFold6
        }
        'Get list of files in each unsuccessful folder

        'For each file, get date modified
        For i = 0 To (listFolders.Count() - 1)
            Dim strFilepath = listFolders(i)  'Specify path details
            Dim directory As New System.IO.DirectoryInfo(strFilepath)
            Dim Files As System.IO.FileInfo() = directory.GetFiles()
            Dim File As System.IO.FileInfo
            For Each File In Files
                Dim datLastModified As System.DateTime
                datLastModified = System.IO.File.GetLastWriteTime(strFilepath & File.ToString())
                'Check curr folder \Archive exists
                System.IO.Directory.CreateDirectory(strFilepath & strArc)
                'Check curr folder \Archive\YYYY exists
                Dim strYYYY As String = datLastModified.Year
                System.IO.Directory.CreateDirectory(strFilepath & strArc & strYYYY)
                'Check curr folder \Archive\MMM exists
                Dim strMMM As String = datLastModified.Month
                System.IO.Directory.CreateDirectory(strFilepath & strArc & strYYYY & "\" & strMMM)
                'Check curr folder \Archive\DD exists
                Dim strDD As String = datLastModified.Day
                System.IO.Directory.CreateDirectory(strFilepath & strArc & strYYYY & "\" & strMMM & "\" & strDD)
                Dim strComplete = strFilepath & strArc & strYYYY & "\" & strMMM & "\" & strDD & "\" & File.Name
                'Check if file exists in destination, kill it if it does
                If (System.IO.File.Exists(strComplete)) Then
                    'Delete existing file
                    System.IO.File.Delete(strComplete)
                End If
                File.MoveTo(strComplete)
                'Next file
            Next
        Next
    End Sub
End Class