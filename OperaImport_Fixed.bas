Sub cb(control As IRibbonControl)
Import_Universal_MONTH_SELECT
End Sub

Function ExtractDateFromFilename(filename As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d{8}"
    re.Global = False
    If re.test(filename) Then
        Set matches = re.Execute(filename)
        ExtractDateFromFilename = matches(0)
    Else
        ExtractDateFromFilename = ""
    End If
End Function

Sub Import_Universal_MONTH_SELECT()
    Dim ws As Worksheet
    Dim wsName As String
    Select_Month.Show
    wsName = Select_Month.ComboBox1.Value
    Unload Select_Month
    If wsName = "" Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & wsName & "' not found.", vbCritical
        Exit Sub
    End If

    ' === Prompt for Manager XML file ===
    Dim mgrPath As String
    mgrPath = Application.GetOpenFilename("XML Files (*.xml;*.txt),*.xml;*.txt", , "Select the Manager XML Report")
    If mgrPath = "False" Then Exit Sub

    ' === Extract day AND YEAR from Manager file name ===
    Dim mgrFileName As String, mgrDateStr As String, mgrDay As Integer, mgrYear As String
    mgrFileName = Mid(mgrPath, InStrRev(mgrPath, "\") + 1)
    mgrDateStr = ExtractDateFromFilename(mgrFileName)
    If Len(mgrDateStr) = 8 Then
        mgrYear = Left(mgrDateStr, 4)  ' Extract YYYY
        mgrDay = CInt(Right(mgrDateStr, 2))
    Else
        MsgBox "Could not parse date from file name: " & mgrFileName, vbCritical
        Exit Sub
    End If

    ' === Find the correct row for this day in the worksheet (search column A for the day) ===
    Dim excelRow As Long, foundRow As Boolean
    foundRow = False
    For excelRow = 1 To ws.Cells(ws.Rows.count, "A").End(xlUp).Row
        If ws.Cells(excelRow, "A").Value = mgrDay Then
            foundRow = True
            Exit For
        End If
    Next excelRow
    If Not foundRow Then
        MsgBox "Day " & mgrDay & " not found in column A of sheet " & wsName, vbCritical
        Exit Sub
    End If

    ' === Parse Manager XML for Room Revenue and Rooms Occupied ===
    Dim mgrDoc As Object
    Set mgrDoc = CreateObject("MSXML2.DOMDocument.6.0")
    mgrDoc.async = False
    mgrDoc.Load mgrPath
    If mgrDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Error in Manager XML file: " & mgrDoc.ParseError.reason, vbCritical
        Exit Sub
    End If

    ' Get Room Revenue (col B) and Rooms Occupied (col E) for CURRENT YEAR, DAY
    Dim mvList As Object, mv As Object, desc As String
    Dim headingList As Object, heading As Object, hYear As String, hType As String
    Dim foundRevenue As Boolean, foundRooms As Boolean
    foundRevenue = False: foundRooms = False

    Set mvList = mgrDoc.SelectNodes("//G_MASTER_VALUE")
    For Each mv In mvList
        desc = ""
        If Not mv.SelectSingleNode("DESCRIPTION") Is Nothing Then
            desc = Trim(mv.SelectSingleNode("DESCRIPTION").Text)
        End If

        If Not foundRevenue And LCase(desc) = "room revenue" Then
            Set headingList = mv.SelectNodes("LIST_G_HEADING_1_ORDER/G_HEADING_1_ORDER")
            For Each heading In headingList
                hYear = ""
                hType = ""
                If Not heading.SelectSingleNode("HEADING_1") Is Nothing Then hYear = heading.SelectSingleNode("HEADING_1").Text
                If Not heading.SelectSingleNode("HEADING_2") Is Nothing Then hType = heading.SelectSingleNode("HEADING_2").Text
                ' *** FIXED: Use mgrYear extracted from filename instead of hardcoded "2025" ***
                If hYear = mgrYear And LCase(hType) = "day" Then
                    If Not heading.SelectSingleNode("LIST_G_SUM_AMOUNT/G_SUM_AMOUNT/FORMATTED_AMOUNT") Is Nothing Then
                        ws.Cells(excelRow, "B").Value = Replace(heading.SelectSingleNode("LIST_G_SUM_AMOUNT/G_SUM_AMOUNT/FORMATTED_AMOUNT").Text, ",", "")
                        foundRevenue = True
                        Exit For
                    End If
                End If
            Next heading
        End If

        If Not foundRooms And LCase(desc) = "rooms occupied" Then
            Set headingList = mv.SelectNodes("LIST_G_HEADING_1_ORDER/G_HEADING_1_ORDER")
            For Each heading In headingList
                hYear = ""
                hType = ""
                If Not heading.SelectSingleNode("HEADING_1") Is Nothing Then hYear = heading.SelectSingleNode("HEADING_1").Text
                If Not heading.SelectSingleNode("HEADING_2") Is Nothing Then hType = heading.SelectSingleNode("HEADING_2").Text
                ' *** FIXED: Use mgrYear extracted from filename instead of hardcoded "2025" ***
                If hYear = mgrYear And LCase(hType) = "day" Then
                    If Not heading.SelectSingleNode("LIST_G_SUM_AMOUNT/G_SUM_AMOUNT/FORMATTED_AMOUNT") Is Nothing Then
                        ws.Cells(excelRow, "E").Value = Replace(heading.SelectSingleNode("LIST_G_SUM_AMOUNT/G_SUM_AMOUNT/FORMATTED_AMOUNT").Text, ",", "")
                        foundRooms = True
                        Exit For
                    End If
                End If
            Next heading
        End If

        If foundRevenue And foundRooms Then Exit For
    Next mv

    ' === Prompt for Journal XML file ===
    Dim xmlPath As String
    xmlPath = Application.GetOpenFilename("XML Files (*.xml;*.txt),*.xml;*.txt", , "Select the Journal XML Report")
    If xmlPath = "False" Then Exit Sub

    ' === Extract day from Journal file name ===
    Dim xmlFileName As String, xmlDateStr As String, xmlDay As Integer
    xmlFileName = Mid(xmlPath, InStrRev(xmlPath, "\") + 1)
    xmlDateStr = ExtractDateFromFilename(xmlFileName)
    If Len(xmlDateStr) = 8 Then
        xmlDay = CInt(Right(xmlDateStr, 2))
    Else
        MsgBox "Could not parse date from file name: " & xmlFileName, vbCritical
        Exit Sub
    End If

    ' === Confirm both reports are for the same day ===
    If xmlDay <> mgrDay Then
        MsgBox "Manager and Journal reports are not for the same day!", vbCritical
        Exit Sub
    End If

    ' === Set up XML parser for Journal ===
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load xmlPath
    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Error in Journal XML file: " & xmlDoc.ParseError.reason, vbCritical
        Exit Sub
    End If

    ' Transaction codes and their corresponding columns
    Dim codeMap As Object
    Set codeMap = CreateObject("Scripting.Dictionary")
    codeMap("9000") = 8   ' CASH (H)
    codeMap("9001") = 9   ' CHECKS (I)
    codeMap("9003") = 10  ' AMEX (J)
    codeMap("9004") = 11  ' VISA (K)
    codeMap("9005") = 11  ' MC (K)
    codeMap("9007") = 12  ' DISC (L)
    codeMap("9002") = 13  ' Direct Bill (M)
    codeMap("9006") = 14  ' MISC (N)
    codeMap("5106") = 15  ' Cupboard (O)

    Dim sums As Object
    Set sums = CreateObject("Scripting.Dictionary")
    Dim code As Variant
    For Each code In codeMap.Keys
        sums(code) = 0
    Next code

    ' Loop through all G_FIRST nodes (each payment type)
    Dim gFirstList As Object
    Set gFirstList = xmlDoc.SelectNodes("//G_FIRST")
    Dim gFirst As Object, trxCode As String
    For Each gFirst In gFirstList
        trxCode = ""
        If Not gFirst.SelectSingleNode("FIRST") Is Nothing Then
            trxCode = Trim(Split(gFirst.SelectSingleNode("FIRST").Text, " ")(0))
        End If
        If codeMap.Exists(trxCode) Then
            Dim totalVal As Double
            totalVal = 0
            If trxCode = "5106" Then
                If Not gFirst.SelectSingleNode("FIRST_DEBIT") Is Nothing Then
                    totalVal = Val(gFirst.SelectSingleNode("FIRST_DEBIT").Text)
                End If
            Else
                If Not gFirst.SelectSingleNode("FIRST_CREDIT") Is Nothing Then
                    totalVal = Val(gFirst.SelectSingleNode("FIRST_CREDIT").Text)
                End If
            End If
            sums(trxCode) = totalVal
        End If
    Next gFirst

    ' Write all totals to worksheet
    ws.Cells(excelRow, 8).Value = sums("9000")    ' CASH (H)
    ws.Cells(excelRow, 9).Value = sums("9001")    ' CHECKS (I)
    ws.Cells(excelRow, 10).Value = sums("9003")   ' AMEX (J)
    ws.Cells(excelRow, 11).Value = sums("9004") + sums("9005")   ' MC/VISA (K)
    ws.Cells(excelRow, 12).Value = sums("9007")   ' DISC (L)
    ws.Cells(excelRow, 13).Value = sums("9002")   ' Direct Bill (M)
    ws.Cells(excelRow, 14).Value = sums("9006")   ' MISC (N)
    ws.Cells(excelRow, 15).Value = sums("5106")   ' Cupboard (O)

    MsgBox "Opera Manager and Journal data imported for day " & mgrDay & " in sheet '" & wsName & "'!"
End Sub
