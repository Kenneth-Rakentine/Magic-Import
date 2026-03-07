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

' ============================================================
' SINGLE IMPORT (original flow)
' ============================================================
Sub Import_Universal_MONTH_SELECT()
    Dim ws As Worksheet
    Dim wsName As String

    Select_Month.Show
    wsName = Select_Month.ComboBox1.Value
    Dim bulkMode As Boolean
    bulkMode = Select_Month.chkBulk.Value
    Unload Select_Month

    If wsName = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & wsName & "' not found.", vbCritical
        Exit Sub
    End If

    If bulkMode Then
        Call BulkImport(ws)
    Else
        Call SingleImport(ws)
    End If
End Sub

' ============================================================
' SINGLE IMPORT
' ============================================================
Sub SingleImport(ws As Worksheet)

    ' === Prompt for Manager XML file ===
    Dim mgrPath As String
    mgrPath = Application.GetOpenFilename("XML Files (*.xml;*.txt),*.xml;*.txt", , "Select the Manager XML Report")
    If mgrPath = "False" Then Exit Sub

    ' === Extract day and year from Manager file name ===
    Dim mgrFileName As String, mgrDateStr As String, mgrDay As Integer, mgrYear As String
    mgrFileName = Mid(mgrPath, InStrRev(mgrPath, "\") + 1)
    mgrDateStr = ExtractDateFromFilename(mgrFileName)
    If Len(mgrDateStr) = 8 Then
        mgrYear = Left(mgrDateStr, 4)
        mgrDay = CInt(Right(mgrDateStr, 2))
    Else
        MsgBox "Could not parse date from file name: " & mgrFileName, vbCritical
        Exit Sub
    End If

    ' === Find the correct row ===
    Dim excelRow As Long, foundRow As Boolean
    foundRow = False
    For excelRow = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If ws.Cells(excelRow, "A").Value = mgrDay Then
            foundRow = True
            Exit For
        End If
    Next excelRow
    If Not foundRow Then
        MsgBox "Day " & mgrDay & " not found in column A of sheet " & ws.Name, vbCritical
        Exit Sub
    End If

    ' === Parse Manager XML ===
    Call ParseAndWriteManager(mgrPath, mgrYear, ws, excelRow)

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

    If xmlDay <> mgrDay Then
        MsgBox "Manager and Journal reports are not for the same day!", vbCritical
        Exit Sub
    End If

    ' === Parse Journal XML ===
    Call ParseAndWriteJournal(xmlPath, ws, excelRow)

    MsgBox "Data imported for day " & mgrDay & " in sheet '" & ws.Name & "'!"
End Sub

' ============================================================
' BULK IMPORT
' ============================================================
Sub BulkImport(ws As Worksheet)

    ' Ask user to multi-select all XML files (both manager and journal)
    Dim files As Variant
    files = Application.GetOpenFilename( _
        "XML Files (*.xml;*.txt),*.xml;*.txt", _
        , _
        "Select ALL Manager and Journal XML files (Ctrl+Click to multi-select)", _
        , _
        True)   ' MultiSelect = True

    If Not IsArray(files) Then Exit Sub  ' User cancelled

    ' === Sort files into manager/journal buckets keyed by date ===
    Dim mgrFiles As Object, jrnFiles As Object
    Set mgrFiles = CreateObject("Scripting.Dictionary")
    Set jrnFiles = CreateObject("Scripting.Dictionary")

    Dim f As Variant, fname As String, dateStr As String
    For Each f In files
        fname = LCase(Mid(f, InStrRev(f, "\") + 1))
        dateStr = ExtractDateFromFilename(fname)
        If Len(dateStr) = 8 Then
            If Left(fname, 7) = "manager" Then
                mgrFiles(dateStr) = f
            ElseIf Left(fname, 7) = "journal" Then
                jrnFiles(dateStr) = f
            End If
        End If
    Next f

    ' === Process only dates where BOTH files are present ===
    Dim successCount As Integer, skipCount As Integer, errorCount As Integer
    Dim missingPairs As String
    successCount = 0: skipCount = 0: errorCount = 0
    missingPairs = ""

    Dim key As Variant
    For Each key In mgrFiles.Keys
        ' Check journal pair exists
        If Not jrnFiles.Exists(key) Then
            missingPairs = missingPairs & "  - " & key & " (missing journal)" & vbCrLf
            errorCount = errorCount + 1
        Else
            Dim dayNum As Integer
            dayNum = CInt(Right(key, 2))

            ' Find row in sheet
            Dim excelRow As Long, foundRow As Boolean
            foundRow = False
            For excelRow = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If ws.Cells(excelRow, "A").Value = dayNum Then
                    foundRow = True
                    Exit For
                End If
            Next excelRow

            If Not foundRow Then
                missingPairs = missingPairs & "  - Day " & dayNum & " not found in sheet" & vbCrLf
                errorCount = errorCount + 1
            Else
                ' === SKIP if row already has data (col B has room revenue) ===
                If ws.Cells(excelRow, "B").Value <> "" And ws.Cells(excelRow, "B").Value <> 0 Then
                    skipCount = skipCount + 1
                Else
                    Dim yearStr As String
                    yearStr = Left(key, 4)
                    Call ParseAndWriteManager(mgrFiles(key), yearStr, ws, excelRow)
                    Call ParseAndWriteJournal(jrnFiles(key), ws, excelRow)
                    successCount = successCount + 1
                End If
            End If
        End If
    Next key

    ' Also flag any journal files with no matching manager
    For Each key In jrnFiles.Keys
        If Not mgrFiles.Exists(key) Then
            missingPairs = missingPairs & "  - " & key & " (missing manager)" & vbCrLf
            errorCount = errorCount + 1
        End If
    Next key

    ' === Summary message ===
    Dim msg As String
    msg = "Bulk Import Complete!" & vbCrLf & vbCrLf
    msg = msg & "  Imported:  " & successCount & " day(s)" & vbCrLf
    msg = msg & "  Skipped:   " & skipCount & " day(s) (already had data)" & vbCrLf
    If errorCount > 0 Then
        msg = msg & "  Errors:    " & errorCount & " day(s)" & vbCrLf & vbCrLf
        msg = msg & "Details:" & vbCrLf & missingPairs
    End If
    MsgBox msg, vbInformation, "Bulk Import Results"
End Sub

' ============================================================
' SHARED: Parse Manager XML and write to sheet
' ============================================================
Sub ParseAndWriteManager(mgrPath As String, mgrYear As String, ws As Worksheet, excelRow As Long)
    Dim mgrDoc As Object
    Set mgrDoc = CreateObject("MSXML2.DOMDocument.6.0")
    mgrDoc.async = False
    mgrDoc.Load mgrPath
    If mgrDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Error in Manager XML: " & mgrDoc.ParseError.reason, vbCritical
        Exit Sub
    End If

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
                hYear = "": hType = ""
                If Not heading.SelectSingleNode("HEADING_1") Is Nothing Then hYear = heading.SelectSingleNode("HEADING_1").Text
                If Not heading.SelectSingleNode("HEADING_2") Is Nothing Then hType = heading.SelectSingleNode("HEADING_2").Text
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
                hYear = "": hType = ""
                If Not heading.SelectSingleNode("HEADING_1") Is Nothing Then hYear = heading.SelectSingleNode("HEADING_1").Text
                If Not heading.SelectSingleNode("HEADING_2") Is Nothing Then hType = heading.SelectSingleNode("HEADING_2").Text
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
End Sub

' ============================================================
' SHARED: Parse Journal XML and write to sheet
' ============================================================
Sub ParseAndWriteJournal(xmlPath As String, ws As Worksheet, excelRow As Long)
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load xmlPath
    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Error in Journal XML: " & xmlDoc.ParseError.reason, vbCritical
        Exit Sub
    End If

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

    ws.Cells(excelRow, 8).Value = sums("9000")
    ws.Cells(excelRow, 9).Value = sums("9001")
    ws.Cells(excelRow, 10).Value = sums("9003")
    ws.Cells(excelRow, 11).Value = sums("9004") + sums("9005")
    ws.Cells(excelRow, 12).Value = sums("9007")
    ws.Cells(excelRow, 13).Value = sums("9002")
    ws.Cells(excelRow, 14).Value = sums("9006")
    ws.Cells(excelRow, 15).Value = sums("5106")
End Sub
