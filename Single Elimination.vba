Option Explicit

Public Const DISCORD_WEBHOOK_URL As String = "PLACEHOLDER URL"

Public Sub createRandomGroup()
    Dim sel As Range, c As Range
    Dim temp As Collection
    Dim randomPlayers() As Variant
    Dim i As Long, n As Long
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells containing the players first.", vbExclamation
        Exit Sub
    End If
    Set sel = Selection
    
    Set temp = New Collection
    For Each c In sel.Cells
        If Trim$(CStr(c.Value)) <> "" Then temp.Add c.Value
    Next c
    If temp.Count = 0 Then
        MsgBox "Your selection has no values.", vbExclamation
        Exit Sub
    End If
    
    ReDim randomPlayers(1 To temp.Count)
    For i = 1 To temp.Count
        randomPlayers(i) = temp(i)
    Next i
    
    Randomize
    Dim j As Long, t As Variant
    For i = UBound(randomPlayers) To LBound(randomPlayers) + 1 Step -1
        j = Int((i - LBound(randomPlayers) + 1) * Rnd) + LBound(randomPlayers)
        t = randomPlayers(i): randomPlayers(i) = randomPlayers(j): randomPlayers(j) = t
    Next i
    
    Dim playersPerGroup As Variant
    playersPerGroup = Application.InputBox( _
        Prompt:="How many players do you want per group ?", _
        Title:="Group Size", Type:=1)
    If playersPerGroup = False Then Exit Sub
    If playersPerGroup < 1 Then
        MsgBox "Please enter a whole number of 1 or greater.", vbExclamation
        Exit Sub
    End If
    playersPerGroup = CLng(playersPerGroup) ' header handled separately
    
    n = UBound(randomPlayers) - LBound(randomPlayers) + 1
    Dim groupCount As Long
    groupCount = Application.WorksheetFunction.RoundUp(n / playersPerGroup, 0)
    
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Groups").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ActiveSheet)
    ws.Name = "Groups"
    ws.Columns("A:A").ColumnWidth = 2.3
    ws.Columns("B:B").ColumnWidth = ws.Columns("A:A").ColumnWidth ' B same as A
    
    Dim startRow As Long: startRow = 4
    Const colB As Long = 2
    Const colC As Long = 3
    
    Dim idx As Long: idx = LBound(randomPlayers)
    Dim g As Long, take As Long, remaining As Long, r As Long
    Dim blockRows As Long
    
    For g = 1 To groupCount
        If idx > UBound(randomPlayers) Then Exit For
        
        remaining = UBound(randomPlayers) - idx + 1
        take = IIf(remaining >= playersPerGroup, playersPerGroup, remaining) ' players in this group
        blockRows = take + 1 ' +1 header
        
        If g = 1 Then
            r = startRow
        Else
            r = r + blockRows + 1 ' spacer row
        End If
        
        With ws.Range(ws.Cells(r, colB), ws.Cells(r, colC))
            .Merge
            .Value = "Lobby " & g
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        Dim k As Long
        For k = 1 To take
            ws.Cells(r + k, colC).Value = randomPlayers(idx)
            idx = idx + 1
        Next k
        
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(r, colB), ws.Cells(r + take, colC))
        StyleGroupBox rng
        
        With ws.Range(ws.Cells(r + 1, colC), ws.Cells(r + take, colC)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next g
    
    ws.Activate
    ws.Range("C2").Select
    
    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
                ws.Range("D2").Left, _
                ws.Range("D2").Top, _
                ws.Range("D2:E2").Width, _
                ws.Range("D2").height)
    btn.TextFrame.Characters.text = "Advance to next round"
    btn.OnAction = "AdvanceRound"
    
    Dim btn2 As Shape
    Set btn2 = ws.Shapes.AddFormControl(xlButtonControl, _
                ws.Range("G2").Left, _
                ws.Range("G2").Top, _
                ws.Range("G2:H2").Width, _
                ws.Range("G2").height)
    btn2.TextFrame.Characters.text = "Send a message"
    btn2.OnAction = "SendToWebhook"
End Sub

Private Sub StyleGroupBox(ByVal rng As Range)
    With rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlMedium
        End With
        
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Weight = xlThin
        End With
    End With
End Sub

Public Sub AdvanceRound()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Groups")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Couldn't find the 'Groups' sheet.", vbExclamation
        Exit Sub
    End If
    
    Const startRow As Long = 4
    Dim ur As Range
    Set ur = ws.UsedRange
    
    Dim headers As Collection: Set headers = New Collection
    Dim f As Range, firstAddr As String, ma As Range
    Set f = ur.Find(What:="Lobby", LookIn:=xlValues, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not f Is Nothing Then
        firstAddr = f.Address
        Do
            If f.MergeCells Then
                Set ma = f.MergeArea
                If ma.Rows.Count = 1 And ma.Columns.Count = 2 Then
                    If f.Row = ma.Row And f.Column = ma.Column Then
                        If InStr(1, CStr(ma.Cells(1, 1).Value), "Lobby", vbTextCompare) > 0 Then
                            headers.Add Array(ma.Row, ma.Column, ma.Column + ma.Columns.Count - 1) ' row, leftCol, rightCol
                        End If
                    End If
                End If
            End If
            Set f = ur.FindNext(f)
        Loop While Not f Is Nothing And f.Address <> firstAddr
    End If
    
    If headers.Count = 0 Then
        MsgBox "No brackets found (looking for merged headers containing 'Bracket').", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long, maxRight As Long: maxRight = 0
    For i = 1 To headers.Count
        If headers(i)(2) > maxRight Then maxRight = headers(i)(2)
    Next i
    Dim posCol As Long, nameCol As Long
    posCol = maxRight - 1
    nameCol = maxRight
    
    Dim lastRow As Long: lastRow = ur.Row + ur.Rows.Count - 1
    Dim headerRows As Collection: Set headerRows = New Collection
    Dim r As Long
    For r = startRow To lastRow
        If ws.Cells(r, posCol).MergeCells Then
            Set ma = ws.Cells(r, posCol).MergeArea
            If ma.Row = r And ma.Column = posCol And ma.Rows.Count = 1 And ma.Columns.Count = 2 Then
                If InStr(1, CStr(ma.Cells(1, 1).Value), "Lobby", vbTextCompare) > 0 Then
                    headerRows.Add r
                End If
            End If
        End If
    Next r
    If headerRows.Count = 0 Then
        MsgBox "No brackets found in the rightmost columns.", vbExclamation
        Exit Sub
    End If
    
    ' --- Colors
    Dim lightGreen As Long, lightRed As Long
    lightGreen = RGB(198, 239, 206)
    lightRed = RGB(255, 199, 206)
    
    Dim qualifiedUsers() As String
    Dim qCount As Long: qCount = 0
    
    Dim rr As Long, posVal As Variant
    For i = 1 To headerRows.Count
        r = headerRows(i)
        rr = r + 1
        
        Do While rr <= lastRow
        
            If Len(Trim$(CStr(ws.Cells(rr, nameCol).Value))) > 0 Then _
                EloCalculation CStr(ws.Cells(rr, nameCol).Value), ws.Cells(rr, posCol).Value
            If ws.Cells(rr, posCol).MergeCells Then
                Set ma = ws.Cells(rr, posCol).MergeArea
                If ma.Rows.Count = 1 And ma.Columns.Count = 2 Then
                    If InStr(1, CStr(ma.Cells(1, 1).Value), "Lobby", vbTextCompare) > 0 Then Exit Do
                End If
            End If
            
            If Len(Trim$(CStr(ws.Cells(rr, posCol).Value))) = 0 And _
               Len(Trim$(CStr(ws.Cells(rr, nameCol).Value))) = 0 Then Exit Do
            
            posVal = ws.Cells(rr, posCol).Value
            If IsNumeric(posVal) And CLng(posVal) >= 1 And CLng(posVal) <= 6 Then
                ws.Cells(rr, nameCol).Interior.Color = lightGreen
                ws.Cells(rr, posCol).Interior.Color = lightGreen
                qCount = qCount + 1
                If qCount = 1 Then
                    ReDim qualifiedUsers(1 To 1)
                Else
                    ReDim Preserve qualifiedUsers(1 To qCount)
                End If
                qualifiedUsers(qCount) = CStr(ws.Cells(rr, nameCol).Value)
            Else
                ws.Cells(rr, nameCol).Interior.Color = lightRed
                ws.Cells(rr, posCol).Interior.Color = lightRed
            End If
            
            rr = rr + 1
        Loop
    Next i
    
    If qCount = 0 Then
        MsgBox "No players with positions 1–6 found in the rightmost brackets.", vbInformation
        Exit Sub
    End If
    
    Dim newPosCol As Long, newNameCol As Long
    newPosCol = nameCol + 2
    newNameCol = newPosCol + 1
    
    ws.Columns(newPosCol).ColumnWidth = ws.Columns(posCol).ColumnWidth
    
    ws.Columns(newPosCol).Resize(, 2).Clear

    Dim nextBracketCount As Long
    nextBracketCount = Application.WorksheetFunction.RoundUp(headerRows.Count / 2, 0)
    If nextBracketCount < 1 Then nextBracketCount = 1
    
    Dim j As Long
    Dim tmp As String
    Randomize
    For i = qCount To 2 Step -1
        j = Int(Rnd * (i - 1)) + 1
        tmp = qualifiedUsers(i): qualifiedUsers(i) = qualifiedUsers(j): qualifiedUsers(j) = tmp
    Next i
    
    Dim perBracket As Long
    perBracket = Application.WorksheetFunction.RoundUp(qCount / nextBracketCount, 0)
    
    Dim idx As Long: idx = 1
    Dim take As Long, g As Long, k As Long
    Dim rng As Range
    
    r = startRow
    For g = 1 To nextBracketCount
        If idx > qCount Then Exit For
        
        take = qCount - idx + 1
        If take > perBracket Then take = perBracket
        
        With ws.Range(ws.Cells(r, newPosCol), ws.Cells(r, newNameCol))
            .Merge
            .Value = "Lobby " & g
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        For k = 1 To take
            ws.Cells(r + k, newNameCol).Value = qualifiedUsers(idx)
            idx = idx + 1
        Next k
        
        Set rng = ws.Range(ws.Cells(r, newPosCol), ws.Cells(r + take, newNameCol))
        StyleGroupBox rng
        
        With ws.Range(ws.Cells(r + 1, newNameCol), ws.Cells(r + take, newNameCol)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        r = r + take + 2
    Next g
    
End Sub

Private Function EloModifierFromPosition(ByVal pos As Long) As Long
    Select Case pos
        Case 1: EloModifierFromPosition = 100
        Case 2: EloModifierFromPosition = 80
        Case 3: EloModifierFromPosition = 60
        Case 4: EloModifierFromPosition = 40
        Case 5: EloModifierFromPosition = 20
        Case 6: EloModifierFromPosition = 10
        Case 7: EloModifierFromPosition = -10
        Case 8: EloModifierFromPosition = -20
        Case 9: EloModifierFromPosition = -40
        Case 10: EloModifierFromPosition = -60
        Case 11: EloModifierFromPosition = -80
        Case Is >= 12: EloModifierFromPosition = -100
        Case Else: EloModifierFromPosition = 0
    End Select
End Function

Public Sub EloCalculation(ByVal playerName As String, ByVal position As Variant)
    Dim nameTrim As String: nameTrim = Trim$(CStr(playerName))
    If Len(nameTrim) = 0 Then Exit Sub

    Dim posNum As Long
    If IsNumeric(position) Then posNum = CLng(position) Else posNum = 0
    Dim eloMod As Long: eloMod = EloModifierFromPosition(posNum)

    Dim wsE As Worksheet
    On Error Resume Next
    Set wsE = ThisWorkbook.Worksheets("ELO Ranking")
    On Error GoTo 0

    If wsE Is Nothing Then
        Set wsE = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsE.Name = "ELO Ranking"
        wsE.Columns("A:A").ColumnWidth = 2
        wsE.Range("B1").Value = "Player"
        wsE.Range("C1").Value = "ELO"
    End If

    Dim lastB As Long
    lastB = wsE.Cells(wsE.Rows.Count, "B").End(xlUp).Row
    If lastB < 2 Then lastB = 1

    Dim searchRng As Range, f As Range
    Set searchRng = wsE.Range(wsE.Cells(2, "B"), wsE.Cells(lastB, "B"))
    If searchRng.Rows.Count > 0 Then
        Set f = searchRng.Find(What:=nameTrim, LookIn:=xlValues, LookAt:=xlWhole, _
                               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End If

    If Not f Is Nothing Then
        wsE.Cells(f.Row, "C").Value = Val(wsE.Cells(f.Row, "C").Value) + eloMod
    Else
        Dim newRow As Long
        newRow = lastB + 1
        wsE.Cells(newRow, "B").Value = nameTrim
        wsE.Cells(newRow, "C").Value = 1000 + eloMod
    End If
End Sub



Public Sub SendToWebhook()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Groups")
    On Error GoTo 0
    
    Dim userInput As Variant
    userInput = Application.InputBox( _
        Prompt:="Please write the message you wanna send", _
        Title:="Send to Discord", Type:=2) ' Type 2 = text
    If userInput = False Then Exit Sub ' Cancel pressed
    
    Dim finalText As String
    finalText = Trim$(CStr(userInput))
    If Len(finalText) = 0 Then Exit Sub
    
    Dim addBrackets As VbMsgBoxResult
    addBrackets = MsgBox("Add the lobby list to the message ?", vbYesNo + vbQuestion, "Include brackets?")
    
    If addBrackets = vbYes Then
        If ws Is Nothing Then
            MsgBox "Groups sheet not found. Sending only your message.", vbInformation
        Else
            Const startRow As Long = 4
            Dim bracketText As String
            bracketText = BuildRightmostBracketsSummary(ws, startRow)
            If Len(bracketText) > 0 Then
                ' Ensure exactly one newline before the first "--- BRACKET # ---"
                If Left$(bracketText, 1) <> vbLf Then finalText = finalText & vbLf
                finalText = finalText & bracketText
            End If
        End If
    End If
    
    SendDiscordInChunks finalText
End Sub


Private Function BuildRightmostBracketsSummary(ByVal ws As Worksheet, ByVal startRow As Long) As String
    Dim ur As Range: Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function
    
    Dim f As Range, firstAddr As String, ma As Range
    Dim rightmostRightCol As Long: rightmostRightCol = 0
    
    Set f = ur.Find(What:="Lobby", LookIn:=xlValues, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not f Is Nothing Then
        firstAddr = f.Address
        Do
            If f.MergeCells Then
                Set ma = f.MergeArea
                If ma.Rows.Count = 1 And ma.Columns.Count = 2 Then
                    If f.Row = ma.Row And f.Column = ma.Column Then
                        If InStr(1, CStr(ma.Cells(1, 1).Value), "Lobby", vbTextCompare) > 0 Then
                            If (ma.Column + 1) > rightmostRightCol Then rightmostRightCol = ma.Column + 1
                        End If
                    End If
                End If
            End If
            Set f = ur.FindNext(f)
        Loop While Not f Is Nothing And f.Address <> firstAddr
    End If
    If rightmostRightCol = 0 Then Exit Function
    
    Dim nameCol As Long: nameCol = rightmostRightCol
    Dim posCol As Long: posCol = nameCol - 1

    Dim lastRow As Long
    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, posCol).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row)
    
    Dim headerRows As New Collection
    Dim r As Long
    For r = startRow To lastRow
        If ws.Cells(r, posCol).MergeCells Then
            Set ma = ws.Cells(r, posCol).MergeArea
            If ma.Row = r And ma.Column = posCol And ma.Rows.Count = 1 And ma.Columns.Count = 2 Then
                If InStr(1, CStr(ma.Cells(1, 1).Value), "Lobby", vbTextCompare) > 0 Then
                    headerRows.Add r
                End If
            End If
        End If
    Next r
    If headerRows.Count = 0 Then Exit Function

    Dim sb As String, i As Long, rr As Long, nm As String
    Dim headerVal As String, brNum As String, line As String
    
    For i = 1 To headerRows.Count
        r = headerRows(i)
        headerVal = Trim$(CStr(ws.Cells(r, nameCol).Value))
        brNum = ExtractFirstNumber(headerVal)
        If Len(brNum) = 0 Then brNum = CStr(i)
        
        sb = sb & vbLf & vbLf & "--- LOBBY " & brNum & " ---" & vbLf
        

        line = ""
        rr = r + 1
        Do While rr <= lastRow
            If ws.Cells(rr, posCol).MergeCells Then
                Set ma = ws.Cells(rr, posCol).MergeArea
                If ma.Rows.Count = 1 And ma.Columns.Count = 2 Then
                    If InStr(1, CStr(ma.Cells(1, 1).Value), "Lobby", vbTextCompare) > 0 Then Exit Do
                End If
            End If
            If Len(Trim$(CStr(ws.Cells(rr, posCol).Value))) = 0 And _
               Len(Trim$(CStr(ws.Cells(rr, nameCol).Value))) = 0 Then Exit Do
            
            nm = Trim$(CStr(ws.Cells(rr, nameCol).Value))
            If Len(nm) > 0 Then
                If Len(line) > 0 Then line = line & ", "
                line = line & nm
            End If
            rr = rr + 1
        Loop
        
        sb = sb & line
    Next i
    
    BuildRightmostBracketsSummary = sb
End Function

Private Sub SendDiscordInChunks(ByVal content As String)
    Const MAX_LEN As Long = 1900 ' headroom for JSON escaping
    Dim pos As Long: pos = 1
    Dim chunk As String, cut As Long
    
    Do While pos <= Len(content)
        chunk = Mid$(content, pos, MAX_LEN)
        If pos + MAX_LEN <= Len(content) Then
            cut = InStrRev(chunk, vbLf)
            If cut > 0 Then chunk = Left$(chunk, cut - 1)
        End If
        If Not PostDiscordText(chunk) Then Exit Do
        pos = pos + Len(chunk)
        If cut = 0 And pos <= Len(content) Then pos = pos + 1
    Loop
End Sub

Private Function PostDiscordText(ByVal text As String) As Boolean
    On Error GoTo Fail
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim payload As String
    payload = "{""content"":""" & JsonEscape(text) & """}"
    http.Open "POST", DISCORD_WEBHOOK_URL, False
    http.SetTimeouts 30000, 30000, 30000, 30000
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send payload
    PostDiscordText = (http.Status = 200 Or http.Status = 204)
    Exit Function
Fail:
    PostDiscordText = False
End Function

Private Function ExtractFirstNumber(ByVal s As String) As String
    Dim i As Long, ch As String, num As String, started As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            num = num & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i
    ExtractFirstNumber = num
End Function

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

