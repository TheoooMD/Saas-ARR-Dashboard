Sub RefreshDashboard()
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("Metrics")

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found in Metrics sheet. Please paste your aggregated CSV data first.", vbExclamation
        Exit Sub
    End If

    MsgBox "Dashboard refreshed! " & (lastRow - 1) & " months of data loaded.", vbInformation, "Refresh Complete"
End Sub

Sub GenerateAICommentary()
    Dim wsM As Worksheet
    Dim wsD As Worksheet
    Set wsM = ThisWorkbook.Sheets("Metrics")
    Set wsD = ThisWorkbook.Sheets("Dashboard")

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data in Metrics sheet. Please paste your data first.", vbExclamation
        Exit Sub
    End If

    ' Build metrics string
    Dim metricsText As String
    metricsText = ""
    Dim i As Long
    For i = 2 To lastRow
        metricsText = metricsText & wsM.Cells(i, 1).Value & "|" & _
            Round(wsM.Cells(i, 2).Value, 0) & "|" & _
            Round(wsM.Cells(i, 3).Value, 0) & "|" & _
            Round(wsM.Cells(i, 4).Value, 0) & "|" & _
            Round(wsM.Cells(i, 5).Value, 0) & "|" & _
            Round(wsM.Cells(i, 6).Value, 0) & "|" & _
            Round(wsM.Cells(i, 7).Value, 0) & "|" & _
            Round(wsM.Cells(i, 10).Value, 1) & "|" & _
            Round(wsM.Cells(i, 11).Value, 1) & Chr(10)
    Next i

    Dim apiKey As String
    apiKey = InputBox("Enter your Anthropic API key (sk-ant-...):", "Claude API Key", "")
    If apiKey = "" Then Exit Sub

    ' Build the curl command as an AppleScript shell command
    Dim promptText As String
    promptText = "You are a CFO-level FP&A analyst at a B2B SaaS company. Based on the following monthly ARR waterfall metrics (Month|StartMRR|NewMRR|Expansion|Contraction|Churn|EndMRR|NRR|GRR), write a concise board-ready executive commentary max 250 words. Structure: (1) ARR Growth Narrative (2) Retention Analysis NRR GRR trends (3) Key Risk spike in churn or contraction (4) Forward Outlook. Use precise numbers. Confident CFO prose no bullet points. Data: " & metricsText

    ' Escape single quotes for shell safety
    promptText = Replace(promptText, "'", "")
    promptText = Replace(promptText, Chr(10), " ")
    promptText = Replace(promptText, Chr(13), " ")
    promptText = Replace(promptText, Chr(34), " ")

    Dim jsonPayload As String
    jsonPayload = "{""model"":""claude-sonnet-4-20250514"",""max_tokens"":800,""messages"":[{""role"":""user"",""content"":""" & promptText & """}]}"
    jsonPayload = Replace(jsonPayload, "'", "")

    Dim curlCmd As String
    curlCmd = "curl -s -X POST 'https://api.anthropic.com/v1/messages' " & _
        "-H 'Content-Type: application/json' " & _
        "-H 'x-api-key: " & apiKey & "' " & _
        "-H 'anthropic-version: 2023-06-01' " & _
        "--data '" & jsonPayload & "'"

    ' Use AppleScript to run curl and capture output
    Dim appleScript As String
    appleScript = "do shell script """ & Replace(curlCmd, """", "\""") & """"

    Dim response As String
    On Error GoTo ApiError
    response = MacScript(appleScript)
    On Error GoTo 0

    If Len(response) = 0 Then
        MsgBox "No response received. Check API key and internet connection.", vbCritical
        Exit Sub
    End If

    ' Parse text from JSON response
    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(response, """text"":""") + 8
    If startPos <= 8 Then
        MsgBox "Unexpected response: " & Left(response, 300), vbCritical
        Exit Sub
    End If
    endPos = InStr(startPos, response, """,""type""")
    If endPos = 0 Then endPos = InStr(startPos, response, """}]}")

    Dim commentary As String
    commentary = Mid(response, startPos, endPos - startPos)
    commentary = Replace(commentary, "\n", Chr(10))
    commentary = Replace(commentary, "\""", Chr(34))

    wsD.Range("A36").Value = commentary
    wsD.Range("A36").Font.Name = "Arial"
    wsD.Range("A36").Font.Size = 10
    wsD.Range("A36").Font.Color = RGB(30, 42, 58)
    wsD.Range("A36").Font.Italic = False

    MsgBox "Commentary generated successfully!", vbInformation, "Claude AI"
    Exit Sub

ApiError:
    MsgBox "Error calling API: " & Err.Description, vbCritical
End Sub
