Imports System
Module Program
    Private apiID As String
    Private apiKey As String
    Private secLabs_token$
    Private slClient As SL_Client
    Private studentS As List(Of slStudent)
    Private VC As VC_Client
    Private uFolder$

    Sub Main(args As String())
        uFolder$ = Environment.GetEnvironmentVariable("UserProfile")
        Dim actionWord$ = ""

        '      GoTo forceContinue

        If args.Count.ToString < 1 Or argExist("help", args) Then
            Call giveHelp()
            End
        Else
            actionWord$ = args(0)
            Console.WriteLine("ACTION: " + actionWord)
        End If

forceContinue:

        apiKey = ""
        apiID = ""
        secLabs_token = ""

        apiID = argValue("--apiID", args)
        apiKey = argValue("--apiKey", args)
        secLabs_token = argValue("--slToken", args)

        If Len(apiID) = 0 Or Len(apiKey) = 0 Then
            Call getCredsFromFile()
            '      If Len(apiID) Then Console.WriteLine("Pulled apiID from UserProfile: " + apiID)
        End If

        Select Case actionWord
            Case "get_seclabs_summary"
                If Len(secLabs_token) = 0 Then
                    Console.WriteLine("ERROR: Use --slToken to provide the Security Labs API Token")
                    End
                End If
                Call getSecLabsSummary(secLabs_token)

            Case "get_app_profiles"
                If Len(apiKey) = 0 Or Len(apiID) = 0 Then
                    Console.WriteLine("ERROR: Must provide --apiID and --apiKEY parameters")
                    End
                End If
                VC = New VC_Client
                Call VC.getAppProfiles(apiID, apiKey)
                Call showAppProfiles()

            Case "get_dast"
                If Len(apiKey) = 0 Or Len(apiID) = 0 Then
                    Console.WriteLine("ERROR: Must provide --apiID and --apiKEY parameters")
                    End
                End If
                VC = New VC_Client
                Call showDASTanalysis()

            Case "new_dast"
                If Len(apiKey) = 0 Or Len(apiID) = 0 Then
                    Console.WriteLine("ERROR: Must provide --apiID and --apiKEY parameters")
                    End
                End If

                VC = New VC_Client
                Dim appUUID$ = argValue("--linkapp_UUID", args)
                If Len(argValue("--linkapp_NAME", args)) Then
                    Call VC.getAppProfiles(apiID, apiKey)
                    For Each P In VC.appProfiles
                        If LCase(P.name) = LCase(argValue("--linkapp_NAME", args)) Then
                            appUUID = P.uuid
                            ' Console.WriteLine(appUUID + "=" + P.uuid)
                        End If
                    Next
                End If

                Call VC.newDAST(argValue("--dast_name", args), argValue("--dast_url", args), apiID, apiKey, appUUID)

        End Select


        End
    End Sub

    Private Sub showDASTanalysis()
        Dim dastAnalyses As List(Of dastAnalysis) = New List(Of dastAnalysis)
        dastAnalyses = VC.getDynamicAnalyses(apiID, apiKey)

        Console.WriteLine("NAME" + spaces(36) + "ANALYSIS_OCC_ID" + spaces(20) + "SCAN_ID" + spaces(30) + "END_TIME" + spaces(20) + "LINKED_APP" + spaces(20) + "FLAWS" + spaces(3) + "TARGET_URL" + spaces(20))
        Console.WriteLine("----" + spaces(36) + "---------------" + spaces(20) + "-------" + spaces(30) + "--------" + spaces(20) + "----------" + spaces(20) + "-----" + spaces(3) + "----------" + spaces(20))

        Dim rowNum As Integer = 1
        Dim a$ = ""
        Dim b$ = ""
        Dim c$ = ""
        Dim targetURL$ = ""

        For Each D In dastAnalyses
            With D
                a$ = "" : b$ = "" : c$ = ""
                a$ = .name + spaces(40 - Len(.name)) ' + .analysis_id + spaces(35 - Len(.analysis_id))
                '                a$ = .name + spaces(40 - Len(.name)) + .number_of_scans.ToString + spaces(Len(10 - .number_of_scans.ToString)) + .analysis_id + spaces(35 - Len(.analysis_id))
                rowNum = 1
                For Each dA In .occurrenceS
                    b$ = dA.analysis_occurrence_id + spaces(35 - Len(dA.analysis_occurrence_id))
                    '                    b$ = dA.analysis_occurrence_id + spaces(35 - Len(dA.analysis_occurrence_id)) + dA.actual_end_date + spaces(25 - Len(dA.actual_end_date))

                    For Each sD In .scansOfOccurrence
                        If dA.analysis_occurrence_id = sD.analysis_occurrence_id Then
                            targetURL$ = Replace(Replace(sD.target_url, "http://", ""), "https://", "")
                            c$ = sD.scan_id + spaces(37 - Len(sD.scan_id)) + sD.end_date + spaces(28 - Len(sD.end_date)) + sD.linked_platform_app_name + spaces(30 - Len(sD.linked_platform_app_name)) + sD.total_flaw_count.ToString + spaces(8 - Len(sD.total_flaw_count.ToString)) + targetURL
                            '                            c$ = sD.scan_id + spaces(40 - Len(sD.scan_id)) + sD.result_import_status + spaces(20 - Len(sD.result_import_status)) + sD.target_url + spaces(20 - Len(sD.target_url)) + sD.linked_platform_app_name
                        End If
                        If rowNum = 1 Then
                            Console.WriteLine(a + b + c)
                        Else
                            Console.WriteLine(spaces(Len(a)) + b + c)
                        End If
                    Next


                    rowNum += 1
                Next
            End With
        Next


    End Sub

    Private Sub showAppProfiles()
        Dim a$ = ""
        Dim b$ = ""
        Dim c$ = ""

        For Each App In VC.appProfiles
            With App
                b = "[" + .id.ToString + "]"
                b += spaces(10 - Len(b)) : a = b
                a += .name + spaces(40 - Len(.name))
                b = .uuid
                b += spaces(40 - Len(b)) : a += b
                c = .linked_scan_target_url
                If Len(c) = 0 Then c = "NO DAST" Else c = "DAST:" + c
                a += c
            End With
            Console.WriteLine(a)
        Next
    End Sub


    Public Sub getSecLabsSummary(apiToken$)

        slClient = New SL_Client(secLabs_token)
        slClient.loadProgressAllUsers()

        Console.WriteLine("# of Lessons: " + slClient.moduleNames.Count.ToString)
        For Each L In slClient.moduleNames
            Dim a$ = L + spaces(50 - Len(L))
            Dim S As stats
            S = lessonStats(L)

            Dim b$ = ""
            b += "[" + S.numStudents.ToString + "/" + slClient.studentList.Count.ToString + "]   " : a += b

            b = "[Min/Avg/Max Mins]:" + S.minTime.ToString + "/" + S.avgTime.ToString + "/" + S.maxTime.ToString
            b += spaces(35 - Len(b)) : a += b

            b = "[Before/After/Avg Improve]:" + S.avgBefore.ToString + "/" + S.avgAfter.ToString + "/" + S.avgImprove.ToString
            b += spaces(45 - Len(b)) : a += b

            b = "[Students/Revisits]:" + S.numStudentsRevisit.ToString + "/" + S.numRevisitsTL.ToString
            b += spaces(25 - Len(b))
            a += b

            a += "FAST: " + S.quickestStudent + " SLOW: " + S.slowestStudent

            Console.WriteLine(a)
        Next


    End Sub



    Private Function lessonStats(moduleName$) As stats
        lessonStats = New stats
        Dim tlTime As Decimal
        Dim tlImprove As Decimal
        Dim numImproveRatings As Integer


        With lessonStats

            For Each U In slClient.studentList
                Dim S As studentLessonStats
                S = gatherStudentLessonStats(U.id, moduleName)


                If S.numVisits > 0 Then
                    If S.numVisits > 1 Then
                        .numStudentsRevisit += 1
                        .numRevisitsTL += S.numVisits - 1
                    End If
                    .numStudents += 1
                    tlTime += S.tlTime
                    .totalPoints += S.points

                    If S.startRating > 0 And S.endRating > 0 Then
                        tlImprove += S.endRating - S.startRating
                        numImproveRatings += 1
                        .avgBefore += S.startRating
                        .avgAfter += S.endRating
                    End If

                    If S.tlTime > .maxTime Then
                        .maxTime = S.tlTime
                        .slowestStudent = U.name + " (" + .maxTime.ToString + ")"
                    End If
                    If S.tlTime > 0 Then
                        If .minTime = 0 Then
                            .minTime = S.tlTime
                            .quickestStudent = U.name + " (" + .minTime.ToString + ")"
                        End If
                        If S.tlTime < .minTime Then
                            .minTime = S.tlTime
                            .quickestStudent = U.name + " (" + .minTime.ToString + ")"
                        End If
                    End If
                End If
            Next

            ' final calcs
            .avgBefore = Math.Round(.avgBefore / numImproveRatings, 2)
            .avgAfter = Math.Round(.avgAfter / numImproveRatings, 2)
            .avgImprove = Math.Round(tlImprove / numImproveRatings, 2)
            .avgTime = Math.Round(tlTime / .numStudents, 2)


        End With
    End Function

    Private Function gatherStudentLessonStats(userID$, lessonName$) As studentLessonStats
        gatherStudentLessonStats = New studentLessonStats
        For Each S In slClient.studentList
            If S.id <> userID Then GoTo skipStudent

            For Each L In S.lessons
                If L.module <> lessonName Then GoTo skipLesson

                With gatherStudentLessonStats
                    .numVisits += 1
                    If L.startRating > 0 Then .startRating = L.startRating
                    If L.endRating > 0 Then .endRating = L.endRating
                    .tlTime += L.minutes
                    .points += L.points
                End With


skipLesson:
            Next
skipStudent:
        Next
    End Function


    Private Sub giveHelp()
        Console.WriteLine("USAGE: VCCLI action --param1 param1_value --param2 param2_value" + vbCrLf)
        Console.WriteLine("ACTIONS:")
        Console.WriteLine("--------")
        Console.WriteLine(fLine("help", "Produces this list of actions and parameters"))
        Console.WriteLine(fLine("get_seclabs_summary", "Returns a summary of Security Labs student & lesson activity"))
        Console.WriteLine(fLine("get_dast", "Returns a summary of DAST analyses and scans"))
        Console.WriteLine(fLine("get_app_profiles", "Returns a summary of Application Profiles"))
        Console.WriteLine(fLine("new_dast", "Creates and starts a DAST scan [req:dast_name & dast_url,optional: linkapp_UUID/NAME]"))
        Console.WriteLine(fLine("link_dast", "Links a DAST scan to an Application Profile [req dast_name/dast_analysis_id & linkapp_UUID/NAME"))

        Console.WriteLine(vbCrLf + "PARAMETERS:")
        Console.WriteLine("-----------")

        Console.WriteLine(fLine("--apiID", "The Veracode API ID (if not inside " + uFolder))
        Console.WriteLine(fLine("--apiKEY", "The Veracode API KEY (if not inside " + uFolder))
        Console.WriteLine(fLine("--slToken", "The Security Labs API KEY"))
        Console.WriteLine(fLine("--dast_name", "Name of the DAST scan"))
        Console.WriteLine(fLine("--dast_url", "URL used for a DAST scan"))
        Console.WriteLine(fLine("--linkapp_UUID", "Used to link a DAST scan to a Profile using APP UUID"))
        Console.WriteLine(fLine("--linkapp_Name", "Used to link a DAST scan to a Profile using APP NAME"))

    End Sub
    Private Function fLine(arg1$, arg2$, Optional ByVal numSpaces As Integer = 25) As String
        Return arg1 + spaces(numSpaces - Len(arg1)) + arg2
    End Function

    Private Sub getCredsFromFile()
        Dim fName$ = uFolder + "\.veracode\credentials"

        If Dir(fName) = "" Then Exit Sub

        Dim a$ = ""
        Dim FF As Integer
        FF = FreeFile()

        FileOpen(FF, fName, OpenMode.Input)

        Do Until EOF(FF) = True
            a$ = LineInput(FF)
            If InStr(a, "veracode_api_key_id") Then
                apiID = Trim(Mid(a, InStr(a, "=") + 1))
            End If
            If InStr(a, "veracode_api_key_secret") Then
                apiKey = Trim(Mid(a, InStr(a, "=") + 1))
            End If
        Loop
        FileClose(FF)
    End Sub
End Module
