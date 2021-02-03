Imports RestSharp
Imports Newtonsoft.Json

Public Class VC_Client
    Public lastError$ = ""
    Public fqdN$ 'ie localhost or myserver.myzone.com
    Public appProfiles As List(Of appProfile)
    '    Public dastAnalyses As List(Of dy)

    Public Sub New()
        fqdN = "api.veracode.com"

    End Sub

    Public Function getDynamicAnalyses(apiID$, apiKEY$) As List(Of dastAnalysis)
        '        getDynamicAnalyses = New List(Of dastAnalysis)

        Dim urI$ = "/was/configservice/v1/analyses?size=500"
        Dim client = New RestClient("https://" + fqdN + urI)
        Dim request = New RestRequest(Method.GET)
        Dim response As IRestResponse

        Dim H As New HmacAuthHeader
        Dim hmacHeader$ = H.CalculateAuthorizationHeader(apiID, apiKEY, fqdN, urI, "", "GET")

        request.AddHeader("Host", fqdN)
        request.AddHeader("Authorization", hmacHeader)

        response = client.Execute(request)

        Dim jsoN$ = response.Content
        jsoN = removePaging(jsoN)
        jsoN = cleanJSON(jsoN)
        jsoN = removeDASTextraCapabilities(jsoN) 'ugh, sloppy json, sloppier editing

        '        Console.WriteLine(jsoN)

        Console.WriteLine("Pulling Analysis Occurrences with Analyses..")
        Dim dastAnalyses As List(Of dastAnalysis) = JsonConvert.DeserializeObject(Of List(Of dastAnalysis))(jsoN)
        Dim dAnalysisOccurrences As List(Of dastAnalOccur) = getDynamicAnalysisOccurrences(apiID, apiKEY)
        Console.WriteLine("Pulling Scans for " + dAnalysisOccurrences.Count.ToString + " Analysis Occurences..")


        For Each D In dastAnalyses
            With D
                D.occurrenceS = New List(Of dastAnalOccur)
                For Each dA In dAnalysisOccurrences
                    If dA.analysis_id = D.analysis_id Then
                        D.occurrenceS.Add(dA)
                        D.scansOfOccurrence = getScansOfAnalysisOccurrences(apiID, apiKEY, dA.analysis_occurrence_id)
                    End If
                Next
            End With
        Next

        Return dastAnalyses


    End Function
    Public Function getDynamicAnalysisOccurrences(apiID$, apiKEY$) As List(Of dastAnalOccur)
        Dim dAnal As New List(Of dastAnalOccur)

        Dim urI$ = "/was/configservice/v1/analysis_occurrences?size=500"
        Dim client = New RestClient("https://" + fqdN + urI)
        Dim request = New RestRequest(Method.GET)
        Dim response As IRestResponse

        Dim H As New HmacAuthHeader
        Dim hmacHeader$ = H.CalculateAuthorizationHeader(apiID, apiKEY, fqdN, urI, "", "GET")

        request.AddHeader("Host", fqdN)
        request.AddHeader("Authorization", hmacHeader)

        response = client.Execute(request)

        Dim jsoN$ = response.Content
        jsoN = removePaging(jsoN)
        jsoN = cleanJSON(jsoN)

        dAnal = JsonConvert.DeserializeObject(Of List(Of dastAnalOccur))(jsoN)

        Return dAnal
    End Function

    Public Function getScansOfAnalysisOccurrences(apiID$, apiKEY$, analID$) As List(Of scanDetail)
        Dim scanOcc As New List(Of scanDetail)

        Dim urI$ = "/was/configservice/v1/analysis_occurrences/" + analID + "/scan_occurrences?size=500"

        Dim client = New RestClient("https://" + fqdN + urI)
        Dim request = New RestRequest(Method.GET)
        Dim response As IRestResponse

        Dim H As New HmacAuthHeader
        Dim hmacHeader$ = H.CalculateAuthorizationHeader(apiID, apiKEY, fqdN, urI, "", "GET")

        request.AddHeader("Host", fqdN)
        request.AddHeader("Authorization", hmacHeader)

        response = client.Execute(request)

        Dim jsoN$ = response.Content
        jsoN = removePaging(jsoN)
        jsoN = cleanJSON(jsoN)

        'Console.WriteLine(jsoN)

        scanOcc = JsonConvert.DeserializeObject(Of List(Of scanDetail))(jsoN)

        Return scanOcc
    End Function

    Public Sub newDAST(dastName$, dastURL$, apiID$, apiKEY$, Optional ByVal linkedAppUUID$ = "", Optional ByVal linkedAppName$ = "")

        Dim appID$ = ""
        Dim c$ = Chr(34)
        Dim jsoN$ = "{" + c$ + "name" + c$ + ":" + c$ + dastName + c$ + "," + c$ + "scans" + c$ + ":[{"

        If Len(linkedAppUUID) Then appID = linkedAppUUID
        If Len(linkedAppName) Then appID = "lookitup"

        If Len(appID) Then jsoN += c$ + "linked_platform_app_uuid" + c$ + ":" + c$ + appID + c$ + ","


        'lame & lazy but obj structure is a heavier lift than this
        jsoN += c$ + "scan_config_request" + c$ + ":{" + c$ + "target_url" + c$ + ":{" + c$ + "url" + c$ + ":" + c$ + dastURL + c$ + "}}}],"
        jsoN += c$ + "schedule" + c$ + ":{" + c$ + "now" + c$ + ":true," + c$ + "duration" + c$ + ":{" + c$ + "length" + c$ + ":1," + c$ + "unit" + c$ + ":" + c$ + "DAY" + c$
        jsoN += "}}}"


        'jsoN += "}," + c$ + "end date" + c$ + ":" + c$ + c$ + "," + c$ + "scan_recurrence_schedule" + c$ + ":{" + c$ + "recurrence_type" + c$ + ":" + c$ + "WEEKLY" + c$ + ","
        'jsoN += c$ + "schedule_end_after" + c$ + ":2," + c$ + "recurrence_interval" + c$ + ":1," + c$ + "day_of_week" + c$ + ":" + c$ + "MONDAY" + c$ + "}}}"

        Dim urI$ = "/was/configservice/v1/analyses"
        Dim client = New RestClient("https://" + fqdN + urI)
        Dim request = New RestRequest(Method.POST)
        Dim response As IRestResponse

        Dim H As New HmacAuthHeader
        Dim hmacHeader$ = H.CalculateAuthorizationHeader(apiID, apiKEY, fqdN, urI, "", "POST")

        request.AddHeader("Host", fqdN)
        request.AddHeader("Content-Type", "application/json")
        request.AddHeader("Authorization", hmacHeader)
        request.AddParameter("text/json", jsoN, ParameterType.RequestBody)
        response = client.Execute(request)

        ' Console.WriteLine("PAYLOAD:" + jsoN)

        Console.WriteLine(response.StatusCode.ToString + "/" + response.Content + " " + response.ErrorMessage + " " + response.StatusDescription)
    End Sub

    Public Sub getAppProfiles(apiID$, apiKEY$)
        Dim urI$ = "/was/configservice/v1/platform_applications?size=500"
        Dim client = New RestClient("https://" + fqdN + urI)
        Dim request = New RestRequest(Method.GET)
        Dim response As IRestResponse

        Dim H As New HmacAuthHeader
        Dim hmacHeader$ = H.CalculateAuthorizationHeader(apiID, apiKEY, fqdN, urI, "", "GET")

        request.AddHeader("Host", fqdN)
        request.AddHeader("Authorization", hmacHeader)

        response = client.Execute(request)

        Dim jsoN$ = cleanJSON(response.Content)
        jsoN = removePaging(jsoN)
        '        Console.WriteLine(jsoN)
        appProfiles = JsonConvert.DeserializeObject(Of List(Of appProfile))(jsoN)


        '        Console.WriteLine(response.Content)

    End Sub

    Private Function removeDASTextraCapabilities(json$) As String
        Do Until Mid(json, Len(json) - 1, 2) = "},"
            json = Mid(json, 1, Len(json) - 1)
        Loop
        json = Mid(json, 1, Len(json) - 2)
        Return json
    End Function


    Private Function removePaging(json$) As String
        removePaging = ""
        If Len(json) = 0 Then Exit Function
        json = Mid(json, 1, InStr(json, Chr(34) + "page" + Chr(34) + " :") - 1)
        Do Until Mid(json, Len(json), 1) = "]"
            json = Mid(json, 1, Len(json) - 1)
        Loop
        Return json
    End Function
    Private Function cleanJSON(json$) As String
        If Len(json) = 0 Then Return ""

        json = Mid(json, InStr(json, "["))
        If Mid(json, Len(json), 1) = "}" Then json = Mid(json, 1, Len(json) - 1)
        json = Replace(json, "null", "0")
        'Console.WriteLine("CLEAN:" + vbCrLf + json)
        Return json


    End Function

End Class
Public Class appProfile
    '          "uuid": "db10a568-7f6e-433e-8ca1-1b5cfd3dd509",
    '              "id": 969980,
    '              "name": "Verademo_Java",
    '              "linked_scan_target_url": "http://verademo.mhorty.vuln.sa.veracode.io/verademo"
    Public uuid As String
    Public id As Long
    Public name As String
    Public linked_scan_target_url As String
End Class

Public Class dastAnalysis
    Public name As String
    Public analysis_id As String
    Public number_of_scans As Integer
    Public occurrenceS As List(Of dastAnalOccur)
    Public scansOfOccurrence As List(Of scanDetail)
End Class
Public Class dastAnalOccur
    Public analysis_occurrence_id$
    Public analysis_id$
    Public actual_start_date$
    Public actual_end_date$

End Class

'Public Class dastDetail
'    Public name$
'    Public analysis_id$
'    Public number_of_scans As Integer
'    Public latest_scanid$
'    Public target_url$
'    Public linked_platform_app_uuid$
'    Public linked_platform_app_name$
'    Public result_import_status$
'End Class

Public Class scanDetail
    Public end_date$
    Public scan_id$
    Public target_url$
    Public analysis_occurrence_id$
    Public linked_platform_app_uuid$
    Public linked_platform_app_name$
    Public result_import_status$
    Public total_flaw_count As Long
    Public app_link_type$
End Class