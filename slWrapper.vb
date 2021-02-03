Imports RestSharp
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class studentLessonStats
    Public numVisits As Integer
    Public tlTime As Decimal
    Public startRating As Integer
    Public endRating As Integer
    Public points As Integer
End Class


Public Class stats
    Public maxTime As Decimal
    Public minTime As Decimal
    Public avgTime As Decimal
    Public numStudents As Integer
    Public avgImprove As Decimal
    Public avgBefore As Decimal
    Public avgAfter As Decimal
    Public numStudentsRevisit As Integer
    Public numRevisitsTL As Integer
    Public totalPoints As Long
    Public quickestStudent$
    Public slowestStudent$

    Public Sub New()
        maxTime = 0 : minTime = 0 : avgTime = 0
        numStudents = 0 : avgImprove = 0
    End Sub
End Class

Public Class slStudent
    Public id$
    Public email$
    Public name$
    Public roles$()
    Public isAdmin As Boolean
    Public isDisabled As Boolean
    Public lastActive As Long
    Public joined As Boolean
    Public points As Integer
    Public lessons As List(Of slLesson)
End Class

Public Class slLesson
    Public module$
    Public lessonId$
    Public lessonName$
    Public lastVisited$
    Public status$
    Public minutes As Long
    Public points As Integer
    Public startRating As Integer
    Public endRating As Integer
End Class


Public Class SL_Client
    Public lastError$ = ""
    Public tokeN$ 'token used for REST calls
    Public fqdN$ 'ie http://localhost https://myserver.myzone.com

    Public studentList As List(Of slStudent)
    Public moduleNames As Collection

    Private slToken$

    Public Sub New(auth_Token$)
        fqdN = "https://securitylabs.veracode.com" '/api/users

        slToken = auth_Token

        Console.WriteLine("New SL_Client activated")

        Dim client = New RestClient(fqdN + "/api/users")
        Dim request = New RestRequest(Method.GET)
        Dim response As IRestResponse

        request.AddHeader("auth", slToken)
        response = client.Execute(request)
        'Console.WriteLine(response.Content)

        Call parseStudents(response.Content)
        Console.WriteLine("# of Students: " + studentList.Count.ToString)
    End Sub


    Public Sub loadProgressAllUsers()
        moduleNames = New Collection
        For Each S In studentList
            Dim client = New RestClient(fqdN + "/api/users/" + S.id + "/progress")
            Dim request = New RestRequest(Method.GET)
            Dim response As IRestResponse

            request.AddHeader("auth", slToken)
            response = client.Execute(request)

            S.lessons = JsonConvert.DeserializeObject(Of List(Of slLesson))(cleanJSON(response.Content))

            For Each L In S.lessons
                Dim a$ = L.module
                If grpNDX(moduleNames, a) = 0 Then moduleNames.Add(a)
                S.points += L.points
            Next
            Console.WriteLine(S.name + ": " + S.lessons.Count.ToString + " Lessons/ Total Points: " + S.points.ToString)

        Next
    End Sub
    Private Function cleanJSON(json$) As String
        If Len(json) = 0 Then Return ""

        json = Mid(json, InStr(json, "["))
        If Mid(json, Len(json), 1) = "}" Then json = Mid(json, 1, Len(json) - 1)
        json = Replace(json, "null", "0")
        'Console.WriteLine("CLEAN:" + vbCrLf + json)
        Return json


    End Function

    Private Sub parseStudents(sList$)
        sList = cleanJSON(sList)
        '        deserializeDeviceResponse = JsonConvert.DeserializeObject(Of List(Of deviceData))(jList)

        studentList = New List(Of slStudent)
        studentList = JsonConvert.DeserializeObject(Of List(Of slStudent))(sList)

    End Sub

    Private Sub setError(ByVal theError$)
        '    MainUI.addLOG("ERROR: " + theError)
        '    lastError = Now.ToLongTimeString + " - " + theError
    End Sub



    Public Function getJSONObject(key$, json$) As String
            On Error GoTo errorcatch

            Dim sObject = JsonConvert.DeserializeObject(json)
            Return sObject(key)

errorcatch:
            Return ""
        End Function



    End Class
