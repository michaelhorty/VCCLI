
Module modCommon

    Public Function csvTOquotedList(ByVal a$) As String
        Dim b$ = ""

        Dim C As Object
        C = Split(a, ",")

        Dim d$
        d = Chr(34)

        Dim K As Integer
        For K = 0 To UBound(C)
            b += d + C(K) + d + ","
        Next

        b = Mid(b, 1, Len(b) - 1)
        Return b$
    End Function


    Public Function argValue(lookForArg$, ByRef theArgs$()) As String
        ' assumes format "--arg value"
        Dim argNum As Integer = 0

        For Each A In theArgs
            argNum += 1
            If LCase(A) = LCase(lookForArg) Then
                If argNum + 1 > theArgs.Count Then Return ""
                Return theArgs(argNum)
            End If
        Next
        Return ""
    End Function

    Public Function argExist(lookForArg$, ByRef theArgs$()) As String
        ' assumes format "--arg value"
        argExist = False

        For Each A In theArgs
            If LCase(A) = LCase(lookForArg) Then
                Return True
            End If
        Next
        Return False
    End Function


    Public Function grpNDX(ByRef C As Collection, ByRef a$, Optional ByVal caseSensitive As Boolean = True) As Integer
        Dim K As Long
        grpNDX = 0

        If C.Count = 0 Then Exit Function
        If caseSensitive = False Then GoTo dontEvalCase

        For K = 1 To C.Count
            If a = C(K) Then
                grpNDX = K
                Exit Function
            End If
        Next
        Exit Function

dontEvalCase:
        For Each S In C
            K += 1
            If LCase(a) = LCase(S) Then
                grpNDX = K
                Exit Function
            End If
        Next


    End Function

    Public Function safeFilename(ByVal a) As String
        safeFilename = Replace(a, "\", "")
        safeFilename = Replace(safeFilename, "..", "")
    End Function


    Public Function listNDX(ByRef C As List(Of String), ByRef a$) As Integer
        Dim K As Integer = 0
        listNDX = 0

        For K = 0 To C.Count - 1
            If C(K).ToString = a Then
                listNDX = K + 1
                Exit Function
            End If
        Next

    End Function

    Public Function arrNDX(ByRef A$(), ByRef matcH$) As Integer
        'returns 0 if not found, otherwise NDX + 1
        Dim K As Long
        arrNDX = 0
        For K = 0 To UBound(A)
            If Trim(Str(A(K))) = matcH Then
                arrNDX = K + 1
                Exit Function
            End If
        Next
    End Function
    Public Function spaces(howmany As Integer) As String
        spaces = ""
        Dim K As Integer
        For K = 1 To howmany
            spaces += " "
        Next
    End Function
    Public Function removeExtraSpaces(a) As String
        removeExtraSpaces = ""
        If Len(a) = 0 Then Exit Function
        Dim lastSpace As Boolean = False

        Dim K As Integer = 0
        For K = 0 To Len(a) - 1
            If lastSpace = False Then
                removeExtraSpaces += Mid(a, K + 1, 1)
            Else
                If Mid(a, K + 1, 1) <> " " Then removeExtraSpaces += Mid(a, K + 1, 1)
            End If
            If Mid(a, K + 1, 1) = " " Then
                lastSpace = True
            End If
        Next

    End Function

    Public Function countChars(a$, chr2Count$) As Integer
        countChars = 0

        Dim K As Integer = 0
        For K = 0 To Len(a) - 1
            If Mid(a, K + 1, 1) = chr2Count Then countChars += 1
        Next
    End Function

    Public Function stripToFilename(ByVal fileN$) As String
        'C:\Program Files\Checkmarx\Checkmarx Jobs Manager\Results\WebGoat.NET.Default 2014-10.9.2016-19.59.35.pdf
        stripToFilename = ""

        Do Until InStr(fileN, "\") = 0
            fileN = Mid(fileN, InStr(fileN, "\") + 1)
        Loop

        stripToFilename = fileN

    End Function

    Public Function addSlash(ByVal a$) As String
        addSlash = a
        If Len(a) = 0 Then Exit Function

        If Mid(a, Len(a), 1) <> "\" Then addSlash += "\"
    End Function

    Public Function getParentGroup(ByVal g$) As String
        Dim a$ = StrReverse(g)
        a = Mid(a, InStr(a, "\") + 1)
        Return StrReverse(a)
    End Function

    Public Function stripLastWord(ByVal g$) As String
        Dim a$ = StrReverse(g)
        a = Mid(a, InStr(a, " ") + 1)
        Return StrReverse(a)
    End Function


    Public Function assembleCollFromCLI(clI$) As Collection
        Dim C As New Collection
        ' takes windows dos-style dir output and makes sense of it for collection storage
        Dim tempStr$ = clI
        Dim K As Integer
        Do Until InStr(tempStr, "  ") = 0
            K = InStr(tempStr, "  ")
            If Len(Mid(tempStr, 1, K - 1)) Then C.Add(Mid(tempStr, 1, K - 1))
            tempStr = Replace(tempStr, Mid(tempStr, 1, K - 1) + "  ", "")
            'Debug.Print(tempStr)
        Loop
        tempStr = LTrim(tempStr)
        C.Add(Mid(tempStr, 1, InStr(tempStr, " ") - 1))
        tempStr = Replace(tempStr, Mid(tempStr, 1, InStr(tempStr, " ") - 1), "")
        C.Add(LTrim(tempStr))
        Return C

    End Function

    Public Function CSVtoCOLL(ByRef csV$) As Collection
        CSVtoCOLL = New Collection

        Dim splitCHR$ = ","
        If InStr(csV, splitCHR) = 0 Then splitCHR = ";"


        Dim longS = Split(csV, splitCHR)

        Dim K As Integer
        For K = 0 To UBound(longS)
            CSVtoCOLL.Add(longS(K))
        Next

    End Function


    Public Function CSVFiletoCOLL(ByRef csV$) As Collection
        CSVFiletoCOLL = New Collection
        If Dir(csV) = "" Then Exit Function

        'use file
        Dim FF As Integer
        FF = FreeFile()

        FileOpen(FF, csV, OpenMode.Input)

        Do Until EOF(FF) = True
            CSVFiletoCOLL.Add(LineInput(FF))
        Loop
        FileClose(FF)

    End Function

    Public Sub safeKILL(ByRef fileN$)
        If Dir(fileN) <> "" Then Kill(fileN)
    End Sub


    Public Function filePROP(fileN$, proP$) As String
        filePROP = ""
        If Dir(fileN) = "" Then Exit Function

        If Len(proP) = 0 Then Exit Function

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""

        a = LineInput(FF)
        Do Until a = "" Or EOF(FF) = True
            If InStr(a, "=") = 0 Then GoTo nextLine

            If UCase(proP) = Mid(a, 1, InStr(a, "=") - 1) Then
                filePROP = Replace(a, proP + "=", "")
            End If
nextLine:
            a = LineInput(FF)
        Loop

        If Len(a) = 0 Then GoTo closeHere

        If UCase(proP) = Mid(a, 1, InStr(a, "=") - 1) Then
            filePROP = Replace(a, proP + "=", "")
        End If

closeHere:

        FileClose(FF)
    End Function

    Public Function allObjectsToList(fileN$) As List(Of String)
        allObjectsToList = New List(Of String)
        Dim C As New Collection
        Call getAllObjNamesFromFile(fileN, C)

        For Each A In C
            allObjectsToList.Add(loadOBJfromFILE(fileN, A))
        Next
    End Function

    Public Sub allObjectsWithProp(ByRef objS As List(Of String), prop$, propValue$, ByRef coll2Fill As Collection)
        coll2Fill = New Collection

        For Each O In objS
            If UCase(objProp(O, UCase(prop))) = UCase(propValue) Then coll2Fill.Add(objProp(O, "NAME"))
        Next
    End Sub



    Public Sub getAllObjNamesFromFile(fileN$, ByRef collOFnames As Collection)
        collOFnames = New Collection

        If Dir(fileN) = "" Then Exit Sub

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""

        a = LineInput(FF)
        Do Until EOF(FF) = True
            If UCase(Mid(a, 1, 5)) = "NAME=" Then
                collOFnames.Add(Replace(a, "NAME" + "=", ""))
            End If
            a = LineInput(FF)
        Loop

        FileClose(FF)

    End Sub

    Public Function loadOBJfromFILE(fileN$, objName$) As String
        loadOBJfromFILE = ""

        If Dir(fileN) = "" Then Exit Function

        If Len(objName) = 0 Then Exit Function

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""
        Dim buildStr$ = ""

        Dim findSTR$ = "NAME=" + UCase(objName)

        a = LineInput(FF)
        Do Until UCase(a) = findSTR Or EOF(FF) = True
nextLine:
            a = LineInput(FF)
        Loop

        If UCase(a) = findSTR Then
            Do Until a = "" Or EOF(FF) = True
                buildStr += a + vbCrLf
                a = LineInput(FF)
            Loop
        End If

        loadOBJfromFILE = buildStr
        FileClose(FF)


    End Function

    Public Function objProp(ByRef ObjString As String, propName$) As String
        objProp = ""
        Dim findS$ = UCase(propName) + "="

        Dim O = Split(ObjString, vbCrLf)

        If UBound(O) = 0 Then Exit Function

        Dim K As Integer

        For K = 0 To UBound(O)
            If Mid(O(K), 1, Len(findS)) = UCase(propName) + "=" Then
                'found object, return property
                objProp = Mid(O(K), InStr(O(K), "=") + 1)
                Exit Function
            End If
        Next

    End Function

    Public Function xlsDataType(dType$) As String
        xlsDataType = "nonefound"
        Select Case dType
            Case "bigint", "int", "numeric", "float"
                xlsDataType = "Numeric"
            Case "datetime", "datetime2"
                xlsDataType = "DateTime"
            Case "date"
                xlsDataType = "Date"
            Case "time"
                xlsDataType = "Time"
            Case "bit"
                xlsDataType = "Boolean"
            Case "ntext", "nvarchar", "nchar", "varchar", "image", "uniqueidentifier", "real"
                xlsDataType = "String"
        End Select
        If xlsDataType = "nonefound" Then
            Debug.Print("No Def: " + dType)
            xlsDataType = "String"
        End If
    End Function

    Public Function xlsColName(colNum As Integer) As String
        Dim d As Integer
        Dim m As Integer
        Dim name As String
        d = colNum
        name = ""
        Do While (d > 0)
            m = (d - 1) Mod 26
            name = Chr(65 + m) + name
            d = Int((d - m) / 26)
        Loop
        xlsColName = name
    End Function


End Module
