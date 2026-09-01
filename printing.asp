<%@ Language=VBScript CodePage="65001"%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, lessonId, targetBits
    Dim teacherRs, teacherId, teacherName, teacherStatus
    Dim sigMap, sigPath, sigUserId, fallbackUsers
    Dim printNames(), printFiles(), personCount, i, j, colCount
    Dim rowCount, lessonTitle, lessonDateText, lessonDate

    Function SafeStr(value)
        If IsNull(value) Or IsEmpty(value) Then
            SafeStr = ""
        Else
            SafeStr = CStr(value)
        End If
    End Function

    Function SafeLng(value)
        If IsNull(value) Or IsEmpty(value) Then
            SafeLng = 0
        ElseIf IsNumeric(value) Then
            SafeLng = CLng(value)
        Else
            SafeLng = 0
        End If
    End Function

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    lessonId = Request.QueryString("id")
    If Not IsNumeric(lessonId) Then
        Response.Status = "400 Bad Request"
        Response.Write "잘못된 연수 ID입니다."
        Response.End
    End If

    ' 1) 현재 연수의 참여 대상 비트값 조회
    sql = "SELECT [참여대상] FROM lesson_list WHERE ID = " & lessonId
    Set rs = conn.Execute(sql)

    targetBits = ""
    If Not rs.EOF Then
        If Not IsNull(rs("참여대상")) Then
            targetBits = SafeStr(rs("참여대상"))
        End If
    End If

    rs.Close
    Set rs = Nothing

    ' 2) 연수 제목/날짜 조회
    lessonTitle = ""
    lessonDateText = ""
    sql = "SELECT [연수제목], [연수일시] FROM lesson_list WHERE ID = " & lessonId
    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        If Not IsNull(rs("연수제목")) Then
            lessonTitle = SafeStr(rs("연수제목"))
        End If
        If Not IsNull(rs("연수일시")) Then
            lessonDate = rs("연수일시")
            If IsDate(lessonDate) Then
                lessonDateText = Year(lessonDate) & "-" & Right("0" & Month(lessonDate), 2) & "-" & Right("0" & Day(lessonDate), 2)
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing

    ' 3) 서명 경로 맵 준비
    Set sigMap = Server.CreateObject("Scripting.Dictionary")
    Set fallbackUsers = Server.CreateObject("Scripting.Dictionary")
    sql = "SELECT [사용자id], [파일경로] FROM signature_data WHERE [연수id] = " & lessonId
    Set rs = conn.Execute(sql)

    Do While Not rs.EOF
        sigUserId = SafeStr(rs("사용자id"))
        sigPath = ""
        If Not IsNull(rs("파일경로")) Then
            sigPath = SafeStr(rs("파일경로"))
        End If
        sigMap(sigUserId) = sigPath
        fallbackUsers(sigUserId) = 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    ' 4) 참여 대상자만 이름 기준 정렬해서 수집
    personCount = 0
    ReDim printNames(0)
    ReDim printFiles(0)

    sql = "SELECT ID, 이름, 재직상태 FROM teachers ORDER BY 이름"
    Set teacherRs = conn.Execute(sql)

    Do While Not teacherRs.EOF
        teacherId = SafeLng(teacherRs("ID"))
        teacherName = ""
        teacherStatus = ""

        If Not IsNull(teacherRs("이름")) Then
            teacherName = SafeStr(teacherRs("이름"))
        End If

        If Not IsNull(teacherRs("재직상태")) Then
            teacherStatus = SafeStr(teacherRs("재직상태"))
        End If

        If teacherId > 0 Then
            Dim isSelected, userKey
            isSelected = False
            userKey = CStr(teacherId)

            If Len(targetBits) >= teacherId Then
                If Mid(targetBits, teacherId, 1) = "1" Then
                    isSelected = True
                End If
            End If

            If Not isSelected And fallbackUsers.Exists(userKey) Then
                isSelected = True
            End If

            If isSelected Then
                personCount = personCount + 1
                ReDim Preserve printNames(personCount)
                ReDim Preserve printFiles(personCount)

                printNames(personCount - 1) = teacherName

                If sigMap.Exists(userKey) Then
                    printFiles(personCount - 1) = sigMap(userKey)
                Else
                    printFiles(personCount - 1) = ""
                End If
            End If
        End If

        teacherRs.MoveNext
    Loop

    teacherRs.Close
    Set teacherRs = Nothing
    conn.Close
    Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>출석 명단</title>
    <style>
        @page {
            size: A4 portrait;
            margin: 8mm;
        }

        body {
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            background: #ffffff;
            color: #000000;
        }

        .lesson-title {
            margin: 0 0 6px 0;
            font-size: 15pt;
            font-weight: bold;
            text-align: center;
            line-height: 1.2;
        }

        .date-line {
            margin: 0 0 8px 0;
            font-size: 10pt;
            text-align: right;
            line-height: 1.1;
        }

        .sheet {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            border: 1px solid #000;
            background: #fff;
        }

        .sheet td {
            height: 6.05mm;
            border: 1px solid #000;
            padding: 0;
            text-align: center;
            vertical-align: middle;
            background: #fff;
            line-height: 1.1;
            box-sizing: border-box;
        }

        .sheet td:nth-child(odd) {
            width: 24.67mm;
            font-size: 10pt;
        }

        .sheet td:nth-child(even) {
            width: 24.67mm;
            font-size: 10pt;
        }

        .sheet td:nth-child(odd) {
            width: 24.67mm;
        }

        .sheet td:nth-child(even) {
            width: 24.67mm;
        }

        .name-cell {
            font-weight: normal;
        }

        .signature-cell {
            min-height: 6.05mm;
        }

        .signature-cell img {
            max-width: 100%;
            max-height: 18px;
            object-fit: contain;
            display: block;
            margin: 0 auto;
        }

        @media print {
            body {
                padding: 0;
            }
        }
    </style>
</head>
<body>
    <h1 class="lesson-title">연수명 : <%= lessonTitle %></h1>
    <div class="date-line">날짜 : <%= lessonDateText %></div>

    <table class="sheet" aria-label="출석 명단 표">
        <tbody>
            <%
                colCount = 6
                rowCount = 30
                For i = 0 To rowCount - 1
            %>
                <tr>
                    <% For j = 0 To colCount - 1 %>
                        <%
                            Dim cellPersonIndex, cellValue, signaturePath
                            cellPersonIndex = -1
                            If j Mod 2 = 0 Then
                                cellPersonIndex = (i * 3) + (j / 2)
                            Else
                                cellPersonIndex = (i * 3) + ((j - 1) / 2)
                            End If

                            cellValue = ""
                            signaturePath = ""
                            If cellPersonIndex >= 0 And cellPersonIndex < personCount Then
                                If j Mod 2 = 0 Then
                                    cellValue = printNames(cellPersonIndex)
                                Else
                                    signaturePath = printFiles(cellPersonIndex)
                                End If
                            End If
                        %>
                        <% If j Mod 2 = 0 Then %>
                            <td class="name-cell"><%= cellValue %>&nbsp;</td>
                        <% Else %>
                            <td class="signature-cell">
                                <% If Len(signaturePath) > 0 Then %>
                                    <img src="<%= signaturePath %>" alt="서명">
                                <% Else %>
                                    &nbsp;
                                <% End If %>
                            </td>
                        <% End If %>
                    <% Next %>
                </tr>
            <%
                Next
            %>
        </tbody>
    </table>
</body>
</html>
