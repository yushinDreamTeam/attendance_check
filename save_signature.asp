<%@ Language="VBScript" CodePage="65001" %>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, rs, sql
    Dim lessonId, userName, signatureData
    Dim fileName, filePath, physicalPath
    Dim base64Data, dataParts
    Dim xmlDoc, encodedNode, stream
    Dim userId, lessonTargetBits

    lessonId = Request.Form("lessonId")
    userName = Request.Form("userName")
    signatureData = Request.Form("signatureData")

    If Not IsNumeric(lessonId) Then
        Response.Write "잘못된 연수 ID입니다."
        Response.End
    End If

    If Len(Trim(userName)) = 0 Then
        Response.Write "이름을 입력해주세요."
        Response.End
    End If

    If Len(Trim(signatureData)) = 0 Then
        Response.Write "서명 이미지가 없습니다."
        Response.End
    End If

    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    ' 이름으로 사용자 ID 찾기
    sql = "SELECT ID FROM teachers WHERE 이름 = '" & Replace(userName, "'", "''") & "'"
    Set rs = conn.Execute(sql)

    If rs.EOF Then
        Response.Write "해당 이름의 교직원을 찾을 수 없습니다."
        conn.Close
        Set conn = Nothing
        Set rs = Nothing
        Response.End
    End If

    userId = rs("ID")
    rs.Close
    Set rs = Nothing

    ' 해당 연수의 참여 대상인지 확인
    sql = "SELECT [참여대상] FROM lesson_list WHERE ID = " & lessonId
    Set rs = conn.Execute(sql)

    If rs.EOF Then
        Response.Write "존재하지 않는 연수입니다."
        conn.Close
        Set conn = Nothing
        Set rs = Nothing
        Response.End
    End If

    lessonTargetBits = ""
    If Not IsNull(rs("참여대상")) Then
        lessonTargetBits = CStr(rs("참여대상"))
    End If

    rs.Close
    Set rs = Nothing

    If Len(lessonTargetBits) < userId Or Mid(lessonTargetBits, userId, 1) <> "1" Then
        Response.Write "이 이름은 현재 연수 참여 명단에 없습니다."
        conn.Close
        Set conn = Nothing
        Response.End
    End If

    fileName = "sig_" & lessonId & "_" & userId & ".png"
    filePath = "signatures/" & fileName
    physicalPath = Server.MapPath("./" & filePath)

    ' data:image/png;base64,... 부분에서 실제 base64 추출
    dataParts = Split(signatureData, ",")
    If UBound(dataParts) > 0 Then
        base64Data = dataParts(1)
    Else
        base64Data = signatureData
    End If

    ' base64 -> binary 변환
    Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument.6.0")
    Set encodedNode = xmlDoc.CreateElement("root")
    encodedNode.DataType = "bin.base64"
    encodedNode.Text = base64Data

    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write encodedNode.NodeTypedValue
    stream.SaveToFile physicalPath, 2
    stream.Close

    Set stream = Nothing
    Set encodedNode = Nothing
    Set xmlDoc = Nothing

    ' 해당 연수 + 사용자 기준으로 signature_data 찾기
    sql = "SELECT ID FROM signature_data WHERE [연수id] = " & lessonId & " AND [사용자id] = " & userId
    Set rs = conn.Execute(sql)

    If rs.EOF Then
        sql = "INSERT INTO signature_data ([연수id], [사용자id], [파일경로]) VALUES (" & _
              lessonId & ", " & userId & ", '" & filePath & "')"
        conn.Execute sql
    Else
        sql = "UPDATE signature_data SET [파일경로] = '" & filePath & "' WHERE ID = " & rs("ID")
        conn.Execute sql
    End If

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

    Response.Write "서명이 저장되었습니다."
%>
