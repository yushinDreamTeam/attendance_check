<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Function StrValue(v)
    Dim value
    If IsObject(v) Then value = v.Value Else value = v
    If IsNull(value) Or IsEmpty(value) Then StrValue = "" Else StrValue = CStr(value)
End Function

Function H(v)
    H = Server.HTMLEncode(StrValue(v))
End Function

Function PositiveID(v)
    Dim s, re
    PositiveID = 0
    s = StrValue(v)
    Set re = New RegExp
    re.Pattern = "^[1-9][0-9]{0,9}$"
    If Not re.Test(s) Then Exit Function
    If CDbl(s) > 2147483647 Then Exit Function
    PositiveID = CLng(s)
End Function

Function BitIsSet(bits, id)
    BitIsSet = False
    If id < 1 Then Exit Function
    If Len(StrValue(bits)) < id Then Exit Function
    BitIsSet = (Mid(bits, id, 1) = "1")
End Function

Sub JsonReply(status, ok, message)
    Response.Status = status
    Response.ContentType = "application/json"
    Dim flag
    flag = "false"
    If ok Then flag = "true"
    Response.Write "{""ok"":" & flag & ",""message"":" & JsonText(message) & "}"
    Response.End
End Sub

Function ParamText(v)
    Dim s, kind
    s = StrValue(v)
    kind = 202
    If Len(s) > 255 Then kind = 203
    ParamText = Array(kind, Len(s) + 1, s)
End Function

Function ParamID(v)
    ParamID = Array(3, 4, CLng(v))
End Function

Function Rows(c, sql, params)
    Dim cmd
    Set cmd = NewCommand(c, sql, params)
    Set Rows = cmd.Execute
End Function

Function RunSQL(c, sql, params)
    Dim cmd, affected
    Set cmd = NewCommand(c, sql, params)
    cmd.Execute affected, , 128
    RunSQL = affected
End Function

Function SignatureParam(c, id)
    Dim r, t
    Set r = c.Execute("SELECT [사용자id] FROM [signature_data] WHERE 1=0")
    t = r.Fields(0).Type
    r.Close
    If t = 202 Or t = 200 Or t = 130 Or t = 129 Then
        SignatureParam = ParamText(CStr(id))
    ElseIf t = 2 Or t = 3 Then
        SignatureParam = ParamID(id)
    Else
        Err.Raise vbObjectError + 101, "DB", "signature_data의 사용자id 필드 형식을 확인해 주세요."
    End If
End Function

Function JsonText(s)
    Dim i, c, code, result
    result = ""
    For i = 1 To Len(StrValue(s))
        c = Mid(s, i, 1)
        code = AscW(c)
        Select Case code
            Case 34: result = result & "\" & Chr(34)
            Case 92: result = result & "\\"
            Case Else
                If code >= 0 And code <= 31 Then
                    result = result & "\u" & Right("0000" & Hex(code), 4)
                Else
                    result = result & c
                End If
        End Select
    Next
    JsonText = Chr(34) & result & Chr(34)
End Function

Function NewCommand(c, sql, params)
    Dim cmd, p, i
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = c
    cmd.CommandType = 1
    cmd.CommandText = sql
    For i = 0 To UBound(params)
        p = params(i)
        cmd.Parameters.Append cmd.CreateParameter("p" & i, p(0), 1, p(1), p(2))
    Next
    Set NewCommand = cmd
End Function
%>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Dim c, lessonID, teacherID, userName, data, path, bytes, fso, stream, oldStream
Dim transactionOpen, changedFile, hadOriginal, msg, restoreError

Function DecodePNG(value)
    Dim doc, node, re, encoded, raw, hexData, imageWidth, imageHeight
    If Left(value,22)<>"data:image/png;base64," Then Err.Raise vbObjectError+10,"서명","PNG 서명 데이터가 아닙니다."
    encoded=Mid(value,23)
    Set re=New RegExp
    re.Pattern="^[A-Za-z0-9+/]+={0,2}$"
    If Len(encoded)=0 Or Len(encoded)>87000 Or Len(encoded) Mod 4<>0 Or Not re.Test(encoded) Then Err.Raise vbObjectError+11,"서명","서명 데이터 크기나 형식을 확인해 주세요."
    Set doc=Server.CreateObject("MSXML2.DOMDocument.6.0")
    Set node=doc.createElement("png")
    node.dataType="bin.base64"
    node.text=encoded
    raw=node.nodeTypedValue
    node.dataType="bin.hex"
    node.nodeTypedValue=raw
    hexData=LCase(node.text)
    If Len(hexData)<90 Or Len(hexData)>130000 Then Err.Raise vbObjectError+12,"서명","서명 이미지 크기를 확인해 주세요."
    If Left(hexData,32)<>"89504e470d0a1a0a0000000d49484452" Or Right(hexData,24)<>"0000000049454e44ae426082" Then Err.Raise vbObjectError+13,"서명","PNG 서명 데이터가 손상되었습니다."
    imageWidth=CLng("&H" & Mid(hexData,33,8))
    imageHeight=CLng("&H" & Mid(hexData,41,8))
    If imageWidth<1 Or imageHeight<1 Or imageWidth>1200 Or imageHeight>600 Then Err.Raise vbObjectError+14,"서명","서명 이미지 가로·세로 크기를 확인해 주세요."
    DecodePNG=raw
End Function

' 대상 연수와 이름을 확인하고 같은 연수/사용자 행에 저장합니다.
Function ResolveTeacher(c, lessonID, userName)
    Dim r, bits, candidates, id
    Set r = Rows(c, "SELECT [참여대상] FROM [lesson_list] WHERE [ID]=?", Array(ParamID(lessonID)))
    If r.EOF Then Err.Raise vbObjectError + 128, "연수", "해당 연수가 없습니다. QR을 다시 확인해 주세요."
    bits = StrValue(r("참여대상")): r.Close
    Set r = Rows(c, "SELECT [ID] FROM [teachers] WHERE [이름]=?", Array(ParamText(userName)))
    candidates = 0: id = 0
    Do Until r.EOF
        If BitIsSet(bits, CLng(r("ID"))) Then
            id = CLng(r("ID")): candidates = candidates + 1
        End If
        r.MoveNext
    Loop
    r.Close
    If candidates = 0 Then Err.Raise vbObjectError + 129, "대상자", "이름을 확인해 주세요. 등록된 이름과 일치하는 연수 대상자가 없습니다."
    If candidates > 1 Then Err.Raise vbObjectError + 130, "대상자", "같은 이름의 대상자가 여러 명입니다. 관리자에게 이름 구분을 요청해 주세요."
    ResolveTeacher = id
End Function
Sub UpdateSignature(c, lessonID, teacherID, path)
    Dim r, count, affected
    Set r = Rows(c, "SELECT COUNT(*) AS n FROM [signature_data] WHERE [연수id]=? AND [사용자id]=?", Array(ParamID(lessonID), SignatureParam(c, teacherID)))
    count = CLng(r("n")): r.Close
    If count > 1 Then Err.Raise vbObjectError + 131, "DB", "중복 서명 행이 있습니다. 관리자에게 확인을 요청해 주세요."
    If count = 0 Then
        affected = RunSQL(c, "INSERT INTO [signature_data] ([연수id],[사용자id],[파일경로]) VALUES (?,?,?)", Array(ParamID(lessonID), SignatureParam(c, teacherID), ParamText(path)))
    Else
        affected = RunSQL(c, "UPDATE [signature_data] SET [파일경로]=? WHERE [연수id]=? AND [사용자id]=?", Array(ParamText(path), ParamID(lessonID), SignatureParam(c, teacherID)))
    End If
    If affected <> 1 Then Err.Raise vbObjectError + 132, "DB", "서명 저장 결과를 확인할 수 없습니다. 다시 시도해 주세요."
End Sub

Sub SaveSignature()
    bytes=DecodePNG(data)
    Set c=Server.CreateObject("ADODB.Connection")
    c.Open "DSN=attendanceDB"
    c.BeginTrans
    transactionOpen=True
    teacherID=ResolveTeacher(c,lessonID,userName)
    path="signatures/sig_" & lessonID & "_" & teacherID & ".png"
    Set fso=Server.CreateObject("Scripting.FileSystemObject")
    hadOriginal=fso.FileExists(Server.MapPath(path))
    ' DB 저장 실패 시 기존 파일을 덮어쓰지 않도록 메모리에 보관합니다.
    If hadOriginal Then
        Set oldStream=Server.CreateObject("ADODB.Stream")
        oldStream.Type=1:oldStream.Open:oldStream.LoadFromFile Server.MapPath(path)
    End If
    Set stream=Server.CreateObject("ADODB.Stream")
    stream.Type=1:stream.Open:stream.Write bytes
    changedFile=True
    stream.SaveToFile Server.MapPath(path),2
    stream.Close
    UpdateSignature c,lessonID,teacherID,path
    c.CommitTrans
    transactionOpen=False
End Sub

If Request.ServerVariables("REQUEST_METHOD")<>"POST" Then JsonReply "405 Method Not Allowed",False,"POST 요청만 사용할 수 있습니다."
If Request.TotalBytes>100000 Then JsonReply "413 Payload Too Large",False,"서명이 너무 큽니다. 지운 뒤 다시 작성해 주세요."
lessonID=PositiveID(Request.Form("lessonId"))
userName=Trim(StrValue(Request.Form("userName")))
data=StrValue(Request.Form("signature"))
If lessonID=0 Or userName="" Or Len(userName)>255 Then JsonReply "400 Bad Request",False,"연수와 이름을 확인해 주세요."

' 동시에 제출해도 동일 대상자의 행을 두 번 추가하지 않도록 저장만 순서대로 처리합니다.
Application.Lock
transactionOpen=False:changedFile=False:hadOriginal=False
On Error Resume Next
SaveSignature
msg=Err.Description
Err.Clear
If msg<>"" Then
    If transactionOpen Then c.RollbackTrans
    Err.Clear
    If changedFile Then
        If hadOriginal Then
            oldStream.Position=0
            oldStream.SaveToFile Server.MapPath(path),2
        ElseIf fso.FileExists(Server.MapPath(path)) Then
            fso.DeleteFile Server.MapPath(path),True
        End If
    End If
    restoreError=Err.Description
    If restoreError<>"" Then msg="서명 파일 복원에 실패했습니다. 관리자에게 확인을 요청해 주세요."
End If
If IsObject(oldStream) Then oldStream.Close
If IsObject(stream) Then stream.Close
If IsObject(c) Then c.Close
Err.Clear
On Error GoTo 0
Application.UnLock
If msg<>"" Then JsonReply "400 Bad Request",False,"저장하지 못했습니다. " & msg
Session("signedLesson")=lessonID
Session("signedTeacher")=teacherID
JsonReply "200 OK",True,"서명이 저장되었습니다."
%>
