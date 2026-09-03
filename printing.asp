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

Function ISODate(d)
    Dim value
    If IsObject(d) Then value = d.Value Else value = d
    ISODate = ""
    If IsNull(value) Then Exit Function
    If Not IsDate(value) Then Exit Function
    ISODate = Year(value) & "-" & Right("0" & Month(value), 2) & "-" & Right("0" & Day(value), 2)
End Function

Function ParamID(v)
    ParamID = Array(3, 4, CLng(v))
End Function

Function Rows(c, sql, params)
    Dim cmd
    Set cmd = NewCommand(c, sql, params)
    Set Rows = cmd.Execute
End Function

Function SafeSignaturePath(path, lessonID, teacherID)
    Dim expected
    expected = "signatures/sig_" & lessonID & "_" & teacherID & ".png"
    SafeSignaturePath = ""
    If LCase(Replace(StrValue(path), "\", "/")) = expected Then SafeSignaturePath = expected
End Function

Sub FailPage(status, message)
    Response.Status = status
    Response.Write "<meta charset='utf-8'><p>" & H(message) & "</p><a href='javascript:history.back()'>돌아가기</a>"
    Response.End
End Sub

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
Dim c,r,msg,lessonID,title,lessonDateValue,bits,signatures,people(),total,id,path,fso,pageCount,p,b,row,index,missing,key,entry,cacheVersion
cacheVersion=Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
lessonID=PositiveID(Request.QueryString("id"))
If lessonID=0 Then FailPage "400 Bad Request","잘못된 연수 ID입니다."
On Error Resume Next
Set c=Server.CreateObject("ADODB.Connection"):c.Open "DSN=attendanceDB":msg=Err.Description
On Error GoTo 0
If msg<>"" Then FailPage "503 Service Unavailable",msg
Set r=Rows(c,"SELECT * FROM [lesson_list] WHERE [ID]=?",Array(ParamID(lessonID)))
If r.EOF Then r.Close:c.Close:FailPage "404 Not Found","해당 연수가 없습니다."
title=StrValue(r("연수제목")):lessonDateValue=ISODate(r("연수일시")):bits=StrValue(r("참여대상")):r.Close
Set signatures=Server.CreateObject("Scripting.Dictionary")
Set r=Rows(c,"SELECT [사용자id],[파일경로] FROM [signature_data] WHERE [연수id]=?",Array(ParamID(lessonID)))
Do Until r.EOF
    id=PositiveID(r("사용자id"))
    If id>0 Then
        key=CStr(id)
        If signatures.Exists(key) Then r.Close:c.Close:FailPage "409 Conflict","동일 대상자의 서명 행이 중복되어 있습니다. DB 자료를 확인해 주세요."
        signatures(key)=StrValue(r("파일경로"))
    End If
    r.MoveNext
Loop
r.Close
Set fso=Server.CreateObject("Scripting.FileSystemObject")
Set r=c.Execute("SELECT [ID],[이름] FROM [teachers] ORDER BY [이름],[ID]")
total=0:missing=0
Do Until r.EOF
    id=CLng(r("ID"))
    If BitIsSet(bits,id) Then
        path=""
        If signatures.Exists(CStr(id)) Then
            If signatures(CStr(id))<>"" Then
                path=SafeSignaturePath(signatures(CStr(id)),lessonID,id)
                If path<>"" Then
                    If Not fso.FileExists(Server.MapPath(path)) Then path=""
                End If
                If path="" Then path="MISSING":missing=missing+1
            End If
        End If
        ReDim Preserve people(total)
        people(total)=Array(id,StrValue(r("이름")),path)
        total=total+1
    End If
    r.MoveNext
Loop
r.Close:c.Close
pageCount=1
If total>0 Then pageCount=(total+89)\90
%>
<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title><%=H(title)%> 참석 명부</title><style>
*{box-sizing:border-box}body{margin:0;font-family:"Malgun Gothic","맑은 고딕",sans-serif}button,.btn{padding:9px 15px;border:1px solid #cbd5e1;border-radius:7px;background:white;color:#334155;text-decoration:none;cursor:pointer}button.primary{background:#2563eb;color:white}button:disabled{opacity:.6}.empty{text-align:center;padding:30px}.missing{color:#b42318}
body.print-page{background:#e8edf3;color:#111}.print-toolbar{display:flex;justify-content:center;align-items:center;gap:18px;flex-wrap:wrap;padding:18px}.print-toolbar p{font-size:13px;color:#526079}.print-warning{max-width:190mm;margin:0 auto 14px;color:#b42318}.print-document{overflow:auto;padding:0 16px 30px}.sheet{width:190mm;min-height:260mm;margin:0 auto 24px;padding:10mm 0 5mm;background:#fff;box-shadow:0 2px 12px #0001}.sheet h1{font-size:18px;text-align:center;margin:0 10mm 8mm;overflow-wrap:anywhere;line-height:1.4}.print-meta{display:flex;justify-content:space-between;font-size:11px;border-bottom:1px solid #333;padding-bottom:3mm;margin-bottom:3mm}.roster-columns{display:flex;align-items:flex-start;gap:3mm}.roster{table-layout:fixed;border-collapse:collapse;width:calc((100% - 6mm)/3);flex:1;font-size:10px}.roster th,.roster td{border:1px solid #555;text-align:center;padding:1px 2px;height:6.8mm;vertical-align:middle;line-height:1.2}.roster th{background:#f1f5f9;font-weight:700}.number-col{width:16%}.name-col{width:40%}.signature-col{width:44%}.person-name{overflow-wrap:anywhere;word-break:normal;font-weight:600}.signature{display:block;max-height:6mm;max-width:100%;width:auto;margin:auto;object-fit:contain}.missing{font-size:9px}.page-number{text-align:center;font-size:10px;margin-top:5mm;color:#64748b}
@page{size:A4 portrait;margin:10mm}
@media print{html,body.print-page{margin:0;padding:0;background:#fff;color:#000}.no-print{display:none!important}.print-document{overflow:visible;padding:0}.sheet{width:100%;min-height:0;margin:0;padding:5mm 0 0;box-shadow:none;break-after:page;page-break-after:always}.sheet:last-child{break-after:auto;page-break-after:auto}.roster th{background:#f1f5f9;print-color-adjust:exact;-webkit-print-color-adjust:exact}.roster tr{break-inside:avoid;page-break-inside:avoid}.sheet h1{break-after:avoid}.print-meta{break-after:avoid}}

</style></head><body class="print-page">
<div class="print-toolbar no-print"><a class="btn" href="administer.asp">돌아가기</a><p id="print-status" role="status">서명 이미지를 확인하고 있습니다…</p><button id="print-button" class="primary" type="button" disabled>인쇄 / PDF 저장</button></div>
<%If missing>0 Then%><p class="print-warning no-print">서명 파일 <%=missing%>개를 찾지 못했습니다. 해당 칸을 확인해 주세요.</p><%End If%>
<main class="print-document">
<%For p=0 To pageCount-1%><section class="sheet"><h1><%=H(title)%> 참석 명부</h1><div class="print-meta"><span>일자: <%=H(lessonDateValue)%></span><span>대상 <%=total%>명</span></div>
<%If total=0 Then%><p class="empty">선택된 연수 대상자가 없습니다.</p><%Else%>
<div class="roster-columns"><%For b=0 To 2%><table class="roster"><colgroup><col class="number-col"><col class="name-col"><col class="signature-col"></colgroup><thead><tr><th>연번</th><th>성명</th><th>서명</th></tr></thead><tbody>
<%For row=0 To 29:index=p*90+b*30+row%><tr><%If index<total Then
entry=people(index)%><td><%=index+1%></td><td class="person-name"><%=H(entry(1))%></td><td>
<%If entry(2)="MISSING" Then%><span class="missing">파일 확인 필요</span><%ElseIf entry(2)<>"" Then%><img class="signature" src="<%=H(entry(2))%>?v=<%=cacheVersion%>" alt="<%=H(entry(1))%> 서명"><%End If%>
</td><%Else%><td>&nbsp;</td><td></td><td></td><%End If%></tr><%Next%>
</tbody></table><%Next%></div><%End If%><div class="page-number"><%=p+1%> / <%=pageCount%></div></section><%Next%>
</main><script>
'use strict';
const button=document.getElementById('print-button');
const status=document.getElementById('print-status');
const images=[...document.querySelectorAll('.signature')];
let failed=0;
Promise.all(images.map(image=>new Promise(resolve=>{
  let settled=false;
  function done(ok){if(settled)return;settled=true;clearTimeout(timer);if(!ok){failed++;image.replaceWith(Object.assign(document.createElement('span'),{className:'missing',textContent:'이미지 확인 필요'}));}resolve();}
  const timer=setTimeout(()=>done(false),20000);
  image.addEventListener('load',()=>done(image.naturalWidth>0),{once:true});
  image.addEventListener('error',()=>done(false),{once:true});
  if(image.complete)done(image.naturalWidth>0);
}))).then(()=>{
  const missing=document.querySelectorAll('.missing').length;
  status.textContent=missing?'서명 파일을 확인해야 하는 칸이 '+missing+'개 있습니다.':'인쇄 준비가 완료되었습니다.';
  button.disabled=false;
});
button.addEventListener('click',()=>{
  if(document.querySelector('.missing')&&!window.confirm('확인되지 않은 서명 칸이 있습니다. 이 상태로 출력할까요?'))return;
  window.print();
});

</script></body></html>
