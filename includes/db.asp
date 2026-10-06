<%
Dim cn
Set cn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath(".") & "\bd\nomina_v2.mdb;Persist Security Info=False;"
If Err.Number <> 0 Then
    Response.Write "<h1>Error de conexión a base de datos</h1>"
    Response.Write "<p>Error: " & Err.Description & "</p>"
    Response.Write "<p>Verifica que:</p>"
    Response.Write "<ul>"
    Response.Write "<li>La carpeta bd/ existe en: " & Server.MapPath(".") & "\bd\</li>"
    Response.Write "<li>El archivo nomina_v2.mdb existe</li>"
    Response.Write "<li>Tienes instalado Microsoft Access Database Engine</li>"
    Response.Write "</ul>"
    Response.End
End If
On Error GoTo 0
%>
