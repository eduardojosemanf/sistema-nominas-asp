<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=nomina_reporte.xls"

Dim desdeEx, hastaEx, empleadoIdEx
desdeEx = Request.QueryString("desde")
hastaEx = Request.QueryString("hasta")
empleadoIdEx = Request.QueryString("empleadoId")

Response.Write "<table border='1'>"
Response.Write "<tr><th>Empleado</th><th>Horas</th><th>Subtotal</th><th>Anticipos</th><th>Neto</th></tr>"

Dim rsEx
Set rsEx = Server.CreateObject("ADODB.Recordset")

Dim sqlEx
If empleadoIdEx <> "0" And Len(empleadoIdEx) > 0 Then
    sqlEx = "SELECT * FROM empleados WHERE id=" & empleadoIdEx & " AND activo=True ORDER BY nombre"
Else
    sqlEx = "SELECT * FROM empleados WHERE activo=True ORDER BY nombre"
End If

rsEx.Open sqlEx, cn, 3, 1

Do While Not rsEx.EOF
    Dim horasEx, subtotalEx, anticiposEx, netoEx
    horasEx = 0
    subtotalEx = 0
    anticiposEx = 0
    netoEx = 0

    Dim rsHorasEx
    Set rsHorasEx = Server.CreateObject("ADODB.Recordset")
    rsHorasEx.Open "SELECT SUM(total_horas) AS horas FROM asistencias WHERE empleado_id=" & rsEx("id") & " AND fecha_entrada BETWEEN #" & desdeEx & "# AND #" & hastaEx & "#", cn, 3, 1
    If Not rsHorasEx.EOF Then
        If IsNull(rsHorasEx("horas")) Then
            horasEx = 0
        Else
            horasEx = CDbl(rsHorasEx("horas"))
        End If
    End If
    rsHorasEx.Close
    Set rsHorasEx = Nothing

    Dim rsAntEx
    Set rsAntEx = Server.CreateObject("ADODB.Recordset")
    rsAntEx.Open "SELECT SUM(monto) AS total FROM anticipos WHERE empleado_id=" & rsEx("id") & " AND fecha BETWEEN #" & desdeEx & "# AND #" & hastaEx & "#", cn, 3, 1
    If Not rsAntEx.EOF Then
        If IsNull(rsAntEx("total")) Then
            anticiposEx = 0
        Else
            anticiposEx = CDbl(rsAntEx("total"))
        End If
    End If
    rsAntEx.Close
    Set rsAntEx = Nothing

    If UCase(rsEx("categoria")) = "SERENO" Then
        subtotalEx = CDbl(rsEx("tarifa_jornada"))
    Else
        subtotalEx = horasEx * CDbl(rsEx("tarifa_hora"))
    End If

    netoEx = subtotalEx - anticiposEx

    Response.Write "<tr>"
    Response.Write "<td>" & rsEx("nombre") & "</td>"
    Response.Write "<td>" & FormatNumber(horasEx, 2) & "</td>"
    Response.Write "<td>" & FormatCurrency(subtotalEx, 2) & "</td>"
    Response.Write "<td>" & FormatCurrency(anticiposEx, 2) & "</td>"
    Response.Write "<td>" & FormatCurrency(netoEx, 2) & "</td>"
    Response.Write "</tr>"

    rsEx.MoveNext
Loop

rsEx.Close
Set rsEx = Nothing

Response.Write "</table>"
%>
