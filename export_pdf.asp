<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Reporte PDF</title>
  <style>
    body { font-family: Arial; font-size: 12px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #666; padding: 8px; }
    h2 { margin-bottom: 20px; }
  </style>
</head>
<body>
  <h2>Reporte de nómina</h2>
  <table>
    <tr>
      <th>Empleado</th>
      <th>Horas</th>
      <th>Subtotal</th>
      <th>Anticipos</th>
      <th>Neto</th>
    </tr>
    <%
    Dim desdePdf, hastaPdf, empleadoIdPdf
    desdePdf = Request.QueryString("desde")
    hastaPdf = Request.QueryString("hasta")
    empleadoIdPdf = Request.QueryString("empleadoId")

    Dim rsPdf
    Set rsPdf = Server.CreateObject("ADODB.Recordset")

    Dim sqlPdf
    If empleadoIdPdf <> "0" And Len(empleadoIdPdf) > 0 Then
        sqlPdf = "SELECT * FROM empleados WHERE id=" & empleadoIdPdf & " AND activo=True ORDER BY nombre"
    Else
        sqlPdf = "SELECT * FROM empleados WHERE activo=True ORDER BY nombre"
    End If

    rsPdf.Open sqlPdf, cn, 3, 1
    Do While Not rsPdf.EOF
        Dim horasPdf, subtotalPdf, anticiposPdf, netoPdf
        horasPdf = 0
        subtotalPdf = 0
        anticiposPdf = 0
        netoPdf = 0

        Dim rsHorasPdf
        Set rsHorasPdf = Server.CreateObject("ADODB.Recordset")
        rsHorasPdf.Open "SELECT SUM(total_horas) AS horas FROM asistencias WHERE empleado_id=" & rsPdf("id") & " AND fecha_entrada BETWEEN #" & desdePdf & "# AND #" & hastaPdf & "#", cn, 3, 1
        If Not rsHorasPdf.EOF Then
            If IsNull(rsHorasPdf("horas")) Then
                horasPdf = 0
            Else
                horasPdf = CDbl(rsHorasPdf("horas"))
            End If
        End If
        rsHorasPdf.Close
        Set rsHorasPdf = Nothing

        Dim rsAntPdf
        Set rsAntPdf = Server.CreateObject("ADODB.Recordset")
        rsAntPdf.Open "SELECT SUM(monto) AS total FROM anticipos WHERE empleado_id=" & rsPdf("id") & " AND fecha BETWEEN #" & desdePdf & "# AND #" & hastaPdf & "#", cn, 3, 1
        If Not rsAntPdf.EOF Then
            If IsNull(rsAntPdf("total")) Then
                anticiposPdf = 0
            Else
                anticiposPdf = CDbl(rsAntPdf("total"))
            End If
        End If
        rsAntPdf.Close
        Set rsAntPdf = Nothing

        If UCase(rsPdf("categoria")) = "SERENO" Then
            subtotalPdf = CDbl(rsPdf("tarifa_jornada"))
        Else
            subtotalPdf = horasPdf * CDbl(rsPdf("tarifa_hora"))
        End If

        netoPdf = subtotalPdf - anticiposPdf

        Response.Write "<tr>"
        Response.Write "<td>" & rsPdf("nombre") & "</td>"
        Response.Write "<td>" & FormatNumber(horasPdf, 2) & "</td>"
        Response.Write "<td>" & FormatCurrency(subtotalPdf, 2) & "</td>"
        Response.Write "<td>" & FormatCurrency(anticiposPdf, 2) & "</td>"
        Response.Write "<td>" & FormatCurrency(netoPdf, 2) & "</td>"
        Response.Write "</tr>"

        rsPdf.MoveNext
    Loop
    rsPdf.Close
    Set rsPdf = Nothing
    %>
  </table>
</body>
</html>
