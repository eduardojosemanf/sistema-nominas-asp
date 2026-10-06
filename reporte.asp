<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Reporte</title>
  <link href="css/bootstrap.min.css" rel="stylesheet">
  <link href="css/estilo.css" rel="stylesheet">
</head>
<body>
  <div class="container-fluid">
    <div class="row">
      <aside class="col-md-2 sidebar">
        <h4>Nomina</h4>
        <ul class="nav flex-column">
          <li><a href="default.asp">Inicio</a></li>
          <li><a href="empleados.asp">Empleados</a></li>
          <li><a href="asistencias.asp">Asistencias</a></li>
          <li><a href="anticipos.asp">Anticipos</a></li>
          <li><a href="quincena.asp">Quincena</a></li>
          <li><a href="reporte.asp">Reportes</a></li>
        </ul>
      </aside>

      <main class="col-md-10 content">
        <h2>Reporte financiero</h2>

        <form method="get" class="mb-4">
          <div class="row">
            <div class="col-md-3">
              <label>Desde</label>
              <input type="date" name="desde" class="form-control" required>
            </div>
            <div class="col-md-3">
              <label>Hasta</label>
              <input type="date" name="hasta" class="form-control" required>
            </div>
            <div class="col-md-3">
              <label>Empleado</label>
              <select name="empleadoId" class="form-control">
                <option value="0">Todos</option>
                <%
                Dim rsERep
                Set rsERep = Server.CreateObject("ADODB.Recordset")
                rsERep.Open "SELECT * FROM empleados WHERE activo=True ORDER BY nombre", cn, 3, 1
                Do While Not rsERep.EOF
                    Response.Write "<option value='" & rsERep("id") & "'>" & rsERep("nombre") & "</option>"
                    rsERep.MoveNext
                Loop
                rsERep.Close
                Set rsERep = Nothing
                %>
              </select>
            </div>
          </div>
          <button type="submit" class="btn btn-success mt-3">Generar</button>
          <a href="export_excel.asp?desde=<%=Request.QueryString("desde")%>&hasta=<%=Request.QueryString("hasta")%>&empleadoId=<%=Request.QueryString("empleadoId")%>" class="btn btn-info mt-3">Exportar Excel</a>
          <a href="export_pdf.asp?desde=<%=Request.QueryString("desde")%>&hasta=<%=Request.QueryString("hasta")%>&empleadoId=<%=Request.QueryString("empleadoId")%>" class="btn btn-secondary mt-3" target="_blank">Ver PDF</a>
        </form>

        <%
        Dim desdeRep, hastaRep, empleadoIdRep
        desdeRep = Request.QueryString("desde")
        hastaRep = Request.QueryString("hasta")
        empleadoIdRep = Request.QueryString("empleadoId")

        If Len(desdeRep) > 0 And Len(hastaRep) > 0 Then
            If CDate(desdeRep) > CDate(hastaRep) Then
                Response.Write "<div class='alert alert-danger'>La fecha desde no puede ser mayor que la fecha hasta.</div>"
            Else
                Dim rsRep
                Set rsRep = Server.CreateObject("ADODB.Recordset")

                Dim sqlRep
                If empleadoIdRep <> "0" And Len(empleadoIdRep) > 0 Then
                    sqlRep = "SELECT e.id, e.nombre, e.categoria, e.tarifa_hora, e.tarifa_jornada FROM empleados e WHERE e.id=" & empleadoIdRep & " AND e.activo=True"
                Else
                    sqlRep = "SELECT * FROM empleados WHERE activo=True ORDER BY nombre"
                End If

                rsRep.Open sqlRep, cn, 3, 1

                Response.Write "<table class='table table-bordered table-striped'>"
                Response.Write "<thead><tr><th>Empleado</th><th>Horas</th><th>Subtotal</th><th>Anticipos</th><th>Neto</th></tr></thead>"

                Do While Not rsRep.EOF
                    Dim horasRep, subtotalRep, anticiposRep, netoRep
                    horasRep = 0
                    subtotalRep = 0
                    anticiposRep = 0
                    netoRep = 0

                    Dim rsHorasRep
                    Set rsHorasRep = Server.CreateObject("ADODB.Recordset")
                    rsHorasRep.Open "SELECT SUM(total_horas) AS horas FROM asistencias WHERE empleado_id=" & rsRep("id") & " AND fecha_entrada BETWEEN #" & desdeRep & "# AND #" & hastaRep & "#", cn, 3, 1
                    If Not rsHorasRep.EOF Then
                        If IsNull(rsHorasRep("horas")) Then
                            horasRep = 0
                        Else
                            horasRep = CDbl(rsHorasRep("horas"))
                        End If
                    End If
                    rsHorasRep.Close
                    Set rsHorasRep = Nothing

                    Dim rsAntRep
                    Set rsAntRep = Server.CreateObject("ADODB.Recordset")
                    rsAntRep.Open "SELECT SUM(monto) AS total FROM anticipos WHERE empleado_id=" & rsRep("id") & " AND fecha BETWEEN #" & desdeRep & "# AND #" & hastaRep & "#", cn, 3, 1
                    If Not rsAntRep.EOF Then
                        If IsNull(rsAntRep("total")) Then
                            anticiposRep = 0
                        Else
                            anticiposRep = CDbl(rsAntRep("total"))
                        End If
                    End If
                    rsAntRep.Close
                    Set rsAntRep = Nothing

                    If UCase(rsRep("categoria")) = "SERENO" Then
                        subtotalRep = CDbl(rsRep("tarifa_jornada"))
                    Else
                        subtotalRep = horasRep * CDbl(rsRep("tarifa_hora"))
                    End If

                    netoRep = subtotalRep - anticiposRep

                    Response.Write "<tr>"
                    Response.Write "<td>" & rsRep("nombre") & "</td>"
                    Response.Write "<td>" & FormatNumber(horasRep, 2) & "</td>"
                    Response.Write "<td>" & FormatCurrency(subtotalRep, 2) & "</td>"
                    Response.Write "<td>" & FormatCurrency(anticiposRep, 2) & "</td>"
                    Response.Write "<td>" & FormatCurrency(netoRep, 2) & "</td>"
                    Response.Write "</tr>"

                    rsRep.MoveNext
                Loop

                Response.Write "</table>"
                rsRep.Close
                Set rsRep = Nothing
            End If
        End If
        %>
      </main>
    </div>
  </div>
</body>
</html>
