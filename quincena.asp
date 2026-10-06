<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/funciones_nomina.asp"-->
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Quincena</title>
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
        <h2>Cálculo de quincena</h2>

        <form method="get">
          <div class="row">
            <div class="col-md-3">
              <label>Empleado</label>
              <select name="empleadoId" class="form-control">
                <option value="0">Todos</option>
                <%
                Dim rsQ
                Set rsQ = Server.CreateObject("ADODB.Recordset")
                rsQ.Open "SELECT * FROM empleados WHERE activo=True ORDER BY nombre", cn, 3, 1
                Do While Not rsQ.EOF
                    Response.Write "<option value='" & rsQ("id") & "'>" & rsQ("nombre") & "</option>"
                    rsQ.MoveNext
                Loop
                rsQ.Close
                Set rsQ = Nothing
                %>
              </select>
            </div>
            <div class="col-md-3">
              <label>Año</label>
              <input type="number" name="anio" class="form-control" value="2026">
            </div>
            <div class="col-md-3">
              <label>Mes</label>
              <input type="number" name="mes" class="form-control" value="10">
            </div>
            <div class="col-md-3">
              <label>Quincena</label>
              <select name="quincena" class="form-control">
                <option value="1">1era</option>
                <option value="2">2da</option>
              </select>
            </div>
          </div>
          <button type="submit" class="btn btn-success mt-3">Calcular</button>
        </form>

        <%
        Dim empleadoIdQ, anioQ, mesQ, quincenaQ
        empleadoIdQ = Request.QueryString("empleadoId")
        anioQ = Request.QueryString("anio")
        mesQ = Request.QueryString("mes")
        quincenaQ = Request.QueryString("quincena")

        If Len(empleadoIdQ) > 0 And Len(anioQ) > 0 And Len(mesQ) > 0 And Len(quincenaQ) > 0 Then
            Response.Write "<table class='table table-bordered table-striped mt-4'>"
            Response.Write "<thead><tr><th>Empleado</th><th>Horas</th><th>Subtotal</th><th>Anticipos</th><th>Neto</th></tr></thead><tbody>"

            Dim rsListaQ
            If empleadoIdQ = "0" Then
                Set rsListaQ = Server.CreateObject("ADODB.Recordset")
                rsListaQ.Open "SELECT * FROM empleados WHERE activo=True ORDER BY nombre", cn, 3, 1
            Else
                Set rsListaQ = Server.CreateObject("ADODB.Recordset")
                rsListaQ.Open "SELECT * FROM empleados WHERE id=" & empleadoIdQ & " AND activo=True ORDER BY nombre", cn, 3, 1
            End If

            Do While Not rsListaQ.EOF
                Dim totalHorasQ, subtotalQ, anticiposQ, netoQ
                totalHorasQ = 0
                subtotalQ = 0
                anticiposQ = 0
                netoQ = 0

                Dim rsHorasQ
                Set rsHorasQ = Server.CreateObject("ADODB.Recordset")
                rsHorasQ.Open "SELECT SUM(total_horas) AS horas FROM asistencias WHERE empleado_id=" & rsListaQ("id") & " AND YEAR(fecha_entrada)=" & anioQ & " AND MONTH(fecha_entrada)=" & mesQ, cn, 3, 1
                If Not rsHorasQ.EOF Then
                    If IsNull(rsHorasQ("horas")) Then
                        totalHorasQ = 0
                    Else
                        totalHorasQ = CDbl(rsHorasQ("horas"))
                    End If
                End If
                rsHorasQ.Close
                Set rsHorasQ = Nothing

                Dim rsAntQ
                Set rsAntQ = Server.CreateObject("ADODB.Recordset")
                rsAntQ.Open "SELECT SUM(monto) AS total FROM anticipos WHERE empleado_id=" & rsListaQ("id") & " AND YEAR(fecha)=" & anioQ & " AND MONTH(fecha)=" & mesQ, cn, 3, 1
                If Not rsAntQ.EOF Then
                    If IsNull(rsAntQ("total")) Then
                        anticiposQ = 0
                    Else
                        anticiposQ = CDbl(rsAntQ("total"))
                    End If
                End If
                rsAntQ.Close
                Set rsAntQ = Nothing

                If UCase(rsListaQ("categoria")) = "SERENO" Then
                    subtotalQ = CDbl(rsListaQ("tarifa_jornada"))
                Else
                    subtotalQ = totalHorasQ * CDbl(rsListaQ("tarifa_hora"))
                End If

                netoQ = subtotalQ - anticiposQ

                Response.Write "<tr>"
                Response.Write "<td>" & rsListaQ("nombre") & "</td>"
                Response.Write "<td>" & FormatNumber(totalHorasQ, 2) & "</td>"
                Response.Write "<td>" & FormatCurrency(subtotalQ, 2) & "</td>"
                Response.Write "<td>" & FormatCurrency(anticiposQ, 2) & "</td>"
                Response.Write "<td>" & FormatCurrency(netoQ, 2) & "</td>"
                Response.Write "</tr>"

                rsListaQ.MoveNext
            Loop
            rsListaQ.Close
            Set rsListaQ = Nothing

            Response.Write "</tbody></table>"
        End If
        %>
      </main>
    </div>
  </div>
</body>
</html>
