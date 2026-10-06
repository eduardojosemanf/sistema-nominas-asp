<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/funciones_nomina.asp"-->
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Asistencias</title>
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
        <h2>Asistencias</h2>
        <a href="asistencia_form.asp?action=add" class="btn btn-primary mb-3">Nueva asistencia</a>

        <table class="table table-striped table-bordered">
          <thead>
            <tr>
              <th>Empleado</th>
              <th>Entrada</th>
              <th>Salida</th>
              <th>Intervalo</th>
              <th>Horas</th>
              <th>Acciones</th>
            </tr>
          </thead>
          <tbody>
            <%
            Dim rsA
            Set rsA = Server.CreateObject("ADODB.Recordset")
            rsA.Open "SELECT a.*, e.nombre FROM asistencias a INNER JOIN empleados e ON e.id = a.empleado_id ORDER BY a.fecha_entrada DESC", cn, 3, 1

            Do While Not rsA.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & rsA("nombre") & "</td>"
                Response.Write "<td>" & rsA("fecha_entrada") & " " & rsA("hora_entrada") & "</td>"
                Response.Write "<td>" & rsA("fecha_salida") & " " & rsA("hora_salida") & "</td>"
                Response.Write "<td>" & rsA("intervalo_inicio") & " - " & rsA("intervalo_fin") & "</td>"
                Response.Write "<td>" & rsA("total_horas") & "</td>"
                Response.Write "<td>"
                Response.Write "<a href='asistencia_form.asp?action=edit&id=" & rsA("id") & "' class='btn btn-sm btn-warning'>Editar</a> "
                Response.Write "<a href='asistencia_form.asp?action=delete&id=" & rsA("id") & "' class='btn btn-sm btn-danger' onclick=\"return confirm('¿Eliminar asistencia?')\">Eliminar</a>"
                Response.Write "</td>"
                Response.Write "</tr>"
                rsA.MoveNext
            Loop
            rsA.Close
            Set rsA = Nothing
            %>
          </tbody>
        </table>
      </main>
    </div>
  </div>
</body>
</html>
