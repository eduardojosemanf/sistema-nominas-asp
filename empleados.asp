<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Empleados</title>
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
        <h2>Empleados</h2>
        <a href="empleado_form.asp?action=add" class="btn btn-primary mb-3">Nuevo empleado</a>

        <table class="table table-striped table-bordered">
          <thead>
            <tr>
              <th>ID</th>
              <th>Nombre</th>
              <th>Documento</th>
              <th>Categoría</th>
              <th>Tarifa H</th>
              <th>Tarifa J</th>
              <th>Activo</th>
              <th>Acciones</th>
            </tr>
          </thead>
          <tbody>
            <%
            Dim rsEmp
            Set rsEmp = Server.CreateObject("ADODB.Recordset")
            rsEmp.Open "SELECT * FROM empleados ORDER BY nombre", cn, 3, 1

            Do While Not rsEmp.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & rsEmp("id") & "</td>"
                Response.Write "<td>" & rsEmp("nombre") & "</td>"
                Response.Write "<td>" & rsEmp("documento") & "</td>"
                Response.Write "<td>" & rsEmp("categoria") & "</td>"
                Response.Write "<td>" & FormatCurrency(rsEmp("tarifa_hora"), 2) & "</td>"
                Response.Write "<td>" & FormatCurrency(rsEmp("tarifa_jornada"), 2) & "</td>"
                Response.Write "<td>" & IIf(rsEmp("activo"), "Activo", "Inactivo") & "</td>"
                Response.Write "<td>"
                Response.Write "<a href='empleado_form.asp?action=edit&id=" & rsEmp("id") & "' class='btn btn-sm btn-warning'>Editar</a> "
                Response.Write "<a href='empleado_form.asp?action=delete&id=" & rsEmp("id") & "' class='btn btn-sm btn-danger' onclick=\"return confirm('¿Eliminar empleado?')\">Eliminar</a>"
                Response.Write "</td>"
                Response.Write "</tr>"
                rsEmp.MoveNext
            Loop
            rsEmp.Close
            Set rsEmp = Nothing
            %>
          </tbody>
        </table>
      </main>
    </div>
  </div>
</body>
</html>
