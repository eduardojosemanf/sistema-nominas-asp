<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/validaciones.asp"-->
<%
Dim accion, id, nombre, documento, categoria, tarifaHora, tarifaJornada, activo, mensaje
accion = Request.QueryString("action")
id = Request.QueryString("id")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    nombre = Request.Form("nombre")
    documento = Request.Form("documento")
    categoria = Request.Form("categoria")
    tarifaHora = Request.Form("tarifaHora")
    tarifaJornada = Request.Form("tarifaJornada")
    activo = Request.Form("activo")

    mensaje = ValidarEmpleado(nombre)

    If mensaje <> "" Then
        Response.Redirect "empleado_form.asp?action=add&error=" & Server.URLEncode(mensaje)
    Else
        If accion = "edit" Then
            Dim sqlUpdate
            sqlUpdate = "UPDATE empleados SET nombre='" & Replace(nombre, "'", "''") & "', documento='" & Replace(documento, "'", "''") & "', categoria='" & categoria & "', tarifa_hora=" & Replace(tarifaHora, ",", ".") & ", tarifa_jornada=" & Replace(tarifaJornada, ",", ".") & ", activo=" & IIf(activo = "on", "True", "False") & " WHERE id=" & id
            cn.Execute(sqlUpdate)
            Response.Redirect "empleados.asp?msg=" & Server.URLEncode("Empleado actualizado.")
        Else
            Dim sqlInsert
            sqlInsert = "INSERT INTO empleados (nombre, documento, categoria, tarifa_hora, tarifa_jornada, activo) VALUES ('" & Replace(nombre, "'", "''") & "', '" & Replace(documento, "'", "''") & "', '" & categoria & "', " & Replace(tarifaHora, ",", ".") & ", " & Replace(tarifaJornada, ",", ".") & ", " & IIf(activo = "on", "True", "False") & ")"
            cn.Execute(sqlInsert)
            Response.Redirect "empleados.asp?msg=" & Server.URLEncode("Empleado registrado correctamente.")
        End If
    End If
End If

If accion = "delete" And id <> "" Then
    cn.Execute "DELETE FROM empleados WHERE id=" & id
    Response.Redirect "empleados.asp?msg=" & Server.URLEncode("Empleado eliminado.")
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Empleado</title>
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
        <h2><%= IIf(accion = "edit", "Editar empleado", "Nuevo empleado") %></h2>

        <%
        If Request.QueryString("error") <> "" Then
            Response.Write "<div class='alert alert-danger'>" & Request.QueryString("error") & "</div>"
        End If
        %>

        <form method="post">
          <div class="mb-3">
            <label>Nombre</label>
            <input type="text" name="nombre" class="form-control" required>
          </div>

          <div class="mb-3">
            <label>Documento</label>
            <input type="text" name="documento" class="form-control">
          </div>

          <div class="mb-3">
            <label>Categoría</label>
            <select name="categoria" class="form-control">
              <option value="normal">Normal</option>
              <option value="sereno">Sereno</option>
            </select>
          </div>

          <div class="mb-3">
            <label>Tarifa por hora</label>
            <input type="text" name="tarifaHora" class="form-control" value="0">
          </div>

          <div class="mb-3">
            <label>Tarifa por jornada</label>
            <input type="text" name="tarifaJornada" class="form-control" value="0">
          </div>

          <div class="mb-3 form-check">
            <input type="checkbox" name="activo" class="form-check-input" checked>
            <label class="form-check-label">Activo</label>
          </div>

          <button type="submit" class="btn btn-success">Guardar</button>
          <a href="empleados.asp" class="btn btn-secondary">Cancelar</a>
        </form>
      </main>
    </div>
  </div>
</body>
</html>
