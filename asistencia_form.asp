<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/funciones_nomina.asp"-->
<!--#include file="includes/validaciones.asp"-->
<%
Dim accion, id, empleadoId, fechaEntrada, horaEntrada, fechaSalida, horaSalida, intervaloInicio, intervaloFin, observaciones, mensaje
accion = Request.QueryString("action")
id = Request.QueryString("id")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    empleadoId = Request.Form("empleadoId")
    fechaEntrada = Request.Form("fechaEntrada")
    horaEntrada = Request.Form("horaEntrada")
    fechaSalida = Request.Form("fechaSalida")
    horaSalida = Request.Form("horaSalida")
    intervaloInicio = Request.Form("intervaloInicio")
    intervaloFin = Request.Form("intervaloFin")
    observaciones = Request.Form("observaciones")

    mensaje = ValidarAsistencia(empleadoId, fechaEntrada, horaEntrada, fechaSalida, horaSalida, intervaloInicio, intervaloFin)

    If mensaje <> "" Then
        Response.Redirect "asistencia_form.asp?action=add&error=" & Server.URLEncode(mensaje)
    Else
        Dim totalHoras
        totalHoras = CalcularHorasTrabajadas(fechaEntrada, horaEntrada, fechaSalida, horaSalida, intervaloInicio, intervaloFin)

        If accion = "edit" Then
            Dim sqlUpdateA
            sqlUpdateA = "UPDATE asistencias SET empleado_id=" & empleadoId & ", fecha_entrada=#" & fechaEntrada & "#, hora_entrada='" & horaEntrada & "', fecha_salida=#" & fechaSalida & "#, hora_salida='" & horaSalida & "', intervalo_inicio='" & intervaloInicio & "', intervalo_fin='" & intervaloFin & "', total_horas=" & Replace(totalHoras, ",", ".") & ", observaciones='" & Replace(observaciones, "'", "''") & "' WHERE id=" & id
            cn.Execute(sqlUpdateA)
            Response.Redirect "asistencias.asp?msg=" & Server.URLEncode("Asistencia actualizada.")
        Else
            Dim sqlInsertA
            sqlInsertA = "INSERT INTO asistencias (empleado_id, fecha_entrada, hora_entrada, fecha_salida, hora_salida, intervalo_inicio, intervalo_fin, total_horas, observaciones) VALUES (" & empleadoId & ", #" & fechaEntrada & "#, '" & horaEntrada & "', #" & fechaSalida & "#, '" & horaSalida & "', '" & intervaloInicio & "', '" & intervaloFin & "', " & Replace(totalHoras, ",", ".") & ", '" & Replace(observaciones, "'", "''") & "')"
            cn.Execute(sqlInsertA)
            Response.Redirect "asistencias.asp?msg=" & Server.URLEncode("Asistencia registrada.")
        End If
    End If
End If

If accion = "delete" And id <> "" Then
    cn.Execute "DELETE FROM asistencias WHERE id=" & id
    Response.Redirect "asistencias.asp?msg=" & Server.URLEncode("Asistencia eliminada.")
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Asistencia</title>
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
        <h2><%= IIf(accion = "edit", "Editar asistencia", "Nueva asistencia") %></h2>

        <% If Request.QueryString("error") <> "" Then Response.Write "<div class='alert alert-danger'>" & Request.QueryString("error") & "</div>" End If %>

        <form method="post">
          <div class="mb-3">
            <label>Empleado</label>
            <select name="empleadoId" class="form-control" required>
              <option value="">Seleccione</option>
              <%
              Dim rsEmpSel
              Set rsEmpSel = Server.CreateObject("ADODB.Recordset")
              rsEmpSel.Open "SELECT * FROM empleados WHERE activo=True ORDER BY nombre", cn, 3, 1
              Do While Not rsEmpSel.EOF
                  Response.Write "<option value='" & rsEmpSel("id") & "'>" & rsEmpSel("nombre") & "</option>"
                  rsEmpSel.MoveNext
              Loop
              rsEmpSel.Close
              Set rsEmpSel = Nothing
              %>
            </select>
          </div>

          <div class="row">
            <div class="col-md-6">
              <label>Fecha entrada</label>
              <input type="date" name="fechaEntrada" class="form-control" required>
            </div>
            <div class="col-md-6">
              <label>Hora entrada</label>
              <input type="time" name="horaEntrada" class="form-control" required>
            </div>
          </div>

          <div class="row mt-3">
            <div class="col-md-6">
              <label>Fecha salida</label>
              <input type="date" name="fechaSalida" class="form-control" required>
            </div>
            <div class="col-md-6">
              <label>Hora salida</label>
              <input type="time" name="horaSalida" class="form-control" required>
            </div>
          </div>

          <div class="row mt-3">
            <div class="col-md-6">
              <label>Intervalo inicio</label>
              <input type="time" name="intervaloInicio" class="form-control" value="12:30">
            </div>
            <div class="col-md-6">
              <label>Intervalo fin</label>
              <input type="time" name="intervaloFin" class="form-control" value="13:30">
            </div>
          </div>

          <div class="mb-3 mt-3">
            <label>Observaciones</label>
            <textarea name="observaciones" class="form-control"></textarea>
          </div>

          <button type="submit" class="btn btn-success">Guardar</button>
          <a href="asistencias.asp" class="btn btn-secondary">Cancelar</a>
        </form>
      </main>
    </div>
  </div>
</body>
</html>
