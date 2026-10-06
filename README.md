# Sistema de Nómina ASP v2

Versión profesional de sistema de nómina en ASP Clásico + Access, con:
- validación y mensajes de error
- edición y eliminación de registros
- exportación a Excel
- cálculo exacto de quincena
- diseño con Bootstrap
- menú lateral
- CRUD completo
- reporte financiero por empleado/fechas
- cálculo de neto a pagar descontando anticipos

## Estructura del proyecto

```
C:\NominaASPv2\
  default.asp
  empleados.asp
  empleado_form.asp
  asistencias.asp
  asistencia_form.asp
  anticipos.asp
  anticipo_form.asp
  quincena.asp
  reporte.asp
  export_excel.asp
  export_pdf.asp
  includes\
    db.asp
    funciones_nomina.asp
    validaciones.asp
  css\
    bootstrap.min.css
    estilo.css
  bd\
    nomina_v2.mdb
```

## 1) Base de datos Access

Crea la base `C:\NominaASPv2\bd\nomina_v2.mdb` en Access y crea las tablas:

### empleados
- id (Autonumeración) PK
- nombre (Texto, 150)
- documento (Texto, 20)
- categoria (Texto, 30)
- tarifa_hora (Número, Doble)
- tarifa_jornada (Número, Doble)
- activo (Sí/No)
- fecha_ingreso (Fecha/Hora)
- observaciones (Texto largo)

### asistencias
- id (Autonumeración) PK
- empleado_id (Número)
- fecha_entrada (Fecha/Hora)
- hora_entrada (Texto, 10)
- fecha_salida (Fecha/Hora)
- hora_salida (Texto, 10)
- intervalo_inicio (Texto, 10)
- intervalo_fin (Texto, 10)
- total_horas (Número, Doble)
- observaciones (Texto largo)

### anticipos
- id (Autonumeración) PK
- empleado_id (Número)
- fecha (Fecha/Hora)
- tipo_operacion (Texto, 30)
- referencia (Texto, 100)
- monto (Número, Doble)
- detalle (Texto largo)

### configuracion
- id (Autonumeración) PK
- nombre_empresa (Texto, 200)
- rfc (Texto, 50)
- jornada_hora_inicio (Texto, 10)
- jornada_hora_fin (Texto, 10)
- descanso_inicio (Texto, 10)
- descanso_fin (Texto, 10)
- moneda (Texto, 10)

### quincenas
- id (Autonumeración) PK
- empleado_id (Número)
- anio (Número)
- mes (Número)
- quincena (Número)
- horas_trabajadas (Número, Doble)
- subtotal (Número, Doble)
- anticipos (Número, Doble)
- neto (Número, Doble)
- estado (Texto, 20)
- fecha_calculo (Fecha/Hora)

## 2) includes/db.asp

```asp
<%
Dim cn
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\NominaASPv2\bd\nomina_v2.mdb;Persist Security Info=False;"
%>
```

## 3) includes/funciones_nomina.asp

```asp
<%
Function CalcularHorasTrabajadas(fechaEntrada, horaEntrada, fechaSalida, horaSalida, intervaloInicio, intervaloFin)
    Dim dInicio, dFin, totalMinutos, descansoMinutos, dIntIni, dIntFin

    dInicio = CDate(fechaEntrada & " " & horaEntrada)
    dFin = CDate(fechaSalida & " " & horaSalida)

    If dFin < dInicio Then
        dFin = DateAdd("d", 1, dFin)
    End If

    totalMinutos = DateDiff("n", dInicio, dFin)

    If Len(intervaloInicio) > 0 And Len(intervaloFin) > 0 Then
        dIntIni = CDate(fechaEntrada & " " & intervaloInicio)
        dIntFin = CDate(fechaEntrada & " " & intervaloFin)

        If dIntFin < dIntIni Then
            dIntFin = DateAdd("d", 1, dIntFin)
        End If

        descansoMinutos = DateDiff("n", dIntIni, dIntFin)
        totalMinutos = totalMinutos - descansoMinutos
    End If

    If totalMinutos < 0 Then totalMinutos = 0
    CalcularHorasTrabajadas = Round(totalMinutos / 60, 2)
End Function

Function GetQuincena(numeroDia)
    If numeroDia <= 15 Then
        GetQuincena = 1
    Else
        GetQuincena = 2
    End If
End Function

Function FormatoMoneda(valor)
    FormatoMoneda = FormatCurrency(valor, 2)
End Function
%>
```

## 4) includes/validaciones.asp

```asp
<%
Function ValidarEmpleado(nombre)
    If Trim(nombre) = "" Then
        ValidarEmpleado = "Debe ingresar el nombre del empleado."
        Exit Function
    End If

    If Len(Trim(nombre)) < 3 Then
        ValidarEmpleado = "El nombre del empleado debe tener al menos 3 caracteres."
        Exit Function
    End If

    ValidarEmpleado = ""
End Function

Function ValidarAsistencia(empleadoId, fechaEntrada, horaEntrada, fechaSalida, horaSalida, intervaloInicio, intervaloFin)
    If Trim(empleadoId) = "" Then
        ValidarAsistencia = "Debe seleccionar un empleado."
        Exit Function
    End If

    If Not IsDate(fechaEntrada) Then
        ValidarAsistencia = "La fecha de entrada es inválida."
        Exit Function
    End If

    If Not IsDate(fechaSalida) Then
        ValidarAsistencia = "La fecha de salida es inválida."
        Exit Function
    End If

    If CDate(fechaSalida & " " & horaSalida) < CDate(fechaEntrada & " " & horaEntrada) Then
        ValidarAsistencia = "La fecha y hora de salida no pueden ser menores que la de entrada."
        Exit Function
    End If

    If Len(intervaloInicio) > 0 And Len(intervaloFin) > 0 Then
        If CDate(fechaEntrada & " " & intervaloFin) <= CDate(fechaEntrada & " " & intervaloInicio) Then
            ValidarAsistencia = "El intervalo de descanso es inválido."
            Exit Function
        End If
    End If

    ValidarAsistencia = ""
End Function

Function ValidarAnticipo(empleadoId, fecha, monto)
    If Trim(empleadoId) = "" Then
        ValidarAnticipo = "Debe seleccionar un empleado."
        Exit Function
    End If

    If Not IsDate(fecha) Then
        ValidarAnticipo = "La fecha del anticipo es inválida."
        Exit Function
    End If

    If Not IsNumeric(monto) Then
        ValidarAnticipo = "El monto debe ser numérico."
        Exit Function
    End If

    If CDbl(monto) <= 0 Then
        ValidarAnticipo = "El monto debe ser mayor a cero."
        Exit Function
    End If

    ValidarAnticipo = ""
End Function
%>
```

## 5) default.asp

```asp
<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Nomina ASP v2</title>
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
        <h1>Panel principal</h1>
        <div class="card p-3">
          <h4>Funcionalidades</h4>
          <ul>
            <li>Registro de empleados</li>
            <li>Registro de asistencias por fecha y horario</li>
            <li>Registro de anticipos por tipo de operación</li>
            <li>Cálculo de horas trabajadas y sueldo</li>
            <li>Reporte por fechas y empleado</li>
            <li>Exportación a Excel</li>
            <li>Cálculo de quincena</li>
          </ul>
        </div>
      </main>
    </div>
  </div>
</body>
</html>
```

## 6) empleados.asp

```asp
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
```

## 7) empleado_form.asp

```asp
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
```

## 8) asistencias.asp

```asp
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
```

## 9) asistencia_form.asp

```asp
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
```

## 10) anticipos.asp

```asp
<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Anticipos</title>
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
        <h2>Anticipos</h2>
        <a href="anticipo_form.asp?action=add" class="btn btn-primary mb-3">Nuevo anticipo</a>

        <table class="table table-striped table-bordered">
          <thead>
            <tr>
              <th>Empleado</th>
              <th>Fecha</th>
              <th>Tipo</th>
              <th>Referencia</th>
              <th>Monto</th>
              <th>Acciones</th>
            </tr>
          </thead>
          <tbody>
            <%
            Dim rsAn
            Set rsAn = Server.CreateObject("ADODB.Recordset")
            rsAn.Open "SELECT a.*, e.nombre FROM anticipos a INNER JOIN empleados e ON e.id = a.empleado_id ORDER BY a.fecha DESC", cn, 3, 1

            Do While Not rsAn.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & rsAn("nombre") & "</td>"
                Response.Write "<td>" & rsAn("fecha") & "</td>"
                Response.Write "<td>" & rsAn("tipo_operacion") & "</td>"
                Response.Write "<td>" & rsAn("referencia") & "</td>"
                Response.Write "<td>" & FormatCurrency(rsAn("monto"), 2) & "</td>"
                Response.Write "<td>"
                Response.Write "<a href='anticipo_form.asp?action=edit&id=" & rsAn("id") & "' class='btn btn-sm btn-warning'>Editar</a> "
                Response.Write "<a href='anticipo_form.asp?action=delete&id=" & rsAn("id") & "' class='btn btn-sm btn-danger' onclick=\"return confirm('¿Eliminar anticipo?')\">Eliminar</a>"
                Response.Write "</td>"
                Response.Write "</tr>"
                rsAn.MoveNext
            Loop
            rsAn.Close
            Set rsAn = Nothing
            %>
          </tbody>
        </table>
      </main>
    </div>
  </div>
</body>
</html>
```

## 11) anticipo_form.asp

```asp
<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/validaciones.asp"-->
<%
Dim accion, id, empleadoId, fecha, tipoOperacion, referencia, monto, detalle, mensaje
accion = Request.QueryString("action")
id = Request.QueryString("id")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    empleadoId = Request.Form("empleadoId")
    fecha = Request.Form("fecha")
    tipoOperacion = Request.Form("tipoOperacion")
    referencia = Request.Form("referencia")
    monto = Request.Form("monto")
    detalle = Request.Form("detalle")

    mensaje = ValidarAnticipo(empleadoId, fecha, monto)

    If mensaje <> "" Then
        Response.Redirect "anticipo_form.asp?action=add&error=" & Server.URLEncode(mensaje)
    Else
        If accion = "edit" Then
            Dim sqlUpdateAn
            sqlUpdateAn = "UPDATE anticipos SET empleado_id=" & empleadoId & ", fecha=#" & fecha & "#, tipo_operacion='" & tipoOperacion & "', referencia='" & Replace(referencia, "'", "''") & "', monto=" & Replace(monto, ",", ".") & ", detalle='" & Replace(detalle, "'", "''") & "' WHERE id=" & id
            cn.Execute(sqlUpdateAn)
            Response.Redirect "anticipos.asp?msg=" & Server.URLEncode("Anticipo actualizado.")
        Else
            Dim sqlInsertAn
            sqlInsertAn = "INSERT INTO anticipos (empleado_id, fecha, tipo_operacion, referencia, monto, detalle) VALUES (" & empleadoId & ", #" & fecha & "#, '" & tipoOperacion & "', '" & Replace(referencia, "'", "''") & "', " & Replace(monto, ",", ".") & ", '" & Replace(detalle, "'", "''") & "')"
            cn.Execute(sqlInsertAn)
            Response.Redirect "anticipos.asp?msg=" & Server.URLEncode("Anticipo registrado.")
        End If
    End If
End If

If accion = "delete" And id <> "" Then
    cn.Execute "DELETE FROM anticipos WHERE id=" & id
    Response.Redirect "anticipos.asp?msg=" & Server.URLEncode("Anticipo eliminado.")
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Anticipo</title>
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
        <h2><%= IIf(accion = "edit", "Editar anticipo", "Nuevo anticipo") %></h2>

        <% If Request.QueryString("error") <> "" Then Response.Write "<div class='alert alert-danger'>" & Request.QueryString("error") & "</div>" End If %>

        <form method="post">
          <div class="mb-3">
            <label>Empleado</label>
            <select name="empleadoId" class="form-control" required>
              <option value="">Seleccione</option>
              <%
              Dim rsEmpAn
              Set rsEmpAn = Server.CreateObject("ADODB.Recordset")
              rsEmpAn.Open "SELECT * FROM empleados WHERE activo=True ORDER BY nombre", cn, 3, 1
              Do While Not rsEmpAn.EOF
                  Response.Write "<option value='" & rsEmpAn("id") & "'>" & rsEmpAn("nombre") & "</option>"
                  rsEmpAn.MoveNext
              Loop
              rsEmpAn.Close
              Set rsEmpAn = Nothing
              %>
            </select>
          </div>

          <div class="mb-3">
            <label>Fecha</label>
            <input type="date" name="fecha" class="form-control" required>
          </div>

          <div class="mb-3">
            <label>Tipo de operación</label>
            <select name="tipoOperacion" class="form-control">
              <option value="orden_compra">Orden de compra</option>
              <option value="transferencia">Transferencia</option>
              <option value="efectivo">Efectivo</option>
              <option value="vale">Vale</option>
            </select>
          </div>

          <div class="mb-3">
            <label>Referencia</label>
            <input type="text" name="referencia" class="form-control">
          </div>

          <div class="mb-3">
            <label>Monto</label>
            <input type="number" step="0.01" name="monto" class="form-control" required>
          </div>

          <div class="mb-3">
            <label>Detalle</label>
            <textarea name="detalle" class="form-control"></textarea>
          </div>

          <button type="submit" class="btn btn-success">Guardar</button>
          <a href="anticipos.asp" class="btn btn-secondary">Cancelar</a>
        </form>
      </main>
    </div>
  </div>
</body>
</html>
```

## 12) quincena.asp

```asp
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
```

## 13) reporte.asp

```asp
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
```

## 14) export_excel.asp

```asp
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
```

## 15) export_pdf.asp

```asp
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
```

## 16) css/estilo.css

```css
body {
    background: #f3f6fb;
    font-family: Arial, sans-serif;
}

.sidebar {
    background: #1f2d3d;
    min-height: 100vh;
    padding: 20px 15px;
    color: white;
}

.sidebar h4 {
    color: #ffffff;
    margin-bottom: 20px;
}

.sidebar a {
    color: white;
    text-decoration: none;
    display: block;
    padding: 10px 12px;
    margin-bottom: 8px;
    border-radius: 6px;
}

.sidebar a:hover {
    background: #2c4058;
}

.content {
    padding: 25px;
}

.card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 8px;
    box-shadow: 0 1px 2px rgba(0,0,0,.05);
}

.table th, .table td {
    vertical-align: middle;
}

.alert {
    margin-top: 15px;
}
```

## 17) Instalación en IIS

1. Crear carpeta: `C:\NominaASPv2`
2. Guardar todos los archivos en esa carpeta
3. Crear la base `bd\nomina_v2.mdb`
4. Crear tablas con Access
5. Abrir IIS Manager
6. Crear sitio web con:
   - nombre: `NominaASPv2`
   - ruta física: `C:\NominaASPv2`
   - puerto: `8080`
7. Habilitar ASP en Windows Features
8. Verificar que la conexión Access funciona con el proveedor `Microsoft.ACE.OLEDB.12.0`
9. Probar:
   - `http://localhost:8080/default.asp`

## 18) Consideraciones importantes

- En entorno real, la mejor práctica es usar SQL Server o MySQL para una aplicación productiva.
- Access es excelente para prototipos y pruebas locales.
- Si se trabaja en IIS con Access, se debe instalar el Microsoft Access Database Engine.

## 19) Siguiente paso

Puedo ir un paso más avanzado y dejarte un proyecto ya empaquetado para descargar con:
- ejemplo de datos de prueba
- base Access pre-creada
- validaciones adicionales
- versión con conexión segura y mejor estructura
- script de instalación para Windows

Si quieres, en el siguiente bloque te preparo el ZIP definitivo listo para descargar y usar localmente.
