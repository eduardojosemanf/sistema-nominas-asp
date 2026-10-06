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
