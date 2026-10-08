# Automatización de la sincronización con Mercado Pago (sin subir CSV a mano)

Esta guía explica cómo conectar tu backend de Google Apps Script con la API de
Mercado Pago para que los datos de suscripciones se actualicen solos, sin tener
que exportar y subir el CSV manualmente.

> **Importante:** esto requiere pegar código en tu Google Apps Script (el mismo
> que ya usás para guardar socios) y cargar un token secreto de Mercado Pago.
> El token NO va en el sitio web (sería inseguro); va guardado en el Apps Script.

---

## Paso 1 — Obtener el Access Token de Mercado Pago

1. Entrá a https://www.mercadopago.com.ar/developers/panel
2. Creá una aplicación (o usá una existente).
3. En "Credenciales de producción", copiá el **Access Token** (empieza con `APP_USR-...`).
4. Guardalo, lo vas a pegar en el Apso Script en el paso 3.

> El Access Token da acceso a tu cuenta de MP. Tratalo como una contraseña.

---

## Paso 2 — Abrir tu Google Apps Script

1. Abrí la planilla de Google Sheets que usa el sistema.
2. Menú **Extensiones → Apps Script**.
3. Vas a ver el código actual (el que maneja `doGet`/`doPost` de los socios).

---

## Paso 3 — Guardar el token de forma segura

En el editor de Apps Script, pegá esta función, ejecutala UNA vez (botón ▷),
y luego borrala (para que el token no quede en el código):

```javascript
function guardarTokenMP() {
  // Pegá tu Access Token entre las comillas:
  PropertiesService.getScriptProperties().setProperty(
    'MP_ACCESS_TOKEN',
    'APP_USR-TU-TOKEN-ACA'
  );
}
```

Al ejecutarla, el token queda guardado en las propiedades del script (seguro,
no visible en el sitio). Después borrá la función del editor.

---

## Paso 4 — Agregar la función que trae las suscripciones de MP

Pegá esto en el Apps Script (es nuevo, no reemplaza lo que ya tenés):

```javascript
// Devuelve las suscripciones (preapprovals) de MP en el mismo formato
// de columnas que el CSV, para que el sitio las procese igual que el archivo.
function obtenerSuscripcionesMP() {
  var token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
  if (!token) throw new Error('Falta MP_ACCESS_TOKEN. Ejecutá guardarTokenMP primero.');

  var resultados = [];
  var offset = 0;
  var limit = 100;
  var seguir = true;

  while (seguir) {
    var url = 'https://api.mercadopago.com/preapproval/search?limit=' + limit + '&offset=' + offset;
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    var items = (data && data.results) ? data.results : [];
    items.forEach(function(it) {
      resultados.push({
        payer_id: it.payer_id || '',
        payer_first_name: '',        // MP no siempre expone el nombre acá
        payer_last_name: '',
        status: it.status || '',
        reason: it.reason || '',
        preapproval_plan_id: it.preapproval_plan_id || '',
        frequency: (it.auto_recurring && it.auto_recurring.frequency) || '',
        frequency_type: (it.auto_recurring && it.auto_recurring.frequency_type) || '',
        billing_status: it.status || '',
        last_charge_date: it.last_charged_date || '',
        start_date: (it.auto_recurring && it.auto_recurring.start_date) || it.date_created || '',
        charged_quantity: it.charged_quantity || ''
      });
    });
    offset += limit;
    seguir = items.length === limit && offset < (data.paging ? data.paging.total : 0);
  }
  return resultados;
}
```

---

## Paso 5 — Exponer esos datos al sitio

En tu función `doGet` del Apps Script, agregá un caso para cuando el sitio pida
las suscripciones. Buscá donde manejás los parámetros (`e.parameter.action`) y
agregá:

```javascript
if (e.parameter.action === 'suscripciones_mp') {
  var subs = obtenerSuscripcionesMP();
  return ContentService
    .createTextOutput(JSON.stringify({ suscripciones: subs }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

Volvé a **Implementar → Administrar implementaciones → Editar → Nueva versión**
para publicar el cambio.

---

## Paso 6 — Sincronización 100% automática (cron, cero intervención)

Con esto, Google ejecuta la sincronización solo (cada día o semana): trae los
datos de MP, calcula qué meses están pagos de cada socio y los marca
directamente en la planilla. Cuando el equipo abre el sistema, ya está todo
tildado, sin subir CSV ni tocar ningún botón.

### 6.1 — Pegá esta función en el Apps Script

Ajustá al principio los **nombres de tu hoja y columnas** si difieren. La
función asume que la hoja de socios tiene una columna con el `payer_id` de MP
(la misma que el sistema llama `USUARIO_MP`) y columnas de meses `ENE..DIC`.

```javascript
function sincronizarMPAutomatico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Socios'); // <-- nombre de tu hoja de socios
  if (!hoja) { Logger.log('No se encontró la hoja Socios'); return; }

  var datos = hoja.getDataRange().getValues();
  var headers = datos[0];
  function col(nombre) { return headers.indexOf(nombre); } // índice de columna por nombre

  var cPayer = col('USUARIO_MP');
  var cMetodo = col('METODO_PAGO');
  var meses = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  var cMes = meses.map(function(m){ return col(m); });
  if (cPayer === -1) { Logger.log('Falta columna USUARIO_MP'); return; }

  var subs = obtenerSuscripcionesMP();          // del Paso 4
  var porPayer = {};
  subs.forEach(function(s){ if (s.payer_id) porPayer[String(s.payer_id)] = s; });

  var anioActual = new Date().getFullYear();
  var cambios = 0;

  for (var r = 1; r < datos.length; r++) {
    var payer = String(datos[r][cPayer] || '').trim();
    if (!payer || !porPayer[payer]) continue;
    var s = porPayer[payer];

    // Solo suscripciones activas
    var st = String(s.status || '').toLowerCase();
    var activa = (st === 'authorized' || st === 'up_to_date');
    if (!activa) continue;

    // Plan
    var freq = parseInt(s.frequency, 10) || 1;
    var ftype = String(s.frequency_type || '').toLowerCase();
    var mesesFrec = (ftype.indexOf('year') !== -1) ? freq * 12 : freq;
    var plan = (mesesFrec >= 11.5) ? 'anual' : (mesesFrec >= 5.5 ? 'semestral' : 'mensual');

    // Fechas
    var inicio = s.start_date ? new Date(s.start_date) : null;
    var ultimo = s.last_charge_date ? new Date(s.last_charge_date) : null;
    if (!ultimo) continue;
    if (ultimo.getFullYear() < anioActual) continue;

    var mesInicio = (inicio && inicio.getFullYear() === anioActual) ? inicio.getMonth() : 0;
    var mesFin = (ultimo.getFullYear() === anioActual) ? ultimo.getMonth() : 11;

    var cubrir = [];
    if (plan === 'anual') { for (var i = mesInicio; i <= 11; i++) cubrir.push(i); }
    else if (plan === 'semestral') { var f = Math.min(11, mesFin + 5); for (var i = mesInicio; i <= f; i++) cubrir.push(i); }
    else { for (var i = mesInicio; i <= mesFin; i++) cubrir.push(i); }

    // Tilda meses (solo agrega, no borra lo cargado a mano)
    cubrir.forEach(function(idx){
      var c = cMes[idx];
      if (c === -1) return;
      var val = String(datos[r][c] || '').trim();
      if (!val) { hoja.getRange(r + 1, c + 1).setValue('MP'); cambios++; }
    });
  }
  Logger.log('Sincronización MP: ' + cambios + ' meses marcados.');
}
```

### 6.2 — Crear el disparador horario

1. En el Apps Script, panel izquierdo → ⏰ **Activadores (Triggers)**.
2. **+ Agregar activador** (abajo a la derecha).
3. Configurá:
   - Función a ejecutar: **sincronizarMPAutomatico**
   - Implementación: **Head**
   - Origen del evento: **Según tiempo**
   - Tipo: **Temporizador por día** (o "por semana" si preferís)
   - Hora: la franja que quieras (ej. 3am–4am).
4. Guardar. Google te pedirá autorizar permisos la primera vez (aceptá).

Listo: a partir de ahí, todos los días (o semanas) a esa hora, Google corre la
función solo, consulta MP y deja los meses tildados en la planilla. **Cero
intervención.** El sistema simplemente lee la planilla ya actualizada.

### 6.3 — Verificar que funciona

- En el Apps Script, ejecutá `sincronizarMPAutomatico` una vez a mano (botón ▷)
  y revisá el **Registro de ejecución** (Ver → Registros): debe decir
  "X meses marcados".
- Abrí el sistema y confirmá que los socios MP tengan sus meses tildados.

### Comparación de las dos vías

| | Botón "Sincronizar con MP" (Paso 7) | Cron automático (Paso 6) |
|---|---|---|
| Quién lo dispara | Una persona, cuando quiere | Google, solo (diario/semanal) |
| Intervención | Un clic | Ninguna |
| Dónde se tilda | En el navegador y se guarda | Directo en la planilla |
| Recomendado para | Control puntual | Mantener todo al día sin pensar |

Podés tener **las dos** a la vez: el cron mantiene todo al día solo, y el botón
te sirve para forzar una actualización en el momento si hace falta.

---

## Paso 7 — Lado del sitio (lo hago yo)

Cuando tengas los pasos 1-5 listos y me confirmes, agrego en el sistema un botón
**"🔄 Sincronizar con MP"** que:
1. Le pide al Apps Script los datos (`?action=suscripciones_mp`).
2. Los procesa con el MISMO análisis que ya usa el CSV (detecta plan, marca
   caídos/activos y tilda los meses pagos).
3. Todo sin subir ningún archivo.

Avisame cuando completes los pasos del Apps Script y lo conecto.

---

## Resumen de seguridad

- El token vive solo en el Apps Script (propiedades del script), nunca en el sitio.
- El sitio nunca ve el token; solo recibe los datos ya procesados.
- Si alguna vez querés revocar el acceso, generás un token nuevo en el panel de MP
  y actualizás la propiedad con `guardarTokenMP`.
