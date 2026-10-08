# Automatización de la sincronización con Mercado Pago (sin subir CSV a mano)

Esta guía deja los socios de Mercado Pago con sus **meses pagos tildados
automáticamente**, consultando los pagos reales de MP. Se ejecuta solo (cron
diario), sin exportar ni subir ningún CSV.

> El código vive en tu Google Apps Script (el mismo de la planilla de socios).
> El token de MP se guarda de forma segura en el script, nunca en el sitio web.
> **Estado: IMPLEMENTADO Y FUNCIONANDO** ✅ (178 suscripciones, pagos reales por mes).

---

## Paso 1 — Access Token de Mercado Pago

1. https://www.mercadopago.com.ar/developers/panel → tu aplicación.
2. "Credenciales de producción" → copiá el **Access Token** (`APP_USR-...`).
3. Tratalo como una contraseña.

## Paso 2 — Guardar el token (una vez)

En el Apps Script, archivo `Código.gs` o uno nuevo, pegá, ejecutá UNA vez y
después borralo:

```javascript
function guardarTokenMP() {
  PropertiesService.getScriptProperties().setProperty("MP_ACCESS_TOKEN", "APP_USR-TU-TOKEN");
}
```

> Importante: el token va en UNA sola línea, entre comillas rectas `"`.
> Si el editor queda en gris, es un error de sintaxis (comilla rota): borrá la
> línea, reescribila a mano y guardá.

## Paso 3 — Código de sincronización

Creá un archivo nuevo en Apps Script (ícono **+** → Secuencia de comandos),
nombralo `MercadoPago`, borrá lo que trae por defecto y pegá TODO esto:

```javascript
// ============ Sincronización con Mercado Pago (pagos reales) ============

function tokenMP_() {
  var t = PropertiesService.getScriptProperties().getProperty("MP_ACCESS_TOKEN");
  if (!t) throw new Error("Falta MP_ACCESS_TOKEN. Ejecutá guardarTokenMP primero.");
  return t;
}

// Meses pagados por payer_id, según los pagos REALES aprobados del año en curso.
// Devuelve: { "152028761": {2:true, 3:true, ...}, ... }  (mes 1-12)
function obtenerMesesPagadosMP() {
  var token = tokenMP_();
  var anio = new Date().getFullYear();
  var desde = anio + "-01-01T00:00:00.000-03:00";
  var hasta = anio + "-12-31T23:59:59.000-03:00";
  var porPayer = {};
  var offset = 0, limit = 50, seguir = true, guard = 0;

  while (seguir && guard < 200) {
    guard++;
    var url = "https://api.mercadopago.com/v1/payments/search"
      + "?sort=date_created&criteria=asc"
      + "&range=date_created&begin_date=" + encodeURIComponent(desde) + "&end_date=" + encodeURIComponent(hasta)
      + "&status=approved&limit=" + limit + "&offset=" + offset;
    var resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    var items = (data && data.results) ? data.results : [];
    items.forEach(function (p) {
      var payer = p.payer && p.payer.id ? String(p.payer.id) : "";
      if (!payer) return;
      var fecha = p.date_approved || p.date_created;
      if (!fecha) return;
      var mes = new Date(fecha).getMonth() + 1;
      if (!porPayer[payer]) porPayer[payer] = {};
      porPayer[payer][mes] = true;
    });
    var total = (data && data.paging) ? data.paging.total : 0;
    offset += limit;
    seguir = items.length === limit && offset < total;
  }
  return porPayer;
}

// Escribe los meses pagados en la planilla de socios (solo marca vacíos; nunca pisa lo manual).
function sincronizarMPAutomatico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = null, headers = null;
  var hojas = ss.getSheets();
  for (var h = 0; h < hojas.length; h++) {
    var hd = hojas[h].getRange(1, 1, 1, hojas[h].getLastColumn()).getValues()[0];
    if (hd.indexOf("USUARIO_MP") !== -1 && hd.indexOf("ENE") !== -1) { hoja = hojas[h]; headers = hd; break; }
  }
  if (!hoja) { Logger.log("No se encontró la hoja de socios (con USUARIO_MP y ENE)."); return; }

  function col(n) { return headers.indexOf(n); }
  var cPayer = col("USUARIO_MP");
  var mesesCol = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
  var cMes = mesesCol.map(col);

  var pagados = obtenerMesesPagadosMP();
  var datos = hoja.getDataRange().getValues();
  var cambios = 0, sociosTocados = 0;

  for (var r = 1; r < datos.length; r++) {
    var payer = String(datos[r][cPayer] || "").trim();
    if (!payer || !pagados[payer]) continue;
    var toco = false;
    for (var m = 0; m < 12; m++) {
      if (pagados[payer][m + 1]) {
        var c = cMes[m];
        if (c === -1) continue;
        var val = String(datos[r][c] || "").trim();
        if (!val) { hoja.getRange(r + 1, c + 1).setValue("MP"); cambios++; toco = true; }
      }
    }
    if (toco) sociosTocados++;
  }
  Logger.log("Sincronización MP: " + cambios + " meses marcados en " + sociosTocados + " socios.");
  return cambios;
}

// Pruebas manuales (Ver > Registros):
function probarPagosMP() {
  var m = obtenerMesesPagadosMP();
  var p = Object.keys(m);
  Logger.log("Payers con pagos este año: " + p.length);
  if (p.length) Logger.log(p[0] + " -> " + JSON.stringify(m[p[0]]));
}
```

Guardá (Ctrl+S). Ejecutá `sincronizarMPAutomatico` una vez a mano y revisá
**Ver → Registros**: debe decir "X meses marcados en Y socios".

## Paso 4 — Cron automático (cero intervención)

1. Apps Script → ícono **⏰ Activadores** (panel izquierdo).
2. **+ Agregar activador**.
3. Configurá:
   - Función: **sincronizarMPAutomatico**
   - Implementación: **Head**
   - Origen del evento: **Según tiempo**
   - Tipo: **Temporizador por día**
   - Hora: una franja de madrugada (ej. 3–4 a. m.)
4. Guardar y autorizar permisos.

A partir de ahí corre solo todas las madrugadas. El equipo abre el sistema y los
meses MP ya están tildados.

---

## Notas importantes

- **Pagos reales:** usa `/v1/payments/search` con `status=approved`. Refleja lo
  que MP efectivamente cobró, mes por mes.
- **Matcheo por payer_id:** cada pago se asocia al socio por el `payer_id`
  guardado en la ficha (`USUARIO_MP`). Los socios deben estar vinculados (badge
  verde "🔗 MP vinculado"). Para vincular en masa, hacé una pasada con el CSV
  (que trae nombres) y el sistema auto-vincula; después el cron funciona por id.
- **Solo agrega, nunca quita:** marca únicamente meses vacíos. Si MP reversa un
  pago, el mes queda tildado hasta que lo destildes a mano.
- **Año en curso:** mira los pagos del año actual; en enero arranca el año nuevo.
- **Seguridad:** el token vive solo en las propiedades del script, nunca en el
  sitio. Para revocarlo, generá uno nuevo en MP y re-ejecutá `guardarTokenMP`.

## (Opcional) Botón "Sincronizar con MP" en el sitio

Si además querés un botón en el sistema para forzar la sync en el momento (sin
esperar al cron), se puede exponer `sincronizarMPAutomatico` vía `doGet` y
agregar el botón en la web. Pedilo cuando quieras y se implementa.
