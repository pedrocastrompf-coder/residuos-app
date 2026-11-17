/**
 * Hallazgos ambientales · módulo del backend (Google Apps Script)
 * Pega este archivo como un archivo NUEVO («Hallazgos») en Extensiones > Apps Script, junto a Code.gs.
 *
 * Sección 1 · reglas puras (sin Sheets ni Drive): se prueban con `node --test tests/`.
 * Sección 2 · persistencia en Sheet y Drive.
 */

// ===================== 1. REGLAS PURAS =====================
const HZ_TIPOS = ["Residuos","Agua y efluentes","Emisiones y olores","Derrames y químicos","Energía","Orden y limpieza","Otro"];
const HZ_RIESGOS = ["Bajo","Medio","Alto"];
const HZ_TIPOS_AREA = ["Producción","Servicio","Común"];
const HZ_MAX_FOTO = 400 * 1024, HZ_MAX_PDF = 1.5 * 1024 * 1024, HZ_MAX_ARCHIVOS = 3;
const HZ_ROLES_MEDIO_AMBIENTE = ["Inspector","Administrador"];

function hzTxt_(v){ return String(v == null ? "" : v).trim().replace(/\s+/g, " "); }
function hzIdValido_(id){ return /^[A-Za-z0-9._-]{4,64}$/.test(String(id || "")); }

// data-URL JPEG (≤ 400 KB) o PDF (≤ 1.5 MB) → {mime, b64, bytes}; cualquier otra cosa → null
function hzDataUrl_(s){
  const m = /^data:(image\/jpeg|application\/pdf);base64,([A-Za-z0-9+\/]+={0,2})$/.exec(String(s || ""));
  if (!m) return null;
  const b64 = m[2], pad = (b64.match(/=+$/) || [""])[0].length, bytes = Math.floor(b64.length * 3 / 4) - pad;
  if (bytes < 1 || bytes > (m[1] === "application/pdf" ? HZ_MAX_PDF : HZ_MAX_FOTO)) return null;
  return {mime: m[1], b64: b64, bytes: bytes};
}

function hzValidarReporte_(b){
  b = b || {};
  if (!hzIdValido_(b.id)) return {ok:false, error:"id_invalido"};
  const area = hzTxt_(b.area);
  if (!area || area.length > 40) return {ok:false, error:"area_invalida"};
  const descripcion = hzTxt_(b.descripcion);
  if (descripcion.length < 5 || descripcion.length > 500) return {ok:false, error:"descripcion_invalida"};
  if (HZ_TIPOS.indexOf(b.tipo) < 0) return {ok:false, error:"tipo_invalido"};
  if (HZ_RIESGOS.indexOf(b.riesgo) < 0) return {ok:false, error:"riesgo_invalido"};
  const fotos = Array.isArray(b.fotos) ? b.fotos : [];
  const parsed = fotos.map(hzDataUrl_);
  if (!fotos.length || fotos.length > HZ_MAX_ARCHIVOS || parsed.some(p => !p || p.mime !== "image/jpeg")) return {ok:false, error:"fotos_invalidas"};
  return {ok:true, v:{id:String(b.id), area:area, descripcion:descripcion, tipo:b.tipo, riesgo:b.riesgo, fotos:fotos}};
}

// yyyy-MM-dd real → número de día (UTC) o NaN
function hzDia_(s){
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "")); if (!m) return NaN;
  const t = Date.UTC(+m[1], +m[2]-1, +m[3]), d = new Date(t);
  return (d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2]-1 && d.getUTCDate() === +m[3]) ? Math.floor(t / 86400000) : NaN;
}
function hzValidarAsignacion_(b, hoy){
  b = b || {};
  if (!hzIdValido_(b.id)) return {ok:false, error:"id_invalido"};
  const ej = hzTxt_(b.areaEjecutora);
  if (!ej || ej.length > 40) return {ok:false, error:"ejecutora_invalida"};
  const p = hzDia_(b.plazo), h = hzDia_(hoy);
  if (isNaN(p) || isNaN(h) || p < h || p - h > 365) return {ok:false, error:"plazo_invalido"};
  return {ok:true, v:{id:String(b.id), areaEjecutora:ej, plazo:String(b.plazo)}};
}

function hzValidarCierre_(b){
  b = b || {};
  if (!hzIdValido_(b.id)) return {ok:false, error:"id_invalido"};
  const comentario = hzTxt_(b.comentario);
  if (comentario.length > 300) return {ok:false, error:"comentario_invalido"};
  const ev = Array.isArray(b.evidencia) ? b.evidencia : [];
  if (!ev.length) return {ok:false, error:"evidencia_requerida"};
  const parsed = ev.map(hzDataUrl_);
  if (ev.length > HZ_MAX_ARCHIVOS || parsed.some(p => !p)) return {ok:false, error:"evidencia_invalida"};
  const pdfs = parsed.filter(p => p.mime === "application/pdf").length;
  if (pdfs && (pdfs !== 1 || ev.length !== 1)) return {ok:false, error:"evidencia_invalida"};   // o 1 a 3 fotos, o un solo PDF
  return {ok:true, v:{id:String(b.id), comentario:comentario, evidencia:ev}};
}

// el estado guardado es Abierto|Asignado|Cerrado; «Vencido» se calcula (Asignado con plazo anterior a hoy)
function hzEstadoVista_(h, hoy){
  if (h.estado === "Cerrado") return "Cerrado";
  if (h.estado === "Asignado") return (h.plazo && hzDia_(h.plazo) < hzDia_(hoy)) ? "Vencido" : "Asignado";
  return "Abierto";
}

function hzEsMedioAmbiente_(u){ return HZ_ROLES_MEDIO_AMBIENTE.indexOf(u.rol) >= 0; }
function hzPuedeVer_(h, u){
  if (hzEsMedioAmbiente_(u)) return true;
  const A = u.areas || [];
  return h.reportadoPor === u.usuario || (!!h.area && A.indexOf(h.area) >= 0) || (!!h.areaEjecutora && A.indexOf(h.areaEjecutora) >= 0);
}
function hzPuedeCerrar_(h, u){
  if (hzEsMedioAmbiente_(u)) return true;
  return u.rol === "Responsable de área" && h.estado === "Asignado" && !!h.areaEjecutora && (u.areas || []).indexOf(h.areaEjecutora) >= 0;
}

// área: nombre 3–40 caracteres, tipo de la lista; ejecutora/activo son booleanos
function hzValidarArea_(b){
  b = b || {};
  const nombre = hzTxt_(b.nombre);
  if (nombre.length < 3 || nombre.length > 40) return {ok:false, error:"nombre_invalido"};
  if (HZ_TIPOS_AREA.indexOf(b.tipo) < 0) return {ok:false, error:"tipo_invalido"};
  return {ok:true, v:{nombre:nombre, tipo:b.tipo, ejecutora:b.ejecutora === true, activo:b.activo !== false}};
}
// motivo de reapertura: 5–300 caracteres
function hzValidarMotivo_(b){
  b = b || {};
  if (!hzIdValido_(b.id)) return {ok:false, error:"id_invalido"};
  const motivo = hzTxt_(b.motivo);
  if (motivo.length < 5 || motivo.length > 300) return {ok:false, error:"motivo_invalido"};
  return {ok:true, v:{id:String(b.id), motivo:motivo}};
}

// áreas de un usuario (lista o texto separado por |): recortadas, sin repetir (sin distinguir mayúsculas), máx. 20 de hasta 40 caracteres
function areasUsr_(v){
  const L = Array.isArray(v) ? v : String(v == null ? "" : v).split("|"), vistos = {}, R = [];
  L.forEach(x => {
    const a = hzTxt_(x), k = a.toLowerCase();
    if (a && a.length <= 40 && !vistos[k] && R.length < 20){ vistos[k] = 1; R.push(a); }
  });
  return R;
}

// ===================== 2. PERSISTENCIA (Sheet y Drive) =====================
// Usa de Code.gs: hoja_, bitacora_, SEDE. Las pestañas son por planta (cada planta tiene su Sheet).
const HZ_COLS = ["id","sede","creado","reportadoPor","reportadoNombre","area","descripcion","tipo","riesgo","fotos","origen","areaEjecutora","plazo","estado","asignadoPor","asignadoEn","cierreComentario","cierreEvidencia","cerradoPor","cerradoEn","actualizado","cerradoNombre"];
const HZ_AREAS_COLS = ["nombre","tipo","ejecutora","activo"];
const HZ_TZ = "America/Lima";
const HZ_RAIZ = "Hallazgos Ambientales";
const HZ_DIAS_CERRADOS = 90;

function hzAhora_(){ return Utilities.formatDate(new Date(), HZ_TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function hzHoy_(){ return Utilities.formatDate(new Date(), HZ_TZ, "yyyy-MM-dd"); }
function hzHoja_(){ return hoja_("Hallazgos", HZ_COLS, HZ_COLS); }
function hzHojaAreas_(){ return hoja_("Áreas", HZ_AREAS_COLS, HZ_AREAS_COLS); }
// el usuario con sesión, con las áreas siempre como lista
function hzU_(u){
  const a = u.areas;
  return {usuario:u.usuario, nombre:u.nombre || u.usuario, rol:u.rol,
    areas:(Array.isArray(a) ? a : String(a || "").split("|")).map(x => String(x).trim()).filter(Boolean)};
}
function hzJson_(s){ try{ const a = JSON.parse(s || "[]"); return Array.isArray(a) ? a : []; }catch(e){ return []; } }

function hzFilas_(sh){
  const n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n-1, HZ_COLS.length).getValues()
    .map((r,i) => { const o = {_fila:i+2}; HZ_COLS.forEach((c,j) => o[c] = String(r[j] == null ? "" : r[j])); return o; })
    .filter(o => o.sede === SEDE);
}
function hzEscribir_(sh, o){ sh.getRange(o._fila, 1, 1, HZ_COLS.length).setValues([HZ_COLS.map(c => o[c])]); }
function hzBuscar_(sh, id){ return hzFilas_(sh).filter(o => o.id === String(id))[0] || null; }

function hzAreasTodas_(){
  const sh = hzHojaAreas_(), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n-1, HZ_AREAS_COLS.length).getValues().map((r,i) => ({
    _fila:i+2, nombre:String(r[0]), tipo:String(r[1]), ejecutora:String(r[2]) === "Sí", activo:String(r[3]) === "Sí"}));
}

// Drive: Hallazgos Ambientales/<año>/<mm> bajo la raíz de Drive de la cuenta; se crea si falta
function hzSub_(padre, nombre){ const it = padre.getFoldersByName(nombre); return it.hasNext() ? it.next() : padre.createFolder(nombre); }
function hzCarpeta_(fechaISO){
  const p = String(fechaISO).split("-");
  return hzSub_(hzSub_(hzSub_(DriveApp.getRootFolder(), HZ_RAIZ), p[0]), p[1]);
}
// guarda los archivos (data-URL ya validados); si uno falla, borra los ya creados para no dejar huérfanos
function hzGuardarArchivos_(lista, carpeta, prefijo){
  const hechos = [];
  try{
    lista.forEach((s,i) => {
      const d = hzDataUrl_(s), ext = d.mime === "application/pdf" ? "pdf" : "jpg";
      const f = carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(d.b64), d.mime, prefijo + "-" + (i+1) + "." + ext));
      hechos.push({id:f.getId(), mime:d.mime});
    });
  }catch(err){
    hechos.forEach(h => { try{ DriveApp.getFileById(h.id).setTrashed(true); }catch(e){} });
    throw err;
  }
  return hechos;
}

function hzCrear_(b, u){
  u = hzU_(u);
  const val = hzValidarReporte_(b); if (!val.ok) return val;
  const v = val.v, sh = hzHoja_();
  if (hzBuscar_(sh, v.id)) return {ok:true, id:v.id, repetido:true};          // reintento por mala señal: no duplica
  const ar = hzAreasTodas_().filter(a => a.nombre === v.area)[0];
  if (!ar || !ar.activo) return {ok:false, error:"area_invalida"};
  const ahora = hzAhora_();
  const fotos = hzGuardarArchivos_(v.fotos, hzCarpeta_(ahora.slice(0,10)), v.id);
  const o = {id:v.id, sede:SEDE, creado:ahora, reportadoPor:u.usuario, reportadoNombre:u.nombre, area:v.area, descripcion:v.descripcion,
    tipo:v.tipo, riesgo:v.riesgo, fotos:JSON.stringify(fotos), origen:"reporte", areaEjecutora:"", plazo:"", estado:"Abierto",
    asignadoPor:"", asignadoEn:"", cierreComentario:"", cierreEvidencia:"", cerradoPor:"", cerradoEn:"", actualizado:ahora, cerradoNombre:""};
  sh.appendRow(HZ_COLS.map(c => o[c]));
  bitacora_(u, "Hallazgo reportado", v.id + " · " + v.area + " · riesgo " + v.riesgo);
  return {ok:true, id:v.id};
}

function hzAsignar_(b, u){
  u = hzU_(u);
  const hoy = hzHoy_(), val = hzValidarAsignacion_(b, hoy); if (!val.ok) return val;
  const v = val.v, sh = hzHoja_(), h = hzBuscar_(sh, v.id);
  if (!h) return {ok:false, error:"no_existe"};
  if (h.estado === "Cerrado") return {ok:false, error:"estado_invalido"};
  const ar = hzAreasTodas_().filter(a => a.nombre === v.areaEjecutora)[0];
  if (!ar || !ar.activo || !ar.ejecutora) return {ok:false, error:"ejecutora_invalida"};
  const ahora = hzAhora_();
  h.areaEjecutora = v.areaEjecutora; h.plazo = v.plazo; h.estado = "Asignado"; h.asignadoPor = u.usuario; h.asignadoEn = ahora; h.actualizado = ahora;
  hzEscribir_(sh, h);
  bitacora_(u, "Hallazgo asignado", v.id + " → " + v.areaEjecutora + " · plazo " + v.plazo);
  return {ok:true, id:v.id};
}

function hzCerrar_(b, u){
  u = hzU_(u);
  const val = hzValidarCierre_(b); if (!val.ok) return val;
  const v = val.v, sh = hzHoja_(), h = hzBuscar_(sh, v.id);
  if (!h) return {ok:false, error:"no_existe"};
  if (!hzPuedeVer_(h, u)) return {ok:false, error:"permiso"};
  if (h.estado === "Cerrado" && h.cerradoPor === u.usuario) return {ok:true, id:v.id, repetido:true};   // reintento del mismo usuario
  if (h.estado !== "Asignado") return {ok:false, error:"estado_invalido"};                              // otro ya lo cerró, o aún sin asignar
  if (!hzPuedeCerrar_(h, u)) return {ok:false, error:"permiso"};
  const ahora = hzAhora_();
  const ev = hzGuardarArchivos_(v.evidencia, hzCarpeta_(ahora.slice(0,10)), v.id + "-cierre");
  h.estado = "Cerrado"; h.cierreComentario = v.comentario; h.cierreEvidencia = JSON.stringify(ev);
  h.cerradoPor = u.usuario; h.cerradoNombre = u.nombre; h.cerradoEn = ahora; h.actualizado = ahora;
  hzEscribir_(sh, h);
  bitacora_(u, "Hallazgo cerrado", v.id + " · " + ev.length + " archivo(s)");
  return {ok:true, id:v.id};
}

function hzReabrir_(b, u){
  u = hzU_(u);
  const val = hzValidarMotivo_(b); if (!val.ok) return val;
  const v = val.v, sh = hzHoja_(), h = hzBuscar_(sh, v.id);
  if (!h) return {ok:false, error:"no_existe"};
  if (!hzEsMedioAmbiente_(u)) return {ok:false, error:"permiso"};
  if (h.estado !== "Cerrado") return {ok:false, error:"estado_invalido"};
  const evidenciaAnterior = h.cierreEvidencia;
  h.estado = "Asignado"; h.cierreComentario = ""; h.cierreEvidencia = ""; h.cerradoPor = ""; h.cerradoNombre = ""; h.cerradoEn = ""; h.actualizado = hzAhora_();
  hzEscribir_(sh, h);   // área y plazo se conservan
  bitacora_(u, "Hallazgo reabierto", v.id + " · motivo: " + v.motivo + " · evidencia anterior: " + evidenciaAnterior);
  return {ok:true, id:v.id};
}

function hzListar_(u){
  u = hzU_(u);
  const hoy = hzHoy_(), hoyD = hzDia_(hoy), RANGO = {Vencido:0, Abierto:1, Asignado:2, Cerrado:3};
  const L = hzFilas_(hzHoja_())
    .filter(h => hzPuedeVer_(h, u))
    .filter(h => h.estado !== "Cerrado" || !h.cerradoEn || hoyD - hzDia_(h.cerradoEn.slice(0,10)) <= HZ_DIAS_CERRADOS)
    .map(h => ({id:h.id, creado:h.creado, reportadoNombre:h.reportadoNombre, area:h.area, descripcion:h.descripcion, tipo:h.tipo, riesgo:h.riesgo,
      areaEjecutora:h.areaEjecutora, plazo:h.plazo, estado:h.estado, estadoVista:hzEstadoVista_(h, hoy),
      nFotos:hzJson_(h.fotos).length, nEvidencia:hzJson_(h.cierreEvidencia).length,
      cierreComentario:h.cierreComentario, cerradoPor:h.cerradoNombre, cerradoEn:h.cerradoEn, puedeCerrar:hzPuedeCerrar_(h, u)}));
  L.sort((a,b) => (RANGO[a.estadoVista] - RANGO[b.estadoVista]) || (a.creado < b.creado ? 1 : a.creado > b.creado ? -1 : 0));
  return {ok:true, hallazgos:L};
}

// una foto o evidencia a la vez, solo para quien puede ver el hallazgo; nunca se comparte por enlace
function hzFoto_(b, u){
  u = hzU_(u); b = b || {};
  if (!hzIdValido_(b.id) || (b.donde !== "hallazgo" && b.donde !== "cierre")) return {ok:false, error:"id_invalido"};
  const h = hzBuscar_(hzHoja_(), b.id);
  if (!h) return {ok:false, error:"no_existe"};
  if (!hzPuedeVer_(h, u)) return {ok:false, error:"permiso"};
  const lista = hzJson_(b.donde === "cierre" ? h.cierreEvidencia : h.fotos), i = Number(b.idx);
  if (!(i >= 0 && i < lista.length && Math.floor(i) === i)) return {ok:false, error:"no_existe"};
  const f = DriveApp.getFileById(lista[i].id);
  return {ok:true, mime:lista[i].mime, dataUrl:"data:" + lista[i].mime + ";base64," + Utilities.base64Encode(f.getBlob().getBytes())};
}

function hzAreasListar_(u){
  u = hzU_(u);
  const todas = hzAreasTodas_().map(a => ({nombre:a.nombre, tipo:a.tipo, ejecutora:a.ejecutora, activo:a.activo}));
  return {ok:true, areas:hzEsMedioAmbiente_(u) ? todas : todas.filter(a => a.activo)};
}
// el nombre no se cambia: para renombrar se desactiva y se crea otra (igual que el Catálogo)
function hzAreaGuardar_(b, u){
  u = hzU_(u);
  const val = hzValidarArea_(b); if (!val.ok) return val;
  const v = val.v, sh = hzHojaAreas_();
  const ex = hzAreasTodas_().filter(a => a.nombre.toLowerCase() === v.nombre.toLowerCase())[0];
  const fila = [ex ? ex.nombre : v.nombre, v.tipo, v.ejecutora ? "Sí" : "No", v.activo ? "Sí" : "No"];
  if (ex) sh.getRange(ex._fila, 1, 1, HZ_AREAS_COLS.length).setValues([fila]); else sh.appendRow(fila);
  bitacora_(u, "Área creada/modificada", fila.join(" · "));
  return {ok:true, nombre:fila[0], creada:!ex};
}

// Para el editor de Apps Script: crea un hallazgo de prueba, lo asigna, lo cierra y deja el resultado en el registro.
function pruebaHallazgos(){
  const u = {usuario:"prueba", nombre:"Prueba", rol:"Administrador", areas:[]};
  const jpg = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=";
  const id = "prueba-" + new Date().getTime(), hoy = hzHoy_();
  Logger.log("Planta: %s · fecha de Lima: %s", SEDE, hoy);
  Logger.log("Áreas: %s", JSON.stringify([hzAreaGuardar_({nombre:"Prueba Producción", tipo:"Producción", ejecutora:false, activo:true}, u),
    hzAreaGuardar_({nombre:"Prueba Mantenimiento", tipo:"Servicio", ejecutora:true, activo:true}, u)]));
  Logger.log("Crear: %s", JSON.stringify(hzCrear_({id:id, area:"Prueba Producción", descripcion:"PRUEBA: borrar esta fila", tipo:"Otro", riesgo:"Bajo", fotos:[jpg]}, u)));
  Logger.log("Crear de nuevo (debe decir repetido): %s", JSON.stringify(hzCrear_({id:id, area:"Prueba Producción", descripcion:"PRUEBA: borrar esta fila", tipo:"Otro", riesgo:"Bajo", fotos:[jpg]}, u)));
  Logger.log("Asignar: %s", JSON.stringify(hzAsignar_({id:id, areaEjecutora:"Prueba Mantenimiento", plazo:hoy}, u)));
  Logger.log("Cerrar: %s", JSON.stringify(hzCerrar_({id:id, comentario:"prueba", evidencia:[jpg]}, u)));
  Logger.log("Cerrar otra vez como otro usuario (debe decir estado_invalido): %s", JSON.stringify(hzCerrar_({id:id, evidencia:[jpg]}, {usuario:"otro", nombre:"Otro", rol:"Administrador", areas:[]})));
  const L = hzListar_(u).hallazgos.filter(h => h.id === id)[0];
  Logger.log("En la lista: %s", JSON.stringify(L));
  Logger.log("Foto servida por la app (largo del dataUrl): %s", hzFoto_({id:id, donde:"hallazgo", idx:0}, u).dataUrl.length);
  Logger.log("Carpeta en Drive: %s/%s/%s", HZ_RAIZ, hoy.slice(0,4), hoy.slice(5,7));
  Logger.log("Listo. Puedes borrar la fila «PRUEBA» en la pestaña Hallazgos y las áreas «Prueba …» en Áreas.");
}

if (typeof module !== "undefined") module.exports = {
  HZ_TIPOS, HZ_RIESGOS, HZ_TIPOS_AREA, HZ_MAX_FOTO, HZ_MAX_PDF, HZ_MAX_ARCHIVOS, HZ_COLS, HZ_AREAS_COLS,
  hzTxt_, areasUsr_, hzIdValido_, hzDataUrl_, hzDia_, hzValidarReporte_, hzValidarAsignacion_, hzValidarCierre_, hzValidarArea_, hzValidarMotivo_,
  hzEstadoVista_, hzEsMedioAmbiente_, hzPuedeVer_, hzPuedeCerrar_,
  hzCrear_, hzAsignar_, hzCerrar_, hzReabrir_, hzListar_, hzFoto_, hzAreasListar_, hzAreaGuardar_
};
