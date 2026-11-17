/* Hallazgos ambientales · pestaña de la app: lista única, reporte ágil con foto marcada y detalle.
   Se carga después del script de index.html y usa sus funciones: $, esc, uuid, toast, postJ, can, SES, CFG, TAB, showView.
   El servidor decide estado, visibilidad y permisos; aquí solo se muestra, se comprime y se marca. */
(function(){
"use strict";

const TIPOS = (CFG.TIPOS_HALLAZGO && CFG.TIPOS_HALLAZGO.length) ? CFG.TIPOS_HALLAZGO : ["Residuos","Agua y efluentes","Emisiones y olores","Derrames y químicos","Energía","Orden y limpieza","Otro"];
const RIESGOS = (CFG.RIESGOS_HALLAZGO && CFG.RIESGOS_HALLAZGO.length) ? CFG.RIESGOS_HALLAZGO : ["Bajo","Medio","Alto"];
const MAX_FOTOS = 3;
const ERR = {
  area_invalida:"El área elegida no existe o está inactiva.", descripcion_invalida:"La descripción debe tener entre 5 y 500 caracteres.",
  tipo_invalido:"Elija el tipo de hallazgo.", riesgo_invalido:"Elija el nivel de riesgo.", fotos_invalidas:"Las fotos no son válidas (1 a 3 fotos JPEG).",
  id_invalido:"El reporte no es válido. Vuelva a empezar.", permiso:"No tiene permiso para esta acción.", no_existe:"El hallazgo ya no existe.",
  estado_invalido:"Otra persona ya cambió este hallazgo. Se actualizó la lista.", ejecutora_invalida:"Elija un área ejecutante activa.",
  plazo_invalido:"El plazo debe ser desde hoy y hasta 365 días.", evidencia_requerida:"Agregue al menos una evidencia (foto o PDF).",
  evidencia_invalida:"La evidencia no es válida: 1 a 3 fotos, o un solo PDF de hasta 1.5 MB.", comentario_invalido:"El comentario admite hasta 300 caracteres.",
  motivo_invalido:"El motivo debe tener entre 5 y 300 caracteres.", nombre_invalido:"El nombre del área debe tener entre 3 y 40 caracteres.", tipo_area_invalido:"Elija el tipo de área."
};
const S = {lista:null, t:0, cargando:null, accion:null, areasAdm:null, err:"", filtro:"pend", vista:"lista", id:null, areas:null, F:null, enviando:false, cache:{}};

/* ---------- estilos propios (prefijo hz-) ---------- */
const st = document.createElement("style");
st.textContent = `
.hz-top{display:flex;gap:8px;margin:8px 0}
.hz-chips{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0}
.hz-chips button{border:1.5px solid var(--taupe);background:#fff;color:var(--brown);border-radius:999px;padding:6px 12px;font-weight:bold;font-size:13px}
.hz-chips button.on{background:var(--brown);border-color:var(--brown);color:#fff}
.hz-row{display:block;width:100%;text-align:left;background:#fff;border:1px solid var(--line);border-left:5px solid var(--taupe);border-radius:12px;padding:10px 12px;margin:6px 0;font:inherit;color:inherit}
.hz-row.venc{border-left-color:var(--red);background:var(--red-bg)}
.hz-row.abi{border-left-color:var(--orange)} .hz-row.asg{border-left-color:#4A6FA5} .hz-row.cer{border-left-color:var(--ok)}
.hz-row .l1{display:flex;gap:8px;align-items:center;justify-content:space-between}
.hz-row b{color:var(--brown)} .hz-row .d{margin:4px 0;font-size:14px} .hz-row .m{font-size:12.5px;color:var(--gray)}
.hz-b{font-size:11.5px;font-weight:bold;border-radius:999px;padding:2px 9px;white-space:nowrap}
.hz-b.venc{background:var(--red);color:#fff} .hz-b.abi{background:#FFE3C2;color:#9a4d00} .hz-b.asg{background:#DCE6F5;color:#2F4C7A} .hz-b.cer{background:#DDEFD5;color:var(--ok)}
.hz-alto{color:var(--red);font-weight:bold}
.hz-thumbs{display:flex;gap:8px;flex-wrap:wrap;margin:6px 0}
.hz-thumb{position:relative;width:96px;height:96px;border-radius:10px;overflow:hidden;border:2px solid var(--taupe)}
.hz-thumb img{width:100%;height:100%;object-fit:cover;display:block}
.hz-thumb button{position:absolute;top:2px;right:2px;width:28px;height:28px;border-radius:50%;border:0;background:rgba(0,0,0,.65);color:#fff;font-size:15px}
.hz-fotos{display:flex;gap:8px;flex-wrap:wrap;margin:6px 0}
.hz-fotos .hz-ph{min-width:120px;min-height:90px;border:1px dashed var(--taupe);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:12.5px;color:var(--gray);padding:6px;text-align:center}
.hz-fotos img{max-width:100%;width:240px;border-radius:10px;border:1px solid var(--line)} .hz-fotos img.g{width:100%}
.hz-kv{display:grid;grid-template-columns:110px 1fr;gap:4px 8px;font-size:14px;margin:8px 0}.hz-kv span:nth-child(odd){color:var(--gray);font-size:12.5px}
.hz-acc{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.hz-acc .btn{flex:1;min-width:140px}
`;
document.head.appendChild(st);

/* ---------- utilidades ---------- */
const box = () => document.getElementById("hall-body");
const fecha = s => s ? String(s).slice(8,10) + "/" + String(s).slice(5,7) + "/" + String(s).slice(0,4) : "";
const CL = {Vencido:"venc", Abierto:"abi", Asignado:"asg", Cerrado:"cer"};
const badge = h => `<span class="hz-b ${CL[h.estadoVista]||""}">${esc(h.estadoVista)}</span>`;
const simple = () => typeof SIMPLE === "function" && SIMPLE();
const porFiltro = {
  pend:h => h.estadoVista !== "Cerrado", abi:h => h.estadoVista === "Abierto", asg:h => h.estadoVista === "Asignado",
  ven:h => h.estadoVista === "Vencido", cer:h => h.estadoVista === "Cerrado"
};
const NOMF = {pend:"Pendientes", abi:"Abiertos", asg:"Asignados", ven:"Vencidos", cer:"Cerrados"};

/* ---------- datos ---------- */
function carga(){
  if(S.cargando) return S.cargando;
  S.cargando = (async () => {
    try{
      const j = await postJ({action:"hallazgos_listar"});
      if(j && j.ok){ S.lista = j.hallazgos; S.t = Date.now(); S.err = ""; }
      else if(!(j && (j.error === "sesion" || j.error === "cambiar_clave"))) S.err = "No se pudo cargar la lista (" + ((j && j.error) || "error") + ").";
    }catch(e){ S.err = "Sin conexión. Se muestra lo último cargado."; }
    finally{ S.cargando = null; repinta(); }
  })();
  return S.cargando;
}
async function cargaAreas(){
  try{
    const j = await postJ({action:"areas_listar"});
    if(j && j.ok){ S.areas = j.areas; if(S.vista === "nuevo") pintaNuevo(); }
  }catch(e){}
}

/* ---------- pantallas ---------- */
function drawHall(){
  if(!box()) return;
  repinta();
  if(S.vista !== "nuevo" && !S.cargando && (S.lista === null || Date.now() - S.t > 15000)) carga();
  if(S.areas === null) cargaAreas();
}
function repinta(){
  if(!box() || TAB !== "hall") return;
  if(S.vista === "nuevo") return pintaNuevo();
  if(S.vista === "detalle") return pintaDetalle();
  pintaLista();
}
function pintaLista(){
  const L = S.lista || [], vis = L.filter(porFiltro[S.filtro] || porFiltro.pend);
  const cnt = k => L.filter(porFiltro[k]).length;
  box().innerHTML =
    (simple() ? `<div class="hz-top"><button class="btn sec sm" data-hz="atras">← Atrás</button></div>` : "") +
    `<div class="hz-top"><button class="btn" data-hz="nuevo">＋ Reportar hallazgo</button></div>` +
    `<div class="hz-chips">${Object.keys(NOMF).map(k => `<button type="button" data-hz-filtro="${k}" class="${S.filtro===k?"on":""}">${NOMF[k]} (${cnt(k)})</button>`).join("")}</div>` +
    (S.err ? `<div class="err">${esc(S.err)}</div>` : "") +
    (S.lista === null ? `<div class="card muted">Cargando…</div>`
      : vis.length ? vis.map(fila).join("")
      : `<div class="card muted">${L.length ? "No hay hallazgos en este filtro." : "Todavía no hay hallazgos. Toque «Reportar hallazgo» para registrar el primero."}</div>`);
}
function fila(h){
  const d = String(h.descripcion || ""), c = CL[h.estadoVista] || "";
  return `<button class="hz-row ${c}" type="button" data-hz-abrir="${esc(h.id)}">
    <div class="l1"><b>${esc(h.area)}</b>${badge(h)}</div>
    <div class="d">${esc(d.length > 90 ? d.slice(0, 90) + "…" : d)}</div>
    <div class="m">${esc(h.tipo)} · <span class="${h.riesgo==="Alto"?"hz-alto":""}">riesgo ${esc(h.riesgo.toLowerCase())}</span> · ${esc(fecha(h.creado))}${h.areaEjecutora ? ` · Ejecuta: <b>${esc(h.areaEjecutora)}</b>` : ""}${h.plazo ? ` · Plazo ${esc(fecha(h.plazo))}` : ""}</div>
  </button>`;
}

/* ---------- reporte ágil ---------- */
function nuevoHallazgo(){
  const propia = (SES && SES.areas && SES.areas[0]) || "";
  S.F = {id:uuid(), area:propia, tipo:"", riesgo:"Medio", desc:"", fotos:[]};   // el id se reutiliza si hay que reintentar: no duplica
  S.vista = "nuevo"; S.enviando = false;
  if(S.areas === null) cargaAreas();
  repinta(); window.scrollTo(0, 0);
}
function leeCampos(){
  const F = S.F; if(!F) return;
  const a = document.getElementById("hz-area"), t = document.getElementById("hz-tipo"), d = document.getElementById("hz-desc");
  if(a) F.area = a.value; if(t) F.tipo = t.value; if(d) F.desc = d.value;
}
function pintaNuevo(){
  const F = S.F; if(!F || !box()) return;
  const areas = (S.areas || []).filter(a => a.activo);
  if(F.area && areas.length && !areas.some(a => a.nombre === F.area)) F.area = "";
  box().innerHTML = `<div class="hz-top"><button class="btn sec sm" data-hz="volver">← Volver</button></div>
    <div class="card"><h2>Reportar hallazgo</h2>
      ${S.areas !== null && !areas.length ? `<div class="err">No hay áreas registradas. Pida a Medio Ambiente que las cree en la pestaña Áreas.</div>` : ""}
      <label class="f" for="hz-area">Área</label>
      <select id="hz-area" class="in"><option value="">Elegir…</option>${areas.map(a => `<option ${a.nombre===F.area?"selected":""}>${esc(a.nombre)}</option>`).join("")}</select>
      <label class="f" for="hz-tipo">Tipo</label>
      <select id="hz-tipo" class="in"><option value="">Elegir…</option>${TIPOS.map(t => `<option ${t===F.tipo?"selected":""}>${esc(t)}</option>`).join("")}</select>
      <label class="f">Riesgo</label>
      <div class="seg">${RIESGOS.map(r => `<button type="button" data-hz-riesgo="${esc(r)}" class="${r===F.riesgo?"on":""}">${esc(r)}</button>`).join("")}</div>
      <label class="f" for="hz-desc">¿Qué se observa?</label>
      <textarea id="hz-desc" class="in" rows="3" maxlength="500" placeholder="Ej.: fuga de agua en el caño del vestuario de hombres">${esc(F.desc)}</textarea>
      <label class="f">Fotos con la observación marcada (1 a ${MAX_FOTOS})</label>
      <div class="hz-thumbs">${F.fotos.map((f, i) => `<div class="hz-thumb"><img src="${f}" alt="Foto ${i+1}"><button type="button" data-hz-quitar="${i}" aria-label="Quitar foto">✕</button></div>`).join("")}</div>
      ${F.fotos.length < MAX_FOTOS ? `<button class="btn sec" type="button" data-hz="foto">📷 ${F.fotos.length ? "Agregar otra foto" : "Tomar o elegir foto"}</button>` : ""}
      <input type="file" id="hz-file" accept="image/*" class="hidden">
      <div class="err" id="hz-err"></div>
      <div style="height:10px"></div>
      <button class="btn" type="button" data-hz="enviar" ${S.enviando?"disabled":""}>${S.enviando ? "Enviando…" : "Enviar hallazgo"}</button>
    </div>`;
}
async function agregaFoto(file){
  if(!file || !S.F || S.F.fotos.length >= MAX_FOTOS) return;
  const e = document.getElementById("hz-err");
  try{
    const c = await comprimirFoto(file);
    const m = await marcarFoto(c.dataUrl);          // pantalla completa: círculo, flecha o trazo; sin marca no se puede continuar
    if(m) S.F.fotos.push(m);
  }catch(x){ toast("No se pudo leer la foto. Intente de nuevo."); }
  leeCampos(); pintaNuevo();
}
async function enviar(){
  const F = S.F; if(!F || S.enviando) return;
  leeCampos();
  const e = document.getElementById("hz-err"); e.textContent = "";
  if(!F.area) return void(e.textContent = "Elija el área.");
  if(!F.tipo) return void(e.textContent = "Elija el tipo de hallazgo.");
  if(F.desc.trim().length < 5) return void(e.textContent = "Describa la observación (mínimo 5 letras).");
  if(!F.fotos.length) return void(e.textContent = "Agregue al menos una foto con la observación marcada.");
  if(!navigator.onLine) return void(e.textContent = "Sin conexión. Su reporte se conserva: toque Enviar cuando tenga señal.");
  S.enviando = true; pintaNuevo();
  try{
    const j = await postJ({action:"hallazgo_crear", id:F.id, area:F.area, tipo:F.tipo, riesgo:F.riesgo, descripcion:F.desc, fotos:F.fotos}, 120000);
    S.enviando = false;
    if(j && j.ok){ toast(j.repetido ? "Ese hallazgo ya estaba registrado." : "Hallazgo reportado."); S.F = null; S.vista = "lista"; S.t = 0; repinta(); carga(); return; }
    pintaNuevo();
    const e2 = document.getElementById("hz-err");
    if(j && (j.error === "sesion" || j.error === "cambiar_clave")) e2.textContent = "Su sesión venció. Ingrese de nuevo: su reporte se conserva.";
    else e2.textContent = ERR[j && j.error] || "No se pudo enviar (" + ((j && j.error) || "error") + ").";
  }catch(x){
    S.enviando = false; pintaNuevo();
    document.getElementById("hz-err").textContent = "Se cortó la señal. Su reporte se conserva: vuelva a tocar Enviar (no se duplicará).";
  }
}

/* ---------- detalle ---------- */
function abrirHallazgo(id){ S.vista = "detalle"; S.id = id; S.accion = null; repinta(); window.scrollTo(0, 0); }
function pintaDetalle(){
  const h = (S.lista || []).find(x => x.id === S.id);
  if(!h){ S.vista = "lista"; return pintaLista(); }
  const ph = (donde, n) => Array.from({length:n}, (_, i) => `<div class="hz-ph" data-hz-ph="${donde}|${i}">Cargando…</div>`).join("");
  box().innerHTML = `<div class="hz-top"><button class="btn sec sm" data-hz="volver">← Volver</button></div>
    <div class="card"><div style="display:flex;justify-content:space-between;align-items:center;gap:8px"><h2 style="margin:0">${esc(h.area)}</h2>${badge(h)}</div>
      <div class="hz-kv">
        <span>Descripción</span><span>${esc(h.descripcion)}</span>
        <span>Tipo</span><span>${esc(h.tipo)}</span>
        <span>Riesgo</span><span class="${h.riesgo==="Alto"?"hz-alto":""}">${esc(h.riesgo)}</span>
        <span>Reportó</span><span>${esc(h.reportadoNombre)} · ${esc(fecha(h.creado))}</span>
        ${h.areaEjecutora ? `<span>Ejecuta</span><span>${esc(h.areaEjecutora)}</span>` : ""}
        ${h.plazo ? `<span>Plazo</span><span>${esc(fecha(h.plazo))}</span>` : ""}
      </div>
      <label class="f">Fotos</label><div class="hz-fotos">${ph("hallazgo", h.nFotos)}</div>
      ${h.estado === "Cerrado" ? `<label class="f">Cierre</label>
        <div class="hz-kv"><span>Cerró</span><span>${esc(h.cerradoPor)} · ${esc(fecha(h.cerradoEn))}</span>${h.cierreComentario ? `<span>Comentario</span><span>${esc(h.cierreComentario)}</span>` : ""}</div>
        <div class="hz-fotos">${ph("cierre", h.nEvidencia)}</div>` : ""}
      ${typeof accionesHall === "function" ? accionesHall(h) : ""}
    </div>`;
  cargaFotos(h);
}
async function cargaFotos(h){
  const slots = [...box().querySelectorAll("[data-hz-ph]")];
  for(const el of slots){
    if(S.vista !== "detalle" || S.id !== h.id) return;                     // salió del detalle: se deja de pedir
    const [donde, idx] = el.dataset.hzPh.split("|"), k = h.id + "|" + donde + "|" + idx;
    try{
      let r = S.cache[k];
      if(!r){ const j = await postJ({action:"hallazgo_foto", id:h.id, donde, idx:+idx}, 60000); if(!(j && j.ok)) throw 0; r = S.cache[k] = {mime:j.mime, dataUrl:j.dataUrl}; }
      if(!el.isConnected) return;
      if(r.mime === "application/pdf"){
        const b = await (await fetch(r.dataUrl)).blob(), u = URL.createObjectURL(b);
        el.outerHTML = `<a class="btn sec sm" href="${u}" target="_blank" rel="noopener">📄 Abrir PDF de evidencia</a>`;
      } else el.outerHTML = `<img src="${r.dataUrl}" alt="Foto" data-hz-zoom="1">`;
    }catch(e){ if(el.isConnected){ el.textContent = "No se pudo cargar. Toque para reintentar."; el.dataset.hzReint = "1"; } }
  }
}


/* ---------- asignar, cerrar y reabrir ---------- */
const hoyStr = () => (typeof todayStr === "function" ? todayStr() : new Date().toISOString().slice(0, 10));
function masDias(f, n){ const d = new Date(f + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
const leerArchivo = f => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = no; r.readAsDataURL(f); });
const ejecutoras = () => (S.areas || []).filter(a => a.activo && a.ejecutora);

function accionesHall(h){
  if(S.accion && S.accion.id === h.id) return panelAccion(h, S.accion);
  const b = [];
  if(can("hallazgo_asignar") && h.estado !== "Cerrado") b.push(`<button class="btn ${h.estado==="Abierto"?"":"sec"}" data-hz="asignar">${h.estado==="Abierto" ? "Asignar área y plazo" : "Reasignar o cambiar plazo"}</button>`);
  if(h.puedeCerrar && h.estado === "Asignado") b.push(`<button class="btn" data-hz="cerrar">Cerrar con evidencia</button>`);
  if(can("hallazgo_reabrir") && h.estado === "Cerrado") b.push(`<button class="btn sec" data-hz="reabrir">Reabrir</button>`);
  if(b.length) return `<div class="hz-acc">${b.join("")}</div>`;
  return h.estado === "Abierto" ? `<div class="muted" style="margin-top:8px">Medio Ambiente asignará el área que lo atenderá y el plazo.</div>` : "";
}
function iniciaAccion(t){
  const h = (S.lista || []).find(x => x.id === S.id); if(!h) return;
  S.accion = {t, id:h.id, ej:h.areaEjecutora || "", plazo:h.plazo || "", com:"", mot:"", ev:[], pdf:""};
  if(S.areas === null) cargaAreas();
  repinta();
}
function panelAccion(h, A){
  const err = `<div class="err" id="hz-aerr"></div>`, bot = txt => `<div class="hz-acc"><button class="btn sec" data-hz="cancelar">Cancelar</button><button class="btn" data-hz="guardar" ${S.enviando?"disabled":""}>${S.enviando ? "Enviando…" : txt}</button></div>`;
  if(A.t === "asignar"){
    const E = ejecutoras(), hoy = hoyStr();
    return `<div class="card" style="background:#FBF7F2"><h2>Asignar</h2>
      ${S.areas !== null && !E.length ? `<div class="err">No hay áreas ejecutantes. Créelas en la pestaña Áreas (marque «Ejecutora»).</div>` : ""}
      <label class="f" for="hz-ej">Área ejecutante</label>
      <select id="hz-ej" class="in"><option value="">Elegir…</option>${E.map(a => `<option ${a.nombre===A.ej?"selected":""}>${esc(a.nombre)}</option>`).join("")}</select>
      <label class="f" for="hz-plazo">Plazo de ejecución</label>
      <input type="date" id="hz-plazo" class="in" min="${hoy}" max="${masDias(hoy, 365)}" value="${esc(A.plazo)}">
      ${err}${bot("Guardar asignación")}</div>`;
  }
  if(A.t === "cerrar"){
    const th = A.ev.map((f, i) => A.pdf ? "" : `<div class="hz-thumb"><img src="${f}" alt="Evidencia ${i+1}"><button type="button" data-hz-evq="${i}" aria-label="Quitar">✕</button><button type="button" data-hz-evm="${i}" style="left:2px;right:auto" aria-label="Marcar">✎</button></div>`).join("");
    return `<div class="card" style="background:#FBF7F2"><h2>Cerrar con evidencia</h2>
      <div class="muted">1 a ${MAX_FOTOS} fotos (puede marcarlas con ✎) o un PDF pequeño de hasta 1.5 MB.</div>
      <div class="hz-thumbs">${th}${A.pdf ? `<div class="hz-b asg">📄 ${esc(A.pdf)} <a href="#" data-hz-evq="0" style="color:inherit">✕</a></div>` : ""}</div>
      ${(!A.pdf && A.ev.length < MAX_FOTOS) ? `<button class="btn sec" type="button" data-hz="evfile">📎 ${A.ev.length ? "Agregar otra" : "Foto o PDF"}</button>` : ""}
      <input type="file" id="hz-file2" accept="image/*,application/pdf" class="hidden">
      <label class="f" for="hz-com">Comentario (opcional)</label>
      <textarea id="hz-com" class="in" rows="2" maxlength="300">${esc(A.com)}</textarea>
      ${err}${bot("Cerrar hallazgo")}</div>`;
  }
  return `<div class="card" style="background:#FBF7F2"><h2>Reabrir</h2>
    <label class="f" for="hz-mot">Motivo (obligatorio)</label>
    <textarea id="hz-mot" class="in" rows="2" maxlength="300" placeholder="Ej.: la fuga continúa">${esc(A.mot)}</textarea>
    ${err}${bot("Reabrir hallazgo")}</div>`;
}
function leeAccion(){
  const A = S.accion; if(!A) return;
  const v = id => { const e = document.getElementById(id); return e ? e.value : null; };
  if(v("hz-ej") !== null) A.ej = v("hz-ej"); if(v("hz-plazo") !== null) A.plazo = v("hz-plazo");
  if(v("hz-com") !== null) A.com = v("hz-com"); if(v("hz-mot") !== null) A.mot = v("hz-mot");
}
async function agregaEvidencia(file){
  const A = S.accion; if(!file || !A || A.t !== "cerrar") return;
  leeAccion();
  try{
    if(file.type === "application/pdf"){
      if(A.ev.length) throw "No mezcle PDF con fotos: quite lo anterior primero.";
      if(file.size > 1.5 * 1024 * 1024) throw "El PDF pesa más de 1.5 MB. Use una foto o un PDF más liviano.";
      A.ev = [await leerArchivo(file)]; A.pdf = file.name;
    } else {
      if(A.pdf) throw "Ya cargó un PDF: quítelo para usar fotos.";
      if(A.ev.length >= MAX_FOTOS) throw "Máximo " + MAX_FOTOS + " fotos.";
      A.ev.push((await comprimirFoto(file)).dataUrl);
    }
  }catch(x){ toast(typeof x === "string" ? x : "No se pudo leer el archivo."); }
  repinta();
}
async function guardarAccion(){
  const A = S.accion; if(!A || S.enviando) return;
  leeAccion();
  const e = document.getElementById("hz-aerr"); e.textContent = "";
  let body, ok;
  if(A.t === "asignar"){
    if(!A.ej) return void(e.textContent = "Elija el área ejecutante.");
    if(!A.plazo) return void(e.textContent = "Elija el plazo.");
    if(A.plazo < hoyStr()) return void(e.textContent = ERR.plazo_invalido);
    body = {action:"hallazgo_asignar", id:A.id, areaEjecutora:A.ej, plazo:A.plazo}; ok = "Hallazgo asignado a " + A.ej + ".";
  } else if(A.t === "cerrar"){
    if(!A.ev.length) return void(e.textContent = ERR.evidencia_requerida);
    body = {action:"hallazgo_cerrar", id:A.id, comentario:A.com, evidencia:A.ev}; ok = "Hallazgo cerrado.";
  } else {
    if(A.mot.trim().length < 5) return void(e.textContent = "Escriba el motivo (mínimo 5 letras).");
    body = {action:"hallazgo_reabrir", id:A.id, motivo:A.mot}; ok = "Hallazgo reabierto.";
  }
  if(!navigator.onLine) return void(e.textContent = "Sin conexión. Los datos se conservan: reintente cuando tenga señal.");
  S.enviando = true; repinta();
  try{
    const j = await postJ(body, 120000);
    S.enviando = false;
    if(j && j.ok){
      Object.keys(S.cache).filter(k => k.indexOf(A.id + "|cierre") === 0).forEach(k => delete S.cache[k]);
      toast(j.repetido ? "Ya estaba registrado." : ok); S.accion = null; await carga(); return;
    }
    if(j && (j.error === "estado_invalido" || j.error === "permiso" || j.error === "no_existe")){   // otra persona actuó antes: se avisa y se refresca
      toast(ERR[j.error]); S.accion = null; await carga(); return;
    }
    repinta();
    const e2 = document.getElementById("hz-aerr");
    if(e2) e2.textContent = (j && (j.error === "sesion" || j.error === "cambiar_clave")) ? "Su sesión venció. Ingrese de nuevo: lo escrito se conserva." : (ERR[j && j.error] || "No se pudo guardar (" + ((j && j.error) || "error") + ").");
  }catch(x){
    S.enviando = false; repinta();
    const e2 = document.getElementById("hz-aerr"); if(e2) e2.textContent = "Se cortó la señal. Vuelva a tocar el botón (no se duplicará).";
  }
}

/* ---------- pestaña Áreas (Inspector y Administrador) ---------- */
const TIPOS_AREA = ["Producción","Servicio","Común"];
async function drawAreas(){
  if(!document.getElementById("are-body")) return;
  pintaAreas();
  try{
    const j = await postJ({action:"areas_listar"});
    if(j && j.ok){ S.areasAdm = j.areas; S.areas = j.areas; pintaAreas(); }
    else if(!(j && (j.error === "sesion" || j.error === "cambiar_clave"))) S.areaErr = "No se pudo cargar (" + ((j && j.error) || "error") + ").";
  }catch(e){ S.areaErr = "Sin conexión con el Google Sheet."; }
}
function pintaAreas(){
  const b = document.getElementById("are-body"); if(!b) return;
  const L = S.areasAdm || [], sel = t => TIPOS_AREA.map(x => `<option ${x===t?"selected":""}>${x}</option>`).join("");
  b.innerHTML = `<div class="card"><h2>Crear área</h2>
      <div class="muted">Las áreas se eligen al reportar un hallazgo. Marque «Ejecutora» si atiende hallazgos (Mantenimiento, Proyectos, Servicios Generales…).</div>
      <label class="f" for="ar-nom">Nombre</label><input id="ar-nom" class="in" maxlength="40" autocomplete="off">
      <div class="row2"><div><label class="f" for="ar-tipo">Tipo</label><select id="ar-tipo" class="in">${sel("")}</select></div>
        <div style="display:flex;align-items:flex-end"><label style="display:flex;gap:8px;align-items:center;font-size:14px;min-height:48px"><input type="checkbox" id="ar-eje"> Ejecutora</label></div></div>
      <div class="err" id="ar-err">${esc(S.areaErr || "")}</div>
      <div style="height:8px"></div><button class="btn" data-hz="area-crear">Crear área</button></div>
    <div class="card"><h2>Áreas de la planta</h2>
      <div class="muted">El nombre no se cambia: para renombrar, desactive el área y cree otra.</div>
      ${S.areasAdm === null ? `<div class="muted" style="margin-top:8px">Cargando…</div>` : L.length ? L.map(a => `<div class="urow"><div class="nm">${esc(a.nombre)} <span class="bdg ${a.activo?"ap":"inact"}">${a.activo?"activa":"inactiva"}</span>${a.ejecutora?' <span class="bdg q">ejecutora</span>':""}</div>
        <div class="acts"><select class="in" data-hz-are-tipo="${esc(a.nombre)}">${sel(a.tipo)}</select>
        <label style="display:flex;gap:6px;align-items:center;font-size:13px"><input type="checkbox" data-hz-are-eje="${esc(a.nombre)}" ${a.ejecutora?"checked":""}> Ejecutora</label>
        <button class="btn sec sm2" data-hz-are-act="${esc(a.nombre)}">${a.activo?"Desactivar":"Activar"}</button></div></div>`).join("") : `<div class="muted" style="margin-top:8px">Todavía no hay áreas.</div>`}
    </div>`;
}
async function guardaArea(a, msg){
  S.areaErr = "";
  try{
    const j = await postJ({action:"area_guardar", nombre:a.nombre, tipo:a.tipo, ejecutora:!!a.ejecutora, activo:a.activo !== false});
    if(j && j.ok){ if(msg) toast(msg); S.areas = null; await drawAreas(); return true; }
    const m = ERR[j && j.error] || ERR[(j && j.error) + ""] || "No se pudo guardar (" + ((j && j.error) || "error") + ").";
    if(a._nuevo){ S.areaErr = j && j.error === "tipo_invalido" ? ERR.tipo_area_invalido : m; pintaAreas(); } else { toast(m); await drawAreas(); }
  }catch(e){ toast("Sin conexión con el Google Sheet."); }
  return false;
}

/* ---------- Usuarios: áreas del «Responsable de área» (lo usa la pestaña Usuarios de index.html) ---------- */
async function hzCargaAreasUsr(){
  try{ const j = await postJ({action:"areas_listar"}); if(j && j.ok) S.areas = j.areas; }catch(e){}
  pintaUsrAreasNuevo();
}
function pintaUsrAreasNuevo(){
  const r = document.getElementById("u-rol"), bx = document.getElementById("u-arbox"), c = document.getElementById("u-areas"); if(!r || !bx || !c) return;
  const marc = new Set([...c.querySelectorAll("input:checked")].map(i => i.value)), A = (S.areas || []).filter(a => a.activo);
  bx.classList.toggle("hidden", r.value !== "Responsable de área");
  c.innerHTML = A.length ? A.map(a => `<label><input type="checkbox" value="${esc(a.nombre)}" ${marc.has(a.nombre)?"checked":""}> ${esc(a.nombre)}</label>`).join("") : `<span class="muted">No hay áreas todavía: créelas en la pestaña Áreas. Puede dejarlo sin áreas y asignarlas después.</span>`;
}
function hzAreasMarcadas(){ return [...document.querySelectorAll("#u-areas input:checked")].map(i => i.value); }
function hzUsrAreas(u){
  if(u.rol !== "Responsable de área") return "";
  const act = (S.areas || []).filter(a => a.activo).map(a => a.nombre), mias = u.areas || [], todas = act.concat(mias.filter(n => act.indexOf(n) < 0));
  return `<div class="plchk" style="margin-top:6px"><span class="muted">Áreas que lidera:</span>${todas.length ? todas.map(n => `<label><input type="checkbox" data-uar="${esc(u.usuario)}" value="${esc(n)}" ${mias.indexOf(n)>=0?"checked":""}> ${esc(n)}</label>`).join("") : `<span class="muted">Sin áreas todavía</span>`}</div>`;
}

/* ---------- eventos ---------- */
document.addEventListener("click", ev => {
  const t = ev.target;
  const ir = t.closest("[data-hz-ir]");
  if(ir){ TAB = "hall"; S.vista = "lista"; showView(); return; }
  if(!t.closest("#v-hall") && !t.closest("#v-are")) return;
  const fl = t.closest("[data-hz-filtro]"); if(fl){ S.filtro = fl.dataset.hzFiltro; pintaLista(); return; }
  const ab = t.closest("[data-hz-abrir]"); if(ab){ abrirHallazgo(ab.dataset.hzAbrir); return; }
  const rg = t.closest("[data-hz-riesgo]"); if(rg){ leeCampos(); S.F.riesgo = rg.dataset.hzRiesgo; pintaNuevo(); return; }
  const qt = t.closest("[data-hz-quitar]"); if(qt){ leeCampos(); S.F.fotos.splice(+qt.dataset.hzQuitar, 1); pintaNuevo(); return; }
  const re = t.closest("[data-hz-reint]"); if(re){ const h = (S.lista || []).find(x => x.id === S.id); if(h) pintaDetalle(); return; }
  const zm = t.closest("[data-hz-zoom]"); if(zm){ zm.classList.toggle("g"); return; }
  const evq = t.closest("[data-hz-evq]"); if(evq){ ev.preventDefault(); leeAccion(); const A = S.accion; if(A.pdf){ A.ev = []; A.pdf = ""; } else A.ev.splice(+evq.dataset.hzEvq, 1); repinta(); return; }
  const evm = t.closest("[data-hz-evm]"); if(evm){ leeAccion(); const i = +evm.dataset.hzEvm; marcarFoto(S.accion.ev[i], {obligatoria:false}).then(m => { if(m && S.accion) S.accion.ev[i] = m; repinta(); }); return; }
  const at = t.closest("[data-hz-are-tipo]"); if(at) return;
  const aa = t.closest("[data-hz-are-act]"); if(aa){ const a = (S.areasAdm || []).find(x => x.nombre === aa.dataset.hzAreAct); if(a) guardaArea({nombre:a.nombre, tipo:a.tipo, ejecutora:a.ejecutora, activo:!a.activo}, a.activo ? "Área desactivada." : "Área activada."); return; }
  const a = t.closest("[data-hz]"); if(!a) return;
  const k = a.dataset.hz;
  if(k === "area-crear"){
    const nombre = document.getElementById("ar-nom").value.trim().replace(/\s+/g, " ");
    if(nombre.length < 3){ S.areaErr = ERR.nombre_invalido; return pintaAreas(); }
    guardaArea({nombre, tipo:document.getElementById("ar-tipo").value, ejecutora:document.getElementById("ar-eje").checked, activo:true, _nuevo:true}, "Área guardada.");
  }
  else if(k === "asignar" || k === "cerrar" || k === "reabrir") iniciaAccion(k);
  else if(k === "cancelar"){ S.accion = null; repinta(); }
  else if(k === "guardar") guardarAccion();
  else if(k === "evfile"){ leeAccion(); const f = document.getElementById("hz-file2"); f.value = ""; f.click(); }
  else if(k === "nuevo") nuevoHallazgo();
  else if(k === "volver"){
    if(S.vista === "nuevo" && S.F && (S.F.fotos.length || S.F.desc.trim()) && !confirm("Se perderá el reporte que empezó. ¿Salir?")) return;
    S.vista = "lista"; S.F = null; S.accion = null; repinta(); if(Date.now() - S.t > 5000) carga();
  }
  else if(k === "atras"){ TAB = "hoja"; showView(); }
  else if(k === "foto"){ leeCampos(); const f = document.getElementById("hz-file"); f.value = ""; f.click(); }
  else if(k === "enviar") enviar();
});
document.addEventListener("change", ev => {
  const t = ev.target; if(!t) return;
  if(t.id === "hz-file") agregaFoto(t.files && t.files[0]);
  else if(t.id === "hz-file2") agregaEvidencia(t.files && t.files[0]);
  else if(t.id === "hz-area" || t.id === "hz-tipo") leeCampos();
  else if(t.id === "hz-ej" || t.id === "hz-plazo") leeAccion();
  else if(t.dataset && t.dataset.hzAreTipo !== undefined){ const a = (S.areasAdm || []).find(x => x.nombre === t.dataset.hzAreTipo); if(a) guardaArea({nombre:a.nombre, tipo:t.value, ejecutora:a.ejecutora, activo:a.activo}, "Tipo actualizado."); }
  else if(t.dataset && t.dataset.hzAreEje !== undefined){ const a = (S.areasAdm || []).find(x => x.nombre === t.dataset.hzAreEje); if(a) guardaArea({nombre:a.nombre, tipo:a.tipo, ejecutora:t.checked, activo:a.activo}, t.checked ? "Marcada como ejecutora." : "Ya no es ejecutora."); }
  else if(t.id === "u-rol") pintaUsrAreasNuevo();
  else if(t.dataset && t.dataset.uar !== undefined){      // áreas de un Responsable existente (cuando el líder rota)
    const u = t.dataset.uar, L = [...document.querySelectorAll('[data-uar]')].filter(i => i.dataset.uar === u && i.checked).map(i => i.value);
    usrOp({action:"usuario_editar", usuario:u, areas:L}, "Áreas actualizadas.");
  }
});
document.addEventListener("input", ev => {
  const id = ev.target && ev.target.id;
  if(id === "hz-desc") leeCampos(); else if(id === "hz-com" || id === "hz-mot") leeAccion();
});

window.drawAreas = drawAreas; window.hzCargaAreasUsr = hzCargaAreasUsr; window.hzUsrAreas = hzUsrAreas; window.hzAreasMarcadas = hzAreasMarcadas;
window.drawHall = drawHall; window.nuevoHallazgo = nuevoHallazgo; window.abrirHallazgo = abrirHallazgo;
window.HZUI = {S, carga, repinta, cargaAreas, ERR, box, fecha, badge};   // lo usa la parte de asignar/cerrar/áreas (misma pestaña)
// si el Responsable de área ya inició sesión antes de que se cargara este archivo, se muestra Hallazgos
if(typeof RESP === "function" && RESP() && typeof showView === "function") showView();
})();
