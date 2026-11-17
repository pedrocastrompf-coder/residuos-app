/* Hallazgos ambientales · fotos en el celular: comprimir y marcar la observación.
   Sin librerías. Se carga como script normal (funciones globales); en Node solo se prueba dimensionesMax. */

// reduce proporcionalmente el lado mayor a `max`; no agranda y nunca devuelve 0
function dimensionesMax(w, h, max){
  max = max || 1280;
  const k = Math.min(1, max / Math.max(w, h));
  return {w: Math.max(1, Math.round(w * k)), h: Math.max(1, Math.round(h * k))};
}

// bytes reales de un data-URL en base64
function bytesDataUrl(s){
  const i = String(s).indexOf(","), b64 = String(s).slice(i + 1), pad = (b64.match(/=+$/) || [""])[0].length;
  return Math.floor(b64.length * 3 / 4) - pad;
}

// exporta el canvas a JPEG bajando la calidad hasta quedar bajo `tope`; si aun así pesa, reduce el tamaño y repite
function jpegBajo_(canvas, tope, q0){
  let c = canvas;
  for (let intento = 0; intento < 4; intento++){
    for (let q = q0; q >= 0.4 - 1e-9; q -= 0.1){
      const d = c.toDataURL("image/jpeg", q), n = bytesDataUrl(d);
      if (n <= tope) return {dataUrl: d, bytes: n};
    }
    const r = document.createElement("canvas");
    r.width = Math.max(1, Math.round(c.width * 0.8)); r.height = Math.max(1, Math.round(c.height * 0.8));
    r.getContext("2d").drawImage(c, 0, 0, r.width, r.height);
    c = r;
  }
  const d = c.toDataURL("image/jpeg", 0.4);
  return {dataUrl: d, bytes: bytesDataUrl(d)};
}

// foto del celular (12 MP, con orientación EXIF) → JPEG derecho de lado mayor `max` y ≤ 250 KB
async function comprimirFoto(file, max){
  max = max || 1280;
  let bmp;
  try{ bmp = await createImageBitmap(file, {imageOrientation: "from-image"}); }
  catch(e){ bmp = await createImageBitmap(file); }   // navegadores que no aceptan la opción ya giran la imagen solos
  const d = dimensionesMax(bmp.width, bmp.height, max), c = document.createElement("canvas");
  c.width = d.w; c.height = d.h;
  c.getContext("2d").drawImage(bmp, 0, 0, d.w, d.h);
  if (bmp.close) bmp.close();
  return jpegBajo_(c, 250 * 1024, 0.7);
}

const ROJO_MARCA = "#e11d2e";
function dibujarMarca_(g, m, W){
  const ancho = Math.max(4, W * 0.008);
  g.strokeStyle = ROJO_MARCA; g.fillStyle = ROJO_MARCA; g.lineWidth = ancho; g.lineCap = "round"; g.lineJoin = "round";
  if (m.t === "o"){ g.beginPath(); g.arc(m.x, m.y, W * 0.07, 0, Math.PI * 2); g.stroke(); }
  else if (m.t === "f"){
    const a = Math.atan2(m.y2 - m.y, m.x2 - m.x), L = Math.max(ancho * 5, W * 0.05);
    g.beginPath(); g.moveTo(m.x, m.y); g.lineTo(m.x2, m.y2); g.stroke();
    g.beginPath(); g.moveTo(m.x2, m.y2);
    g.lineTo(m.x2 - L * Math.cos(a - 0.45), m.y2 - L * Math.sin(a - 0.45));
    g.lineTo(m.x2 - L * Math.cos(a + 0.45), m.y2 - L * Math.sin(a + 0.45));
    g.closePath(); g.fill();
  } else if (m.t === "t" && m.p.length){
    g.beginPath(); g.moveTo(m.p[0][0], m.p[0][1]);
    m.p.slice(1).forEach(q => g.lineTo(q[0], q[1]));
    if (m.p.length === 1) g.lineTo(m.p[0][0] + 0.1, m.p[0][1]);
    g.stroke();
  }
}

/* Abre la foto a pantalla completa para marcar la observación (círculo, flecha o trazo, en rojo, con Deshacer).
   Devuelve el dataUrl JPEG con la marca integrada, o null si se cancela.
   opts.obligatoria (por omisión true): «Listo» solo se habilita con al menos una marca. */
function marcarFoto(dataUrl, opts){
  const obligatoria = !(opts && opts.obligatoria === false);
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onerror = () => reject(new Error("foto_invalida"));
    img.onload = () => {
      const W = img.naturalWidth, H = img.naturalHeight, marcas = [];
      let herr = "o", actual = null;
      const ov = document.createElement("div");
      ov.style.cssText = "position:fixed;inset:0;z-index:99999;background:#111;display:flex;flex-direction:column;font-family:system-ui,sans-serif;color:#fff";
      const btn = (txt, id) => '<button type="button" data-mk="' + id + '" style="flex:1;min-height:48px;border:0;border-radius:10px;background:#333;color:#fff;font-size:15px">' + txt + '</button>';
      ov.innerHTML =
        '<div data-h style="padding:10px 12px;font-size:14px;text-align:center">' + (obligatoria ? "Marca la observación en la foto" : "Marca la foto (opcional)") + '</div>' +
        '<div style="flex:1;min-height:0;display:flex;align-items:center;justify-content:center;padding:0 6px"><canvas style="max-width:100%;max-height:100%;touch-action:none;background:#000"></canvas></div>' +
        '<div style="display:flex;gap:8px;padding:8px 10px 4px">' + btn("◯ Círculo", "o") + btn("➚ Flecha", "f") + btn("✎ Trazo", "t") + btn("↶ Deshacer", "u") + '</div>' +
        '<div style="display:flex;gap:8px;padding:4px 10px 14px">' + btn("Cancelar", "x") + btn("Listo", "ok") + '</div>';
      document.body.appendChild(ov);
      const cv = ov.querySelector("canvas"), g = cv.getContext("2d");
      cv.width = W; cv.height = H;
      const B = {}; ov.querySelectorAll("button").forEach(b => B[b.dataset.mk] = b);
      function pintar(){
        g.drawImage(img, 0, 0, W, H);
        marcas.forEach(m => dibujarMarca_(g, m, W));
        if (actual) dibujarMarca_(g, actual, W);
        B.ok.disabled = obligatoria && !marcas.length; B.ok.style.opacity = B.ok.disabled ? 0.4 : 1; B.ok.style.background = B.ok.disabled ? "#333" : ROJO_MARCA;
        B.u.disabled = !marcas.length; B.u.style.opacity = B.u.disabled ? 0.4 : 1;
        ["o", "f", "t"].forEach(k => B[k].style.background = k === herr ? ROJO_MARCA : "#333");
      }
      const punto = ev => { const r = cv.getBoundingClientRect(); return [(ev.clientX - r.left) * W / r.width, (ev.clientY - r.top) * H / r.height]; };
      cv.addEventListener("pointerdown", ev => {
        ev.preventDefault(); try{ cv.setPointerCapture(ev.pointerId); }catch(e){}
        const p = punto(ev);
        actual = herr === "o" ? {t:"o", x:p[0], y:p[1]} : herr === "f" ? {t:"f", x:p[0], y:p[1], x2:p[0], y2:p[1]} : {t:"t", p:[p]};
        pintar();
      });
      cv.addEventListener("pointermove", ev => {
        if (!actual) return; ev.preventDefault(); const p = punto(ev);
        if (actual.t === "o"){ actual.x = p[0]; actual.y = p[1]; }      // el círculo se puede ajustar antes de soltar
        else if (actual.t === "f"){ actual.x2 = p[0]; actual.y2 = p[1]; }
        else actual.p.push(p);
        pintar();
      });
      const soltar = () => { if (!actual) return; if (!(actual.t === "f" && Math.hypot(actual.x2 - actual.x, actual.y2 - actual.y) < W * 0.02)) marcas.push(actual); actual = null; pintar(); };
      cv.addEventListener("pointerup", soltar); cv.addEventListener("pointercancel", () => { actual = null; pintar(); });
      ["o", "f", "t"].forEach(k => B[k].onclick = () => { herr = k; pintar(); });
      B.u.onclick = () => { marcas.pop(); pintar(); };
      const cerrar = v => { ov.remove(); resolve(v); };
      B.x.onclick = () => cerrar(null);
      B.ok.onclick = () => { if (B.ok.disabled) return; cerrar(jpegBajo_(cv, 300 * 1024, 0.85).dataUrl); };
      pintar();
    };
    img.src = dataUrl;
  });
}

if (typeof module !== "undefined") module.exports = {dimensionesMax, bytesDataUrl};
