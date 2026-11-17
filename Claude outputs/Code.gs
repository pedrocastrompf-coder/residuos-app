/**
 * Pesaje de residuos sólidos  |  Backend en Google Apps Script
 * Pega este archivo en Extensiones > Apps Script de TU hoja de cálculo y publica una NUEVA VERSIÓN
 * de la implementación (Implementar > Administrar implementaciones > Editar > Versión nueva).
 *
 * Pestañas que usa (las crea solas):
 *   Hojas, Pesadas      -> lo que registran los operarios (como antes)
 *   Despachos           -> datos del despacho (empresa, conductor, manejo...) y su estado Pendiente/Aprobado
 *   Registro Interno    -> una fila por despacho APROBADO (mismas columnas que la Tabla3 del Excel)
 *   Data                -> totales mensuales por residuo (mismas columnas que la hoja "Data" del Excel; alimenta el dashboard)
 *   Usuarios            -> personas con acceso (usuario, rol, contraseña CIFRADA). Se administran desde la app (solo Administrador).
 *   Catálogo            -> la lista de residuos que ve la app. El administrador agrega, desactiva o corrige residuos AQUÍ
 *                          (la app, Registro Interno y Data se actualizan solos).
 */
const CLAVE = "MEDIOAMBIENTE";                 // llave de la app: debe ser IGUAL a CLAVE en config.js (el acceso de cada persona es su usuario y contraseña)
const SEDE_BASE = "Callao 1";      // planta del Sheet principal (donde vive este script, los Usuarios y la lista de Plantas)
let SEDE = SEDE_BASE;              // planta de la petición en curso
let SS_ = null;                    // Sheet de la planta en curso (null = el principal)
function ss_(){ return SS_ || SpreadsheetApp.getActive(); }
function central_(fn){ const k = SS_; SS_ = null; try{ return fn(); } finally { SS_ = k; } }
const ANO_DATA = 2026;

const COLS_H = ["id","fecha","puerta","placa","responsable","estado","nPesadas","kgNeto","actualizado","sede","viaje","usuario"];
const COLS_P = ["hojaId","fecha","puerta","placa","codSAP","categoria","residuo","n","bruto","tara","taraPendiente","neto","cantidad","unidad","sede"];
const COLS_D = ["clave","sede","fecha","placa","empresa","conductor","brevete","hora","manejoNP","manejoP","obs","estado","aprobadoPor","aprobadoEn","actualizado","viaje","precintos","numero"];
const TXT_H = ["id","fecha","puerta","placa"];
const TXT_P = ["hojaId","fecha","puerta","placa","codSAP"];
const TXT_D = ["clave","fecha","placa","hora","brevete","precintos"];

const RI = "Registro Interno", DT = "Data";
const RI_FIJAS = ["Sede","Fecha de registro","Responsable de registro","Tipo de residuo","Tipo de Manejo","Empresa Operadora","Placa de vehículo"];
const RI_RESIDUOS = ["Residuo chocolate","Residuo chocolate FSTB","Residuo wafer","Residuo galleta","Residuo galleta de chocolate","Residuo galleta oreo multipack","Residuo galleta oreo crumbs","Residuo galleta oreo tripack original","Residuo galleta oreo tripack chocolate","Residuo galleta oreo bañado dark","Residuo galleta oreo bañado white","Residuo masa fermentada","Residuo orgánico de comedor","Residuo polietileno (stretch film y manga plast)","Residuo polietileno sucio","Residuo saco de polipropileno (rafia)","Residuo plástico duro","Residuo galonera plástica 25kg","Residuo bidón azul 280 litros","Residuo balde plástico","Residuo molde de plástico","Residuo parihuela en desuso","Residuo parihuela leña","Residuo cartón","Residuo bolsa de papel","Residuo papel bond","Residuo chatarra","Residuo cables de cobre","Residuo cilindro de metal","Residuo acero","Residuo aceite usado","Residuo saco de yute","Residuo bobina en rollo","Residuo varios","Residuo bobina inutilizada","Residuo semisólido (lodo)","Residuo orgánico de comedor*","Residuo de aparato eléctrico y electrónico (RAEE)","Neumático fuera de uso (NFU)","Residuo de construcción y demolición (RCD)","Residuo de laboratorio (autoclavado)","Residuo envase vacío de tinta (citronix)","Residuo biocontaminado","Residuo saco vacío de bicarbonato de amonio/bicarbonato de sodio","Residuo saco vacío de carbonato de potasio/soda cáustica","Residuos varios contaminados (trapos, waypes, otros)","Residuo envase vacío de producto químico (plástico)","Residuo envase vacío de producto químico (vidrio, metal)","Residuo aceite contaminado","Residuo batería usada","Residuo balón vacío de refrigerante","Residuo de asbesto","Residuo envase con producto químico","Residuo luminaria rota","Residuo fibra de vidrio","Residuo fosfuro de aluminio","Otro: Agua contaminada con hidrocarburo","Otro: Acelerador de concreto","Residuo bolitas crispy de arroz","Residuo insumo vencido"];
const RI_ORIGEN = "Origen";
const RI_NUM = "N.° de registro";   // código del despacho (C1-000123): para filtrar en el Registro Interno y en el Excel
const RI_COLS = RI_FIJAS.concat(RI_RESIDUOS, [RI_ORIGEN]);
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Setiembre","Octubre","Noviembre","Diciembre"];
const DATA_COLS = ["Sede","Clasificación según su peligrosidad","Clasificación según su aprovechamiento","Clasificación según su composición","Descripción","Tipo de manejo","Año"].concat(MESES, ["Columna en Registro Interno"]);
// Filas de Data cuya descripción tiene otro nombre en el Registro Interno (igual que tu Excel)
const DATA_ALIAS = {"Residuo común":"Residuo varios"};
// Celdas que en tu Excel están en 0 fijo, no con fórmula: [índice de fila en DATA_FILAS, mes 1-12]
const DATA_CEROS = [[5,1],[5,2],[5,3],[5,4],[5,5],[5,6],[5,7],[5,8],[5,9],[5,10],[5,11],[5,12],[15,1],[15,2],[15,3],[15,4],[15,5],[15,6],[15,7],[15,8],[24,1],[24,2],[24,3],[24,4],[24,5],[24,6],[24,7],[24,8],[24,9],[24,10],[24,11],[24,12],[25,1],[25,2],[25,3],[25,4],[25,5],[25,6],[25,7],[25,8],[25,9],[25,10],[25,11],[25,12],[45,5],[46,5],[47,5],[48,5],[49,5],[50,5],[51,5],[70,1],[70,2],[70,3],[70,4],[70,5],[70,6],[70,7],[70,8],[70,9],[70,10],[70,11],[70,12],[71,1],[71,2],[71,3],[71,4],[71,5],[71,6],[71,7],[71,8],[71,9],[71,10],[71,11],[71,12]];
// ---------- Catálogo de residuos (pestaña "Catálogo")
const CAT_HOJA = "Catálogo";
const CAT_COLS = ["Residuo","Categoría","Código SAP","Manejo","Activo","Estado"];
const MANEJOS_DATA = ["Valorización","Tratamiento","Disposición final"];
// categoría del catálogo -> [peligrosidad, aprovechamiento, composición] tal como van en la hoja Data
const CATEGORIAS = {
  "Orgánico - producción":["No Peligroso","Aprovechable","Orgánico"], "Orgánico - comedor":["No Peligroso","Aprovechable","Orgánico"],
  "Plástico":["No Peligroso","Aprovechable","Plástico"], "Madera":["No Peligroso","Aprovechable","Madera"],
  "Papel y cartón":["No Peligroso","Aprovechable","Papel y cartón"], "Metal":["No Peligroso","Aprovechable","Metal"],
  "Vidrio":["No Peligroso","Aprovechable","Vidrio"], "Otros aprovechables":["No Peligroso","Aprovechable","Otro aprovechable"],
  "No aprovechable":["No Peligroso","No aprovechable","No aprovechable"], "Especial":["No Peligroso","Aprovechable","Especial"],
  "Peligrosos":["Peligroso","Peligroso","Peligroso"]
};
const RI_ANCHO = 200;   // columnas reservadas en Registro Interno (hasta la GR): las fórmulas de Data leen H:GR
// Lista inicial (se copia a la pestaña Catálogo la primera vez): [residuo, categoría, código SAP, manejo, activo]
const CATALOGO_INICIAL = [["Residuo chocolate", "Orgánico - producción", "6001000226", "Valorización", "Sí"], ["Residuo chocolate FSTB", "Orgánico - producción", "6001000297", "Valorización", "Sí"], ["Residuo wafer", "Orgánico - producción", "6001000227", "Valorización", "Sí"], ["Residuo galleta", "Orgánico - producción", "6001000247", "Valorización", "Sí"], ["Residuo galleta de chocolate", "Orgánico - producción", "6001000260", "Valorización", "Sí"], ["Residuo galleta oreo multipack", "Orgánico - producción", "6001000271", "Valorización", "Sí"], ["Residuo galleta oreo crumbs", "Orgánico - producción", "6001000273", "Valorización", "Sí"], ["Residuo galleta oreo tripack original", "Orgánico - producción", "6001000277", "Valorización", "Sí"], ["Residuo galleta oreo tripack chocolate", "Orgánico - producción", "6001000276", "Valorización", "Sí"], ["Residuo galleta oreo bañado dark", "Orgánico - producción", "6001000296", "Valorización", "Sí"], ["Residuo galleta oreo bañado white", "Orgánico - producción", "6001000295", "Valorización", "Sí"], ["Residuo masa fermentada", "Orgánico - producción", "6001000318", "Valorización", "Sí"], ["Residuo orgánico de comedor", "Orgánico - comedor", "", "Valorización", "Sí"], ["Residuo polietileno (stretch film y manga plast)", "Plástico", "6001000238", "Valorización", "Sí"], ["Residuo polietileno sucio", "Plástico", "6001000258", "Valorización", "Sí"], ["Residuo saco de polipropileno (rafia)", "Plástico", "6001000207", "Valorización", "Sí"], ["Residuo plástico duro", "Plástico", "6001000229", "Valorización", "Sí"], ["Residuo galonera plástica 25kg", "Plástico", "6001000214", "Valorización", "Sí"], ["Residuo bidón azul 280 litros", "Plástico", "6001000249", "Valorización", "Sí"], ["Residuo balde plástico", "Plástico", "6001000212", "Valorización", "Sí"], ["Residuo molde de plástico", "Plástico", "6001000335", "Valorización", "Sí"], ["Residuo parihuela en desuso", "Madera", "6001000208", "Valorización", "Sí"], ["Residuo parihuela leña", "Madera", "6001000210", "Valorización", "Sí"], ["Residuo cartón", "Papel y cartón", "6001000231", "Valorización", "Sí"], ["Residuo bolsa de papel", "Papel y cartón", "6001000232", "Valorización", "Sí"], ["Residuo papel bond", "Papel y cartón", "6001000233", "Valorización", "Sí"], ["Residuo chatarra", "Metal", "5003000032", "Valorización", "Sí"], ["Residuo cables de cobre", "Metal", "6001000236", "Valorización", "Sí"], ["Residuo cilindro de metal", "Metal", "6001000213", "Valorización", "Sí"], ["Residuo acero", "Metal", "6001000259", "Valorización", "Sí"], ["Residuo aceite usado", "Otros aprovechables", "6001000240", "Valorización", "Sí"], ["Residuo saco de yute", "Otros aprovechables", "6001000237", "Valorización", "Sí"], ["Residuo bobina en rollo", "Otros aprovechables", "6001000305", "Valorización", "Sí"], ["Residuo varios", "No aprovechable", "", "Disposición final", "Sí"], ["Residuo bobina inutilizada", "No aprovechable", "", "Disposición final", "Sí"], ["Residuo semisólido (lodo)", "No aprovechable", "", "Disposición final", "Sí"], ["Residuo de aparato eléctrico y electrónico (RAEE)", "Especial", "", "Valorización", "Sí"], ["Neumático fuera de uso (NFU)", "Especial", "", "Valorización", "Sí"], ["Residuo de construcción y demolición (RCD)", "Especial", "", "Valorización", "Sí"], ["Residuo de laboratorio (autoclavado)", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuo envase vacío de tinta (citronix)", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuo biocontaminado", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuo saco vacío de bicarbonato de amonio/bicarbonato de sodio", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuo saco vacío de carbonato de potasio/soda cáustica", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuos varios contaminados (trapos, waypes, otros)", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuo envase vacío de producto químico (plástico)", "Peligrosos", "", "Tratamiento, Disposición final", "Sí"], ["Residuo envase vacío de producto químico (vidrio, metal)", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo aceite contaminado", "Peligrosos", "", "Valorización", "Sí"], ["Residuo batería usada", "Peligrosos", "", "Valorización", "Sí"], ["Residuo balón vacío de refrigerante", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo de asbesto", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo envase con producto químico", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo luminaria rota", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo fibra de vidrio", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo fosfuro de aluminio", "Peligrosos", "", "Disposición final", "Sí"], ["Residuo orgánico de comedor*", "No aprovechable", "", "Disposición final", "No"], ["Otro: Agua contaminada con hidrocarburo", "Peligrosos", "", "Disposición final", "No"], ["Otro: Acelerador de concreto", "Peligrosos", "", "Disposición final", "No"], ["Residuo bolitas crispy de arroz", "Orgánico - producción", "", "Valorización", "No"], ["Residuo insumo vencido", "Orgánico - producción", "", "Valorización", "No"]];
// [sede, peligrosidad, aprovechamiento, composición, descripción, tipo de manejo, año]
const DATA_FILAS = [["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo chocolate","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo chocolate FSTB","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo wafer","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta de chocolate","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo orgánico con bobina","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta oreo multipack","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta oreo crumbs","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta oreo tripack original","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta oreo tripack chocolate","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta oreo bañado dark","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo galleta oreo bañado white","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo masa fermentada","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo bolitas crispy de arroz","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo insumo vencido","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Orgánico","Residuo orgánico de comedor","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo polietileno (stretch film y manga plast)","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo polietileno sucio","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo saco de polipropileno (rafia)","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo plástico duro","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo galonera plástica 25kg","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo bidón azul 280 litros","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo balde plástico","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo molde de plástico","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo galonera plástica","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Plástico","Residuo galonera plástica 50kg","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Madera","Residuo parihuela en desuso","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Madera","Residuo parihuela leña","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Papel y cartón","Residuo cartón","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Papel y cartón","Residuo bolsa de papel","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Papel y cartón","Residuo papel bond","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Metal","Residuo chatarra","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Metal","Residuo cables de cobre","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Metal","Residuo cilindro de metal","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Metal","Residuo acero","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Otro aprovechable","Residuo aceite usado","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Otro aprovechable","Residuo saco de yute","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Otro aprovechable","Residuo bobina en rollo","Valorización",2026],["Callao 1","No Peligroso","No aprovechable","No aprovechable","Residuo común","Disposición final",2026],["Callao 1","No Peligroso","No aprovechable","No aprovechable","Residuo bobina inutilizada","Disposición final",2026],["Callao 1","No Peligroso","No aprovechable","No aprovechable","Residuo semisólido (lodo)","Disposición final",2026],["Callao 1","No Peligroso","No aprovechable","No aprovechable","Residuo orgánico de comedor*","Disposición final",2026],["Callao 1","No Peligroso","Aprovechable","Especial","Residuo de aparato eléctrico y electrónico (RAEE)","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Especial","Neumático fuera de uso (NFU)","Valorización",2026],["Callao 1","No Peligroso","Aprovechable","Especial","Residuo de construcción y demolición (RCD)","Valorización",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo de laboratorio (autoclavado)","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo envase vacío de tinta (citronix)","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo biocontaminado","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo saco vacío de bicarbonato de amonio/bicarbonato de sodio","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo saco vacío de carbonato de potasio/soda cáustica","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuos varios contaminados (trapos, waypes, otros)","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo envase vacío de producto químico (plástico)","Tratamiento",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo aceite contaminado","Valorización",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo batería usada","Valorización",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo envase vacío de producto químico (vidrio, metal)","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo balón vacío de refrigerante","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo de asbesto","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo envase con producto químico","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo luminaria rota","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo fibra de vidrio","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo fosfuro de aluminio","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo de laboratorio (autoclavado)","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo envase vacío de tinta (citronix)","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo biocontaminado","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo saco vacío de bicarbonato de amonio/bicarbonato de sodio","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo saco vacío de carbonato de potasio/soda cáustica","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuos varios contaminados (trapos, waypes, otros)","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Residuo envase vacío de producto químico (plástico)","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Otro: Agua contaminada con hidrocarburo","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Otro: Acelerador de concreto","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Otro: Residuo lodo de PTARI","Disposición final",2026],["Callao 1","Peligroso","Peligroso","Peligroso","Otro: Residuo desengrase electrolítico JE610","Disposición final",2026]];

function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function norm_(v){ return (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(v==null?"":v); }
// la hora de llegada: Sheets puede convertirla en valor de hora/fecha; siempre se devuelve como "HH:mm"
function horaTxt_(v){
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), "HH:mm");
  const m = String(v==null?"":v).match(/^\s*(\d{1,2}):(\d{2})/);
  return m ? ("0"+m[1]).slice(-2)+":"+m[2] : "";
}
function placa_(s){ return String(s||"").toUpperCase().replace(/\s/g,""); }
function r2_(x){ return Math.round(Number(x||0)*100)/100; }

// ---------- pestañas
function hoja_(name, cols, txtCols){
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight("bold").setBackground("#EFE1D4");
    sh.setFrozenRows(1);
    (txtCols||[]).forEach(c => sh.getRange(1, cols.indexOf(c)+1, sh.getMaxRows(), 1).setNumberFormat("@"));
  } else {
    // pestaña creada con una versión anterior: agrega al final las columnas nuevas
    const have = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String).filter(x => x!=="");
    if (have.length < cols.length && cols.slice(0,have.length).join("|") === have.join("|"))
      sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight("bold").setBackground("#EFE1D4");
  }
  return sh;
}

function riHoja_(){
  const ss = ss_();
  let sh = ss.getSheetByName(RI);
  if (!sh){
    sh = ss.insertSheet(RI);
    sh.getRange(1,1,1,RI_COLS.length).setValues([RI_COLS]).setFontWeight("bold").setBackground("#EFE1D4");
    sh.setFrozenRows(1);
    sh.getRange(2,2,sh.getMaxRows()-1,1).setNumberFormat("dd/mm/yyyy");
  }
  ampliar_(sh, RI_ANCHO);
  return sh;
}
function ampliar_(sh, n){ const m = sh.getMaxColumns(); if (m < n) sh.insertColumnsAfter(m, n-m); }
// encabezado real de Registro Interno (las columnas de residuos nuevos se agregan al final)
function riEncabezado_(sh){
  const n = Math.max(sh.getLastColumn(),1);
  const h = sh.getRange(1,1,1,n).getValues()[0].map(String);
  while (h.length && h[h.length-1]==="") h.pop();
  return h;
}
function asegurarColumnasRI_(nombres){
  const sh = riHoja_(), h = riEncabezado_(sh);
  const nuevos = nombres.filter((n,i) => n && h.indexOf(n)<0 && nombres.indexOf(n)===i);
  if (nuevos.length) sh.getRange(1,h.length+1,1,nuevos.length).setValues([nuevos]).setFontWeight("bold").setBackground("#EFE1D4");
}
function riContexto_(){
  let sh = riHoja_(), h = riEncabezado_(sh);
  if (h.indexOf(RI_ORIGEN)<0){ asegurarColumnasRI_([RI_ORIGEN]); h = riEncabezado_(sh); }
  if (h.indexOf(RI_NUM)<0){ asegurarColumnasRI_([RI_NUM]); h = riEncabezado_(sh); }
  const idx = {}; h.forEach((n,i) => { if (i>=7 && n!==RI_ORIGEN && n!==RI_NUM) idx[n] = i; });
  return {sh:sh, hdr:h, ncol:h.length, iOri:h.indexOf(RI_ORIGEN), iNum:h.indexOf(RI_NUM), idx:idx};
}

function dataFormulas_(r){
  const f = [];
  for (let m=1; m<=12; m++){
    // SUBSTITUTE escapa el "*" de "Residuo orgánico de comedor*" (en MATCH el * es comodín); IFERROR: residuos sin columna dan 0
    f.push("=IFERROR(SUMIFS(INDEX('"+RI+"'!$H$2:$GR,0,MATCH(SUBSTITUTE($T"+r+",\"*\",\"~*\"),'"+RI+"'!$H$1:$GR$1,0)),'"+RI+"'!$A$2:$A,$A"+r+
           ",'"+RI+"'!$B$2:$B,\">=\"&DATE($G"+r+","+m+",1),'"+RI+"'!$B$2:$B,\"<=\"&EOMONTH(DATE($G"+r+","+m+",1),0)),0)");
  }
  return f;
}

function dataHoja_(){
  const ss = ss_();
  let sh = ss.getSheetByName(DT);
  if (!sh){
    sh = ss.insertSheet(DT);
    sh.getRange(1,1,1,DATA_COLS.length).setValues([DATA_COLS]).setFontWeight("bold").setBackground("#EFE1D4");
    sh.setFrozenRows(1);
    const ceros = {}; if (SEDE === SEDE_BASE) DATA_CEROS.forEach(c => ceros[c[0]+"|"+c[1]] = 1);
    const base = DATA_FILAS.map((d,i) => {
      const fila = d.slice(0,7).slice(), fm = dataFormulas_(i+2); fila[0] = SEDE;
      for (let m=1; m<=12; m++) fila.push(ceros[i+"|"+m] ? 0 : fm[m-1]);
      fila.push(DATA_ALIAS[d[4]] || d[4]);
      return fila;
    });
    const extras = filasNuevasData_({}).map((d,k) => d.concat(dataFormulas_(DATA_FILAS.length+2+k), [d[4]]));
    const f = base.concat(extras);
    sh.getRange(2,1,f.length,DATA_COLS.length).setValues(f);   // setValues: el texto queda como texto y lo que empieza con "=" como fórmula
    sh.getRange(2,8,f.length,12).setNumberFormat("#,##0.00");
  }
  return sh;
}

// Si la pestaña Data quedó mal, ejecútala desde el editor: la borra y la vuelve a crear
function rehacerData(){
  const ss = ss_(), sh = ss.getSheetByName(DT);
  if (sh) ss.deleteSheet(sh);
  dataHoja_();
}

// Ejecútala UNA vez desde el editor (o se ejecuta sola con el primer envío de la app)
// la estructura se verifica como mucho una vez cada 6 horas (antes: en cada guardado)
function estructura_(){
  const C = CacheService.getScriptCache(), k = "est_ok" + (SS_ ? "_"+SS_.getId() : "");
  if (C.get(k) === "1") return;
  crearEstructura(); C.put(k, "1", 21600);
}
function crearEstructura(){
  hoja_("Hojas", COLS_H, TXT_H); hoja_("Pesadas", COLS_P, TXT_P); hoja_("Despachos", COLS_D, TXT_D);
  usuariosHoja_(); adminInicial_(); plantasHoja_(); hoja_("Bitácora", BIT_COLS, []);
  catalogoHoja_(); riHoja_(); dataHoja_(); cambiosHoja_(); formatoHoja_(); hzHoja_(); hzHojaAreas_();
}

// ---------- Usuarios, sesión y permisos por rol
const USU_COLS = ["usuario","nombre","rol","activo","hash","sal","cambiar","creado","actualizado","plantas","areas"];
const ROLES = ["Operador","Responsable de área","Inspector","Administrador"];
const SESION_DIAS = 14, HASH_VUELTAS = 400, MAX_FALLOS = 5, BLOQUEO_SEG = 600;
// quién puede hacer qué (se valida aquí, en el servidor; la app solo oculta los botones)
const PERMISOS = {
  aprobar:["Operador","Inspector","Administrador"], reabrir:["Inspector","Administrador"],
  manejo:["Operador","Inspector","Administrador"], borrarHoja:["Operador","Inspector","Administrador"],
  catalogo:["Inspector","Administrador"], plantas:["Inspector","Administrador"], usuarios:["Administrador"], actividad:["Administrador"],
  // hallazgos ambientales (el rol «Responsable de área» no pesa ni despacha: solo tiene estos permisos)
  hallazgo_reportar:["Operador","Responsable de área","Inspector","Administrador"], hallazgo_asignar:["Inspector","Administrador"],
  hallazgo_cerrar:["Responsable de área","Inspector","Administrador"], hallazgo_reabrir:["Inspector","Administrador"], areas_admin:["Inspector","Administrador"]
};
function puede_(rol, p){ return (PERMISOS[p]||[]).indexOf(rol) >= 0; }
function usuariosHoja_(){ return central_(() => hoja_("Usuarios", USU_COLS, ["usuario"])); }
function hex_(bytes){ return bytes.map(b => ("0"+((b+256)%256).toString(16)).slice(-2)).join(""); }
function hash_(clave, sal){
  let h = sal + "|" + clave;
  for (let i=0;i<HASH_VUELTAS;i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + sal));
  return h;
}
function azar_(n){
  const A = "abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let s = "";
  const u = (Utilities.getUuid()+Utilities.getUuid()).replace(/-/g,"");
  for (let i=0;i<n;i++) s += A[parseInt(u.substr(i*2,2),16) % A.length];
  return s;
}
function secreto_(){
  const P = PropertiesService.getScriptProperties();
  let s = P.getProperty("SECRETO");
  if (!s){ s = Utilities.getUuid()+Utilities.getUuid(); P.setProperty("SECRETO", s); }
  return s;
}
function firma_(texto){ return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, secreto_())); }
function usuariosTodos_(){
  const sh = usuariosHoja_(), V = sh.getDataRange().getValues().slice(1);
  return V.filter(r => String(r[0])!=="").map((r,i) => ({
    _fila:i+2, usuario:String(r[0]), nombre:String(r[1]), rol:String(r[2]), activo:String(r[3])!=="No",
    hash:String(r[4]), sal:String(r[5]), cambiar:String(r[6])==="Sí", creado:norm_(r[7]), actualizado:norm_(r[8]),
    plantas:plantasUsr_(r[9]), areas:areasUsr_(r[10])
  }));
}
function usuarioDe_(usuario){ const k = String(usuario||"").trim().toLowerCase(); return usuariosTodos_().filter(x => x.usuario===k)[0] || null; }
function escribirUsuario_(x){
  const sh = usuariosHoja_();
  const fila = [x.usuario, x.nombre, x.rol, x.activo?"Sí":"No", x.hash, x.sal, x.cambiar?"Sí":"No", x.creado||new Date(), new Date(), (x.plantas||[SEDE_BASE]).join(" | "), areasUsr_(x.areas).join(" | ")];
  if (x._fila) sh.getRange(x._fila,1,1,USU_COLS.length).setValues([fila]); else sh.appendRow(fila);
}
function claveValida_(c){ return typeof c==="string" && c.length>=4 && c.length<=64; }
// la primera vez (Usuarios vacía) se crea el Administrador con una contraseña temporal que aparece en el registro de ejecución
function adminInicial_(){
  if (usuariosTodos_().length) return;
  const tmp = azar_(10), sal = azar_(16);
  escribirUsuario_({usuario:"pcastro", nombre:"Pedro Castro", rol:"Administrador", activo:true, sal:sal, hash:hash_(tmp,sal), cambiar:true});
  Logger.log("ADMIN INICIAL -> usuario: pcastro | contraseña temporal: " + tmp + " (la app pedirá cambiarla en el primer ingreso)");
}
function tokenDe_(x){
  const exp = new Date().getTime() + SESION_DIAS*86400000;
  const cuerpo = Utilities.base64EncodeWebSafe(JSON.stringify({u:x.usuario, e:exp, k:x.hash.substr(0,10)}));
  return {token: cuerpo + "." + firma_(cuerpo), exp: exp};
}
function sesion_(token){
  try{
    const p = String(token||"").split(".");
    if (p.length!==2 || firma_(p[0])!==p[1]) return null;
    const d = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString());
    if (!d || d.e < new Date().getTime()) return null;
    const x = usuarioDe_(d.u);
    if (!x || !x.activo || x.hash.substr(0,10)!==d.k) return null;   // desactivado o contraseña cambiada: la sesión deja de valer
    return x;
  }catch(err){ return null; }
}
// reconfirma la clave de quien aprueba o reabre (comparte el contador de intentos fallidos con el ingreso)
function confirmarClave_(u, clave){
  const C = CacheService.getScriptCache(), ck = "f_"+u.usuario;
  if (parseInt(C.get(ck)||"0",10) >= MAX_FALLOS) return {ok:false, error:"bloqueado"};
  if (!clave) return {ok:false, error:"clave_requerida"};
  if (hash_(String(clave), u.sal) !== u.hash){ C.put(ck, String(parseInt(C.get(ck)||"0",10)+1), BLOQUEO_SEG); return {ok:false, error:"clave_incorrecta"}; }
  C.remove(ck); return {ok:true};
}
function pub_(x){ return {usuario:x.usuario, nombre:x.nombre, rol:x.rol, activo:x.activo, cambiar:x.cambiar, plantas:x.rol==="Administrador" ? ["*"] : x.plantas, areas:x.areas||[]}; }
function login_(b){
  const k = String(b.usuario||"").trim().toLowerCase(), C = CacheService.getScriptCache(), ck = "f_"+k;
  if (parseInt(C.get(ck)||"0",10) >= MAX_FALLOS) return {ok:false, error:"bloqueado"};
  const x = usuarioDe_(k);
  if (!x || !x.activo || hash_(String(b.clave||""), x.sal) !== x.hash){
    C.put(ck, String(parseInt(C.get(ck)||"0",10)+1), BLOQUEO_SEG);
    return {ok:false, error:"credenciales"};
  }
  C.remove(ck);
  const t = tokenDe_(x);
  return Object.assign({ok:true, token:t.token, exp:t.exp}, pub_(x));
}
function cambiarClave_(b, u){
  if (hash_(String(b.actual||""), u.sal) !== u.hash) return {ok:false, error:"clave_actual"};
  if (!claveValida_(b.nueva)) return {ok:false, error:"clave_corta"};
  if (b.nueva === b.actual) return {ok:false, error:"clave_igual"};
  u.sal = azar_(16); u.hash = hash_(b.nueva, u.sal); u.cambiar = false; escribirUsuario_(u);
  const t = tokenDe_(u);
  return Object.assign({ok:true, token:t.token, exp:t.exp}, pub_(u));
}
function usuariosListar_(){ return {ok:true, usuarios:usuariosTodos_().map(pub_)}; }
function usuarioCrear_(b){
  const k = String(b.usuario||"").trim().toLowerCase(), nombre = String(b.nombre||"").trim().replace(/\s+/g," ");
  if (!/^[a-z0-9._]{3,30}$/.test(k)) return {ok:false, error:"usuario_invalido"};
  if (nombre.split(" ").length<2) return {ok:false, error:"nombre_invalido"};
  if (ROLES.indexOf(b.rol)<0) return {ok:false, error:"rol_invalido"};
  if (!claveValida_(b.clave)) return {ok:false, error:"clave_corta"};
  if (usuarioDe_(k)) return {ok:false, error:"usuario_existe"};
  const sal = azar_(16);
  const pl = plantasValidas_(b.plantas); if (!pl) return {ok:false, error:"planta_invalida"};
  escribirUsuario_({usuario:k, nombre:nombre, rol:b.rol, activo:true, sal:sal, hash:hash_(b.clave,sal), cambiar:b.cambiar!==false, plantas:pl, areas:areasUsr_(b.areas)});
  return {ok:true};
}
function usuarioEditar_(b, yo){
  const x = usuarioDe_(b.usuario); if (!x) return {ok:false, error:"usuario_no_existe"};
  const rol = b.rol!==undefined ? b.rol : x.rol, activo = b.activo!==undefined ? !!b.activo : x.activo;
  const nombre = b.nombre!==undefined ? String(b.nombre).trim().replace(/\s+/g," ") : x.nombre;
  if (ROLES.indexOf(rol)<0) return {ok:false, error:"rol_invalido"};
  if (nombre.split(" ").length<2) return {ok:false, error:"nombre_invalido"};
  // para no quedarse sin acceso: nadie se quita su propio rol de Administrador ni se desactiva, y siempre queda un Administrador activo
  if (x.usuario===yo.usuario && (rol!==x.rol || !activo)) return {ok:false, error:"propio"};
  const admins = usuariosTodos_().filter(a => a.activo && a.rol==="Administrador" && a.usuario!==x.usuario).length;
  if (x.rol==="Administrador" && x.activo && (rol!=="Administrador" || !activo) && admins<1) return {ok:false, error:"ultimo_admin"};
  if (b.plantas !== undefined){ const pl = plantasValidas_(b.plantas); if (!pl) return {ok:false, error:"planta_invalida"}; x.plantas = pl; }
  if (b.areas !== undefined) x.areas = areasUsr_(b.areas);
  x.rol = rol; x.activo = activo; x.nombre = nombre; escribirUsuario_(x);
  return {ok:true};
}
// borrar un perfil: no se borra uno mismo ni el último Administrador activo (lo registrado conserva su nombre)
function usuarioBorrar_(b, yo){
  const x = usuarioDe_(b.usuario); if (!x) return {ok:false, error:"usuario_no_existe"};
  if (x.usuario === yo.usuario) return {ok:false, error:"propio"};
  const admins = usuariosTodos_().filter(a => a.activo && a.rol==="Administrador" && a.usuario!==x.usuario).length;
  if (x.rol==="Administrador" && x.activo && admins<1) return {ok:false, error:"ultimo_admin"};
  usuariosHoja_().deleteRow(x._fila);
  return {ok:true};
}
function usuarioResetear_(b){
  const x = usuarioDe_(b.usuario); if (!x) return {ok:false, error:"usuario_no_existe"};
  if (!claveValida_(b.clave)) return {ok:false, error:"clave_corta"};
  x.sal = azar_(16); x.hash = hash_(b.clave, x.sal); x.cambiar = b.cambiar!==false; escribirUsuario_(x);   // cambia el hash: sus sesiones abiertas dejan de valer
  return {ok:true};
}

// ---------- plantas: cada una tiene su propio Google Sheet con la misma estructura (la lista vive en el Sheet principal)
const PLA_COLS = ["planta","sheetId","lugares","activo","creado","creadoPor"];
const LUGARES_BASE = ["Puerta 2","Puerta 3","Puerta 6"];
function plantasUsr_(v){ const L = String(v==null?"":v).split("|").map(x => x.trim()).filter(x => x); return L.length ? L : [SEDE_BASE]; }
function lugares_(v){
  const L = (Array.isArray(v) ? v : String(v==null?"":v).split("|")).map(x => String(x).trim().replace(/\s+/g," ")).filter(x => x);
  return L.filter((x,i) => L.indexOf(x)===i && x.length<=40).slice(0,20);
}
function plantasHoja_(){
  return central_(() => {
    const sh = hoja_("Plantas", PLA_COLS, ["sheetId"]);
    if (sh.getLastRow() < 2) sh.appendRow([SEDE_BASE, SpreadsheetApp.getActive().getId(), LUGARES_BASE.join(" | "), "Sí", new Date(), "sistema"]);
    return sh;
  });
}
function plantasTodas_(){
  const C = CacheService.getScriptCache(), c = C.get("plantas");
  if (c) try{ return JSON.parse(c); }catch(err){}
  const V = plantasHoja_().getDataRange().getValues().slice(1);
  const L = V.map((r,i) => ({_fila:i+2, planta:String(r[0]).trim(), sheetId:String(r[1]).trim(), lugares:lugares_(r[2]), activo:String(r[3])!=="No"})).filter(p => p.planta && p.sheetId);
  C.put("plantas", JSON.stringify(L), 300);
  return L;
}
function plantasOlvidar_(){ CacheService.getScriptCache().remove("plantas"); }
function plantaDe_(nombre){ const k = String(nombre||"").trim(); return plantasTodas_().filter(p => p.planta===k)[0] || null; }
// alcance: "*" = todas las plantas (también las que se creen después)
function puedePlanta_(u, planta){ const L = u.plantas||[SEDE_BASE]; return u.rol==="Administrador" || L.indexOf("*") >= 0 || L.indexOf(planta) >= 0; }
function plantasValidas_(v){
  if (v === undefined || v === null) return [SEDE];
  const L = (Array.isArray(v) ? v : String(v).split("|")).map(x => String(x).trim()).filter(x => x), T = plantasTodas_().map(p => p.planta);
  if (L.indexOf("*") >= 0) return ["*"];
  if (!L.length || L.some(x => T.indexOf(x) < 0)) return null;
  return L.filter((x,i) => L.indexOf(x)===i);
}
// deja lista la planta de la petición: valida el acceso y abre su Sheet
function usarPlanta_(sede, u){
  const p = plantaDe_(sede || SEDE_BASE);
  if (!p || !p.activo) return {ok:false, error:"planta"};
  if (u && !puedePlanta_(u, p.planta)) return {ok:false, error:"planta_no_asignada"};
  SEDE = p.planta;
  SS_ = p.sheetId === SpreadsheetApp.getActive().getId() ? null : SpreadsheetApp.openById(p.sheetId);
  return {ok:true};
}
function plantasListar_(u){
  const adm = puede_(u.rol, "plantas");
  return {ok:true, plantas:plantasTodas_().filter(p => (adm || p.activo) && puedePlanta_(u, p.planta))
    .map(p => Object.assign({planta:p.planta, lugares:p.lugares, activo:p.activo}, adm ? {url:"https://docs.google.com/spreadsheets/d/"+p.sheetId+"/edit"} : {}))};
}
function plantaCrear_(b, u){
  const nombre = String(b.planta||"").trim().replace(/\s+/g," ");
  if (!/^[A-Za-zÁÉÍÓÚÑáéíóúñ0-9 .-]{3,30}$/.test(nombre)) return {ok:false, error:"nombre_invalido"};
  if (plantasTodas_().some(p => norm2_(p.planta) === norm2_(nombre))) return {ok:false, error:"planta_existe"};
  const lug = lugares_(b.lugares); if (!lug.length) return {ok:false, error:"lugares_invalidos"};
  const base = SpreadsheetApp.getActive();
  const ss = SpreadsheetApp.create("Pesaje RRSS " + nombre);
  ss.setSpreadsheetTimeZone(base.getSpreadsheetTimeZone()); ss.setSpreadsheetLocale(base.getSpreadsheetLocale());
  plantasHoja_().appendRow([nombre, ss.getId(), lug.join(" | "), "Sí", new Date(), u.usuario]);
  plantasOlvidar_();
  const kS = SEDE, kSS = SS_;
  try{
    SEDE = nombre; SS_ = ss;
    hoja_("Hojas", COLS_H, TXT_H); hoja_("Pesadas", COLS_P, TXT_P); hoja_("Despachos", COLS_D, TXT_D); hoja_("Bitácora", BIT_COLS, []); hzHoja_(); hzHojaAreas_();
    catalogoHoja_(); riHoja_(); dataHoja_(); cambiosHoja_(); formatoHoja_(); sincronizarCatalogo_();
    const vacia = ss.getSheets().filter(x => ["Hoja 1","Sheet1","Hoja1"].indexOf(x.getName()) >= 0)[0]; if (vacia) ss.deleteSheet(vacia);
    cambio_(u, "Creación del archivo", "Planta "+nombre+" · lugares de despacho: "+lug.join(", "), false);
  } finally { SEDE = kS; SS_ = kSS; }
  // quien la crea queda con acceso a ella
  if (u.rol !== "Administrador" && u.plantas.indexOf("*") < 0 && u.plantas.indexOf(nombre) < 0){ u.plantas = u.plantas.concat([nombre]); escribirUsuario_(u); }
  return {ok:true, planta:nombre, url:ss.getUrl()};
}
function plantaEditar_(b, u){
  const p = plantaDe_(b.planta); if (!p) return {ok:false, error:"planta"};
  if (!puedePlanta_(u, p.planta)) return {ok:false, error:"planta_no_asignada"};
  const lug = b.lugares !== undefined ? lugares_(b.lugares) : p.lugares; if (!lug.length) return {ok:false, error:"lugares_invalidos"};
  const activo = b.activo !== undefined ? !!b.activo : p.activo;
  if (p.planta === SEDE_BASE && !activo) return {ok:false, error:"planta_principal"};
  plantasHoja_().getRange(p._fila, 3, 1, 2).setValues([[lug.join(" | "), activo ? "Sí" : "No"]]);
  plantasOlvidar_();
  return {ok:true};
}

// ---------- control de cambios del archivo de cada planta (pestaña propia en su Sheet)
const CAM_HOJA = "Control de cambios", CAM_COLS = ["versión","fecha","usuario","cambio","detalle"];
function cambiosHoja_(){
  const sh = hoja_(CAM_HOJA, CAM_COLS, []);
  if (sh.getLastRow() < 2) sh.appendRow([1, new Date(), "sistema", "Versión inicial", "Estructura base del archivo ("+SEDE+")"]);
  return sh;
}
function versionArchivo_(sh){ const n = sh.getLastRow(); return n < 2 ? 1 : (parseInt(sh.getRange(n,1).getValue(),10) || 1); }
// sube = el cambio modifica la estructura del archivo (p. ej. una columna nueva en Registro Interno): pasa a la versión siguiente
function cambio_(u, cambio, detalle, sube){
  try{ const sh = cambiosHoja_(); sh.appendRow([versionArchivo_(sh) + (sube ? 1 : 0), new Date(), u.usuario, cambio, String(detalle||"")]); }catch(err){}
}
// ---------- cuadro del formato SGA.R.G.001 de cada planta (código, versión, fecha de aprobación)
// Un cambio en el catálogo deja una revisión «Pendiente»; el Inspector decide si amerita nueva versión.
const FMT_HOJA = "Formato SGA", FMT_COLS = ["estado","código","versión","fechaAprobación","motivo","usuario","registrado"];
const FMT_BASE = {codigo:"SGA.R.G.001", version:"05", aprobacion:"06/04/2026"};
function formatoHoja_(){ return central_(formatoHojaAqui_); }
function formatoHojaAqui_(){
  const sh = hoja_(FMT_HOJA, FMT_COLS, ["código","versión","fechaAprobación"]);
  if (sh.getLastRow() < 2){
    sh.getRange(2,2,1,3).setNumberFormat("@");
    sh.getRange(2,1,1,FMT_COLS.length).setValues([["Vigente", FMT_BASE.codigo, FMT_BASE.version, FMT_BASE.aprobacion, "Versión vigente al iniciar el registro en la app", "sistema", new Date()]]);
  }
  return sh;
}
function fmtFecha_(v){ return (v instanceof Date) ? Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), "dd/MM/yyyy") : String(v==null?"":v); }
function formatoEstado_(){
  const sh = formatoHoja_(), V = sh.getDataRange().getValues().slice(1);
  const ver = v => { const t = String(v==null?"":v); return /^\d$/.test(t) ? "0"+t : t; };   // «05» guardado como número 5
  const fila = r => ({estado:String(r[0]), codigo:String(r[1]), version:ver(r[2]), aprobacion:fmtFecha_(r[3]), motivo:String(r[4]), usuario:String(r[5]), registrado:(r[6] instanceof Date)?r[6].getTime():String(r[6])});
  const L = V.filter(r => r[0] !== "").map(fila);
  const vig = L.filter(x => x.estado === "Vigente"), act = vig.length ? vig[vig.length-1] : Object.assign({estado:"Vigente"}, FMT_BASE);
  return {vigente:{codigo:act.codigo, version:act.version, aprobacion:act.aprobacion}, pendientes:L.filter(x => x.estado === "Pendiente").map(x => x.motivo),
    historial:L.filter(x => x.estado !== "Pendiente").slice(-30).reverse()};
}
function formatoPendiente_(u, motivo){ if (motivo) try{ formatoHoja_().appendRow(["Pendiente", "", "", "", motivo, u.usuario, new Date()]); }catch(err){} }
function siguienteVersion_(v){ const m = String(v).match(/^(\D*)(\d+)(.*)$/); if (!m) return String(v); const n = String(parseInt(m[2],10)+1); return m[1]+(n.length < m[2].length ? ("0".repeat(m[2].length-n.length))+n : n)+m[3]; }
// cierra las revisiones pendientes: «Aplicado» (nueva versión) o «Descartado» (no amerita)
function formatoCerrarPend_(sh, estado){
  const V = sh.getDataRange().getValues();
  for (let i=1;i<V.length;i++) if (String(V[i][0]) === "Pendiente") sh.getRange(i+1,1).setValue(estado);
}
function formatoGuardar_(b, u){
  const codigo = String(b.codigo||"").trim(), version = String(b.version||"").trim(), aprob = String(b.aprobacion||"").trim();
  if (!/^[0-9A-Za-z.\-]{3,30}$/.test(codigo)) return {ok:false, error:"codigo_invalido"};
  if (!/^[0-9A-Za-z.\-]{1,10}$/.test(version)) return {ok:false, error:"version_invalida"};
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(aprob)) return {ok:false, error:"fecha_invalida"};
  const est = formatoEstado_(), sh = formatoHoja_();
  const motivo = String(b.motivo||"").trim().slice(0,300) || (est.pendientes.length ? est.pendientes.join(" · ") : "Actualización del cuadro del formato");
  formatoCerrarPend_(sh, "Aplicado");
  const fl = sh.getLastRow()+1;
  sh.getRange(fl,2,1,3).setNumberFormat("@");
  sh.getRange(fl,1,1,FMT_COLS.length).setValues([["Vigente", codigo, version, aprob, motivo, u.usuario, new Date()]]);
  return {ok:true, formato:{codigo:codigo, version:version, aprobacion:aprob}};
}
function formatoDescartar_(b, u){
  const est = formatoEstado_(); if (!est.pendientes.length) return {ok:true};
  const sh = formatoHoja_(); formatoCerrarPend_(sh, "Descartado");
  sh.appendRow(["Sin cambio de versión", est.vigente.codigo, est.vigente.version, est.vigente.aprobacion, "No amerita nueva versión: "+est.pendientes.join(" · ")+(b.motivo ? " — "+String(b.motivo).slice(0,200) : ""), u.usuario, new Date()]);
  return {ok:true};
}
function cambiosListar_(n){
  const sh = cambiosHoja_(), V = sh.getDataRange().getValues().slice(1).filter(r => r[0] !== "");
  return {version:versionArchivo_(sh), cambios:V.slice(-n).reverse().map(r => ({version:r[0], fecha:(r[1] instanceof Date)?r[1].getTime():String(r[1]), usuario:String(r[2]), cambio:String(r[3]), detalle:String(r[4])}))};
}

// ---------- catálogo de residuos
function norm2_(s){ return String(s==null?"":s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim(); }
function manejosDe_(txt){
  const out = [];
  String(txt||"").split(/[,;+\/]/).forEach(p => {
    const k = norm2_(p); if (!k) return;
    MANEJOS_DATA.forEach(m => { if (norm2_(m)===k && out.indexOf(m)<0) out.push(m); });
  });
  return out;
}
function nombresBase_(){ const b = {}; DATA_FILAS.forEach(d => { b[DATA_ALIAS[d[4]]||d[4]] = 1; }); return b; }

// el catálogo y el cuadro del formato son corporativos: viven en el Sheet principal y valen para todas las plantas
function catalogoHoja_(){ return central_(catalogoHojaAqui_); }
function catalogoHojaAqui_(){
  const ss = ss_();
  let sh = ss.getSheetByName(CAT_HOJA);
  if (!sh){
    sh = ss.insertSheet(CAT_HOJA);
    sh.getRange(1,1,1,CAT_COLS.length).setValues([CAT_COLS]).setFontWeight("bold").setBackground("#EFE1D4");
    sh.setFrozenRows(1);
    sh.getRange(2,3,Math.max(sh.getMaxRows()-1,1),1).setNumberFormat("@");       // código SAP como texto
    sh.getRange(2,1,CATALOGO_INICIAL.length,5).setValues(CATALOGO_INICIAL);
    const lista = (vals, permitirOtro) => SpreadsheetApp.newDataValidation().requireValueInList(vals,true).setAllowInvalid(!!permitirOtro).build();
    const filas = Math.max(sh.getMaxRows()-1,1);
    sh.getRange(2,2,filas,1).setDataValidation(lista(Object.keys(CATEGORIAS)));
    sh.getRange(2,4,filas,1).setDataValidation(lista(MANEJOS_DATA, true));
    sh.getRange(2,5,filas,1).setDataValidation(lista(["Sí","No"]));
    sh.getRange(1,1).setNote("Nombre del residuo. No lo cambies si ya se usó: crea uno nuevo y desactiva el anterior.");
    sh.getRange(1,4).setNote("Valorización, Tratamiento o Disposición final (si tiene dos, sepáralos con coma). Solo se usa al CREAR un residuo nuevo; los de la lista original ya tienen su manejo fijo en Data.");
    sh.getRange(1,5).setNote("Sí = aparece en la app. No = se retira de la app pero se conserva su historial.");
    sh.getRange(1,6).setNote("Lo llena solo el sistema.");
    sh.getRange(1,8,6,1).setValues([["CÓMO USAR ESTE CATÁLOGO"],
      ["• Agregar un residuo: escribe su nombre en la primera fila libre y elige Categoría y Manejo; el Código SAP es opcional."],
      ["• Retirar un residuo de la app: cambia Activo a No (el historial se conserva)."],
      ["• No cambies el nombre de un residuo que ya se usó: crea uno nuevo y desactiva el anterior."],
      ["• Los tipos de residuo (colores de la NTP 900.058) son fijos: la Categoría decide en qué bloque aparece."],
      ["• La columna Estado dice \"Listo\" cuando el residuo ya funciona en la app, Registro Interno y Data (tarda hasta 1–2 minutos)."]]);
    sh.getRange(1,8).setFontWeight("bold");
    sh.setColumnWidth(1,330);
  }
  return sh;
}

function leerCatalogo_(){
  const sh = catalogoHoja_(), n = sh.getLastRow();
  if (n < 2) return [];
  const V = sh.getRange(2,1,n-1,CAT_COLS.length).getValues(), vistos = {}, out = [];
  V.forEach((r,i) => {
    const name = String(r[0]==null?"":r[0]).trim(); if (!name) return;
    const dup = !!vistos[name]; vistos[name] = 1;
    out.push({fila:i+2, name:name, cat:String(r[1]||"").trim(), cod:String(r[2]==null?"":r[2]).trim(), manejo:String(r[3]||""),
              activo:!/^(no|n|false|0)$/.test(norm2_(r[4])), dup:dup, estado:String(r[5]||"")});
  });
  return out;
}
function estadoCat_(c, base){
  if (c.dup) return "Nombre repetido: no se usa";
  if (!CATEGORIAS[c.cat]) return "Elige una categoría de la lista";
  if (!base[c.name] && !manejosDe_(c.manejo).length) return "Falta el manejo";
  return "Listo";
}
// filas de Data que faltan para los residuos nuevos (los de la lista original ya tienen sus filas); pares = {"residuo|manejo":1}
function filasNuevasData_(pares){
  const base = nombresBase_(), out = [];
  leerCatalogo_().forEach(c => {
    if (!c.activo || base[c.name] || estadoCat_(c, base)!=="Listo") return;
    const t = CATEGORIAS[c.cat];
    manejosDe_(c.manejo).forEach(m => {
      const k = c.name+"|"+m; if (pares[k]) return; pares[k] = 1;
      out.push([SEDE, t[0], t[1], t[2], c.name, m, ANO_DATA]);
    });
  });
  return out;
}
// deja listo lo que el catálogo pide: columna en Registro Interno y filas en Data; y escribe la columna Estado
function sincronizarCatalogo_(){
  const base = nombresBase_(), cat = leerCatalogo_(), estados = {}, listos = [];
  cat.forEach(c => { const e = estadoCat_(c, base); estados[c.fila] = e; if (e==="Listo" && c.activo) listos.push(c.name); });
  asegurarColumnasRI_(listos);
  const sd = dataHoja_(), pares = {}, n = sd.getLastRow();
  if (n > 1) sd.getRange(2,1,n-1,6).getValues().forEach(r => { pares[(DATA_ALIAS[String(r[4])]||String(r[4]))+"|"+String(r[5])] = 1; });
  const nuevas = filasNuevasData_(pares);
  if (nuevas.length){
    const ini = sd.getLastRow()+1;
    const f = nuevas.map((d,k) => d.concat(dataFormulas_(ini+k), [d[4]]));
    sd.getRange(ini,1,f.length,DATA_COLS.length).setValues(f);
    sd.getRange(ini,8,f.length,12).setNumberFormat("#,##0.00");
  }
  const sh = catalogoHoja_(), last = sh.getLastRow();
  if (last >= 2){
    const act = sh.getRange(2,6,last-1,1).getValues(), nuevo = [];
    let cambia = false;
    for (let i=0;i<last-1;i++){ const e = estados[i+2] || ""; nuevo.push([e]); if (String(act[i][0]) !== e) cambia = true; }
    if (cambia) sh.getRange(2,6,nuevo.length,1).setValues(nuevo);
  }
}
function catalogoParaApp_(){
  const base = nombresBase_();
  return leerCatalogo_().filter(c => c.activo && estadoCat_(c, base)==="Listo").map(c => ({name:c.name, cat:c.cat, cod:c.cod}));
}

// ---------- administración del catálogo desde la app (Inspector y Administrador)
function catalogoAdmin_(){
  const base = nombresBase_();
  const cc = cambiosListar_(40);
  const fm = formatoEstado_();
  return {ok:true, categorias:Object.keys(CATEGORIAS), manejos:MANEJOS_DATA, version:cc.version, cambios:cc.cambios,
    formato:fm.vigente, formatoPend:fm.pendientes, formatoHist:fm.historial, formatoSig:siguienteVersion_(fm.vigente.version),
    residuos:leerCatalogo_().map(c => ({name:c.name, cat:c.cat, cod:c.cod, manejo:c.manejo, activo:c.activo, estado:estadoCat_(c, base), base:!!base[c.name]}))};
}
// la categoría decide la clasificación del residuo en la hoja Data: se actualiza en sus filas
// ejecuta fn en el Sheet de cada planta activa (Registro Interno, Data y Pesadas son de cada planta)
function enCadaPlanta_(fn){
  const kS = SEDE, kSS = SS_, base = SpreadsheetApp.getActive().getId();
  try{
    plantasTodas_().filter(p => p.activo).forEach(p => {
      SEDE = p.planta; SS_ = p.sheetId === base ? null : SpreadsheetApp.openById(p.sheetId);
      try{ fn(p); }catch(err){}
    });
  } finally { SEDE = kS; SS_ = kSS; }
}
// una sola vez: los residuos que se crearon en el catálogo propio de otra planta pasan al catálogo corporativo
function fusionarCatalogos_(){
  const P = PropertiesService.getScriptProperties(); if (P.getProperty("cat_fusion_v1")) return;
  const central = catalogoHoja_(), base = SpreadsheetApp.getActive().getId();
  plantasTodas_().filter(p => p.sheetId !== base).forEach(p => {
    try{
      const ss = SpreadsheetApp.openById(p.sheetId), sh = ss.getSheetByName(CAT_HOJA); if (!sh) return;
      const n = sh.getLastRow();
      if (n >= 2){
        const ya = {}; leerCatalogo_().forEach(c => ya[norm2_(c.name)] = 1);
        sh.getRange(2,1,n-1,5).getValues().forEach(r => {
          const name = String(r[0]||"").trim(); if (!name || ya[norm2_(name)]) return; ya[norm2_(name)] = 1;
          const fila = Math.max(central.getLastRow(),1)+1;
          central.getRange(fila,3).setNumberFormat("@");
          central.getRange(fila,1,1,5).setValues([[name, String(r[1]||""), String(r[2]==null?"":r[2]), String(r[3]||""), String(r[4]||"Sí")]]);
        });
      }
      sh.setName(CAT_HOJA+" (antiguo, ya no se usa)");
    }catch(err){}
  });
  P.setProperty("cat_fusion_v1", "1");
  enCadaPlanta_(sincronizarCatalogo_);      // columnas y filas de Data de los residuos fusionados en todas las plantas
}
function categoriaEnData_(name, cat){
  const t = CATEGORIAS[cat]; if (!t) return;
  const sd = dataHoja_(), n = sd.getLastRow(); if (n < 2) return;
  const V = sd.getRange(2,5,n-1,1).getValues(), out = [];
  V.forEach((r,i) => { if ((DATA_ALIAS[String(r[0])]||String(r[0])) === name) out.push(i+2); });
  out.forEach(f => sd.getRange(f,2,1,3).setValues([[t[0],t[1],t[2]]]));
}
// renombrar un residuo creado en la app: se cambia en Catálogo, en su columna del Registro Interno, en sus filas de Data y en las pesadas
function catalogoRenombrar_(viejo, nuevo){
  const lista = leerCatalogo_(), prev = lista.filter(c => c.name === viejo)[0];
  if (!prev) return {ok:false, error:"no_existe"};
  if (nombresBase_()[viejo]) return {ok:false, error:"nombre_base"};
  if (!nuevo || nuevo.length > 120) return {ok:false, error:"nombre_invalido"};
  if (lista.some(c => c.name !== viejo && norm2_(c.name) === norm2_(nuevo)) || nombresBase_()[nuevo]) return {ok:false, error:"nombre_existe"};
  catalogoHoja_().getRange(prev.fila,1).setValue(nuevo);
  enCadaPlanta_(() => renombrarEnPlanta_(viejo, nuevo));
  return {ok:true};
}
function renombrarEnPlanta_(viejo, nuevo){
  const ri = riHoja_(), h = riEncabezado_(ri), ic = h.indexOf(viejo);
  if (ic >= 7) ri.getRange(1,ic+1).setValue(nuevo);
  const sd = dataHoja_(), nd = sd.getLastRow();
  if (nd > 1){ const V = sd.getRange(2,5,nd-1,1).getValues();
    V.forEach((r,i) => { if (String(r[0]) === viejo){ sd.getRange(i+2,5).setValue(nuevo); sd.getRange(i+2,DATA_COLS.length).setValue(nuevo); } }); }
  const sp = ss_().getSheetByName("Pesadas");
  if (sp && sp.getLastRow() > 1){ const iR = COLS_P.indexOf("residuo")+1, R = sp.getRange(2,iR,sp.getLastRow()-1,1), V = R.getValues();
    let c = false; V.forEach(r => { if (String(r[0]) === viejo){ r[0] = nuevo; c = true; } }); if (c) R.setValues(V); }
}
function catalogoGuardar_(b){
  if (b.renombrar){
    const viejo = String(b.renombrar).trim(), nuevo = String(b.name||"").trim().replace(/\s+/g," ");
    if (viejo !== nuevo){ const rr = catalogoRenombrar_(viejo, nuevo); if (!rr.ok) return rr; }
  }
  const name = String(b.name||"").trim().replace(/\s+/g," ");
  if (!name || name.length > 120) return {ok:false, error:"nombre_invalido"};
  if (!CATEGORIAS[b.cat]) return {ok:false, error:"categoria_invalida"};
  const cod = String(b.cod==null?"":b.cod).trim();
  if (cod && !/^[0-9A-Za-z.-]{1,20}$/.test(cod)) return {ok:false, error:"codigo_invalido"};
  const lista = leerCatalogo_(), prev = lista.filter(c => c.name === name)[0];
  const nuevo = !prev;
  if (nuevo && lista.some(c => norm2_(c.name) === norm2_(name))) return {ok:false, error:"nombre_existe"};
  const manejo = String(b.manejo||"").trim();
  if (nuevo && !manejosDe_(manejo).length) return {ok:false, error:"manejo_invalido"};
  if (nuevo && b.activo === undefined) b.activo = true;
  const activo = b.activo === undefined ? prev.activo : !!b.activo;
  const sh = catalogoHoja_();
  const fila = nuevo ? Math.max(sh.getLastRow(), 1) + 1 : prev.fila;
  sh.getRange(fila,3,1,1).setNumberFormat("@");
  sh.getRange(fila,1,1,5).setValues([[name, b.cat, cod, nuevo ? manejo : (manejo || prev.manejo), activo ? "Sí" : "No"]]);
  enCadaPlanta_(() => { sincronizarCatalogo_(); if (!nuevo || nombresBase_()[name]) categoriaEnData_(name, b.cat); });
  // ¿cambia lo que se imprime en el formato SGA.R.G.001? (una fila nueva o retirada, otro código SAP u otra categoría)
  let motivo = "";
  if (b.renombrar && String(b.renombrar).trim() !== name) motivo = "Residuo renombrado en el formato: "+String(b.renombrar).trim()+" → "+name;
  else if (nuevo) motivo = activo ? "Residuo agregado al formato: "+name : "";
  else if (prev.activo !== activo) motivo = (activo ? "Residuo reincorporado al formato: " : "Residuo retirado del formato: ")+name;
  else if (activo && (String(prev.cod||"") !== cod || prev.cat !== b.cat)) motivo = "Residuo modificado en el formato: "+name+(prev.cat !== b.cat ? " (categoría "+b.cat+")" : "")+(String(prev.cod||"") !== cod ? " (código SAP "+(cod||"sin código")+")" : "");
  return {ok:true, nuevo:nuevo, formato:motivo};
}

// ---------- despachos
// Cada viaje del vehículo (misma fecha y placa) es un despacho independiente. Viaje 1 conserva la clave de siempre.
function viaje_(v){ const n = parseInt(v,10); return n>=1 ? n : 1; }
function claveD_(sede, fecha, placa, viaje){ const v = viaje_(viaje); return [sede, fecha, placa_(placa)].concat(v>1 ? [String(v)] : []).join("|"); }
function origen_(viaje){ const v = viaje_(viaje); return v>1 ? "app#"+v : "app"; }
function esOrigenApp_(o){ return /^app(#\d+)?$/.test(String(o)); }

function leerDespacho_(sede, fecha, placa, viaje){
  const sd = hoja_("Despachos", COLS_D, TXT_D);
  const k = claveD_(sede, fecha, placa, viaje);
  const V = sd.getDataRange().getValues();
  for (let i=1;i<V.length;i++) if (String(V[i][0])===k){
    const o = {}; COLS_D.forEach((c,j)=>o[c]=norm_(V[i][j])); o.hora = horaTxt_(V[i][COLS_D.indexOf("hora")]); o._fila=i+1; return o;
  }
  return null;
}

function guardarDespacho_(b, u){
  const sede = b.sede || SEDE, fecha = String(b.fecha), placa = placa_(b.placa), viaje = viaje_(b.viaje);
  const sd = hoja_("Despachos", COLS_D, TXT_D);
  const cur = leerDespacho_(sede, fecha, placa, viaje);
  if (cur && cur.estado === "Aprobado") return {ok:false, error:"aprobado"};
  // manejo, observaciones y precintos: según el permiso «manejo» (hoy lo tienen los tres roles)
  if (!puede_(u.rol, "manejo")){ b.manejoNP = cur ? cur.manejoNP : ""; b.manejoP = cur ? cur.manejoP : ""; b.obs = cur ? cur.obs : ""; b.precintos = cur ? cur.precintos : ""; }
  const fila = [claveD_(sede,fecha,placa,viaje), sede, fecha, placa, b.empresa||"", b.conductor||"", b.brevete||"", b.hora||"", b.manejoNP||"", b.manejoP||"", b.obs||"", "Pendiente", "", "", new Date(), viaje, precintosTxt_(b.precintos), cur ? (cur.numero||"") : ""];
  if (cur) sd.getRange(cur._fila,1,1,COLS_D.length).setValues([fila]); else sd.appendRow(fila);
  const cel = sd.getRange(cur ? cur._fila : sd.getLastRow(), COLS_D.indexOf("hora")+1);
  cel.setNumberFormat("@"); cel.setValue(horaTxt_(b.hora));   // texto plano: Sheets no la convierte en fecha
  return {ok:true};
}

// kg por residuo (todas las puertas) de un despacho = misma fecha y placa
function consolidar_(sede, fecha, placa, viaje){
  const sh = hoja_("Hojas", COLS_H, TXT_H), sp = hoja_("Pesadas", COLS_P, TXT_P);
  const ids = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (norm_(r[1])===fecha && placa_(r[3])===placa && String(r[9]||SEDE)===sede && viaje_(r[10])===viaje_(viaje)) ids[String(r[0])] = 1;
  });
  const NP = {}, P = {}; let n = 0;
  pesadasDe_(sp, ids).forEach(p => {
    const peligroso = String(p[5]) === "Peligrosos";
    const dst = peligroso ? P : NP, res = String(p[6]);
    dst[res] = (dst[res]||0) + Number(p[11]||0); n++;
  });
  return {NP:NP, P:P, hojas:Object.keys(ids).length, pesadas:n};
}

// precintos: lista separada por comas, sin vacíos ni repetidos
function precintosTxt_(v){
  const a = (Array.isArray(v) ? v : String(v||"").split(/[,;\n]/)).map(x => String(x).trim()).filter(Boolean);
  return a.filter((x,i) => a.indexOf(x) === i).join(", ");
}

// siguiente N.° de registro de la planta; la primera vez numera también los despachos ya aprobados, por fecha de aprobación
function siguienteNumero_(sd){
  const V = sd.getDataRange().getValues().slice(1), iN = COLS_D.indexOf("numero"), iE = COLS_D.indexOf("estado"), iA = COLS_D.indexOf("aprobadoEn");
  let max = 0; V.forEach(r => { const n = parseInt(r[iN],10); if (n > max) max = n; });
  V.map((r,i) => ({r:r, i:i})).filter(o => String(o.r[iE]) === "Aprobado" && !parseInt(o.r[iN],10))
   .sort((a,b) => new Date(a.r[iA]).getTime() - new Date(b.r[iA]).getTime())
   .forEach(o => { max++; sd.getRange(o.i+2, iN+1).setValue(max); });
  return max + 1;
}

// código del N.° de registro: inicial + número de la planta (Callao 1 → C1, Pisco 2 → P2) + correlativo de 6 dígitos
function codigoRegistro_(sede, n){
  const m = String(sede||"").trim().match(/^(\S)\S*\s*(\d*)/);
  return n ? (m ? (m[1]+m[2]).toUpperCase() : "") + "-" + ("000000"+n).slice(-6) : "";
}
function numerosDespachos_(sd){
  const V = sd.getDataRange().getValues().slice(1), ix = {}; COLS_D.forEach((c,i) => ix[c]=i);
  const out = {};
  V.forEach(r => { const n = parseInt(r[ix.numero],10); if (n) out[String(r[ix.sede])+"|"+norm_(r[ix.fecha])+"|"+placa_(r[ix.placa])+"|"+origen_(viaje_(r[ix.viaje]))] = codigoRegistro_(r[ix.sede], n); });
  return out;
}

function aprobar_(b, u){
  const sede = b.sede || SEDE, fecha = String(b.fecha), placa = placa_(b.placa), viaje = viaje_(b.viaje), origen = origen_(viaje);
  const d = leerDespacho_(sede, fecha, placa, viaje);
  if (!d) return {ok:false, error:"sin_datos_despacho"};
  if (d.estado === "Aprobado") return {ok:false, error:"ya_aprobado"};
  const c = consolidar_(sede, fecha, placa, viaje);
  const grupos = [["No peligrosos", c.NP, d.manejoNP], ["Peligrosos", c.P, d.manejoP]].filter(g => Object.keys(g[1]).some(k => g[1][k] > 0));
  if (!grupos.length) return {ok:false, error:"sin_pesadas"};
  const faltan = [];
  if (!d.empresa) faltan.push("empresa operadora");
  grupos.forEach(g => { if (!g[2]) faltan.push("tipo de manejo (" + g[0].toLowerCase() + ")"); });
  if (!String(d.precintos||"").trim()) faltan.push("precintos");
  if (faltan.length) return {ok:false, error:"faltan", faltan:faltan};
  const resp = String(u.nombre||"").trim();   // quien aprueba es el usuario con sesión iniciada
  if (!resp) return {ok:false, error:"faltan", faltan:["nombre de quien aprueba"]};

  let ctx = riContexto_();
  const pedidos = []; grupos.forEach(g => Object.keys(g[1]).forEach(r => { if (g[1][r] > 0) pedidos.push(r); }));
  if (pedidos.some(r => ctx.idx[r] == null)){ sincronizarCatalogo_(); ctx = riContexto_(); }   // por si el residuo es nuevo en el catálogo
  const idx = ctx.idx, ncol = ctx.ncol, iOri = ctx.iOri, iNum = ctx.iNum;
  const sd = hoja_("Despachos", COLS_D, TXT_D);
  const numero = parseInt(d.numero,10) || siguienteNumero_(sd);   // N.° de registro correlativo por planta; al reabrir y volver a aprobar se conserva
  const sinCol = [];
  const nuevas = grupos.map(g => {
    const f = new Array(ncol).fill("");
    f[0]=sede; f[1]=Utilities.parseDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd"); f[2]=resp; f[3]=g[0]; f[4]=g[2]; f[5]=d.empresa; f[6]=placa;
    Object.keys(g[1]).forEach(res => {
      const kg = r2_(g[1][res]); if (!(kg > 0)) return;
      if (idx[res] == null) sinCol.push(res); else f[idx[res]] = kg;
    });
    f[iOri] = origen; f[iNum] = codigoRegistro_(sede, numero);
    return f;
  });
  if (sinCol.length) return {ok:false, error:"residuo_sin_columna", residuos:sinCol};

  const sr = ctx.sh;
  const V = sr.getRange(1,1,Math.max(sr.getLastRow(),1),ncol).getValues();
  const keep = [V[0]];
  let aviso = "";
  for (let i=1;i<V.length;i++){
    const r = V[i], mismo = String(r[0])===sede && norm_(r[1])===fecha && placa_(r[6])===placa;
    if (mismo && String(r[iOri])===origen) continue;       // se reemplaza (mismo viaje)
    if (mismo && !esOrigenApp_(r[iOri])) aviso = "Ya existía una fila de este despacho en el historial (no se tocó). Revisa que no se duplique.";
    keep.push(r);
  }
  // filas de la app aprobadas antes sin N.° de registro: se completa con el de su despacho
  const mapa = numerosDespachos_(sd);
  keep.forEach((r,i) => { if (i>0 && esOrigenApp_(r[iOri]) && !r[iNum]){ const c = mapa[String(r[0])+"|"+norm_(r[1])+"|"+placa_(r[6])+"|"+String(r[iOri])]; if (c) r[iNum] = c; } });
  const todo = keep.concat(nuevas);
  sr.getRange(1,1,Math.max(sr.getMaxRows(),todo.length),ncol).clearContent();
  sr.getRange(1,1,todo.length,ncol).setValues(todo);
  sr.getRange(2,2,Math.max(todo.length-1,1),1).setNumberFormat("dd/mm/yyyy");

  sd.getRange(d._fila, COLS_D.indexOf("numero")+1).setValue(numero);
  sd.getRange(d._fila, COLS_D.indexOf("estado")+1, 1, 4).setValues([["Aprobado", resp, new Date(), new Date()]]);
  dataHoja_();
  return {ok:true, filas:nuevas.length, aviso:aviso, numero:numero};
}

function reabrir_(b, u){
  const sede = b.sede || SEDE, fecha = String(b.fecha), placa = placa_(b.placa), viaje = viaje_(b.viaje), origen = origen_(viaje);
  const d = leerDespacho_(sede, fecha, placa, viaje);
  if (!d) return {ok:false, error:"sin_datos_despacho"};
  const ctx = riContexto_(), sr = ctx.sh, ncol = ctx.ncol, iOri = ctx.iOri;
  const V = sr.getRange(1,1,Math.max(sr.getLastRow(),1),ncol).getValues();
  const keep = V.filter((r,i) => i===0 || !(String(r[0])===sede && norm_(r[1])===fecha && placa_(r[6])===placa && String(r[iOri])===origen));
  sr.getRange(1,1,Math.max(sr.getMaxRows(),keep.length),ncol).clearContent();
  sr.getRange(1,1,keep.length,ncol).setValues(keep);
  const sd = hoja_("Despachos", COLS_D, TXT_D);
  sd.getRange(d._fila, COLS_D.indexOf("estado")+1, 1, 4).setValues([["Pendiente","","",new Date()]]);
  return {ok:true};
}

// ---------- bitácora: quién hizo qué (solo el Administrador la consulta desde la app)
const BIT_COLS = ["fecha","usuario","rol","accion","detalle"];
function bitacora_(u, accion, detalle){
  try{ hoja_("Bitácora", BIT_COLS, []).appendRow([new Date(), u.usuario, u.rol, accion, String(detalle||"")]); }catch(err){}
}
function actividadListar_(n){
  const sh = hoja_("Bitácora", BIT_COLS, []), V = sh.getDataRange().getValues().slice(1);
  return {ok:true, filas:V.slice(-n).reverse().map(r => ({fecha:(r[0] instanceof Date)?r[0].getTime():String(r[0]), usuario:String(r[1]), rol:String(r[2]), accion:String(r[3]), detalle:String(r[4])}))};
}
// hojas de un día (todas las puertas y vehículos): la app las baja para que cada equipo vea lo registrado en los demás
function hojasDia_(sede, fecha){
  const sh = hoja_("Hojas", COLS_H, TXT_H), sp = hoja_("Pesadas", COLS_P, TXT_P);
  const H = sh.getDataRange().getValues().slice(1).filter(r => norm_(r[1])===fecha && String(r[9]||SEDE)===sede);
  const ids = {}; H.forEach(r => ids[String(r[0])] = 1);
  const P = pesadasDe_(sp, ids);
  return H.map(r => ({
    id:String(r[0]), viaje:viaje_(r[10]), usuario:String(r[11]||""), fecha:norm_(r[1]), puerta:String(r[2]), placa:String(r[3]), responsable:String(r[4]), estado:String(r[5]),
    pesadas: P.filter(p => String(p[0])===String(r[0])).map(p => ({cod:String(p[4]), cat:String(p[5]), res:String(p[6]), bruto:p[8], tara:p[9], taraPend:p[10]==="SI", neto:p[11], uni:p[12], tu:String(p[13])}))
  }));
}
// días (hasta N atrás) con hojas cuyo despacho aún no está aprobado: para que lo pendiente no desaparezca al cambiar de día
function fechasPendientes_(sede, dias){
  const lim = Utilities.formatDate(new Date(new Date().getTime() - dias*86400000), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const sd = hoja_("Despachos", COLS_D, TXT_D), D = sd.getDataRange().getValues().slice(1), ix = {}; COLS_D.forEach((c,i) => ix[c]=i);
  const apr = {}; D.forEach(r => { if (String(r[ix.estado]) === "Aprobado") apr[norm_(r[ix.fecha])+"|"+placa_(r[ix.placa])+"|"+viaje_(r[ix.viaje])] = 1; });
  const out = {};
  hoja_("Hojas", COLS_H, TXT_H).getDataRange().getValues().slice(1).forEach(r => {
    const f = norm_(r[1]); if (!f || f < lim || String(r[9]||SEDE) !== sede) return;
    if (!apr[f+"|"+placa_(r[3])+"|"+viaje_(r[10])]) out[f] = 1;
  });
  return Object.keys(out).sort();
}
// llave privada del script de la PC (Excel y PDF). Se crea sola y vive en las Propiedades del script, nunca en la app
// versión de los datos: cambia al reiniciar los datos de prueba, para que los equipos borren lo que tenían guardado
function datosVersion_(){ return PropertiesService.getScriptProperties().getProperty("DATOS_VER") || "1"; }
function claveServicio_(){
  const P = PropertiesService.getScriptProperties();
  let k = P.getProperty("CLAVE_SERVICIO");
  if (!k){ k = Utilities.getUuid().replace(/-/g,""); P.setProperty("CLAVE_SERVICIO", k); }
  return k;
}
// ejecutar UNA vez desde el editor (botón ▶ Ejecutar) para ver la llave y copiarla en config_excel.json
function verClaveServicio(){ const k = claveServicio_(); Logger.log("Llave privada (clave_servicio): " + k); return k; }
// despachos aprobados (con N.° de registro) de los últimos días, con todo lo necesario para armar el PDF del SGA.R.G.001
function registrosPdf_(sede, dias){
  const lim = Utilities.formatDate(new Date(new Date().getTime() - dias*86400000), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const sd = hoja_("Despachos", COLS_D, TXT_D), D = sd.getDataRange().getValues().slice(1), ix = {}; COLS_D.forEach((c,i) => ix[c]=i);
  const ap = D.filter(r => String(r[ix.estado]) === "Aprobado" && parseInt(r[ix.numero],10) && String(r[ix.sede]||SEDE) === sede && norm_(r[ix.fecha]) >= lim);
  if (!ap.length) return [];
  const Hs = hoja_("Hojas", COLS_H, TXT_H).getDataRange().getValues().slice(1).filter(r => String(r[9]||SEDE) === sede && norm_(r[1]) >= lim);
  const ids = {}; Hs.forEach(r => ids[String(r[0])] = 1);
  const P = pesadasDe_(hoja_("Pesadas", COLS_P, TXT_P), ids);
  const orden = {}; leerCatalogo_().forEach((c,i) => orden[c.name] = i);
  return ap.map(r => {
    const f = norm_(r[ix.fecha]), pl = placa_(r[ix.placa]), v = viaje_(r[ix.viaje]), n = parseInt(r[ix.numero],10);
    const hid = {}; Hs.filter(h => norm_(h[1]) === f && placa_(h[3]) === pl && viaje_(h[10]) === v).forEach(h => hid[String(h[0])] = 1);
    const M = {};
    P.filter(q => hid[String(q[0])]).forEach(q => {
      const k = String(q[4]||"") || "n:" + String(q[6]).toLowerCase();
      if (!M[k]) M[k] = {cod:String(q[4]||""), res:String(q[6]), cat:String(q[5]), total:0, uni:{}};
      M[k].total = r2_(M[k].total + (Number(q[11])||0));
      const un = Number(q[12])||0; if (un){ const t = String(q[13]||"UD"); M[k].uni[t] = (M[k].uni[t]||0) + un; }
    });
    const filas = Object.keys(M).map(k => M[k]).filter(x => x.total > 0)
      .sort((a,b) => (orden[a.res]==null?9999:orden[a.res]) - (orden[b.res]==null?9999:orden[b.res]));
    const apEn = r[ix.aprobadoEn] instanceof Date ? r[ix.aprobadoEn].getTime() : (new Date(r[ix.aprobadoEn]).getTime() || 0);
    return {numero:n, codigo:codigoRegistro_(sede, n), fecha:f, placa:pl, viaje:v, empresa:String(r[ix.empresa]||""), conductor:String(r[ix.conductor]||""),
      brevete:String(r[ix.brevete]||""), hora:horaTxt_(r[ix.hora]), manejoNP:String(r[ix.manejoNP]||""), manejoP:String(r[ix.manejoP]||""),
      obs:String(r[ix.obs]||""), precintos:String(r[ix.precintos]||""), aprobadoPor:String(r[ix.aprobadoPor]||""), aprobadoEn:apEn, filas:filas};
  });
}
function despachosDia_(sede, fecha){
  const sd = hoja_("Despachos", COLS_D, TXT_D), V = sd.getDataRange().getValues().slice(1);
  const ix = {}; COLS_D.forEach((c,i) => ix[c]=i);
  return V.filter(r => norm_(r[ix.fecha])===fecha && String(r[ix.sede]||SEDE)===sede && r[ix.placa]).map(r => ({placa:placa_(r[ix.placa]), viaje:viaje_(r[ix.viaje]), estado:String(r[ix.estado]), empresa:String(r[ix.empresa]||""), conductor:String(r[ix.conductor]||""), brevete:String(r[ix.brevete]||""), hora:horaTxt_(r[ix.hora]), numero:parseInt(r[ix.numero],10)||""}));
}
// vehículos habituales: empresa, conductor y brevete del último despacho de cada placa (para no volver a escribirlos)
function vehiculosHabituales_(){
  const sd = hoja_("Despachos", COLS_D, TXT_D), V = sd.getDataRange().getValues().slice(1);
  const iP = COLS_D.indexOf("placa"), iE = COLS_D.indexOf("empresa"), iC = COLS_D.indexOf("conductor"), iB = COLS_D.indexOf("brevete");
  const M = {};
  V.slice(-600).forEach(r => { const pl = placa_(r[iP]); if (pl && r[iE] && r[iC]) M[pl] = {placa:pl, empresa:String(r[iE]), conductor:String(r[iC]), brevete:String(r[iB]||"")}; });
  return {ok:true, vehiculos:Object.keys(M).map(k => M[k])};
}

// ---------- hojas de pesaje
// el Administrador puede registrar una hoja a nombre de otra persona (para pruebas o apoyo); los demás roles siempre usan su propio nombre
function respNombre_(x, u){
  if (u.rol !== "Administrador") return u.nombre;
  const n = String(x||"").trim().replace(/\s+/g," ");
  return (n.length<=60 && n.split(" ").length>=2) ? n : u.nombre;
}
// filas de la pestaña Pesadas (2..n) cuyo id de hoja está en ids
function filasPesadas_(sp, ids){
  const last = sp.getLastRow(), out = [];
  if (last < 2) return out;
  const col = sp.getRange(2,1,last-1,1).getValues();
  for (let i=0;i<col.length;i++) if (ids[String(col[i][0])]) out.push(i+2);
  return out;
}
// pesadas de un grupo de hojas: lee solo el tramo de la pestaña donde están (no todo el historial)
function pesadasDe_(sp, ids){
  const f = filasPesadas_(sp, ids);
  if (!f.length) return [];
  return sp.getRange(f[0],1,f[f.length-1]-f[0]+1,COLS_P.length).getValues().filter(p => ids[String(p[0])]);
}
function asegurarFilas_(sh, hasta){ const falta = hasta - sh.getMaxRows(); if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta + 200); }
// borra filas (números ascendentes) por tramos, de abajo hacia arriba
function borrarFilas_(sh, filas){
  if (!filas.length) return;
  if (sh.getMaxRows() - filas.length < 2) sh.insertRowsAfter(sh.getMaxRows(), 50);   // Sheets no permite borrar todas las filas
  for (let k=filas.length-1; k>=0; ){ let j=k; while (j>0 && filas[j-1]===filas[j]-1) j--; sh.deleteRows(filas[j], k-j+1); k=j-1; }
}
// Reemplaza las pesadas de UNA hoja sin tocar las demás filas.
// Primero se escribe lo nuevo y después se borra lo anterior: si la ejecución se corta, nunca se pierden pesadas.
function escribirPesadas_(sp, id, nuevas){
  const ids = {}; ids[id] = 1;
  const viejas = filasPesadas_(sp, ids), last = sp.getLastRow(), n = nuevas.length, m = viejas.length;
  const bloque = m > 0 && viejas[m-1]-viejas[0]+1 === m;
  if (bloque && n === m){ if (n) sp.getRange(viejas[0],1,n,COLS_P.length).setValues(nuevas); return; }
  if (bloque && viejas[m-1] === last){                        // la hoja está al final de la pestaña: se escribe en su sitio
    if (n){ asegurarFilas_(sp, viejas[0]+n-1); sp.getRange(viejas[0],1,n,COLS_P.length).setValues(nuevas); }
    if (n < m) sp.getRange(viejas[0]+n,1,m-n,COLS_P.length).clearContent();
    return;
  }
  if (n){ asegurarFilas_(sp, last+n); sp.getRange(last+1,1,n,COLS_P.length).setValues(nuevas); }
  borrarFilas_(sp, viejas);
}
// Inspector/Administrador: elimina una hoja sin pesadas (creada por error)
function borrarHoja_(b, u){
  const id = String(b.id), sh = hoja_("Hojas", COLS_H, TXT_H), sp = hoja_("Pesadas", COLS_P, TXT_P);
  const idsH = sh.getRange(1,1,Math.max(sh.getLastRow(),1),1).getValues().map(r=>String(r[0])), iH = idsH.indexOf(id);
  if (iH <= 0) return {ok:true};
  const r = sh.getRange(iH+1,1,1,COLS_H.length).getValues()[0], ids = {}; ids[id] = 1;
  if (filasPesadas_(sp, ids).length) return {ok:false, error:"con_pesadas"};
  const dsp = leerDespacho_(String(r[9]||SEDE), norm_(r[1]), r[3], viaje_(r[10]));
  if (dsp && dsp.estado === "Aprobado") return {ok:false, error:"aprobado"};
  borrarFilas_(sh, [iH+1]);
  bitacora_(u, "Hoja eliminada", String(r[2])+" · "+placa_(r[3]));
  return {ok:true};
}
function guardarHoja_(b, u){
  const id = String(b.id), sede = b.sede || SEDE;
  const sp = hoja_("Pesadas", COLS_P, TXT_P), sh = hoja_("Hojas", COLS_H, TXT_H);
  // si el despacho ya está aprobado, no se aceptan cambios en sus hojas
  const viaje = viaje_(b.viaje), dsp = leerDespacho_(sede, String(b.fecha), b.placa, viaje);
  if (dsp && dsp.estado === "Aprobado") return {ok:false, error:"aprobado"};
  // el Operador solo modifica sus propias hojas
  const idsH = sh.getRange(1,1,Math.max(sh.getLastRow(),1),1).getValues().map(r=>String(r[0]));
  const iH = idsH.indexOf(id);
  const prev = iH>0 ? sh.getRange(iH+1,1,1,COLS_H.length).getValues()[0] : null;
  const dueno = prev ? String(prev[11]||"") : "";
  if (dueno && dueno !== u.usuario && u.rol === "Operador") return {ok:false, error:"ajena"};

  const nuevas = (b.pesadas||[]).map(p => [id, b.fecha, b.puerta, placa_(b.placa), p.cod||"", p.cat, p.res, p.n, p.bruto, p.tara, p.taraPend?"SI":"NO", p.neto, p.uni, p.tu, sede]);
  escribirPesadas_(sp, id, nuevas);

  const kg = (b.pesadas||[]).reduce((a,p)=>a+Number(p.neto||0),0);
  const fila = [id, b.fecha, b.puerta, placa_(b.placa), (prev && String(prev[4])) ? String(prev[4]) : respNombre_(b.responsable, u), b.estado, (b.pesadas||[]).length, r2_(kg), new Date(), sede, viaje, dueno || u.usuario];
  if (iH>0) sh.getRange(iH+1,1,1,COLS_H.length).setValues([fila]); else sh.appendRow(fila);
  const det = b.puerta+" · "+placa_(b.placa)+(viaje>1?" · viaje "+viaje:"");
  if (!prev) bitacora_(u, "Hoja creada", det);
  else if (String(prev[5]) !== String(b.estado)) bitacora_(u, b.estado==="cerrada" ? "Hoja cerrada" : "Hoja reabierta", det);
  return {ok:true};
}

const SIN_TURNO = ["usuarios_listar","actividad_listar","plantas","hallazgos_listar","hallazgo_foto","areas_listar"];   // solo lectura: no esperan su turno
function doPost(e){
  let b0 = null; try{ b0 = JSON.parse(e.postData.contents); }catch(err){ return json_({ok:false, error:"formato"}); }
  const lock = SIN_TURNO.indexOf(b0.action) >= 0 ? {releaseLock:function(){}} : LockService.getScriptLock();
  if (lock.waitLock) lock.waitLock(25000);
  try{
    const b = b0;
    if (b.key !== CLAVE) return json_({ok:false, error:"clave"});
    estructura_();
    if (b.action === "login") return json_(login_(b));
    const u = sesion_(b.token);
    if (!u) return json_({ok:false, error:"sesion"});
    if (b.action === "cambiarClave") return json_(cambiarClave_(b, u));
    if (u.cambiar) return json_({ok:false, error:"cambiar_clave"});
    const A = b.action, no = json_({ok:false, error:"permiso"});
    // lo que vive en el Sheet principal: usuarios y plantas
    if (A === "plantas") return json_(plantasListar_(u));
    if (A === "planta_crear" || A === "planta_editar"){
      if (!puede_(u.rol,"plantas")) return no;
      const r = A === "planta_crear" ? plantaCrear_(b, u) : plantaEditar_(b, u);
      if (r && r.ok) bitacora_(u, A === "planta_crear" ? "Planta creada" : "Planta modificada", String(b.planta)+(b.lugares ? " · "+lugares_(b.lugares).join(", ") : "")+(b.activo===false?" · desactivada":""));
      return json_(r);
    }
    // el «Responsable de área» solo puede usar lo de hallazgos (el servidor lo hace cumplir, no solo la app)
    if (u.rol === "Responsable de área" && A.indexOf("hallazgo") !== 0 && A.indexOf("area") !== 0) return no;
    if (A.indexOf("usuario") !== 0){ const up = usarPlanta_(b.sede, u); if (!up.ok) return json_(up); estructura_(); }
    // hallazgos ambientales (la lógica vive en el archivo Hallazgos)
    if (A === "hallazgo_crear") return puede_(u.rol,"hallazgo_reportar") ? json_(hzCrear_(b, u)) : no;
    if (A === "hallazgo_asignar") return puede_(u.rol,"hallazgo_asignar") ? json_(hzAsignar_(b, u)) : no;
    if (A === "hallazgo_cerrar") return puede_(u.rol,"hallazgo_cerrar") ? json_(hzCerrar_(b, u)) : no;
    if (A === "hallazgo_reabrir") return puede_(u.rol,"hallazgo_reabrir") ? json_(hzReabrir_(b, u)) : no;
    if (A === "hallazgos_listar") return json_(hzListar_(u));
    if (A === "hallazgo_foto") return json_(hzFoto_(b, u));
    if (A === "areas_listar") return json_(hzAreasListar_(u));
    if (A === "area_guardar") return puede_(u.rol,"areas_admin") ? json_(hzAreaGuardar_(b, u)) : no;
    if (A === "hoja") return json_(guardarHoja_(b, u));
    if (A === "hoja_borrar") return puede_(u.rol,"borrarHoja") ? json_(borrarHoja_(b, u)) : no;
    if (A === "despacho") return json_(guardarDespacho_(b, u));
    if (A === "aprobar" || A === "reabrir"){
      if (!puede_(u.rol, A)) return no;
      if (A === "reabrir"){ const cc = confirmarClave_(u, b.claveConfirma); if (!cc.ok) return json_(cc); }   // aprobar ya no pide clave (la sesión identifica a quien aprueba)
      const r = A === "aprobar" ? aprobar_(b, u) : reabrir_(b, u);
      if (r && r.ok) bitacora_(u, A==="aprobar" ? "Despacho aprobado" : "Despacho reabierto", placa_(b.placa)+" · "+b.fecha+(viaje_(b.viaje)>1?" · viaje "+viaje_(b.viaje):""));
      return json_(r);
    }
    if (A === "actividad_listar") return puede_(u.rol,"actividad") ? json_(actividadListar_(150)) : no;
    if (A === "catalogo_admin" || A === "catalogo_guardar"){
      if (!puede_(u.rol,"catalogo")) return no;
      if (A === "catalogo_admin"){ fusionarCatalogos_(); sincronizarCatalogo_(); return json_(catalogoAdmin_()); }
      const r = catalogoGuardar_(b);
      if (r && r.ok) enCadaPlanta_(() => {
      if (b.renombrar && String(b.renombrar).trim() !== String(b.name).trim()) cambio_(u, "Residuo renombrado (Catálogo, columna del Registro Interno, Data y pesadas)", String(b.renombrar)+" → "+String(b.name), false);
      cambio_(u, r.nuevo ? "Residuo agregado (columna nueva en Registro Interno y filas en Data)" : (b.activo===false ? "Residuo retirado de la app" : "Residuo modificado"), String(b.name)+" · "+b.cat, !!r.nuevo); });
      if (r && r.ok) formatoPendiente_(u, r.formato);
      if (r && r.ok) bitacora_(u, r.nuevo ? "Residuo agregado" : "Residuo modificado", String(b.name)+" · "+b.cat+(b.activo===false?" · retirado":""));
      return json_(r);
    }
    if (A === "formato_guardar" || A === "formato_descartar"){
      if (!puede_(u.rol,"catalogo")) return no;
      const r = A === "formato_guardar" ? formatoGuardar_(b, u) : formatoDescartar_(b, u);
      if (r && r.ok) bitacora_(u, A === "formato_guardar" ? "Formato actualizado" : "Cambio de formato descartado", A === "formato_guardar" ? r.formato.codigo+" · versión "+r.formato.version+" · "+r.formato.aprobacion : "");
      return json_(r);
    }
    if (A === "usuarios_listar" || A === "usuario_crear" || A === "usuario_editar" || A === "usuario_resetear" || A === "usuario_borrar"){
      if (!puede_(u.rol,"usuarios")) return no;
      if (A === "usuarios_listar") return json_(usuariosListar_());
      const r = A==="usuario_crear" ? usuarioCrear_(b) : A==="usuario_editar" ? usuarioEditar_(b, u) : A==="usuario_borrar" ? usuarioBorrar_(b, u) : usuarioResetear_(b);
      if (r && r.ok) bitacora_(u, A==="usuario_crear" ? "Usuario creado" : A==="usuario_editar" ? "Usuario modificado" : A==="usuario_borrar" ? "Usuario eliminado" : "Contraseña restablecida", String(b.usuario)+(b.rol?" · "+b.rol:"")+(b.activo===false?" · desactivado":""));
      return json_(r);
    }
    return json_({ok:false, error:"accion"});
  }catch(err){
    return json_({ok:false, error:String(err)});
  }finally{ lock.releaseLock(); }
}

function doGet(e){
  const q = e.parameter||{};
  if (q.key !== CLAVE) return json_({ok:false, error:"clave"});
  if (q.action === "ping") return json_({ok:true});
  if (q.action === "plantas_lista") return json_({ok:true, plantas:plantasTodas_().filter(p => p.activo).map(p => p.planta)});   // para la rutina que actualiza los Excel
  if (["catalogo","despacho","hojas","vehiculos","data","registro","buscar_numero","pendientes","registros_pdf"].indexOf(q.action) >= 0){
    // sin sesión: «data» (solo totales, para el dashboard) y, con la llave privada del script de la PC, «registro» y «registros_pdf»
    const svc = !!q.servicio && q.servicio === claveServicio_();
    const libre = svc && (q.action === "data" || q.action === "registro" || q.action === "registros_pdf");
    const u = libre && !q.token ? null : sesion_(q.token);
    if (!u && !libre) return json_({ok:false, error: (q.action === "data" || q.action === "registro" || q.action === "registros_pdf") ? "servicio" : "sesion"});
    if (q.action === "registros_pdf" && !svc) return json_({ok:false, error:"servicio"});
    if (u && u.cambiar) return json_({ok:false, error:"cambiar_clave"});
    if (u && u.rol === "Responsable de área") return json_({ok:false, error:"permiso"});   // solo usa hallazgos: tampoco lee pesajes ni catálogo
    const up = usarPlanta_(q.sede, u); if (!up.ok) return json_(up);
  }
  if (q.action === "catalogo"){   // la lista de residuos que usa la app (pestaña Catálogo)
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try{
      estructura_(); fusionarCatalogos_(); sincronizarCatalogo_();
      return json_({ok:true, residuos:catalogoParaApp_(), formato:formatoEstado_().vigente});
    }catch(err){ return json_({ok:false, error:String(err)}); }
    finally{ lock.releaseLock(); }
  }
  if (q.action === "data"){   // para el script del dashboard: misma forma que la hoja "Data" del Excel
    const rows = dataHoja_().getDataRange().getValues().slice(1).filter(r => r[0]).map(r => r.map(v => (v instanceof Date) ? norm_(v) : v));
    return json_({ok:true, cols:DATA_COLS, rows:rows, ahora:new Date().getTime()});
  }
  if (q.action === "registro"){   // Registro Interno completo de la planta (para actualizar su Excel)
    const ctx = riContexto_(), n = ctx.sh.getLastRow();
    const rows = n < 2 ? [] : ctx.sh.getRange(2,1,n-1,ctx.ncol).getValues().filter(r => r[0] !== "").map(r => r.map(v => (v instanceof Date) ? norm_(v) : v));
    return json_({ok:true, planta:SEDE, cols:ctx.hdr, rows:rows, version:versionArchivo_(cambiosHoja_()), ahora:new Date().getTime()});
  }
  if (q.action === "hojas") { const sede = q.sede || SEDE, fecha = String(q.fecha||""); return json_({ok:true, hojas:hojasDia_(sede, fecha), despachos:despachosDia_(sede, fecha)}); }
  if (q.action === "vehiculos") return json_(vehiculosHabituales_());
  if (q.action === "registros_pdf") return json_({ok:true, planta:SEDE, formato:formatoEstado_().vigente, registros:registrosPdf_(SEDE, parseInt(q.dias,10) || 60)});
  if (q.action === "pendientes") return json_({ok:true, fechas:fechasPendientes_(q.sede || SEDE, 30), ver:datosVersion_()});
  if (q.action === "buscar_numero"){   // N.° de registro -> fecha, placa y viaje del despacho (en la planta elegida)
    const n = parseInt(q.numero,10), sd = hoja_("Despachos", COLS_D, TXT_D), V = sd.getDataRange().getValues().slice(1);
    const ix = {}; COLS_D.forEach((c,i) => ix[c]=i);
    const r = n ? V.find(x => parseInt(x[ix.numero],10) === n) : null;
    return json_(r ? {ok:true, fecha:norm_(r[ix.fecha]), placa:placa_(r[ix.placa]), viaje:viaje_(r[ix.viaje]), estado:String(r[ix.estado])} : {ok:false, error:"no_existe"});
  }
  if (q.action === "despacho"){
    const sede = q.sede || SEDE, fecha = q.fecha, placa = placa_(q.placa), viaje = viaje_(q.viaje);
    const sh = hoja_("Hojas", COLS_H, TXT_H), sp = hoja_("Pesadas", COLS_P, TXT_P);
    const H = sh.getDataRange().getValues().slice(1).filter(r => norm_(r[1])===fecha && placa_(r[3])===placa && String(r[9]||SEDE)===sede && viaje_(r[10])===viaje);
    const idsD = {}; H.forEach(r => idsD[String(r[0])] = 1);
    const P = pesadasDe_(sp, idsD);
    const hojas = H.map(r => ({
      id:String(r[0]), viaje:viaje_(r[10]), usuario:String(r[11]||""), fecha:norm_(r[1]), puerta:String(r[2]), placa:String(r[3]), responsable:String(r[4]), estado:String(r[5]),
      pesadas: P.filter(p => String(p[0])===String(r[0])).map(p => ({cod:String(p[4]), cat:String(p[5]), res:String(p[6]), bruto:p[8], tara:p[9], taraPend:p[10]==="SI", neto:p[11], uni:p[12], tu:String(p[13])}))
    }));
    const d = leerDespacho_(sede, fecha, placa, viaje);
    return json_({ok:true, hojas, despacho: d ? {empresa:d.empresa, conductor:d.conductor, brevete:d.brevete, hora:d.hora, manejoNP:d.manejoNP, manejoP:d.manejoP, obs:d.obs, precintos:d.precintos||"", numero:d.numero||"", estado:d.estado, aprobadoPor:d.aprobadoPor, aprobadoEn:d.aprobadoEn} : null});
  }
  return json_({ok:false, error:"accion"});
}
/**
 * REINICIAR DATOS DE PRUEBA (se ejecuta una sola vez, a mano, desde el editor de Apps Script)
 * Borra en TODAS las plantas: hojas, pesadas, despachos, bitácora y las filas del Registro Interno
 * que vinieron de la app (Origen = app). NO toca: usuarios, plantas, catálogo, formato SGA,
 * ni las filas históricas del Registro Interno (las que no tienen Origen).
 */
function reiniciarDatosDePrueba(){
  const base = SpreadsheetApp.getActive().getId(), out = [];
  plantasTodas_().forEach(p => {
    const ss = p.sheetId === base ? SpreadsheetApp.getActive() : SpreadsheetApp.openById(p.sheetId);
    ["Hojas","Pesadas","Despachos","Bitácora"].forEach(n => {
      const sh = ss.getSheetByName(n); if (!sh) return;
      const f = sh.getLastRow(); if (f > 1) sh.deleteRows(2, f - 1);
      out.push(p.planta + " · " + n + ": " + Math.max(f - 1, 0) + " filas borradas");
    });
    const ri = ss.getSheetByName(RI);
    if (ri && ri.getLastRow() > 1){
      const h = ri.getRange(1,1,1,ri.getLastColumn()).getValues()[0].map(String), io = h.indexOf(RI_ORIGEN);
      if (io >= 0){
        const V = ri.getRange(2, io+1, ri.getLastRow()-1, 1).getValues(); let n = 0;
        for (let i = V.length-1; i >= 0; i--) if (/^app(#\d+)?$/.test(String(V[i][0]))){ ri.deleteRow(i+2); n++; }
        out.push(p.planta + " · Registro Interno: " + n + " filas de la app borradas");
      }
    }
  });
  CacheService.getScriptCache().removeAll(["plantas"]);
  PropertiesService.getScriptProperties().setProperty("DATOS_VER", String(new Date().getTime()));   // los celulares y PC borran sus copias al abrir la app
  Logger.log(out.join("\n"));
}