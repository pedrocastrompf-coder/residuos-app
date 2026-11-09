/* ============================================================
   CONFIGURACIÓN — lo único que se edita para conectar la app
   ============================================================ */
window.APP_CONFIG = {
  // 1) URL de la "aplicación web" de Google Apps Script (termina en /exec).
  //    Mientras esté vacía, la app funciona en modo local (solo en este equipo).
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbzVOzRnrNZlxjMUHIlnPRWwKHPOTOU_ygAUzQvH8l48qenilap-SZjV8N9Yw2c--EfUwQ/exec",

  // 2) Clave simple. Debe ser IGUAL a la CLAVE de Code.gs.
  CLAVE: "MEDIOAMBIENTE",

  ZONA_HORARIA: "America/Lima",
  PLANTA: "Callao 1",
  PUERTAS: ["Puerta 2", "Puerta 3", "Puerta 6"],

  // Placas de los vehículos de despacho (se puede escribir otra).
  PLACAS: ["ABZ-937","AKU-870","ATP-727","ATX-813","BCP-909","BHR-926","BKQ-735","BML-865","BMM-742","BNW-858","BTB-776"],

  // Taras (kg del coche) que aparecen al inicio como botones rápidos.
  // La app va recordando las últimas usadas.
  TARAS_INICIALES: [62, 68.5, 73],

  // Datos del formato oficial que se imprime
  FORMATO: { codigo: "SGA.R.G.001", version: "05", aprobacion: "06/04/2026" },

  // Tipos de manejo del despacho (los del SGA.R.G.001)
  TIPOS_MANEJO: ["Comercialización - Disposición final","Comercialización","Disposición final","Donación (Aniquem)","Valorización","Tratamiento - Valorización - Disp. final"],

  // Unidades de conteo (opcional en cada pesada)
  UNIDADES: [["B","Bolsas"],["C","Cajas"],["UD","Unidades"]],

  // Catálogo = SGA.R.G.001 Vs 05. [código SAP o "", nombre]
  CATALOGO: [
    { cat: "Orgánico - producción", items: [
      ["6001000226","Residuo chocolate"],["6001000297","Residuo chocolate FSTB"],["6001000227","Residuo wafer"],
      ["6001000247","Residuo galleta"],["6001000260","Residuo galleta de chocolate"],["6001000271","Residuo galleta oreo multipack"],
      ["6001000273","Residuo galleta oreo crumbs"],["6001000277","Residuo galleta oreo tripack original"],
      ["6001000276","Residuo galleta oreo tripack chocolate"],["6001000296","Residuo galleta oreo bañado dark"],
      ["6001000295","Residuo galleta oreo bañado white"],["6001000318","Residuo masa fermentada"] ]},
    { cat: "Orgánico - comedor", items: [ ["","Residuo orgánico de comedor"] ]},
    { cat: "Plástico", items: [
      ["6001000238","Residuo polietileno (stretch film y manga plast)"],["6001000258","Residuo polietileno sucio"],
      ["6001000207","Residuo saco de polipropileno (rafia)"],["6001000229","Residuo plástico duro"],
      ["6001000214","Residuo galonera plástica 25kg"],["6001000249","Residuo bidón azul 280 litros"],
      ["6001000212","Residuo balde plástico"],["6001000335","Residuo molde de plástico"] ]},
    { cat: "Madera", items: [ ["6001000208","Residuo parihuela en desuso"],["6001000210","Residuo parihuela leña"] ]},
    { cat: "Papel y cartón", items: [ ["6001000231","Residuo cartón"],["6001000232","Residuo bolsa de papel"],["6001000233","Residuo papel bond"] ]},
    { cat: "Metal", items: [
      ["5003000032","Residuo chatarra"],["6001000236","Residuo cables de cobre"],["6001000213","Residuo cilindro de metal"],["6001000259","Residuo acero"] ]},
    { cat: "Otros aprovechables", items: [ ["6001000240","Residuo aceite usado"],["6001000237","Residuo saco de yute"],["6001000305","Residuo bobina en rollo"] ]},
    { cat: "No aprovechable", items: [ ["","Residuo varios"],["","Residuo bobina inutilizada"],["","Residuo semisólido (lodo)"] ]},
    { cat: "Especial", items: [
      ["","Residuo de aparato eléctrico y electrónico (RAEE)"],["","Neumático fuera de uso (NFU)"],["","Residuo de construcción y demolición (RCD)"] ]},
    { cat: "Peligrosos", peligroso: true, items: [
      ["","Residuo de laboratorio (autoclavado)"],["","Residuo envase vacío de tinta (citronix)"],["","Residuo biocontaminado"],
      ["","Residuo saco vacío de bicarbonato de amonio/bicarbonato de sodio"],["","Residuo saco vacío de carbonato de potasio/soda cáustica"],
      ["","Residuos varios contaminados (trapos, waypes, otros)"],["","Residuo envase vacío de producto químico (plástico)"],
      ["","Residuo envase vacío de producto químico (vidrio, metal)"],["","Residuo aceite contaminado"],["","Residuo batería usada"],
      ["","Residuo balón vacío de refrigerante"],["","Residuo de asbesto"],["","Residuo envase con producto químico"],
      ["","Residuo luminaria rota"],["","Residuo fibra de vidrio"],["","Residuo fosfuro de aluminio"] ]}
  ]
};
