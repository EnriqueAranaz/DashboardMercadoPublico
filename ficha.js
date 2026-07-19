/* ============================================================
   Modulo "ficha" — modal de detalle de una licitación.
   Autonomo: no lee ni modifica el estado de app.js. app.js solo llama
   a `Ficha.abrir(fila)` pasandole la fila completa del excel filtrado;
   este modulo se encarga de construir su propio HTML (fetch de
   ficha.html), pintarlo y manejar apertura/cierre.

   Para no depender del orden de carga de <script> ni de variables
   globales de app.js, duplica aqui las utilidades minimas que necesita
   (son ~15 lineas, se prefiere ese pequeño duplicado a un acoplamiento
   entre modulos).
   ============================================================ */

const Ficha = (function () {
  const KEYWORDS_DASHBOARD = ["jardin", "jardines", "area verde", "areas verdes", "paisajismo"];
  const RUTA_FRAGMENTO = "ficha.html";

  let listo = false; // true cuando ficha.html ya fue inyectado en el DOM

  /* ---------- Utilidades (copia local, ver nota arriba) ---------- */

  function normalizarTexto(texto) {
    const combining = new RegExp(
      String.fromCharCode(0x5b, 0x5c, 0x75, 0x30, 0x33, 0x30, 0x30, 0x2d, 0x5c, 0x75, 0x30, 0x33, 0x36, 0x66, 0x5d),
      "g"
    );
    return (texto || "").toString().normalize("NFD").replace(combining, "").toLowerCase().trim();
  }

  function escapeHtml(texto) {
    return (texto || "")
      .toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function formatearMonto(valor) {
    const n = Number(valor);
    if (!valor || isNaN(n) || n <= 0) return "Sin monto informado";
    return "$" + n.toLocaleString("es-CL");
  }

  function formatearFechaHora(valor) {
    if (!valor) return "—";
    const d = new Date(valor);
    if (isNaN(d.getTime())) return String(valor);
    return (
      d.toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric" }) +
      " " +
      d.toLocaleTimeString("es-CL", { hour: "2-digit", minute: "2-digit" })
    );
  }

  /** Un valor "vacio" en el excel puede llegar como "", null, undefined, " " o "None". */
  function valorVacio(valor) {
    if (valor === null || valor === undefined) return true;
    const texto = String(valor).trim();
    return texto === "" || texto.toLowerCase() === "none" || texto === "nan";
  }

  function claseEstado(estado) {
    const e = normalizarTexto(estado);
    if (e.includes("public")) return "estado-badge--publicada";
    if (e.includes("cerrad")) return "estado-badge--cerrada";
    if (e.includes("adjudic")) return "estado-badge--adjudicada";
    if (e.includes("revocad")) return "estado-badge--revocada";
    if (e.includes("desiert")) return "estado-badge--desierta";
    return "estado-badge--otro";
  }

  function palabrasClaveDe(fila) {
    if (fila.PalabrasClave) {
      return String(fila.PalabrasClave).split(";").map((s) => s.trim()).filter(Boolean);
    }
    const texto = normalizarTexto(fila.Nombre);
    return KEYWORDS_DASHBOARD.filter((k) => texto.includes(normalizarTexto(k)));
  }

  /* ---------- Construccion de la ficha ---------- */

  /** Devuelve un bloque .ficha__field si el valor no esta vacio, o "" si no hay nada que mostrar. */
  function campoFicha(label, valor, opciones) {
    opciones = opciones || {};
    if (valorVacio(valor)) return "";
    let valorHtml;
    if (opciones.esFecha) valorHtml = escapeHtml(formatearFechaHora(valor));
    else if (opciones.esMonto) valorHtml = escapeHtml(formatearMonto(valor));
    else if (opciones.esLink) valorHtml = '<a href="' + escapeHtml(valor) + '" target="_blank" rel="noopener">Ver acta de adjudicación ↗</a>';
    else valorHtml = escapeHtml(valor);

    return (
      '<div class="ficha__field">' +
      '<span class="ficha__field-label">' + escapeHtml(label) + '</span>' +
      '<span class="ficha__field-value">' + valorHtml + '</span>' +
      '</div>'
    );
  }

  /** Agrupa varios campoFicha en una seccion; si ninguno tiene contenido, no se muestra la seccion. */
  function seccionFicha(titulo, camposHtml) {
    const camposConContenido = camposHtml.filter(Boolean);
    if (!camposConContenido.length) return "";
    return (
      '<div class="ficha__section">' +
      '<h3 class="ficha__section-title">' + escapeHtml(titulo) + '</h3>' +
      '<div class="ficha__grid">' + camposConContenido.join("") + '</div>' +
      '</div>'
    );
  }

  function renderFicha(f) {
    const kws = palabrasClaveDe(f);
    const kwsHtml = kws.length
      ? '<div class="ficha__keywords">' + kws.map((k) => '<span class="keyword-tag">' + escapeHtml(k) + "</span>").join("") + "</div>"
      : "";

    const header =
      '<div class="ficha__header">' +
      '<div class="ficha__top-row">' +
      '<h2 class="ficha__nombre">' + escapeHtml(f.Nombre) + '</h2>' +
      '<span class="estado-badge ' + claseEstado(f.Estado) + '">' + escapeHtml(f.Estado) + '</span>' +
      '</div>' +
      '<p class="ficha__codigo">Código Mercado Público: <code>' + escapeHtml(f.CodigoExterno) + '</code></p>' +
      kwsHtml +
      '</div>';

    const descripcion = valorVacio(f.Descripcion)
      ? ""
      : '<div class="ficha__section"><h3 class="ficha__section-title">Descripción</h3>' +
        '<p class="ficha__descripcion">' + escapeHtml(f.Descripcion) + '</p></div>';

    const general = seccionFicha("Datos generales", [
      campoFicha("Organismo", f.Organismo),
      campoFicha("Tipo de licitación", f.Tipo),
      campoFicha("Región", f.Region),
      campoFicha("Comuna", f.Comuna),
      campoFicha("Monto estimado", f.MontoEstimado, { esMonto: true }),
      campoFicha("Moneda", f.Moneda),
      campoFicha("Categorías", f.Categorias),
      campoFicha("Cantidad de reclamos al organismo", f.CantidadReclamos),
    ]);

    const contacto = seccionFicha("Unidad de compra y contacto", [
      campoFicha("Unidad de compra", f.NombreUnidadCompra),
      campoFicha("Dirección unidad de compra", f.DireccionUnidadCompra),
      campoFicha("RUT organismo", f.RutOrganismo),
      campoFicha("Responsable", f.NombreUsuarioResponsable),
      campoFicha("Cargo responsable", f.CargoUsuarioResponsable),
    ]);

    const fechas = seccionFicha("Fechas del proceso", [
      campoFicha("Publicación", f.FechaPublicacion, { esFecha: true }),
      campoFicha("Cierre de recepción de ofertas", f.FechaCierre, { esFecha: true }),
      campoFicha("Días para el cierre", f.DiasCierreLicitacion),
      campoFicha("Visita a terreno", f.FechaVisitaTerreno, { esFecha: true }),
      campoFicha("Dirección de la visita", f.DireccionVisitaTerreno),
      campoFicha("Publicación de respuestas", f.FechaPublicacionRespuestas, { esFecha: true }),
      campoFicha("Apertura técnica", f.FechaAperturaTecnica, { esFecha: true }),
      campoFicha("Apertura económica", f.FechaAperturaEconomica, { esFecha: true }),
      campoFicha("Entrega de antecedentes", f.FechaEntregaAntecedentes, { esFecha: true }),
      campoFicha("Adjudicación estimada", f.FechaEstimadaAdjudicacion, { esFecha: true }),
      campoFicha("Adjudicación", f.FechaAdjudicacion, { esFecha: true }),
      campoFicha("Firma estimada", f.FechaEstimadaFirma, { esFecha: true }),
    ]);

    const condiciones = seccionFicha("Condiciones contractuales", [
      campoFicha("Modalidad de pago", f.ModalidadPago),
      campoFicha(
        "Duración del contrato",
        valorVacio(f.TiempoDuracionContrato) ? "" : f.TiempoDuracionContrato + " " + (f.UnidadTiempoDuracionContrato || "")
      ),
      campoFicha("Tipo de duración", f.TipoDuracionContrato),
      campoFicha("¿Renovable?", f.EsRenovable),
      campoFicha("¿Permite subcontratación?", f.PermiteSubcontratacion),
      campoFicha("¿Extensión automática de plazo?", f.ExtensionPlazoAutomatica),
      campoFicha("Responsable de pago", f.NombreResponsablePago),
      campoFicha("Email responsable de pago", f.EmailResponsablePago),
      campoFicha("Responsable de contrato", f.NombreResponsableContrato),
      campoFicha("Email responsable de contrato", f.EmailResponsableContrato),
      campoFicha("Teléfono responsable de contrato", f.FonoResponsableContrato),
    ]);

    const items = seccionFicha("Ítems licitados", [
      campoFicha("Cantidad de ítems", f.CantidadItems),
      campoFicha("Productos / servicios", f.NombresItems),
    ]);

    const adjudicacionCampos = [
      campoFicha("Tipo de acto administrativo", f.TipoActoAdjudicacion),
      campoFicha("Número de acto", f.NumeroActoAdjudicacion),
      campoFicha("Fecha del acto", f.FechaActoAdjudicacion, { esFecha: true }),
      campoFicha("Número de oferentes", f.NumeroOferentes),
      campoFicha("Proveedor(es) adjudicado(s)", f.ProveedoresAdjudicados),
      campoFicha("Monto adjudicado total", f.MontoAdjudicadoTotal, { esMonto: true }),
    ];
    let adjudicacion = seccionFicha("Adjudicación", adjudicacionCampos);
    if (!valorVacio(f.UrlActaAdjudicacion)) {
      adjudicacion +=
        '<div class="ficha__section">' +
        '<a class="ficha__acta" href="' + escapeHtml(f.UrlActaAdjudicacion) + '" target="_blank" rel="noopener">Ver acta de adjudicación en Mercado Público ↗</a>' +
        '</div>';
    }

    return header + descripcion + general + contacto + fechas + condiciones + items + adjudicacion;
  }

  /* ---------- Ciclo de vida del modal ---------- */

  function inicializar() {
    return fetch(RUTA_FRAGMENTO)
      .then((resp) => {
        if (!resp.ok) throw new Error("HTTP " + resp.status);
        return resp.text();
      })
      .then((html) => {
        document.getElementById("fichaRoot").innerHTML = html;
        const overlay = document.getElementById("fichaOverlay");
        document.getElementById("fichaCerrar").addEventListener("click", cerrar);
        overlay.addEventListener("click", (e) => {
          if (e.target === overlay) cerrar();
        });
        document.addEventListener("keydown", (e) => {
          if (e.key === "Escape" && !overlay.classList.contains("hidden")) cerrar();
        });
        listo = true;
      })
      .catch((err) => {
        console.error("No se pudo cargar el modulo de ficha (" + RUTA_FRAGMENTO + ")", err);
      });
  }

  function abrir(fila) {
    if (!listo) {
      // Si el usuario hace clic antes de que termine de cargar ficha.html, reintenta al terminar.
      inicializar().then(() => abrir(fila));
      return;
    }
    if (!fila) return;
    document.getElementById("fichaContenido").innerHTML = renderFicha(fila);
    document.getElementById("fichaOverlay").classList.remove("hidden");
    document.body.style.overflow = "hidden";
  }

  function cerrar() {
    const overlay = document.getElementById("fichaOverlay");
    if (overlay) overlay.classList.add("hidden");
    document.body.style.overflow = "";
  }

  document.addEventListener("DOMContentLoaded", inicializar);

  const api = { abrir, cerrar };

  // Gancho solo para pruebas automatizadas con Node (no existe `module` en el navegador).
  if (typeof module !== "undefined") {
    module.exports = Object.assign({}, api, { renderFicha, campoFicha, seccionFicha, valorVacio });
  }

  return api;
})();
