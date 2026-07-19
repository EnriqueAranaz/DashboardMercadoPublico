/* ============================================================
   Cuadro de mando de licitaciones — SocIA Tech
   Prototipo: todo corre en el navegador, sin backend.
   Lee dos Excel (SheetJS) exportados desde el notebook de Python:
     - licitaciones_todas_5dias.xlsx  -> panel izquierdo
     - licitaciones_dashboard.xlsx    -> panel derecho (filtrado + ficha)

   La ficha de detalle (modal al hacer clic en una tarjeta) vive en su
   propio modulo: ficha.html / ficha.css / ficha.js. Este archivo solo
   llama a `Ficha.abrir(fila)`, no conoce como se pinta la ficha.
   ============================================================ */

const KEYWORDS_DASHBOARD = ["jardin", "jardines", "area verde", "areas verdes", "paisajismo"];

const state = {
  todos: [],        // filas del excel "todos los proyectos"
  filtrados: [],     // filas del excel "proyectos filtrados"
  estadosTodosSeleccionados: new Set(),   // filtro de estado, panel izquierdo
  filtros: {
    palabrasClave: new Set(),   // seleccionadas = activas (vacio al iniciar = todas)
    regiones: new Set(),
    estados: new Set(),
    montoMin: null,
    montoMax: null,
    incluirSinMonto: true,
  },
};

/* ---------- Utilidades ---------- */

function normalizarTexto(texto) {
  const combining = new RegExp(
    String.fromCharCode(0x5b, 0x5c, 0x75, 0x30, 0x33, 0x30, 0x30, 0x2d, 0x5c, 0x75, 0x30, 0x33, 0x36, 0x66, 0x5d),
    "g"
  );
  return (texto || "")
    .toString()
    .normalize("NFD")
    .replace(combining, "")
    .toLowerCase()
    .trim();
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

function formatearFecha(valor) {
  if (!valor) return "—";
  const d = new Date(valor);
  if (isNaN(d.getTime())) return String(valor);
  return d.toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric" });
}

const MESES_ABREV = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

// FechaConsulta viene como "ddmmaaaa" (ej. "13072026"), no es un formato
// que Date() entienda solo — la parseamos a mano.
function parsearDDMMAAAA(valor) {
  const s = String(valor || "").trim();
  const m = s.match(/^(\d{2})(\d{2})(\d{4})$/);
  if (!m) return null;
  return { dia: m[1], mes: parseInt(m[2], 10), anio: m[3] };
}

// Formato compacto para las tarjetas de resumen por día: "13-Jul"
function formatearFechaCortaDDMMAAAA(valor) {
  const p = parsearDDMMAAAA(valor);
  if (!p) return formatearFecha(valor);
  return p.dia + "-" + MESES_ABREV[p.mes];
}

// Formato para cada línea de proyecto: "13-Jul-26"
function formatearFechaDDMMAAAA(valor) {
  const p = parsearDDMMAAAA(valor);
  if (!p) return formatearFecha(valor);
  return p.dia + "-" + MESES_ABREV[p.mes] + "-" + p.anio.slice(-2);
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

// Misma paleta que claseEstado, pero como variante de chip (para los filtros de Estado)
function claseChipEstado(estado) {
  const e = normalizarTexto(estado);
  if (e.includes("public")) return "chip--publicada";
  if (e.includes("cerrad")) return "chip--cerrada";
  if (e.includes("adjudic")) return "chip--adjudicada";
  if (e.includes("revocad")) return "chip--revocada";
  if (e.includes("desiert")) return "chip--desierta";
  return "chip--otro";
}

function palabrasClaveDe(fila) {
  if (fila.PalabrasClave) {
    return String(fila.PalabrasClave)
      .split(";")
      .map((s) => s.trim())
      .filter(Boolean);
  }
  const texto = normalizarTexto(fila.Nombre);
  return KEYWORDS_DASHBOARD.filter((k) => texto.includes(normalizarTexto(k)));
}

/* ---------- Carga de datos ---------- */

const RUTA_TODOS = "licitaciones_todas_5dias.xlsx";
const RUTA_FILTRADOS = "licitaciones_dashboard.xlsx";

function cargarExcelDesdeUrl(url, callback, onError) {
  fetch(url)
    .then((resp) => {
      if (!resp.ok) throw new Error("HTTP " + resp.status);
      return resp.arrayBuffer();
    })
    .then((buffer) => {
      const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
      const primeraHoja = workbook.SheetNames[0];
      const filas = XLSX.utils.sheet_to_json(workbook.Sheets[primeraHoja], { defval: "" });
      callback(filas);
    })
    .catch((err) => {
      console.error("Error cargando " + url, err);
      if (onError) onError(err);
    });
}

function cargarPanelTodos() {
  cargarExcelDesdeUrl(
    RUTA_TODOS,
    (filas) => {
      state.todos = filas;
      document.getElementById("todosVacio").classList.add("hidden");
      document.getElementById("todosContenido").classList.remove("hidden");
      renderChipsEstadoTodos();
      renderStatsPorDia();
      renderListaTodos();
      document.getElementById("buscarTodos").addEventListener("input", () => {
        renderStatsPorDia();
        renderListaTodos();
      });
    },
    () => {
      document.getElementById("todosVacio").innerHTML =
        '<p>No se pudo cargar <code>' + RUTA_TODOS + '</code>. Verifica que el archivo esté en la carpeta del dashboard y que lo estés abriendo desde un servidor local (no con doble clic).</p>';
    }
  );
}

function cargarPanelFiltrados() {
  cargarExcelDesdeUrl(
    RUTA_FILTRADOS,
    (filas) => {
      state.filtrados = filas;
      document.getElementById("filtradosVacio").classList.add("hidden");
      document.getElementById("filtradosContenido").classList.remove("hidden");
      inicializarFiltros();
      aplicarFiltrosYRenderizar();
    },
    () => {
      document.getElementById("filtradosVacio").innerHTML =
        '<p>No se pudo cargar <code>' + RUTA_FILTRADOS + '</code>. Verifica que el archivo esté en la carpeta del dashboard y que lo estés abriendo desde un servidor local (no con doble clic).</p>';
    }
  );
}

document.addEventListener("DOMContentLoaded", () => {
  cargarPanelTodos();
  cargarPanelFiltrados();
});

/* ---------- Panel izquierdo: todos los proyectos ---------- */

function filasTodosFiltradas() {
  const busqueda = normalizarTexto(document.getElementById("buscarTodos").value);
  return state.todos.filter((f) => {
    if (busqueda && !normalizarTexto(f.Nombre).includes(busqueda)) return false;
    if (!state.estadosTodosSeleccionados.has(f.Estado)) return false;
    return true;
  });
}

function renderStatsPorDia() {
  const cont = document.getElementById("statsPorDia");
  const filas = filasTodosFiltradas();
  const porDia = {};
  filas.forEach((fila) => {
    const fecha = formatearFechaCortaDDMMAAAA(fila.FechaConsulta) || formatearFecha(fila.FechaPublicacion);
    porDia[fecha] = (porDia[fecha] || 0) + 1;
  });

  let html = '<div class="stat-card stat-card--total"><div class="stat-card__value">' +
    filas.length + '</div><div class="stat-card__label">Total 5 días</div></div>';

  Object.keys(porDia).forEach((fecha) => {
    html += '<div class="stat-card"><div class="stat-card__value">' + porDia[fecha] +
      '</div><div class="stat-card__label">' + escapeHtml(fecha) + '</div></div>';
  });

  cont.innerHTML = html;
}

function renderChipsEstadoTodos() {
  const wrap = document.getElementById("chipsEstadoTodosWrap");
  const estados = Array.from(new Set(state.todos.map((f) => f.Estado).filter(Boolean))).sort();

  if (!estados.length) {
    wrap.classList.add("hidden");
    return;
  }
  wrap.classList.remove("hidden");

  // Por defecto, todos los estados quedan marcados (= se listan todos los proyectos)
  if (state.estadosTodosSeleccionados.size === 0) {
    estados.forEach((e) => state.estadosTodosSeleccionados.add(e));
  }

  renderChips(
    "chipsEstadoTodos",
    estados,
    state.estadosTodosSeleccionados,
    (valor) => {
      toggleSetValor(state.estadosTodosSeleccionados, valor);
      renderStatsPorDia();
      renderListaTodos();
    },
    claseChipEstado
  );
}

function renderListaTodos() {
  const cont = document.getElementById("listaTodos");
  const filas = filasTodosFiltradas();

  if (!filas.length) {
    cont.innerHTML = '<div class="list-empty">No hay proyectos que coincidan con la búsqueda o los filtros.</div>';
    return;
  }

  cont.innerHTML = filas
    .map((f) => {
      const fecha = formatearFechaDDMMAAAA(f.FechaConsulta) || formatearFecha(f.FechaPublicacion);
      const estadoHtml = f.Estado
        ? '<span class="estado-badge ' + claseEstado(f.Estado) + '">' + escapeHtml(f.Estado) + '</span>'
        : "";
      return (
        '<div class="list-item">' +
        '<div class="list-item__nombre">' + escapeHtml(f.Nombre) + '</div>' +
        '<div class="list-item__meta">' +
        estadoHtml +
        '<span class="day-badge">' + escapeHtml(fecha) + '</span>' +
        '<span class="list-item__fecha">Cierre: ' + formatearFecha(f.FechaCierre) + '</span>' +
        '</div>' +
        '</div>'
      );
    })
    .join("");
}

/* ---------- Panel derecho: filtros y resultados ---------- */

// Si un set de filtro llega vacio (primera carga, o justo despues de "Limpiar
// filtros"), se rellena con todos los valores disponibles: por convencion,
// todos los filtros arrancan con TODO marcado (= sin exclusiones).
function marcarTodosPorDefecto(set, valores) {
  if (set.size === 0) {
    valores.forEach((v) => set.add(v));
  }
}

function inicializarFiltros() {
  marcarTodosPorDefecto(state.filtros.palabrasClave, KEYWORDS_DASHBOARD);
  renderChips(
    "chipsPalabraClave",
    KEYWORDS_DASHBOARD,
    state.filtros.palabrasClave,
    (valor) => {
      toggleSetValor(state.filtros.palabrasClave, valor);
      aplicarFiltrosYRenderizar();
    }
  );

  const regiones = Array.from(new Set(state.filtrados.map((f) => f.Region).filter(Boolean))).sort();
  marcarTodosPorDefecto(state.filtros.regiones, regiones);
  renderChips("chipsRegion", regiones, state.filtros.regiones, (valor) => {
    toggleSetValor(state.filtros.regiones, valor);
    aplicarFiltrosYRenderizar();
  });

  const estados = Array.from(new Set(state.filtrados.map((f) => f.Estado).filter(Boolean))).sort();
  marcarTodosPorDefecto(state.filtros.estados, estados);
  renderChips(
    "chipsEstado",
    estados,
    state.filtros.estados,
    (valor) => {
      toggleSetValor(state.filtros.estados, valor);
      aplicarFiltrosYRenderizar();
    },
    claseChipEstado
  );

  document.getElementById("montoMin").addEventListener("input", (e) => {
    state.filtros.montoMin = e.target.value ? Number(e.target.value) : null;
    aplicarFiltrosYRenderizar();
  });
  document.getElementById("montoMax").addEventListener("input", (e) => {
    state.filtros.montoMax = e.target.value ? Number(e.target.value) : null;
    aplicarFiltrosYRenderizar();
  });
  document.getElementById("incluirSinMonto").addEventListener("change", (e) => {
    state.filtros.incluirSinMonto = e.target.checked;
    aplicarFiltrosYRenderizar();
  });
  document.getElementById("resetFiltros").addEventListener("click", () => {
    state.filtros.palabrasClave.clear();
    state.filtros.regiones.clear();
    state.filtros.estados.clear();
    state.filtros.montoMin = null;
    state.filtros.montoMax = null;
    state.filtros.incluirSinMonto = true;
    document.getElementById("montoMin").value = "";
    document.getElementById("montoMax").value = "";
    document.getElementById("incluirSinMonto").checked = true;
    inicializarFiltros();
    aplicarFiltrosYRenderizar();
  });
}

function toggleSetValor(set, valor) {
  if (set.has(valor)) set.delete(valor);
  else set.add(valor);
}

function renderChips(contenedorId, valores, seleccionados, onClick, claseFn) {
  const cont = document.getElementById(contenedorId);
  cont.innerHTML = valores
    .map((v) => {
      const activo = seleccionados.has(v) ? " is-active" : "";
      const colorClase = claseFn ? " " + claseFn(v) : "";
      return '<span class="chip' + colorClase + activo + '" data-valor="' + escapeHtml(v) + '">' + escapeHtml(v) + "</span>";
    })
    .join("");

  cont.querySelectorAll(".chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      onClick(chip.dataset.valor);
      chip.classList.toggle("is-active");
    });
  });
}

function proyectoPasaFiltros(fila) {
  const f = state.filtros;

  if (f.palabrasClave.size > 0) {
    const kws = palabrasClaveDe(fila);
    const coincide = kws.some((k) => f.palabrasClave.has(k));
    if (!coincide) return false;
  }

  if (f.regiones.size > 0 && !f.regiones.has(fila.Region)) return false;
  if (f.estados.size > 0 && !f.estados.has(fila.Estado)) return false;

  const monto = Number(fila.MontoEstimado);
  const sinMonto = !fila.MontoEstimado || isNaN(monto) || monto <= 0;

  if (sinMonto) {
    if (!f.incluirSinMonto) return false;
  } else {
    if (f.montoMin != null && monto < f.montoMin) return false;
    if (f.montoMax != null && monto > f.montoMax) return false;
  }

  return true;
}

function aplicarFiltrosYRenderizar() {
  const resultados = state.filtrados.filter(proyectoPasaFiltros);
  renderAggStats(resultados);
  renderListaFiltrados(resultados);
}

function renderAggStats(filas) {
  const cont = document.getElementById("aggStats");

  const conMonto = filas.filter((f) => {
    const m = Number(f.MontoEstimado);
    return f.MontoEstimado && !isNaN(m) && m > 0;
  });
  const sinMonto = filas.length - conMonto.length;
  const montoTotal = conMonto.reduce((acc, f) => acc + Number(f.MontoEstimado), 0);

  const html =
    '<div class="stat-card stat-card--total"><div class="stat-card__value">' + filas.length +
    '</div><div class="stat-card__label">Proyectos totales</div></div>' +

    '<div class="stat-card"><div class="stat-card__value">' + conMonto.length +
    '</div><div class="stat-card__label">Con monto informado</div>' +
    '<div class="stat-card__sub">' + formatearMonto(montoTotal) + '</div></div>' +

    '<div class="stat-card"><div class="stat-card__value">' + sinMonto +
    '</div><div class="stat-card__label">Sin monto informado</div></div>';

  cont.innerHTML = html;
}

function renderListaFiltrados(filas) {
  const cont = document.getElementById("listaFiltrados");

  if (!filas.length) {
    cont.innerHTML = '<div class="list-empty">Ningún proyecto coincide con los filtros seleccionados.</div>';
    return;
  }

  cont.innerHTML = filas
    .map((f, idx) => {
      const kws = palabrasClaveDe(f);
      const kwsHtml = kws.length
        ? '<div class="result-card__keywords">' +
          kws.map((k) => '<span class="keyword-tag">' + escapeHtml(k) + "</span>").join("") +
          "</div>"
        : "";

      return (
        '<div class="result-card" data-idx="' + idx + '" tabindex="0" role="button" ' +
        'aria-label="Ver ficha completa de ' + escapeHtml(f.Nombre) + '">' +
        '<div class="result-card__top">' +
        '<div class="result-card__nombre">' + escapeHtml(f.Nombre) + '</div>' +
        '<span class="estado-badge ' + claseEstado(f.Estado) + '">' + escapeHtml(f.Estado) + '</span>' +
        "</div>" +
        '<div class="result-card__meta">' +
        '<span><strong>Organismo:</strong> ' + escapeHtml(f.Organismo || "—") + '</span>' +
        '<span><strong>Región:</strong> ' + escapeHtml(f.Region || "—") + '</span>' +
        '<span><strong>Monto:</strong> ' + formatearMonto(f.MontoEstimado) + '</span>' +
        '<span><strong>Cierre:</strong> ' + formatearFecha(f.FechaCierre) + '</span>' +
        "</div>" +
        kwsHtml +
        "</div>"
      );
    })
    .join("");

  // Al hacer clic se delega por completo en el modulo "ficha" (ver ficha.js).
  cont.querySelectorAll(".result-card").forEach((card) => {
    const abrir = () => Ficha.abrir(filas[Number(card.dataset.idx)]);
    card.addEventListener("click", abrir);
    card.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        abrir();
      }
    });
  });
}
