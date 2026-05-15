import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const SK_USUARIOS = "senae_usuarios_v3";
const SK_PAGO = "senae_pago_v3";
const SK_GARANTIA = "senae_garantia_v3";

const SK_CODIGOS = "senae_codigos_acceso_v1";
const CODIGOS_DEFAULT = ["SENAE#JCO2026"];

const USUARIOS_DEFAULT = [
  {
    id: "u1",
    usuario: "analista1",
    contrasena: "senae2026",
    nombre: "Ana Gomez",
    rol: "Analista",
  },
  {
    id: "u2",
    usuario: "analista2",
    contrasena: "senae2026",
    nombre: "Carlos Vera",
    rol: "Analista",
  },
  {
    id: "u3",
    usuario: "analista3",
    contrasena: "senae2026",
    nombre: "Maria Torres",
    rol: "Analista",
  },
  {
    id: "u4",
    usuario: "jefe",
    contrasena: "jefe2026",
    nombre: "Luis Andrade",
    rol: "Jefe JCO",
  },
  {
    id: "u5",
    usuario: "admin",
    contrasena: "admin2026",
    nombre: "Administrador",
    rol: "Admin",
  },
];

const generateId = () =>
  `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;

async function sharedGet(key: string, def: any) {
  try {
    // @ts-ignore
    const r = await window.storage.get(key, true);
    if (r && r.value) {
      const parsed = JSON.parse(r.value);
      if (parsed !== null && parsed !== undefined) {
        localStorage.setItem(key, r.value);
        return parsed;
      }
    }
  } catch {}
  try {
    const local = localStorage.getItem(key);
    if (local) return JSON.parse(local);
  } catch {}
  return def;
}

async function sharedSet(key: string, val: any) {
  const str = JSON.stringify(val);
  try {
    localStorage.setItem(key, str);
  } catch {}
  try {
    // @ts-ignore
    await window.storage.set(key, str, true);
  } catch {}
}

const EMPTY_PAGO = {
  id: "",
  tipo: "",
  ciudad: "",
  memorando: "",
  fechaRecepcion: "",
  fechaEntregaAnalista: "",
  tipoContratacion: "",
  contratoNo: "",
  certificacionPresupuestaria: "",
  item: "",
  descripcion: "",
  ruc: "",
  proveedorBeneficiario: "",
  periodo: "",
  fact: "",
  valor: "",
  curDev: "",
  fechaRecibeJCO: "",
  fechaDePago: "",
  analistaPresupuesto: "",
  estadoTramite: "Archivo",
  novedades: "",
  fechaDevolucion: "",
  noMemorandoFinal: "",
  creadoPor: "",
  ultimaModif: "",
  modificadoPor: "",
};

const EMPTY_GARANTIA = {
  id: "",
  tipoPago: "",
  direccion: "",
  memorando: "",
  fechaIngreso: "",
  fechaEntregaAnalista: "",
  descripcion: "",
  nroLiquidacion: "",
  beneficiario: "",
  valor: "",
  curContable: "",
  fechaRecibeJCO: "",
  fechaRecibeJAC: "",
  fechaSolPago: "",
  servidorPresupuesto: "",
  estadoTramite: "Archivo",
  novedades: "",
  fechaDevolucion: "",
  nroMemorandum: "",
  creadoPor: "",
  ultimaModif: "",
  modificadoPor: "",
};

const ESTADOS_PAGO = [
  "Archivo",
  "DFI",
  "Factura",
  "Contabilidad",
  "Devuelto",
  "En revision",
  "Observado",
  "Presupuesto",
];
const ESTADOS_GARANTIA = [
  "Archivo",
  "Factura",
  "Contabilidad",
  "Devuelto",
  "En revision",
  "Observado",
  "Presupuesto",
];
const ESTADOS_ATENDIDOS = ["Archivo", "DFI", "Contabilidad"];
const ROLES = ["Analista", "Jefe JCO", "Admin"];

const ESTADO_COLORS = {
  Archivo: { bg: "#F8FAFC", color: "#475569", border: "#CBD5E1" },
  DFI: { bg: "#FFF7ED", color: "#C2410C", border: "#FDBA74" },
  Factura: { bg: "#ECFDF5", color: "#065F46", border: "#6EE7B7" },
  Contabilidad: { bg: "#EFF6FF", color: "#1D4ED8", border: "#93C5FD" },
  Devuelto: { bg: "#FFF1F2", color: "#BE123C", border: "#FDA4AF" },
  "En revision": { bg: "#FFFBEB", color: "#B45309", border: "#FCD34D" },
  Observado: { bg: "#FDF4FF", color: "#7E22CE", border: "#E9D5FF" },
  Presupuesto: { bg: "#F0FDF4", color: "#15803D", border: "#86EFAC" },
};

// ─── MAPEO DE COLUMNAS EXCEL → CAMPOS INTERNOS ───────────────────────────────
// Cada entrada: [campo_interno, [...alias_posibles_en_excel]]
// Se normalizan a minusculas sin tildes para comparacion flexible
const COLUMN_MAP_PAGO = [
  ["tipo", ["tipo", "type"]],
  ["ciudad", ["ciudad", "city"]],
  [
    "memorando",
    ["memorando", "memorandum", "memo", "no memorando", "no. memorando"],
  ],
  [
    "fechaRecepcion",
    ["fecha recepcion", "fecha de recepcion", "recepcion", "fecha recep"],
  ],
  [
    "fechaEntregaAnalista",
    ["fecha entrega analista", "entrega analista", "fecha entrega al analista"],
  ],
  [
    "tipoContratacion",
    ["tipo contratacion", "tipo de contratacion", "contratacion"],
  ],
  [
    "contratoNo",
    ["contrato no", "contrato no.", "no contrato", "no. contrato", "contrato"],
  ],
  [
    "certificacionPresupuestaria",
    [
      "certificacion presupuestaria",
      "cert presupuestaria",
      "certificacion",
      "cert. presupuestaria",
    ],
  ],
  ["item", ["item", "item presupuestario"]],
  ["descripcion", ["descripcion", "description", "desc"]],
  ["ruc", ["ruc", "ruc proveedor"]],
  [
    "proveedorBeneficiario",
    [
      "proveedor beneficiario",
      "proveedor / beneficiario",
      "proveedor",
      "beneficiario proveedor",
    ],
  ],
  ["periodo", ["periodo", "period"]],
  ["fact", ["fact", "factura", "no factura", "no. factura"]],
  ["valor", ["valor", "value", "monto", "importe"]],
  ["curDev", ["cur dev", "cur/dev", "cur", "dev", "cur-dev"]],
  ["fechaRecibeJCO", ["fecha recibe jco", "recibe jco", "fecha jco"]],
  ["fechaDePago", ["fecha de pago", "fecha pago", "pago fecha"]],
  [
    "analistaPresupuesto",
    ["analista presupuesto", "analista de presupuesto", "analista"],
  ],
  [
    "estadoTramite",
    ["estado tramite", "estado del tramite", "estado", "status"],
  ],
  ["novedades", ["novedades", "observaciones", "novedad", "obs"]],
  [
    "fechaDevolucion",
    ["fecha devolucion", "fecha de devolucion", "devolucion"],
  ],
  [
    "noMemorandoFinal",
    [
      "no memorando final",
      "no. memorando final",
      "memorando final",
      "memo final",
    ],
  ],
  ["creadoPor", ["creado por", "created by", "creado"]],
  [
    "ultimaModif",
    ["ultima modif", "ultima modificacion", "ultima mod", "modif"],
  ],
  ["modificadoPor", ["modificado por", "modified by"]],
];

const COLUMN_MAP_GARANTIA = [
  ["tipoPago", ["tipo pago", "tipo de pago", "tipopago"]],
  ["direccion", ["direccion", "direccion", "direction", "dir"]],
  [
    "memorando",
    ["memorando", "memorandum", "memo", "no memorando", "no. memorando"],
  ],
  [
    "fechaIngreso",
    [
      "fecha ingreso",
      "fecha de ingreso",
      "ingreso",
      "fecha ing",
      "fecha  de ingreso",
    ],
  ],
  [
    "fechaEntregaAnalista",
    [
      "fecha entrega analista",
      "entrega analista",
      "fecha entrega al analista",
      "fecha de entrega al analista",
    ],
  ],
  ["descripcion", ["descripcion", "descripcion", "description", "desc"]],
  [
    "nroLiquidacion",
    [
      "nro liquidacion",
      "nro. liquidacion",
      "liquidacion",
      "no liquidacion",
      "nro. liquidacion",
    ],
  ],
  ["beneficiario", ["beneficiario", "beneficiary"]],
  ["valor", ["valor", "value", "monto", "importe", " valor "]],
  ["curContable", ["cur contable", "curcontable", "contable cur"]],
  ["fechaRecibeJCO", ["fecha recibe jco", "recibe jco", "fecha jco"]],
  ["fechaRecibeJAC", ["fecha recibe jac", "recibe jac", "fecha jac"]],
  [
    "fechaSolPago",
    [
      "fecha sol pago",
      "fecha sol. pago",
      "sol pago",
      "solicitud pago",
      "fecha de  sol. pago",
      "fecha de sol. pago",
      "fecha sol. de pago",
    ],
  ],
  [
    "servidorPresupuesto",
    ["servidor presupuesto", "servidor de presupuesto", "servidor"],
  ],
  [
    "estadoTramite",
    ["estado tramite", "estado del tramite", "estado", "status"],
  ],
  ["novedades", ["novedades", "observaciones", "novedad", "obs"]],
  [
    "fechaDevolucion",
    [
      "fecha devolucion",
      "fecha de devolucion",
      "devolucion",
      "fecha de devolucion",
    ],
  ],
  [
    "nroMemorandum",
    [
      "nro memorandum",
      "nro. memorandum",
      "memorandum",
      "nro memo",
      "nro. memo",
      "nro. memorando",
      "nro memorando",
    ],
  ],
];

function normalizeStr(s) {
  if (!s) return "";
  return String(s)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function mapExcelRowToRecord(row, colMap, emptyState) {
  const record = { ...emptyState, id: generateId() };
  const rowNorm = {};
  Object.keys(row).forEach((k) => {
    rowNorm[normalizeStr(k)] = row[k];
  });
  colMap.forEach(([campo, aliases]) => {
    for (const alias of aliases) {
      const val = rowNorm[alias];
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        let strVal = String(val).trim();

        // ── CASO 1: fecha como objeto Date (cuando cellDates: true) ──
        if (val instanceof Date && campo.toLowerCase().includes("fecha")) {
          const d = val.getDate();
          const m = val.getMonth() + 1;
          const y = val.getFullYear();
          strVal = `${d}/${m}/${y}`;
        }
        // ── CASO 2: fecha como numero serial de Excel ──
        else if (
          typeof val === "number" &&
          campo.toLowerCase().includes("fecha")
        ) {
          try {
            const date = XLSX.SSF.parse_date_code(val);
            if (date) {
              strVal = `${date.d}/${date.m}/${date.y}`;
            }
          } catch {}
        }

        // Normalizar estado
        if (campo === "estadoTramite") {
          const allStates = [...ESTADOS_PAGO, ...ESTADOS_GARANTIA];
          const match = allStates.find(
            (e) => normalizeStr(e) === normalizeStr(strVal)
          );
          strVal = match || "Archivo";
        }

        record[campo] = strVal;
        break;
      }
    }
  });
  return record;
}

// ─── MODAL IMPORTAR EXCEL ─────────────────────────────────────────────────────
function ImportModal({
  onClose,
  onImportar,
  nombreHoja,
  colMap,
  emptyState,
  estadosValidos,
}) {
  const [fase, setFase] = useState("upload"); // upload | preview | done
  const [registrosPreview, setRegistrosPreview] = useState([]);
  const [modoImport, setModoImport] = useState("agregar"); // agregar | reemplazar
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [colsDetectadas, setColsDetectadas] = useState([]);
  const fileRef = useRef();

  const handleFile = (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls", "csv"].includes(ext)) {
      setError("Solo se aceptan archivos .xlsx, .xls o .csv");
      return;
    }
    setError("");
    setLoading(true);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {
          type: "array",
          cellDates: true,
        });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (rawRows.length === 0) {
          setError("El archivo no contiene datos o la hoja esta vacia.");
          setLoading(false);
          return;
        }
        // Detectar columnas del archivo
        const cols = Object.keys(rawRows[0]);
        setColsDetectadas(cols);
        // Mapear registros
        const mapped = rawRows
          .filter((row) => {
            // Descartar filas completamente vacias
            return Object.values(row).some((v) => String(v).trim() !== "");
          })
          .map((row) => mapExcelRowToRecord(row, colMap, emptyState));
        if (mapped.length === 0) {
          setError("No se encontraron filas con datos validos.");
          setLoading(false);
          return;
        }
        setRegistrosPreview(mapped);
        setFase("preview");
      } catch (err) {
        setError(
          "Error al leer el archivo. Asegurate de que es un Excel valido."
        );
      }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const confirmarImport = async () => {
    setLoading(true);
    const now = new Date().toLocaleString("es-EC");
    const registrosFinal = registrosPreview.map((r) => ({
      ...r,
      creadoPor: r.creadoPor || "Importacion SENAE",
      ultimaModif: r.ultimaModif || now,
    }));
    await onImportar(registrosFinal, modoImport);
    setFase("done");
    setLoading(false);
  };

  const estadoCount = estadosValidos.reduce((acc, e) => {
    acc[e] = registrosPreview.filter((r) => r.estadoTramite === e).length;
    return acc;
  }, {});
  const sinMemorando = registrosPreview.filter(
    (r) => !r.memorando?.trim()
  ).length;
  const conValor = registrosPreview.filter((r) => r.valor?.trim()).length;

  return (
    <div
      style={S.overlay}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div style={{ ...S.modal, maxWidth: 640 }}>
        {/* Header */}
        <div style={{ ...S.modalHeader, background: "#065F46" }}>
          <div>
            <div style={S.modalTitle}>Importar desde Excel</div>
            <div style={S.modalSub}>
              Hoja: {nombreHoja} — Carga masiva de registros SENAE
            </div>
          </div>
          <button style={S.modalClose} onClick={onClose}>
            <Ic.Close />
          </button>
        </div>

        {/* Barra de pasos */}
        <div
          style={{
            background: "#F8FAFC",
            borderBottom: "1px solid #E2E8F0",
            padding: "12px 24px",
            display: "flex",
            alignItems: "center",
            gap: 0,
          }}
        >
          {[
            { key: "upload", label: "1. Subir archivo" },
            { key: "preview", label: "2. Vista previa" },
            { key: "done", label: "3. Listo" },
          ].map((step, i) => {
            const isActive = fase === step.key;
            const isPast = (fase === "preview" && i === 0) || fase === "done";
            return (
              <React.Fragment key={step.key}>
                {i > 0 && (
                  <div
                    style={{
                      width: 24,
                      height: 2,
                      background: isPast ? "#065F46" : "#E2E8F0",
                      flexShrink: 0,
                    }}
                  />
                )}
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <div
                    style={{
                      width: 22,
                      height: 22,
                      borderRadius: "50%",
                      fontSize: 11,
                      fontWeight: 800,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      flexShrink: 0,
                      background: isActive
                        ? "#065F46"
                        : isPast
                        ? "#ECFDF5"
                        : "#E2E8F0",
                      color: isActive
                        ? "#FFFFFF"
                        : isPast
                        ? "#065F46"
                        : "#94A3B8",
                      border: isActive
                        ? "2px solid #065F46"
                        : isPast
                        ? "2px solid #6EE7B7"
                        : "2px solid #E2E8F0",
                    }}
                  >
                    {isPast && !isActive ? (
                      <svg
                        width="11"
                        height="11"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="3"
                      >
                        <polyline points="20 6 9 17 4 12" />
                      </svg>
                    ) : (
                      i + 1
                    )}
                  </div>
                  <span
                    style={{
                      fontSize: 12,
                      fontWeight: isActive ? 700 : 500,
                      color: isActive ? "#065F46" : "#94A3B8",
                    }}
                  >
                    {step.label}
                  </span>
                </div>
              </React.Fragment>
            );
          })}
        </div>

        <div style={{ padding: "24px", maxHeight: "60vh", overflowY: "auto" }}>
          {/* ── FASE UPLOAD ── */}
          {fase === "upload" && (
            <div>
              <div
                onDrop={handleDrop}
                onDragOver={(e) => e.preventDefault()}
                onClick={() => fileRef.current.click()}
                style={{
                  border: "2.5px dashed #6EE7B7",
                  borderRadius: 14,
                  padding: "40px 24px",
                  textAlign: "center",
                  cursor: "pointer",
                  background: "#F0FDF4",
                  transition: "all 0.2s",
                }}
              >
                <input
                  ref={fileRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  style={{ display: "none" }}
                  onChange={(e) => handleFile(e.target.files[0])}
                />
                <div style={{ marginBottom: 16 }}>
                  <svg
                    width="48"
                    height="48"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="#059669"
                    strokeWidth="1.5"
                    style={{ margin: "0 auto", display: "block" }}
                  >
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                    <polyline points="14 2 14 8 20 8" />
                    <line x1="12" y1="18" x2="12" y2="12" />
                    <polyline points="9 15 12 12 15 15" />
                  </svg>
                </div>
                <div
                  style={{
                    fontSize: 15,
                    fontWeight: 700,
                    color: "#065F46",
                    marginBottom: 6,
                  }}
                >
                  Arrastra tu archivo Excel aqui o haz clic para seleccionar
                </div>
                <div style={{ fontSize: 13, color: "#64748B" }}>
                  Formatos aceptados: .xlsx, .xls, .csv
                </div>
                {loading && (
                  <div
                    style={{
                      marginTop: 12,
                      fontSize: 13,
                      color: "#059669",
                      fontWeight: 600,
                    }}
                  >
                    Procesando archivo...
                  </div>
                )}
              </div>

              {error && (
                <div
                  style={{
                    marginTop: 12,
                    padding: "10px 14px",
                    borderRadius: 8,
                    background: "#FFF1F2",
                    border: "1px solid #FDA4AF",
                    color: "#BE123C",
                    fontSize: 13,
                    fontWeight: 600,
                  }}
                >
                  {error}
                </div>
              )}

              {/* Guia de columnas esperadas */}
              <div
                style={{
                  marginTop: 20,
                  background: "#F8FAFC",
                  border: "1.5px solid #E2E8F0",
                  borderRadius: 12,
                  padding: 16,
                }}
              >
                <div
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#64748B",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                    marginBottom: 12,
                  }}
                >
                  Columnas reconocidas automaticamente para {nombreHoja}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {colMap.map(([campo, aliases]) => (
                    <div
                      key={campo}
                      style={{
                        background: "#EFF6FF",
                        border: "1px solid #BFDBFE",
                        borderRadius: 6,
                        padding: "3px 10px",
                        fontSize: 11,
                        color: "#1D4ED8",
                        fontWeight: 600,
                      }}
                    >
                      {aliases[0]}
                    </div>
                  ))}
                </div>
                <div
                  style={{
                    marginTop: 10,
                    fontSize: 11,
                    color: "#94A3B8",
                    lineHeight: 1.5,
                  }}
                >
                  El sistema reconoce variantes de los nombres (con o sin
                  tildes, mayusculas/minusculas). Las columnas no reconocidas se
                  ignoran sin error.
                </div>
              </div>
            </div>
          )}

          {/* ── FASE PREVIEW ── */}
          {fase === "preview" && (
            <div>
              {/* Resumen de la importacion */}
              <div
                style={{
                  background: "#ECFDF5",
                  border: "1.5px solid #6EE7B7",
                  borderRadius: 12,
                  padding: 16,
                  marginBottom: 20,
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    marginBottom: 12,
                  }}
                >
                  <svg
                    width="18"
                    height="18"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="#059669"
                    strokeWidth="2.5"
                  >
                    <polyline points="20 6 9 17 4 12" />
                  </svg>
                  <div
                    style={{ fontSize: 14, fontWeight: 800, color: "#065F46" }}
                  >
                    Archivo leido correctamente:{" "}
                    <span style={{ color: "#059669" }}>{fileName}</span>
                  </div>
                </div>
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr 1fr",
                    gap: 10,
                  }}
                >
                  {[
                    {
                      label: "Total filas detectadas",
                      val: registrosPreview.length,
                      color: "#1D4ED8",
                      bg: "#EFF6FF",
                      border: "#93C5FD",
                    },
                    {
                      label: "Con valor monetario",
                      val: conValor,
                      color: "#065F46",
                      bg: "#ECFDF5",
                      border: "#6EE7B7",
                    },
                    {
                      label: "Sin memorando",
                      val: sinMemorando,
                      color: sinMemorando > 0 ? "#B45309" : "#065F46",
                      bg: sinMemorando > 0 ? "#FFFBEB" : "#ECFDF5",
                      border: sinMemorando > 0 ? "#FCD34D" : "#6EE7B7",
                    },
                  ].map((c) => (
                    <div
                      key={c.label}
                      style={{
                        background: c.bg,
                        border: `1.5px solid ${c.border}`,
                        borderRadius: 10,
                        padding: "12px 10px",
                        textAlign: "center",
                      }}
                    >
                      <div
                        style={{
                          fontSize: 24,
                          fontWeight: 800,
                          color: c.color,
                        }}
                      >
                        {c.val}
                      </div>
                      <div
                        style={{
                          fontSize: 10,
                          fontWeight: 700,
                          color: c.color,
                          textTransform: "uppercase",
                          letterSpacing: "0.5px",
                          marginTop: 2,
                        }}
                      >
                        {c.label}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Columnas del Excel detectadas */}
              <div style={{ marginBottom: 16 }}>
                <div
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#64748B",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                    marginBottom: 8,
                  }}
                >
                  Columnas detectadas en el archivo ({colsDetectadas.length})
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                  {colsDetectadas.map((col) => {
                    const mapped = colMap.find(([, aliases]) =>
                      aliases.some((a) => normalizeStr(a) === normalizeStr(col))
                    );
                    return (
                      <span
                        key={col}
                        style={{
                          padding: "3px 10px",
                          borderRadius: 6,
                          fontSize: 11,
                          fontWeight: 600,
                          background: mapped ? "#ECFDF5" : "#F8FAFC",
                          border: `1px solid ${mapped ? "#6EE7B7" : "#E2E8F0"}`,
                          color: mapped ? "#065F46" : "#94A3B8",
                        }}
                      >
                        {col} {mapped ? "" : "(ignorada)"}
                      </span>
                    );
                  })}
                </div>
              </div>

              {/* Distribucion por estado */}
              <div style={{ marginBottom: 16 }}>
                <div
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#64748B",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                    marginBottom: 8,
                  }}
                >
                  Distribucion por estado
                </div>
                <div
                  style={{ display: "flex", flexDirection: "column", gap: 5 }}
                >
                  {estadosValidos.map((est) => {
                    const col = ESTADO_COLORS[est] || ESTADO_COLORS["Archivo"];
                    const count = estadoCount[est] || 0;
                    const pct =
                      registrosPreview.length > 0
                        ? Math.round((count / registrosPreview.length) * 100)
                        : 0;
                    return (
                      <div
                        key={est}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                        }}
                      >
                        <span
                          style={{
                            width: 8,
                            height: 8,
                            borderRadius: "50%",
                            background: col.color,
                            flexShrink: 0,
                            display: "inline-block",
                          }}
                        />
                        <span
                          style={{
                            fontSize: 12,
                            fontWeight: 600,
                            color: "#475569",
                            minWidth: 110,
                          }}
                        >
                          {est}
                        </span>
                        <div
                          style={{
                            flex: 1,
                            background: "#F1F5F9",
                            borderRadius: 99,
                            height: 6,
                            overflow: "hidden",
                          }}
                        >
                          <div
                            style={{
                              width: `${pct}%`,
                              height: "100%",
                              background: col.color,
                              borderRadius: 99,
                              transition: "width 0.3s",
                            }}
                          />
                        </div>
                        <span
                          style={{
                            fontSize: 12,
                            fontWeight: 700,
                            color: col.color,
                            minWidth: 28,
                            textAlign: "right",
                          }}
                        >
                          {count}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Modo de importacion */}
              <div
                style={{
                  background: "#FFFBEB",
                  border: "1.5px solid #FCD34D",
                  borderRadius: 12,
                  padding: 16,
                }}
              >
                <div
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    color: "#B45309",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                    marginBottom: 12,
                  }}
                >
                  Modo de importacion
                </div>
                <div
                  style={{ display: "flex", flexDirection: "column", gap: 8 }}
                >
                  {[
                    {
                      val: "agregar",
                      title: "Agregar a los existentes",
                      desc: "Los registros actuales se conservan. Los importados se suman al final.",
                      icon: "plus",
                    },
                    {
                      val: "reemplazar",
                      title: "Reemplazar todo",
                      desc: "ATENCION: Se eliminan todos los registros actuales y se cargan solo los del archivo.",
                      icon: "replace",
                    },
                  ].map((op) => (
                    <button
                      key={op.val}
                      onClick={() => setModoImport(op.val)}
                      style={{
                        display: "flex",
                        alignItems: "flex-start",
                        gap: 12,
                        padding: "12px 14px",
                        borderRadius: 10,
                        cursor: "pointer",
                        textAlign: "left",
                        fontFamily: "inherit",
                        background:
                          modoImport === op.val
                            ? op.val === "reemplazar"
                              ? "#FFF1F2"
                              : "#EFF6FF"
                            : "#FFFFFF",
                        border:
                          modoImport === op.val
                            ? `2px solid ${
                                op.val === "reemplazar" ? "#FDA4AF" : "#93C5FD"
                              }`
                            : "2px solid #E2E8F0",
                        color: "#475569",
                      }}
                    >
                      <div
                        style={{
                          width: 28,
                          height: 28,
                          borderRadius: 8,
                          flexShrink: 0,
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          background:
                            modoImport === op.val
                              ? op.val === "reemplazar"
                                ? "#FFF1F2"
                                : "#EFF6FF"
                              : "#F8FAFC",
                          border: `1.5px solid ${
                            modoImport === op.val
                              ? op.val === "reemplazar"
                                ? "#FDA4AF"
                                : "#93C5FD"
                              : "#E2E8F0"
                          }`,
                          color:
                            modoImport === op.val
                              ? op.val === "reemplazar"
                                ? "#BE123C"
                                : "#1D4ED8"
                              : "#94A3B8",
                        }}
                      >
                        <svg
                          width="14"
                          height="14"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.2"
                        >
                          {op.icon === "plus" ? (
                            <>
                              <line x1="12" y1="5" x2="12" y2="19" />
                              <line x1="5" y1="12" x2="19" y2="12" />
                            </>
                          ) : (
                            <>
                              <polyline points="1 4 1 10 7 10" />
                              <path d="M3.51 15a9 9 0 1 0 .49-3.5" />
                            </>
                          )}
                        </svg>
                      </div>
                      <div>
                        <div
                          style={{
                            fontSize: 13,
                            fontWeight: 700,
                            color: "#1E293B",
                            marginBottom: 2,
                          }}
                        >
                          {op.title}
                        </div>
                        <div
                          style={{
                            fontSize: 11,
                            color: "#64748B",
                            lineHeight: 1.5,
                          }}
                        >
                          {op.desc}
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Mini tabla preview primeros 3 registros */}
              {registrosPreview.length > 0 && (
                <div style={{ marginTop: 16 }}>
                  <div
                    style={{
                      fontSize: 11,
                      fontWeight: 700,
                      color: "#64748B",
                      textTransform: "uppercase",
                      letterSpacing: "0.5px",
                      marginBottom: 8,
                    }}
                  >
                    Vista previa (primeros 3 registros)
                  </div>
                  <div
                    style={{
                      overflowX: "auto",
                      borderRadius: 8,
                      border: "1px solid #E2E8F0",
                    }}
                  >
                    <table
                      style={{
                        borderCollapse: "collapse",
                        width: "100%",
                        fontSize: 11,
                      }}
                    >
                      <thead>
                        <tr>
                          {[
                            "N",
                            "Memorando",
                            "Descripcion",
                            "Valor",
                            "Estado",
                          ].map((h) => (
                            <th
                              key={h}
                              style={{
                                background: "#003366",
                                color: "#fff",
                                padding: "8px 10px",
                                fontWeight: 700,
                                textAlign: "left",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {registrosPreview.slice(0, 3).map((r, i) => {
                          const ec =
                            ESTADO_COLORS[r.estadoTramite] ||
                            ESTADO_COLORS["Archivo"];
                          return (
                            <tr
                              key={r.id}
                              style={{
                                background: i % 2 === 0 ? "#fff" : "#F8FAFC",
                              }}
                            >
                              <td
                                style={{
                                  padding: "7px 10px",
                                  color: "#94A3B8",
                                  fontWeight: 700,
                                }}
                              >
                                {i + 1}
                              </td>
                              <td
                                style={{
                                  padding: "7px 10px",
                                  color: "#334155",
                                  maxWidth: 150,
                                  overflow: "hidden",
                                  textOverflow: "ellipsis",
                                  whiteSpace: "nowrap",
                                }}
                              >
                                {r.memorando || "-"}
                              </td>
                              <td
                                style={{
                                  padding: "7px 10px",
                                  color: "#334155",
                                  maxWidth: 180,
                                  overflow: "hidden",
                                  textOverflow: "ellipsis",
                                  whiteSpace: "nowrap",
                                }}
                              >
                                {r.descripcion || "-"}
                              </td>
                              <td
                                style={{
                                  padding: "7px 10px",
                                  color: "#334155",
                                  whiteSpace: "nowrap",
                                }}
                              >
                                {r.valor ? `$${r.valor}` : "-"}
                              </td>
                              <td style={{ padding: "7px 10px" }}>
                                <span
                                  style={{
                                    display: "inline-block",
                                    padding: "2px 8px",
                                    borderRadius: 20,
                                    fontSize: 10,
                                    fontWeight: 700,
                                    background: ec.bg,
                                    color: ec.color,
                                    border: `1px solid ${ec.border}`,
                                  }}
                                >
                                  {r.estadoTramite || "Archivo"}
                                </span>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  {registrosPreview.length > 3 && (
                    <div
                      style={{
                        textAlign: "center",
                        marginTop: 6,
                        fontSize: 11,
                        color: "#94A3B8",
                      }}
                    >
                      ...y {registrosPreview.length - 3} registro
                      {registrosPreview.length - 3 !== 1 ? "s" : ""} mas
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {/* ── FASE DONE ── */}
          {fase === "done" && (
            <div style={{ textAlign: "center", padding: "24px 0" }}>
              <div
                style={{
                  width: 64,
                  height: 64,
                  borderRadius: "50%",
                  background: "#ECFDF5",
                  border: "3px solid #6EE7B7",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  margin: "0 auto 20px",
                }}
              >
                <svg
                  width="32"
                  height="32"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="#059669"
                  strokeWidth="2.5"
                >
                  <polyline points="20 6 9 17 4 12" />
                </svg>
              </div>
              <div
                style={{
                  fontSize: 20,
                  fontWeight: 800,
                  color: "#065F46",
                  marginBottom: 8,
                }}
              >
                Importacion completada
              </div>
              <div style={{ fontSize: 14, color: "#64748B", marginBottom: 24 }}>
                Se cargaron{" "}
                <strong style={{ color: "#065F46" }}>
                  {registrosPreview.length} registros
                </strong>{" "}
                en la hoja de {nombreHoja}.
              </div>
              <div
                style={{
                  background: "#ECFDF5",
                  border: "1.5px solid #6EE7B7",
                  borderRadius: 12,
                  padding: "14px 20px",
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 10,
                  fontSize: 13,
                  color: "#065F46",
                  fontWeight: 600,
                }}
              >
                <svg
                  width="16"
                  height="16"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                >
                  <circle cx="12" cy="12" r="10" />
                  <line x1="12" y1="8" x2="12" y2="12" />
                  <line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
                Los datos ya estan disponibles para todos los usuarios
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={S.modalFooter}>
          {fase === "done" ? (
            <button style={S.btnGuardar} onClick={onClose}>
              Cerrar
            </button>
          ) : fase === "preview" ? (
            <>
              <button
                style={S.btnCancelarModal}
                onClick={() => {
                  setFase("upload");
                  setRegistrosPreview([]);
                  setFileName("");
                }}
              >
                Volver
              </button>
              <button
                style={{
                  ...S.btnGuardar,
                  background: "#065F46",
                  opacity: loading ? 0.6 : 1,
                }}
                onClick={confirmarImport}
                disabled={loading}
              >
                {loading
                  ? "Importando..."
                  : `Importar ${registrosPreview.length} registros`}
              </button>
            </>
          ) : (
            <button style={S.btnCancelarModal} onClick={onClose}>
              Cancelar
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

const CAMPOS_FORM_PAGO = [
  { key: "tipo", label: "TIPO", placeholder: "Ej: Pago, Anticipo..." },
  { key: "ciudad", label: "CIUDAD", placeholder: "Ej: Guayaquil" },
  {
    key: "memorando",
    label: "MEMORANDO",
    placeholder: "No. Memorando",
    required: true,
  },
  { key: "fechaRecepcion", label: "FECHA RECEPCION", type: "date" },
  {
    key: "fechaEntregaAnalista",
    label: "FECHA ENTREGA AL ANALISTA",
    type: "date",
  },
  {
    key: "tipoContratacion",
    label: "TIPO CONTRATACION",
    placeholder: "Ej: Menor Cuantia...",
  },
  { key: "contratoNo", label: "CONTRATO No.", placeholder: "No. de contrato" },
  {
    key: "certificacionPresupuestaria",
    label: "CERTIFICACION PRESUPUESTARIA",
    placeholder: "No. certificacion",
  },
  { key: "item", label: "ITEM", placeholder: "Item presupuestario" },
  {
    key: "descripcion",
    label: "DESCRIPCION",
    placeholder: "Descripcion del tramite",
  },
  { key: "ruc", label: "RUC", placeholder: "RUC del proveedor" },
  {
    key: "proveedorBeneficiario",
    label: "PROVEEDOR / BENEFICIARIO",
    placeholder: "Nombre o razon social",
  },
  { key: "periodo", label: "PERIODO", placeholder: "Ej: Enero 2026" },
  { key: "fact", label: "FACT", placeholder: "No. Factura" },
  { key: "valor", label: "VALOR", placeholder: "0.00" },
  { key: "curDev", label: "CUR/DEV", placeholder: "No. CUR/DEV" },
  { key: "fechaRecibeJCO", label: "FECHA RECIBE JCO", type: "date" },
  { key: "fechaDePago", label: "FECHA DE PAGO", type: "date" },
  {
    key: "analistaPresupuesto",
    label: "ANALISTA DE PRESUPUESTO",
    placeholder: "Nombre del analista",
  },
  {
    key: "novedades",
    label: "NOVEDADES",
    placeholder: "Observaciones o novedades",
  },
  { key: "fechaDevolucion", label: "FECHA DE DEVOLUCION", type: "date" },
  {
    key: "noMemorandoFinal",
    label: "No. MEMORANDO FINAL",
    placeholder: "No. memorando de cierre",
  },
];

const ENCABEZADOS_PAGO = [
  "N",
  "TIPO",
  "CIUDAD",
  "MEMORANDO",
  "FECHA RECEPCION",
  "FECHA ENTREGA ANALISTA",
  "TIPO CONTRATACION",
  "CONTRATO No.",
  "CERT. PRESUPUESTARIA",
  "ITEM",
  "DESCRIPCION",
  "RUC",
  "PROVEEDOR / BENEFICIARIO",
  "PERIODO",
  "FACT",
  "VALOR",
  "CUR/DEV",
  "FECHA RECIBE JCO",
  "FECHA DE PAGO",
  "ANALISTA PRESUPUESTO",
  "ESTADO",
  "NOVEDADES",
  "FECHA DEVOLUCION",
  "No. MEMORANDO FINAL",
  "CREADO POR",
  "ULTIMA MODIF.",
  "ACCIONES",
];
const CAMPOS_TABLA_PAGO = [
  "tipo",
  "ciudad",
  "memorando",
  "fechaRecepcion",
  "fechaEntregaAnalista",
  "tipoContratacion",
  "contratoNo",
  "certificacionPresupuestaria",
  "item",
  "descripcion",
  "ruc",
  "proveedorBeneficiario",
  "periodo",
  "fact",
  "valor",
  "curDev",
  "fechaRecibeJCO",
  "fechaDePago",
  "analistaPresupuesto",
  "estadoTramite",
  "novedades",
  "fechaDevolucion",
  "noMemorandoFinal",
  "creadoPor",
  "ultimaModif",
];

const CAMPOS_FORM_GARANTIA = [
  { key: "tipoPago", label: "TIPO PAGO", placeholder: "Tipo de pago" },
  {
    key: "direccion",
    label: "DIRECCION",
    placeholder: "Ej: Direccion Distrital",
  },
  {
    key: "memorando",
    label: "MEMORANDO",
    placeholder: "No. Memorando",
    required: true,
  },
  { key: "fechaIngreso", label: "FECHA DE INGRESO", type: "date" },
  {
    key: "fechaEntregaAnalista",
    label: "FECHA ENTREGA ANALISTA",
    type: "date",
  },
  {
    key: "descripcion",
    label: "DESCRIPCION",
    placeholder: "Descripcion del tramite",
  },
  {
    key: "nroLiquidacion",
    label: "Nro. LIQUIDACION",
    placeholder: "No. de liquidacion",
  },
  {
    key: "beneficiario",
    label: "BENEFICIARIO",
    placeholder: "Nombre o razon social",
  },
  { key: "valor", label: "VALOR", placeholder: "0.00" },
  {
    key: "curContable",
    label: "CUR CONTABLE",
    placeholder: "No. CUR Contable",
  },
  { key: "fechaRecibeJCO", label: "FECHA RECIBE JCO", type: "date" },
  { key: "fechaRecibeJAC", label: "FECHA RECIBE JAC", type: "date" },
  { key: "fechaSolPago", label: "FECHA SOL. PAGO", type: "date" },
  {
    key: "servidorPresupuesto",
    label: "SERVIDOR DE PRESUPUESTO",
    placeholder: "Nombre del servidor",
  },
  {
    key: "novedades",
    label: "NOVEDADES",
    placeholder: "Observaciones o novedades",
  },
  { key: "fechaDevolucion", label: "FECHA DE DEVOLUCION", type: "date" },
  {
    key: "nroMemorandum",
    label: "NRO. MEMORANDO",
    placeholder: "No. memorando final",
  },
];

const ENCABEZADOS_GARANTIA = [
  "N",
  "TIPO PAGO",
  "DIRECCION",
  "MEMORANDO",
  "FECHA INGRESO",
  "FECHA ENTREGA ANALISTA",
  "DESCRIPCION",
  "Nro. LIQUIDACION",
  "BENEFICIARIO",
  "VALOR",
  "CUR CONTABLE",
  "FECHA RECIBE JCO",
  "FECHA RECIBE JAC",
  "FECHA SOL. PAGO",
  "SERVIDOR PRESUPUESTO",
  "ESTADO",
  "NOVEDADES",
  "FECHA DEVOLUCION",
  "NRO. MEMORANDO",
  "CREADO POR",
  "ULTIMA MODIF.",
  "ACCIONES",
];
const CAMPOS_TABLA_GARANTIA = [
  "tipoPago",
  "direccion",
  "memorando",
  "fechaIngreso",
  "fechaEntregaAnalista",
  "descripcion",
  "nroLiquidacion",
  "beneficiario",
  "valor",
  "curContable",
  "fechaRecibeJCO",
  "fechaRecibeJAC",
  "fechaSolPago",
  "servidorPresupuesto",
  "estadoTramite",
  "novedades",
  "fechaDevolucion",
  "nroMemorandum",
  "creadoPor",
  "ultimaModif",
];

const MESES = [
  { val: "01", label: "Enero" },
  { val: "02", label: "Febrero" },
  { val: "03", label: "Marzo" },
  { val: "04", label: "Abril" },
  { val: "05", label: "Mayo" },
  { val: "06", label: "Junio" },
  { val: "07", label: "Julio" },
  { val: "08", label: "Agosto" },
  { val: "09", label: "Septiembre" },
  { val: "10", label: "Octubre" },
  { val: "11", label: "Noviembre" },
  { val: "12", label: "Diciembre" },
];
const ANIOS = ["2024", "2025", "2026", "2027"];

const Ic = {
  Close: () => (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.5"
    >
      <line x1="18" y1="6" x2="6" y2="18" />
      <line x1="6" y1="6" x2="18" y2="18" />
    </svg>
  ),
  Download: ({ s = 16 }) => (
    <svg
      width={s}
      height={s}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.5"
    >
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </svg>
  ),
  Search: () => (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="#94A3B8"
      strokeWidth="2"
    >
      <circle cx="11" cy="11" r="8" />
      <line x1="21" y1="21" x2="16.65" y2="16.65" />
    </svg>
  ),
  Info: () => (
    <svg
      width="15"
      height="15"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <circle cx="12" cy="12" r="10" />
      <line x1="12" y1="8" x2="12" y2="12" />
      <line x1="12" y1="16" x2="12.01" y2="16" />
    </svg>
  ),
  Edit: () => (
    <svg
      width="13"
      height="13"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
      <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
    </svg>
  ),
  Trash: () => (
    <svg
      width="13"
      height="13"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <polyline points="3 6 5 6 21 6" />
      <path d="M19 6l-1 14H6L5 6" />
      <path d="M10 11v6M14 11v6" />
      <path d="M9 6V4h6v2" />
    </svg>
  ),
  User: () => (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
      <circle cx="12" cy="7" r="4" />
    </svg>
  ),
  Lock: () => (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
      <path d="M7 11V7a5 5 0 0 1 10 0v4" />
    </svg>
  ),
  Key: () => (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4" />
    </svg>
  ),
  Logout: () => (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
      <polyline points="16 17 21 12 16 7" />
      <line x1="21" y1="12" x2="9" y2="12" />
    </svg>
  ),
  Settings: () => (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
    </svg>
  ),
  UserPlus: () => (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
      <circle cx="9" cy="7" r="4" />
      <line x1="19" y1="8" x2="19" y2="14" />
      <line x1="22" y1="11" x2="16" y2="11" />
    </svg>
  ),
  EyeShow: () => (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
      <circle cx="12" cy="12" r="3" />
    </svg>
  ),
  EyeHide: () => (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94" />
      <path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19" />
      <line x1="1" y1="1" x2="23" y2="23" />
    </svg>
  ),
  Refresh: () => (
    <svg
      width="14"
      height="14"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <polyline points="23 4 23 10 17 10" />
      <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10" />
    </svg>
  ),
  Shield: () => (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
    </svg>
  ),
  Import: () => (
    <svg
      width="15"
      height="15"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
    >
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="17 8 12 3 7 8" />
      <line x1="12" y1="3" x2="12" y2="15" />
    </svg>
  ),
};

function formatEstadoExcel(estado) {
  const prefijos = {
    Archivo: "[ ARCHIVO ]",
    DFI: "[ DFI ]",
    Factura: "[ FACTURA ]",
    Contabilidad: "[ CONTABILIDAD ]",
    Devuelto: "[ DEVUELTO ]",
    "En revision": "[ EN REVISION ]",
    Observado: "[ OBSERVADO ]",
    Presupuesto: "[ PRESUPUESTO ]",
  };
  return prefijos[estado] || `[ ${(estado || "").toUpperCase()} ]`;
}

function generarExcelProfesional({
  lista,
  camposTabla,
  encabSinExtras,
  nombreHoja,
  labelPeriodo,
  generadoPor,
}) {
  const wb = XLSX.utils.book_new();
  const fechaGen = new Date().toLocaleString("es-EC");
  const totalRegistros = lista.length;
  const atendidos = lista.filter((r) =>
    ESTADOS_ATENDIDOS.includes(r.estadoTramite)
  ).length;
  const pendientes = totalRegistros - atendidos;
  const aoa = [];
  aoa.push(["SERVICIO NACIONAL DE ADUANA DEL ECUADOR - SENAE"]);
  aoa.push([`BITACORA DE TRAMITES - ${nombreHoja.toUpperCase()}`]);
  aoa.push([]);
  aoa.push([
    "Periodo:",
    labelPeriodo,
    "",
    "Generado por:",
    generadoPor,
    "",
    "Fecha:",
    fechaGen,
  ]);
  aoa.push([
    "Total registros:",
    totalRegistros,
    "",
    "Atendidos:",
    atendidos,
    "",
    "Pendientes:",
    pendientes,
  ]);
  aoa.push([]);
  const encabezadosTabla = ["N", ...encabSinExtras];
  aoa.push(encabezadosTabla);
  lista.forEach((r, idx) => {
    const fila = [idx + 1];
    camposTabla.forEach((campo) => {
      let val = r[campo] || "";
      if (campo === "valor" && val) val = `$${val}`;
      if (campo === "estadoTramite") val = formatEstadoExcel(val || "Archivo");
      fila.push(val);
    });
    aoa.push(fila);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const colWidths = encabezadosTabla.map((enc, i) => {
    if (i === 0) return { wch: 5 };
    const c = camposTabla[i - 1] || "";
    if (c.includes("fecha") || c.includes("Fecha")) return { wch: 14 };
    if (c === "estadoTramite") return { wch: 18 };
    if (c === "descripcion" || c === "novedades") return { wch: 35 };
    if (c === "proveedorBeneficiario" || c === "beneficiario")
      return { wch: 28 };
    if (c === "memorando" || c === "nroMemorandum") return { wch: 20 };
    if (enc.length > 16) return { wch: enc.length + 4 };
    return { wch: Math.max(enc.length + 2, 14) };
  });
  ws["!cols"] = colWidths;
  ws["!rows"] = [
    { hpt: 22 },
    { hpt: 18 },
    { hpt: 8 },
    { hpt: 16 },
    { hpt: 16 },
    { hpt: 8 },
    { hpt: 18 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, nombreHoja.slice(0, 31));
  return wb;
}

function generarExcelCantidades({
  registrosPorFecha,
  estados,
  nombreHoja,
  labelPeriodo,
  generadoPor,
}) {
  const wb = XLSX.utils.book_new();
  const fechaGen = new Date().toLocaleString("es-EC");
  const atendidos = registrosPorFecha.filter((r) =>
    ESTADOS_ATENDIDOS.includes(r.estadoTramite)
  ).length;
  const pendientes = registrosPorFecha.length - atendidos;
  const aoa = [];
  aoa.push(["SERVICIO NACIONAL DE ADUANA DEL ECUADOR - SENAE"]);
  aoa.push([`RESUMEN DE CANTIDADES - ${nombreHoja.toUpperCase()}`]);
  aoa.push([]);
  aoa.push(["Periodo:", labelPeriodo, "", "Generado por:", generadoPor]);
  aoa.push(["Fecha de generacion:", fechaGen]);
  aoa.push([]);
  aoa.push(["RESUMEN GENERAL"]);
  aoa.push(["CONCEPTO", "CANTIDAD", "PORCENTAJE"]);
  const total = registrosPorFecha.length;
  const pct = (n) => (total > 0 ? `${Math.round((n / total) * 100)}%` : "0%");
  aoa.push(["Total Ingresados", total, "100%"]);
  aoa.push(["Total Atendidos", atendidos, pct(atendidos)]);
  aoa.push(["Total Pendientes", pendientes, pct(pendientes)]);
  aoa.push([]);
  aoa.push(["DETALLE POR ESTADO"]);
  aoa.push(["ESTADO", "CANTIDAD", "PORCENTAJE", "CLASIFICACION"]);
  estados.forEach((est) => {
    const count = registrosPorFecha.filter(
      (r) => r.estadoTramite === est
    ).length;
    const clasificacion = ESTADOS_ATENDIDOS.includes(est)
      ? "Atendido"
      : "Pendiente";
    aoa.push([formatEstadoExcel(est), count, pct(count), clasificacion]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 22 }, { wch: 12 }, { wch: 14 }, { wch: 14 }];
  ws["!rows"] = [
    { hpt: 22 },
    { hpt: 18 },
    { hpt: 8 },
    { hpt: 16 },
    { hpt: 16 },
    { hpt: 8 },
    { hpt: 16 },
    { hpt: 16 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, "Resumen");
  return wb;
}

function ExportModal({
  registros,
  encabezados,
  camposTabla,
  nombreHoja,
  estados,
  campoFecha,
  onClose,
  usuario,
}) {
  const [tipoReporte, setTipoReporte] = useState("detalle");
  const [modoFecha, setModoFecha] = useState("todo");
  const [hastaFecha, setHastaFecha] = useState("");
  const [mesSel, setMesSel] = useState(new Date().toISOString().slice(5, 7));
  const [anioSel, setAnioSel] = useState(String(new Date().getFullYear()));
  const [estadoSel, setEstadoSel] = useState("");

  const encabSinExtras = encabezados.filter(
    (e) => e !== "ACCIONES" && e !== "N"
  );

  const registrosPorFecha = (() => {
    if (modoFecha === "todo") return registros;
    if (modoFecha === "hasta") {
      if (!hastaFecha) return registros;
      return registros.filter((r) => {
        const f = r[campoFecha];
        return f && f <= hastaFecha;
      });
    }
    if (modoFecha === "mes") {
      return registros.filter((r) => {
        const f = r[campoFecha];
        if (!f) return false;
        return f.slice(0, 4) === anioSel && f.slice(5, 7) === mesSel;
      });
    }
    return registros;
  })();

  const conteosPorEstado = estados.reduce((acc, est) => {
    acc[est] = registrosPorFecha.filter((r) => r.estadoTramite === est).length;
    return acc;
  }, {});
  const registrosFiltradosEstado =
    estadoSel === ""
      ? registrosPorFecha
      : registrosPorFecha.filter((r) => r.estadoTramite === estadoSel);
  const atendidos = registrosPorFecha.filter((r) =>
    ESTADOS_ATENDIDOS.includes(r.estadoTramite)
  ).length;
  const pendientes = registrosPorFecha.length - atendidos;

  const labelPeriodo = (() => {
    if (modoFecha === "todo") return "Todos los registros";
    if (modoFecha === "hasta")
      return hastaFecha ? `Hasta ${hastaFecha}` : "Sin fecha limite definida";
    const mesLabel = MESES.find((m) => m.val === mesSel)?.label || mesSel;
    return `${mesLabel} ${anioSel}`;
  })();

  const puedeExportar = () => {
    if (tipoReporte === "detalle_estado" && estadoSel === "") return false;
    if (modoFecha === "hasta" && !hastaFecha) return false;
    return true;
  };

  const handleExportar = () => {
    const generadoPor = usuario?.nombre || "Sistema";
    if (tipoReporte === "cantidades") {
      if (registrosPorFecha.length === 0) {
        alert("No hay registros para el periodo seleccionado.");
        return;
      }
      const wb = generarExcelCantidades({
        registrosPorFecha,
        estados,
        nombreHoja,
        labelPeriodo,
        generadoPor,
      });
      XLSX.writeFile(
        wb,
        `SENAE_${nombreHoja}_Cantidades_${
          new Date().toISOString().split("T")[0]
        }.xlsx`
      );
      onClose();
    } else {
      const lista =
        tipoReporte === "detalle_estado"
          ? registrosFiltradosEstado
          : registrosPorFecha;
      if (lista.length === 0) {
        alert("No hay registros para exportar con los filtros seleccionados.");
        return;
      }
      const wb = generarExcelProfesional({
        lista,
        camposTabla,
        encabSinExtras,
        nombreHoja,
        labelPeriodo,
        generadoPor,
      });
      XLSX.writeFile(
        wb,
        `SENAE_${nombreHoja}_Detalle_${
          new Date().toISOString().split("T")[0]
        }.xlsx`
      );
      onClose();
    }
  };

  const SelectorFecha = () => (
    <div style={{ marginTop: 16 }}>
      <div style={S.filtroLabel}>
        Periodo (basado en{" "}
        {campoFecha === "fechaIngreso"
          ? "Fecha de Ingreso"
          : "Fecha de Recepcion"}
        )
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        {[
          { val: "todo", label: "Todos" },
          { val: "hasta", label: "Hasta la fecha" },
          { val: "mes", label: "Por mes" },
        ].map((op) => (
          <button
            key={op.val}
            onClick={() => setModoFecha(op.val)}
            style={{
              ...S.chipBtn,
              background: modoFecha === op.val ? "#003366" : "#F1F5F9",
              color: modoFecha === op.val ? "#FFFFFF" : "#475569",
              border:
                modoFecha === op.val
                  ? "2px solid #003366"
                  : "2px solid #E2E8F0",
            }}
          >
            {op.label}
          </button>
        ))}
      </div>
      {modoFecha === "hasta" && (
        <div style={{ marginTop: 12 }}>
          <label style={S.filtroLabel}>Exportar registros hasta:</label>
          <input
            type="date"
            value={hastaFecha}
            onChange={(e) => setHastaFecha(e.target.value)}
            style={{
              ...S.input,
              width: "100%",
              boxSizing: "border-box",
              marginTop: 4,
            }}
          />
        </div>
      )}
      {modoFecha === "mes" && (
        <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
          <div style={{ flex: 1 }}>
            <label style={S.filtroLabel}>Mes</label>
            <select
              value={mesSel}
              onChange={(e) => setMesSel(e.target.value)}
              style={{
                ...S.select,
                width: "100%",
                boxSizing: "border-box",
                marginTop: 4,
              }}
            >
              {MESES.map((m) => (
                <option key={m.val} value={m.val}>
                  {m.label}
                </option>
              ))}
            </select>
          </div>
          <div style={{ flex: 1 }}>
            <label style={S.filtroLabel}>Anio</label>
            <select
              value={anioSel}
              onChange={(e) => setAnioSel(e.target.value)}
              style={{
                ...S.select,
                width: "100%",
                boxSizing: "border-box",
                marginTop: 4,
              }}
            >
              {ANIOS.map((a) => (
                <option key={a} value={a}>
                  {a}
                </option>
              ))}
            </select>
          </div>
        </div>
      )}
      <div style={{ ...S.previewBox, marginTop: 12 }}>
        <Ic.Info />
        <span>
          Periodo: <strong>{labelPeriodo}</strong> &mdash;{" "}
          <strong>{registrosPorFecha.length}</strong> registro
          {registrosPorFecha.length !== 1 ? "s" : ""} coinciden
        </span>
      </div>
    </div>
  );

  return (
    <div
      style={S.overlay}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div style={{ ...S.modal, maxWidth: 580 }}>
        <div style={S.modalHeader}>
          <div>
            <div style={S.modalTitle}>Exportar a Excel</div>
            <div style={S.modalSub}>
              Hoja: {nombreHoja} &mdash; {registros.length} registro
              {registros.length !== 1 ? "s" : ""} totales
            </div>
          </div>
          <button style={S.modalClose} onClick={onClose}>
            <Ic.Close />
          </button>
        </div>
        <div
          style={{
            background: "#F0FDF4",
            borderBottom: "1px solid #BBF7D0",
            padding: "10px 24px",
            display: "flex",
            alignItems: "center",
            gap: 8,
            fontSize: 12,
            color: "#065F46",
          }}
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <polyline points="20 6 9 17 4 12" />
          </svg>
          <span>
            El archivo incluye encabezado institucional SENAE, metadata de
            generacion y estados formateados.
          </span>
        </div>
        <div style={{ ...S.modalBody, maxHeight: "65vh", overflowY: "auto" }}>
          <div style={S.modoLabel}>Tipo de reporte</div>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr 1fr",
              gap: 10,
            }}
          >
            {[
              {
                val: "detalle",
                icon: "doc",
                title: "Detalle general",
                sub: "Todos los campos de cada registro",
              },
              {
                val: "detalle_estado",
                icon: "filter",
                title: "Detalle por estado",
                sub: "Registros de un estado especifico",
              },
              {
                val: "cantidades",
                icon: "chart",
                title: "Resumen cantidades",
                sub: "Ingresados, atendidos y pendientes",
              },
            ].map((op) => {
              const isActive = tipoReporte === op.val;
              return (
                <button
                  key={op.val}
                  onClick={() => setTipoReporte(op.val)}
                  style={{
                    ...S.modoBtn,
                    flexDirection: "column",
                    alignItems: "flex-start",
                    gap: 8,
                    ...(isActive ? S.modoBtnActive : {}),
                  }}
                >
                  <svg
                    width="20"
                    height="20"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="1.8"
                  >
                    {op.icon === "doc" && (
                      <>
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                        <polyline points="14 2 14 8 20 8" />
                        <line x1="16" y1="13" x2="8" y2="13" />
                        <line x1="16" y1="17" x2="8" y2="17" />
                      </>
                    )}
                    {op.icon === "filter" && (
                      <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
                    )}
                    {op.icon === "chart" && (
                      <>
                        <line x1="18" y1="20" x2="18" y2="10" />
                        <line x1="12" y1="20" x2="12" y2="4" />
                        <line x1="6" y1="20" x2="6" y2="14" />
                      </>
                    )}
                  </svg>
                  <div>
                    <div style={{ ...S.modoBtnTitle, fontSize: 12 }}>
                      {op.title}
                    </div>
                    <div style={{ ...S.modoBtnSub, fontSize: 11 }}>
                      {op.sub}
                    </div>
                  </div>
                </button>
              );
            })}
          </div>
          {tipoReporte === "detalle" && <SelectorFecha />}
          {tipoReporte === "detalle_estado" && (
            <>
              <SelectorFecha />
              <div style={{ ...S.filtroPanel, marginTop: 16 }}>
                <div style={S.filtroLabel}>Selecciona el estado a exportar</div>
                <div style={S.estadoGrid}>
                  {estados.map((est) => {
                    const col = ESTADO_COLORS[est] || ESTADO_COLORS["Archivo"];
                    const isSelected = estadoSel === est;
                    const count = conteosPorEstado[est] || 0;
                    return (
                      <button
                        key={est}
                        onClick={() => setEstadoSel(isSelected ? "" : est)}
                        style={{
                          ...S.estadoBtn,
                          background: isSelected ? col.bg : "#FFFFFF",
                          border: isSelected
                            ? `2px solid ${col.border}`
                            : "2px solid #E2E8F0",
                          color: isSelected ? col.color : "#64748B",
                          boxShadow: isSelected
                            ? `0 0 0 3px ${col.border}55`
                            : "none",
                        }}
                      >
                        <span
                          style={{ ...S.estadoBtnDot, background: col.color }}
                        />
                        <span style={S.estadoBtnNombre}>{est}</span>
                        <span
                          style={{
                            ...S.estadoBtnCount,
                            background: isSelected ? col.border : "#F1F5F9",
                            color: isSelected ? col.color : "#94A3B8",
                          }}
                        >
                          {count}
                        </span>
                      </button>
                    );
                  })}
                </div>
                {estadoSel && (
                  <div
                    style={{
                      ...S.previewBox,
                      marginTop: 12,
                      background:
                        (ESTADO_COLORS[estadoSel] || {}).bg || "#EFF6FF",
                      border: `1px solid ${
                        (ESTADO_COLORS[estadoSel] || {}).border || "#BFDBFE"
                      }`,
                      color:
                        (ESTADO_COLORS[estadoSel] || {}).color || "#1D4ED8",
                    }}
                  >
                    <Ic.Info />
                    <span>
                      Se exportaran{" "}
                      <strong>{registrosFiltradosEstado.length}</strong>{" "}
                      registro{registrosFiltradosEstado.length !== 1 ? "s" : ""}{" "}
                      con estado <strong>{estadoSel}</strong>
                    </span>
                  </div>
                )}
              </div>
            </>
          )}
          {tipoReporte === "cantidades" && (
            <>
              <SelectorFecha />
              <div
                style={{
                  marginTop: 16,
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr 1fr",
                  gap: 10,
                }}
              >
                {[
                  {
                    label: "Ingresados",
                    val: registrosPorFecha.length,
                    color: "#1D4ED8",
                    bg: "#EFF6FF",
                    border: "#93C5FD",
                  },
                  {
                    label: "Atendidos",
                    val: atendidos,
                    color: "#065F46",
                    bg: "#ECFDF5",
                    border: "#6EE7B7",
                  },
                  {
                    label: "Pendientes",
                    val: pendientes,
                    color: "#B45309",
                    bg: "#FFFBEB",
                    border: "#FCD34D",
                  },
                ].map((card) => (
                  <div
                    key={card.label}
                    style={{
                      background: card.bg,
                      border: `1.5px solid ${card.border}`,
                      borderRadius: 10,
                      padding: "14px 12px",
                      textAlign: "center",
                    }}
                  >
                    <div
                      style={{
                        fontSize: 26,
                        fontWeight: 800,
                        color: card.color,
                      }}
                    >
                      {card.val}
                    </div>
                    <div
                      style={{
                        fontSize: 11,
                        fontWeight: 700,
                        color: card.color,
                        marginTop: 2,
                        textTransform: "uppercase",
                        letterSpacing: "0.5px",
                      }}
                    >
                      {card.label}
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ marginTop: 12 }}>
                <div style={S.filtroLabel}>Detalle por estado</div>
                <div
                  style={{ display: "flex", flexDirection: "column", gap: 6 }}
                >
                  {estados.map((est) => {
                    const col = ESTADO_COLORS[est] || ESTADO_COLORS["Archivo"];
                    const count = conteosPorEstado[est] || 0;
                    const pct =
                      registrosPorFecha.length > 0
                        ? Math.round((count / registrosPorFecha.length) * 100)
                        : 0;
                    return (
                      <div
                        key={est}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                        }}
                      >
                        <span
                          style={{
                            ...S.estadoBtnDot,
                            background: col.color,
                            width: 9,
                            height: 9,
                          }}
                        />
                        <span
                          style={{
                            fontSize: 12,
                            fontWeight: 600,
                            color: "#475569",
                            minWidth: 100,
                          }}
                        >
                          {est}
                        </span>
                        <div
                          style={{
                            flex: 1,
                            background: "#F1F5F9",
                            borderRadius: 99,
                            height: 6,
                            overflow: "hidden",
                          }}
                        >
                          <div
                            style={{
                              width: `${pct}%`,
                              height: "100%",
                              background: col.color,
                              borderRadius: 99,
                              transition: "width 0.3s",
                            }}
                          />
                        </div>
                        <span
                          style={{
                            fontSize: 12,
                            fontWeight: 700,
                            color: col.color,
                            minWidth: 30,
                            textAlign: "right",
                          }}
                        >
                          {count}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </>
          )}
        </div>
        <div style={S.modalFooter}>
          <button style={S.btnCancelarModal} onClick={onClose}>
            Cancelar
          </button>
          <button
            style={{
              ...S.btnExportar,
              opacity: puedeExportar() ? 1 : 0.45,
              cursor: puedeExportar() ? "pointer" : "not-allowed",
            }}
            onClick={handleExportar}
            disabled={!puedeExportar()}
          >
            <Ic.Download s={16} /> Descargar Excel
          </button>
        </div>
      </div>
    </div>
  );
}

function PerfilModal({ usuario, onClose, onActualizar }) {
  const [nombre, setNombre] = useState(usuario.nombre);
  const [passActual, setPassActual] = useState("");
  const [passNueva, setPassNueva] = useState("");
  const [passConfirm, setPassConfirm] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [codigos, setCodigos] = useState([]);
  const [nuevoCodigo, setNuevoCodigo] = useState("");
  const [showCodigos, setShowCodigos] = useState(false);
  const [msg, setMsg] = useState(null);
  const [loading, setLoading] = useState(false);

  const guardar = async () => {
    if (!nombre.trim()) {
      setMsg({ type: "err", text: "El nombre no puede estar vacio." });
      return;
    }
    if (passNueva && passNueva !== passConfirm) {
      setMsg({ type: "err", text: "Las contrasenas nuevas no coinciden." });
      return;
    }
    if (passNueva && !passActual) {
      setMsg({
        type: "err",
        text: "Ingresa tu contrasena actual para cambiarla.",
      });
      return;
    }
    if (passNueva && passActual !== usuario.contrasena) {
      setMsg({ type: "err", text: "La contrasena actual es incorrecta." });
      return;
    }
    if (passNueva && passNueva.length < 6) {
      setMsg({
        type: "err",
        text: "La nueva contrasena debe tener al menos 6 caracteres.",
      });
      return;
    }
    setLoading(true);
    const usuarios = await sharedGet(SK_USUARIOS, USUARIOS_DEFAULT);
    const updated = usuarios.map((u) =>
      u.id === usuario.id
        ? { ...u, nombre: nombre.trim(), contrasena: passNueva || u.contrasena }
        : u
    );
    await sharedSet(SK_USUARIOS, updated);
    const updatedUser = updated.find((u) => u.id === usuario.id);
    setMsg({ type: "ok", text: "Datos actualizados correctamente." });
    setLoading(false);
    setTimeout(() => {
      onActualizar(updatedUser);
      onClose();
    }, 1200);
  };

  return (
    <div
      style={S.overlay}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div style={{ ...S.modal, maxWidth: 440 }}>
        <div style={S.modalHeader}>
          <div>
            <div style={S.modalTitle}>Mi Perfil</div>
            <div style={S.modalSub}>
              {usuario.usuario} &mdash; {usuario.rol}
            </div>
          </div>
          <button style={S.modalClose} onClick={onClose}>
            <Ic.Close />
          </button>
        </div>
        <div
          style={{
            padding: "24px",
            display: "flex",
            flexDirection: "column",
            gap: 16,
          }}
        >
          <div style={S.formGroup}>
            <label style={S.label}>NOMBRE COMPLETO</label>
            <input
              style={{ ...S.input, width: "100%", boxSizing: "border-box" }}
              value={nombre}
              onChange={(e) => setNombre(e.target.value)}
            />
          </div>
          <div style={{ borderTop: "1px solid #E2E8F0", paddingTop: 16 }}>
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: "#64748B",
                marginBottom: 12,
                textTransform: "uppercase",
                letterSpacing: "0.5px",
              }}
            >
              Cambiar Contrasena (opcional)
            </div>
            {[
              {
                label: "CONTRASENA ACTUAL",
                val: passActual,
                set: setPassActual,
              },
              { label: "NUEVA CONTRASENA", val: passNueva, set: setPassNueva },
              {
                label: "CONFIRMAR NUEVA CONTRASENA",
                val: passConfirm,
                set: setPassConfirm,
              },
            ].map(({ label, val, set }) => (
              <div key={label} style={{ ...S.formGroup, marginBottom: 10 }}>
                <label style={S.label}>{label}</label>
                <div style={{ position: "relative" }}>
                  <input
                    type={showPass ? "text" : "password"}
                    style={{
                      ...S.input,
                      width: "100%",
                      boxSizing: "border-box",
                      paddingRight: 38,
                    }}
                    value={val}
                    onChange={(e) => set(e.target.value)}
                  />
                  <button
                    onClick={() => setShowPass((p) => !p)}
                    style={{
                      position: "absolute",
                      right: 10,
                      top: "50%",
                      transform: "translateY(-50%)",
                      background: "none",
                      border: "none",
                      cursor: "pointer",
                      color: "#94A3B8",
                      display: "flex",
                      padding: 0,
                    }}
                  >
                    {showPass ? <Ic.EyeShow /> : <Ic.EyeHide />}
                  </button>
                </div>
              </div>
            ))}
          </div>
          {msg && (
            <div
              style={{
                padding: "10px 14px",
                borderRadius: 8,
                fontSize: 13,
                fontWeight: 600,
                background: msg.type === "ok" ? "#ECFDF5" : "#FFF1F2",
                color: msg.type === "ok" ? "#065F46" : "#BE123C",
                border: `1px solid ${
                  msg.type === "ok" ? "#6EE7B7" : "#FDA4AF"
                }`,
              }}
            >
              {msg.text}
            </div>
          )}
        </div>
        <div style={S.modalFooter}>
          <button style={S.btnCancelarModal} onClick={onClose}>
            Cancelar
          </button>
          <button
            style={{ ...S.btnGuardar, opacity: loading ? 0.6 : 1 }}
            onClick={guardar}
            disabled={loading}
          >
            {loading ? "Guardando..." : "Guardar Cambios"}
          </button>
        </div>
      </div>
    </div>
  );
}

function AdminUsuariosModal({ onClose }) {
  const [usuarios, setUsuarios] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState({
    usuario: "",
    nombre: "",
    contrasena: "",
    rol: "Analista",
  });
  const [msg, setMsg] = useState(null);
  const [showPass, setShowPass] = useState(false);
  const [codigos, setCodigos] = useState([]);
  const [nuevoCodigo, setNuevoCodigo] = useState("");
  const [showCodigos, setShowCodigos] = useState(false);

  useEffect(() => {
    sharedGet(SK_USUARIOS, USUARIOS_DEFAULT).then((u) => {
      setUsuarios(Array.isArray(u) ? u : USUARIOS_DEFAULT);
      setLoading(false);
    });
    sharedGet(SK_CODIGOS, CODIGOS_DEFAULT).then((c) => {
      setCodigos(Array.isArray(c) ? c : CODIGOS_DEFAULT);
    });
  }, []);

  const agregar = async () => {
    if (
      !form.usuario.trim() ||
      !form.nombre.trim() ||
      !form.contrasena.trim()
    ) {
      setMsg({ type: "err", text: "Todos los campos son obligatorios." });
      return;
    }
    if (form.contrasena.length < 6) {
      setMsg({
        type: "err",
        text: "La contrasena debe tener al menos 6 caracteres.",
      });
      return;
    }
    if (usuarios.find((u) => u.usuario === form.usuario.trim().toLowerCase())) {
      setMsg({ type: "err", text: "Ese nombre de usuario ya existe." });
      return;
    }
    const nuevo = {
      ...form,
      id: generateId(),
      usuario: form.usuario.trim().toLowerCase(),
    };
    const updated = [...usuarios, nuevo];
    await sharedSet(SK_USUARIOS, updated);
    setUsuarios(updated);
    setForm({ usuario: "", nombre: "", contrasena: "", rol: "Analista" });
    setMsg({
      type: "ok",
      text: `Usuario "${nuevo.nombre}" creado correctamente.`,
    });
  };

  const generarCodigo = () => {
    const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    let code = "SENAE#";
    for (let i = 0; i < 7; i++) {
      code += chars[Math.floor(Math.random() * chars.length)];
    }
    setNuevoCodigo(code);
  };

  const agregarCodigo = async () => {
    if (!nuevoCodigo.trim()) return;
    if (codigos.includes(nuevoCodigo.trim())) {
      setMsg({ type: "err", text: "Ese código ya existe." });
      return;
    }
    const updated = [...codigos, nuevoCodigo.trim()];
    await sharedSet(SK_CODIGOS, updated);
    setCodigos(updated);
    setNuevoCodigo("");
    setMsg({ type: "ok", text: `Código "${nuevoCodigo.trim()}" agregado.` });
  };

  const eliminarCodigo = async (cod) => {
    if (codigos.length <= 1) {
      setMsg({ type: "err", text: "Debe quedar al menos un código activo." });
      return;
    }
    if (!window.confirm(`¿Eliminar el código "${cod}"?`)) return;
    const updated = codigos.filter((c) => c !== cod);
    await sharedSet(SK_CODIGOS, updated);
    setCodigos(updated);
  };

  const eliminar = async (id) => {
    if (!window.confirm("Eliminar este usuario?")) return;
    const updated = usuarios.filter((u) => u.id !== id);
    await sharedSet(SK_USUARIOS, updated);
    setUsuarios(updated);
  };

  return (
    <div
      style={S.overlay}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div style={{ ...S.modal, maxWidth: 620 }}>
        <div style={S.modalHeader}>
          <div>
            <div style={S.modalTitle}>Administrar Usuarios</div>
            <div style={S.modalSub}>
              Los cambios son visibles para todos al instante
            </div>
          </div>
          <button style={S.modalClose} onClick={onClose}>
            <Ic.Close />
          </button>
        </div>
        <div style={{ padding: "24px", maxHeight: "70vh", overflowY: "auto" }}>
          {loading ? (
            <div style={{ color: "#94A3B8", textAlign: "center", padding: 20 }}>
              Cargando...
            </div>
          ) : (
            <div style={{ marginBottom: 24 }}>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: "#64748B",
                  letterSpacing: "0.5px",
                  textTransform: "uppercase",
                  marginBottom: 10,
                }}
              >
                Usuarios registrados ({usuarios.length})
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {usuarios.map((u) => (
                  <div
                    key={u.id}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 10,
                      background: "#F8FAFC",
                      border: "1px solid #E2E8F0",
                      borderRadius: 8,
                      padding: "10px 14px",
                    }}
                  >
                    <div
                      style={{
                        width: 32,
                        height: 32,
                        borderRadius: "50%",
                        background: "#C9A84C",
                        color: "#003366",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 12,
                        fontWeight: 800,
                        flexShrink: 0,
                      }}
                    >
                      {u.nombre.slice(0, 2).toUpperCase()}
                    </div>
                    <div style={{ flex: 1 }}>
                      <div
                        style={{
                          fontSize: 13,
                          fontWeight: 700,
                          color: "#1E293B",
                        }}
                      >
                        {u.nombre}
                      </div>
                      <div style={{ fontSize: 11, color: "#94A3B8" }}>
                        @{u.usuario} &mdash; {u.rol}
                      </div>
                    </div>
                    <button
                      onClick={() => eliminar(u.id)}
                      style={{ ...S.btnDel, padding: "5px 8px" }}
                    >
                      <Ic.Trash />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
          <div style={{ borderTop: "1px solid #E2E8F0", paddingTop: 20 }}>
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: "#64748B",
                letterSpacing: "0.5px",
                textTransform: "uppercase",
                marginBottom: 12,
              }}
            >
              Agregar nuevo usuario
            </div>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: "12px 16px",
              }}
            >
              {[
                {
                  key: "nombre",
                  label: "NOMBRE COMPLETO",
                  placeholder: "Ej: Pedro Sanchez",
                },
                {
                  key: "usuario",
                  label: "USUARIO (para login)",
                  placeholder: "Ej: pedro.sanchez",
                },
              ].map(({ key, label, placeholder }) => (
                <div key={key} style={S.formGroup}>
                  <label style={S.label}>{label}</label>
                  <input
                    style={{
                      ...S.input,
                      width: "100%",
                      boxSizing: "border-box",
                    }}
                    placeholder={placeholder}
                    value={form[key]}
                    onChange={(e) =>
                      setForm((p) => ({ ...p, [key]: e.target.value }))
                    }
                  />
                </div>
              ))}
              <div style={S.formGroup}>
                <label style={S.label}>CONTRASENA</label>
                <div style={{ position: "relative" }}>
                  <input
                    type={showPass ? "text" : "password"}
                    style={{
                      ...S.input,
                      width: "100%",
                      boxSizing: "border-box",
                      paddingRight: 38,
                    }}
                    placeholder="Min. 6 caracteres"
                    value={form.contrasena}
                    onChange={(e) =>
                      setForm((p) => ({ ...p, contrasena: e.target.value }))
                    }
                  />
                  <button
                    onClick={() => setShowPass((p) => !p)}
                    style={{
                      position: "absolute",
                      right: 10,
                      top: "50%",
                      transform: "translateY(-50%)",
                      background: "none",
                      border: "none",
                      cursor: "pointer",
                      color: "#94A3B8",
                      display: "flex",
                      padding: 0,
                    }}
                  >
                    {showPass ? <Ic.EyeShow /> : <Ic.EyeHide />}
                  </button>
                </div>
              </div>
              <div style={S.formGroup}>
                <label style={S.label}>ROL</label>
                <select
                  style={{
                    ...S.select,
                    width: "100%",
                    boxSizing: "border-box",
                  }}
                  value={form.rol}
                  onChange={(e) =>
                    setForm((p) => ({ ...p, rol: e.target.value }))
                  }
                >
                  {ROLES.map((r) => (
                    <option key={r} value={r}>
                      {r}
                    </option>
                  ))}
                </select>
              </div>
            </div>
            {msg && (
              <div
                style={{
                  marginTop: 12,
                  padding: "10px 14px",
                  borderRadius: 8,
                  fontSize: 13,
                  fontWeight: 600,
                  background: msg.type === "ok" ? "#ECFDF5" : "#FFF1F2",
                  color: msg.type === "ok" ? "#065F46" : "#BE123C",
                  border: `1px solid ${
                    msg.type === "ok" ? "#6EE7B7" : "#FDA4AF"
                  }`,
                }}
              >
                {msg.text}
              </div>
            )}
            <div
              style={{
                marginTop: 16,
                display: "flex",
                justifyContent: "flex-end",
              }}
            >
              <button style={S.btnGuardar} onClick={agregar}>
                <span style={{ display: "flex", alignItems: "center", gap: 7 }}>
                  <Ic.UserPlus /> Agregar Usuario
                </span>
              </button>
            </div>

            {/* CODIGOS DE VERIFICACION */}
            <div
              style={{
                borderTop: "1px solid #E2E8F0",
                paddingTop: 20,
                marginTop: 20,
              }}
            >
              <button
                onClick={() => setShowCodigos((p) => !p)}
                style={{
                  width: "100%",
                  background: showCodigos ? "#1E293B" : "#F8FAFC",
                  border: "1.5px solid #E2E8F0",
                  borderRadius: 10,
                  padding: "11px 16px",
                  cursor: "pointer",
                  fontFamily: "inherit",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                }}
              >
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    color: showCodigos ? "#FFFFFF" : "#475569",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                  }}
                >
                  Códigos de verificación ({codigos.length} activos)
                </span>
                <span
                  style={{
                    fontSize: 16,
                    color: showCodigos ? "#C9A84C" : "#94A3B8",
                  }}
                >
                  {showCodigos ? "▲" : "▼"}
                </span>
              </button>

              {showCodigos && (
                <div
                  style={{
                    marginTop: 14,
                    display: "flex",
                    flexDirection: "column",
                    gap: 10,
                  }}
                >
                  <div
                    style={{ fontSize: 11, color: "#94A3B8", lineHeight: 1.6 }}
                  >
                    Entrega un código distinto a cada empleado. Si alguien se
                    va, elimina solo su código sin afectar a los demás.
                  </div>

                  <div
                    style={{ display: "flex", flexDirection: "column", gap: 5 }}
                  >
                    {codigos.map((cod) => (
                      <div
                        key={cod}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 10,
                          background: "#FFFBEB",
                          border: "1px solid #FCD34D",
                          borderRadius: 8,
                          padding: "9px 14px",
                        }}
                      >
                        <Ic.Key />
                        <span
                          style={{
                            flex: 1,
                            fontSize: 13,
                            fontWeight: 800,
                            color: "#92400E",
                            letterSpacing: "1px",
                            fontFamily: "monospace",
                          }}
                        >
                          {cod}
                        </span>
                        <button
                          onClick={() =>
                            navigator.clipboard
                              ?.writeText(cod)
                              .then(() =>
                                setMsg({ type: "ok", text: "Código copiado." })
                              )
                          }
                          style={{
                            background: "#FEF3C7",
                            border: "1px solid #FCD34D",
                            borderRadius: 6,
                            padding: "3px 10px",
                            cursor: "pointer",
                            fontSize: 11,
                            fontWeight: 700,
                            color: "#B45309",
                            fontFamily: "inherit",
                          }}
                        >
                          Copiar
                        </button>
                        <button
                          onClick={() => eliminarCodigo(cod)}
                          style={{ ...S.btnDel, padding: "4px 8px" }}
                        >
                          <Ic.Trash />
                        </button>
                      </div>
                    ))}
                  </div>

                  <div
                    style={{
                      background: "#F8FAFC",
                      border: "1.5px solid #E2E8F0",
                      borderRadius: 10,
                      padding: 14,
                    }}
                  >
                    <div
                      style={{
                        fontSize: 11,
                        fontWeight: 700,
                        color: "#64748B",
                        textTransform: "uppercase",
                        letterSpacing: "0.5px",
                        marginBottom: 10,
                      }}
                    >
                      Nuevo código
                    </div>
                    <div style={{ display: "flex", gap: 8 }}>
                      <input
                        value={nuevoCodigo}
                        onChange={(e) => setNuevoCodigo(e.target.value)}
                        placeholder="Ej: SENAE#ABC1234"
                        style={{
                          ...S.input,
                          flex: 1,
                          fontFamily: "monospace",
                          letterSpacing: "1px",
                        }}
                      />
                      <button
                        onClick={generarCodigo}
                        style={{
                          background: "#003366",
                          color: "#FFF",
                          border: "none",
                          borderRadius: 8,
                          padding: "8px 14px",
                          cursor: "pointer",
                          fontSize: 12,
                          fontWeight: 700,
                          fontFamily: "inherit",
                          whiteSpace: "nowrap",
                        }}
                      >
                        Generar
                      </button>
                      <button
                        onClick={agregarCodigo}
                        style={{
                          background: "#065F46",
                          color: "#FFF",
                          border: "none",
                          borderRadius: 8,
                          padding: "8px 14px",
                          cursor: "pointer",
                          fontSize: 12,
                          fontWeight: 700,
                          fontFamily: "inherit",
                          whiteSpace: "nowrap",
                        }}
                      >
                        Guardar
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
        <div style={S.modalFooter}>
          <button style={S.btnCancelarModal} onClick={onClose}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [modo, setModo] = useState("login");
  const [usuario, setUsuario] = useState("");
  const [password, setPassword] = useState("");
  const [nombre, setNombre] = useState("");
  const [codigoAcceso, setCodigoAcceso] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [showCodigo, setShowCodigo] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const intentarLogin = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError("");
    const usuarios = await sharedGet(SK_USUARIOS, USUARIOS_DEFAULT);
    const lista = Array.isArray(usuarios) ? usuarios : USUARIOS_DEFAULT;
    const found = lista.find(
      (u) =>
        u.usuario === usuario.trim().toLowerCase() && u.contrasena === password
    );
    if (found) {
      onLogin(found);
    } else {
      setError("Usuario o contrasena incorrectos.");
    }
    setLoading(false);
  };

  const registrarse = async (e) => {
    e.preventDefault();
    if (
      !nombre.trim() ||
      !usuario.trim() ||
      !password ||
      !codigoAcceso.trim()
    ) {
      setError("Todos los campos son obligatorios.");
      return;
    }
    const codigos = await sharedGet(SK_CODIGOS, CODIGOS_DEFAULT);
    if (!codigos.includes(codigoAcceso.trim())) {
      setError(
        "Codigo de verificacion incorrecto. Contacta al administrador del sistema."
      );
      return;
    }
    if (password.length < 6) {
      setError("La contrasena debe tener al menos 6 caracteres.");
      return;
    }
    setLoading(true);
    setError("");
    const usuarios = await sharedGet(SK_USUARIOS, USUARIOS_DEFAULT);
    const lista = Array.isArray(usuarios) ? usuarios : USUARIOS_DEFAULT;
    if (lista.find((u) => u.usuario === usuario.trim().toLowerCase())) {
      setError("Ese nombre de usuario ya esta en uso.");
      setLoading(false);
      return;
    }
    const nuevo = {
      id: generateId(),
      usuario: usuario.trim().toLowerCase(),
      contrasena: password,
      nombre: nombre.trim(),
      rol: "Analista",
    };
    await sharedSet(SK_USUARIOS, [...lista, nuevo]);
    onLogin(nuevo);
    setLoading(false);
  };

  return (
    <div style={L.root}>
      <div style={L.bg} />
      <div
        style={{
          position: "absolute",
          inset: 0,
          zIndex: 0,
          overflow: "hidden",
          opacity: 0.06,
        }}
      >
        {Array.from({ length: 8 }).map((_, i) => (
          <div
            key={i}
            style={{
              position: "absolute",
              width: 180,
              height: 180,
              border: "2px solid #C9A84C",
              borderRadius: 16,
              top: `${10 + i * 12}%`,
              left: `${-5 + i * 14}%`,
              transform: `rotate(${i * 7}deg)`,
            }}
          />
        ))}
      </div>
      <div style={L.card}>
        <div style={L.logoArea}>
          <svg width="52" height="52" viewBox="0 0 52 52" fill="none">
            <rect width="52" height="52" rx="12" fill="#003366" />
            <rect x="9" y="9" width="14" height="14" fill="#C9A84C" />
            <rect x="29" y="9" width="14" height="14" fill="#C9A84C" />
            <rect x="9" y="29" width="14" height="14" fill="#C9A84C" />
            <rect
              x="29"
              y="29"
              width="14"
              height="14"
              fill="rgba(201,168,76,0.4)"
            />
          </svg>
          <div>
            <div style={L.logoTitle}>SENAE</div>
            <div style={L.logoSub}>Bitacora de Pagos 2026</div>
          </div>
        </div>
        <div style={L.divider} />
        <div
          style={{
            display: "flex",
            gap: 0,
            marginBottom: 24,
            background: "#F1F5F9",
            borderRadius: 10,
            padding: 4,
          }}
        >
          {[
            ["login", "Iniciar Sesion"],
            ["registro", "Registrarse"],
          ].map(([m, label]) => (
            <button
              key={m}
              onClick={() => {
                setModo(m);
                setError("");
              }}
              style={{
                flex: 1,
                padding: "9px 0",
                border: "none",
                borderRadius: 8,
                fontWeight: 700,
                fontSize: 14,
                cursor: "pointer",
                fontFamily: "inherit",
                transition: "all 0.2s",
                background: modo === m ? "#FFFFFF" : "transparent",
                color: modo === m ? "#003366" : "#94A3B8",
                boxShadow: modo === m ? "0 1px 4px rgba(0,0,0,0.1)" : "none",
              }}
            >
              {label}
            </button>
          ))}
        </div>
        <form onSubmit={modo === "login" ? intentarLogin : registrarse}>
          {modo === "registro" && (
            <div style={L.field}>
              <label style={L.label}>NOMBRE COMPLETO</label>
              <div style={L.inputWrap}>
                <span style={L.inputIcon}>
                  <Ic.User />
                </span>
                <input
                  type="text"
                  value={nombre}
                  onChange={(e) => {
                    setNombre(e.target.value);
                    setError("");
                  }}
                  placeholder="Tu nombre completo"
                  style={L.input}
                />
              </div>
            </div>
          )}
          <div style={{ ...L.field, marginTop: modo === "registro" ? 14 : 0 }}>
            <label style={L.label}>USUARIO</label>
            <div style={L.inputWrap}>
              <span style={L.inputIcon}>
                <Ic.User />
              </span>
              <input
                type="text"
                value={usuario}
                onChange={(e) => {
                  setUsuario(e.target.value);
                  setError("");
                }}
                placeholder="Ingresa tu usuario"
                style={L.input}
                autoComplete="username"
              />
            </div>
          </div>
          <div style={{ ...L.field, marginTop: 14 }}>
            <label style={L.label}>CONTRASENA</label>
            <div style={L.inputWrap}>
              <span style={L.inputIcon}>
                <Ic.Lock />
              </span>
              <input
                type={showPass ? "text" : "password"}
                value={password}
                onChange={(e) => {
                  setPassword(e.target.value);
                  setError("");
                }}
                placeholder={
                  modo === "registro"
                    ? "Min. 6 caracteres"
                    : "Ingresa tu contrasena"
                }
                style={{ ...L.input, paddingRight: 40 }}
                autoComplete="current-password"
              />
              <button
                type="button"
                onClick={() => setShowPass((p) => !p)}
                style={L.eyeBtn}
              >
                {showPass ? <Ic.EyeShow /> : <Ic.EyeHide />}
              </button>
            </div>
          </div>
          {modo === "registro" && (
            <div style={{ ...L.field, marginTop: 14 }}>
              <label style={L.label}>CODIGO DE VERIFICACION</label>
              <div style={L.inputWrap}>
                <span style={L.inputIcon}>
                  <Ic.Key />
                </span>
                <input
                  type={showCodigo ? "text" : "password"}
                  value={codigoAcceso}
                  onChange={(e) => {
                    setCodigoAcceso(e.target.value);
                    setError("");
                  }}
                  placeholder="Codigo proporcionado por SENAE"
                  style={{ ...L.input, paddingRight: 40 }}
                />
                <button
                  type="button"
                  onClick={() => setShowCodigo((p) => !p)}
                  style={L.eyeBtn}
                >
                  {showCodigo ? <Ic.EyeShow /> : <Ic.EyeHide />}
                </button>
              </div>
              <div
                style={{
                  marginTop: 6,
                  fontSize: 11,
                  color: "#94A3B8",
                  display: "flex",
                  alignItems: "center",
                  gap: 5,
                }}
              >
                <Ic.Shield />
                <span>
                  Este codigo es provisto por el administrador del sistema
                  SENAE.
                </span>
              </div>
            </div>
          )}
          {error && <div style={L.errorBox}>{error}</div>}
          <button
            type="submit"
            style={{ ...L.btnLogin, opacity: loading ? 0.6 : 1 }}
            disabled={loading}
          >
            {loading
              ? "Procesando..."
              : modo === "login"
              ? "Ingresar al Sistema"
              : "Crear Cuenta"}
          </button>
        </form>
        {modo === "registro" && (
          <div
            style={{
              ...L.infoBox,
              marginTop: 16,
              background: "#FFFBEB",
              border: "1px solid #FCD34D",
              color: "#B45309",
            }}
          >
            Los nuevos usuarios se registran con rol "Analista". Solo personal
            autorizado puede registrarse.
          </div>
        )}
        <div style={L.infoBox}>
          Todos los usuarios comparten los mismos datos en tiempo real desde
          cualquier dispositivo.
        </div>
      </div>
    </div>
  );
}

// ─── COMPONENTE HOJA ──────────────────────────────────────────────────────────
function Hoja({
  emptyState,
  estados,
  camposForm,
  encabezados,
  camposTabla,
  nombreHoja,
  campoFecha,
  visible,
  skKey,
  usuario,
  colMap,
}) {
  const [registros, setRegistros] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState({ ...emptyState });
  const [editandoId, setEditandoId] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [showExport, setShowExport] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [historialVisible, setHistorialVisible] = useState(null);
  const [seleccionados, setSeleccionados] = useState([]);
  const pollRef = useRef(null);

  const cargar = useCallback(async () => {
    const data = await sharedGet(skKey, []);
    setRegistros(Array.isArray(data) ? data : []);
    setLoading(false);
  }, [skKey]);

  useEffect(() => {
    cargar();
    pollRef.current = setInterval(cargar, 10000);
    return () => clearInterval(pollRef.current);
  }, [cargar]);

  const guardarEnStorage = async (nuevos) => {
    setRegistros(nuevos);
    await sharedSet(skKey, nuevos);
  };

  const guardar = async () => {
    if (!form.memorando.trim()) {
      alert("El campo MEMORANDO es obligatorio.");
      return;
    }
    const duplicado = registros.find(
      (r) =>
        r.memorando.trim().toLowerCase() ===
          form.memorando.trim().toLowerCase() &&
        (r.proveedorBeneficiario || r.beneficiario || "")
          .trim()
          .toLowerCase() ===
          (form.proveedorBeneficiario || form.beneficiario || "")
            .trim()
            .toLowerCase() &&
        (r.valor || "").trim() === (form.valor || "").trim() &&
        r.id !== editandoId
    );
    if (duplicado) {
      const confirmar = window.confirm(
        `Ya existe un registro con el memorando "${form.memorando.trim()}", mismo proveedor y mismo valor. ¿Deseas guardarlo de todas formas?`
      );
      if (!confirmar) return;
    }
    const now = new Date().toLocaleString("es-EC");
    let updated;
    if (editandoId) {
      updated = registros.map((r) => {
        if (r.id === editandoId) {
          const historialPrevio = Array.isArray(r.historial) ? r.historial : [];
          return {
            ...form,
            id: editandoId,
            ultimaModif: now,
            modificadoPor: usuario.nombre,
            historial: [
              ...historialPrevio,
              {
                usuario: usuario.nombre,
                fecha: now,
                accion: "Modificacion",
              },
            ],
          };
        }
        return r;
      });
    } else {
      updated = [
        ...registros,
        {
          ...form,
          id: generateId(),
          creadoPor: usuario.nombre,
          ultimaModif: now,
        },
      ];
    }
    await guardarEnStorage(updated);
    setForm({ ...emptyState });
    setEditandoId(null);
    setShowForm(false);
  };

  const editar = (r) => {
    setForm({ ...r });
    setEditandoId(r.id);
    setShowForm(true);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const eliminar = async (id) => {
    if (!window.confirm("Seguro que deseas eliminar este registro?")) return;
    await guardarEnStorage(registros.filter((r) => r.id !== id));
    if (editandoId === id) {
      setForm({ ...emptyState });
      setEditandoId(null);
      setShowForm(false);
    }
  };
  const cancelar = () => {
    setForm({ ...emptyState });
    setEditandoId(null);
    setShowForm(false);
  };

  const handleImportar = async (nuevosRegistros, modo) => {
    let updated;
    if (modo === "reemplazar") {
      updated = nuevosRegistros;
    } else {
      updated = [...registros, ...nuevosRegistros];
    }
    await guardarEnStorage(updated);
  };

  const filtrados = registros.filter((r) => {
    const q = busqueda.toLowerCase();
    return Object.values(r).some((v) =>
      v?.toString().toLowerCase().includes(q)
    );
  });

  if (!visible) return null;

  return (
    <div>
      <div style={S.actionBar}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={S.totalBadge}>
            {registros.length} registro{registros.length !== 1 ? "s" : ""}
          </span>
          <button
            onClick={cargar}
            style={{
              background: "none",
              border: "1px solid #E2E8F0",
              borderRadius: 6,
              padding: "4px 8px",
              cursor: "pointer",
              color: "#64748B",
              display: "flex",
              alignItems: "center",
              gap: 5,
              fontSize: 12,
              fontFamily: "inherit",
            }}
          >
            <Ic.Refresh /> Actualizar
          </button>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {/* NUEVO BOTON IMPORTAR */}
          {seleccionados.length > 0 && (
            <button
              style={{
                background: "#FFF1F2",
                color: "#BE123C",
                border: "1.5px solid #FDA4AF",
                borderRadius: 8,
                padding: "8px 14px",
                cursor: "pointer",
                fontSize: 13,
                fontWeight: 700,
                fontFamily: "inherit",
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
              onClick={async () => {
                if (
                  !window.confirm(
                    `¿Eliminar ${seleccionados.length} registro(s) seleccionado(s)?`
                  )
                )
                  return;
                await guardarEnStorage(
                  registros.filter((r) => !seleccionados.includes(r.id))
                );
                setSeleccionados([]);
              }}
            >
              <Ic.Trash /> Eliminar seleccionados ({seleccionados.length})
            </button>
          )}
          <button style={S.btnImportarBar} onClick={() => setShowImport(true)}>
            <Ic.Import /> Importar Excel
          </button>
          {registros.length > 0 && (
            <button
              style={S.btnExportarBar}
              onClick={() => setShowExport(true)}
            >
              <Ic.Download s={15} /> Exportar Excel
            </button>
          )}
          <button
            style={showForm ? S.btnCancelarHeader : S.btnNuevo}
            onClick={() => (showForm ? cancelar() : setShowForm(true))}
          >
            {showForm ? "Cancelar" : "+ Nuevo Tramite"}
          </button>
        </div>
      </div>

      {showForm && (
        <div style={S.formCard}>
          <div style={S.formHeader}>
            <span style={S.formTitle}>
              {editandoId ? "Editar Registro" : "Nuevo Tramite"}
            </span>
            {editandoId && <span style={S.editingBadge}>Editando</span>}
            <span style={S.hintBadge}>Solo MEMORANDO es obligatorio</span>
          </div>
          <div style={S.formGrid}>
            {camposForm.map(({ key, label, placeholder, type, required }) => (
              <div key={key} style={S.formGroup}>
                <label style={S.label}>
                  {label}
                  {required && (
                    <span style={{ color: "#C2410C", marginLeft: 3 }}>*</span>
                  )}
                </label>
                <input
                  type={type || "text"}
                  value={form[key] || ""}
                  onChange={(e) =>
                    setForm((p) => ({ ...p, [key]: e.target.value }))
                  }
                  placeholder={placeholder || ""}
                  style={{ ...S.input, width: "100%", boxSizing: "border-box" }}
                />
              </div>
            ))}
            <div style={S.formGroup}>
              <label style={S.label}>ESTADO DEL TRAMITE</label>
              <select
                value={form.estadoTramite}
                onChange={(e) =>
                  setForm((p) => ({ ...p, estadoTramite: e.target.value }))
                }
                style={{ ...S.select, width: "100%", boxSizing: "border-box" }}
              >
                {estados.map((e) => (
                  <option key={e} value={e}>
                    {e}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <div style={S.formFooter}>
            <button style={S.btnCancelar} onClick={cancelar}>
              Cancelar
            </button>
            <button style={S.btnGuardar} onClick={guardar}>
              {editandoId ? "Guardar Cambios" : "Guardar Tramite"}
            </button>
          </div>
        </div>
      )}

      {loading ? (
        <div style={{ textAlign: "center", padding: 40, color: "#94A3B8" }}>
          Cargando datos compartidos...
        </div>
      ) : registros.length === 0 ? (
        <div style={S.empty}>
          <div style={S.emptyIcon}>
            <svg width="48" height="48" viewBox="0 0 48 48" fill="none">
              <rect
                x="8"
                y="8"
                width="32"
                height="36"
                rx="4"
                stroke="#CBD5E1"
                strokeWidth="2"
                fill="none"
              />
              <line
                x1="14"
                y1="18"
                x2="34"
                y2="18"
                stroke="#CBD5E1"
                strokeWidth="2"
              />
              <line
                x1="14"
                y1="24"
                x2="34"
                y2="24"
                stroke="#CBD5E1"
                strokeWidth="2"
              />
              <line
                x1="14"
                y1="30"
                x2="26"
                y2="30"
                stroke="#CBD5E1"
                strokeWidth="2"
              />
            </svg>
          </div>
          <div style={S.emptyText}>No hay registros aun</div>
          <div style={S.emptySub}>
            Haz clic en "+ Nuevo Tramite" para agregar uno, o usa "Importar
            Excel" para cargar datos masivamente
          </div>
        </div>
      ) : (
        <>
          <div style={S.searchBar}>
            <Ic.Search />
            <input
              style={S.searchInput}
              placeholder="Buscar en cualquier campo..."
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
            />
            {busqueda && (
              <span style={S.searchCount}>
                {filtrados.length} resultado{filtrados.length !== 1 ? "s" : ""}
              </span>
            )}
          </div>
          <div style={S.tableWrap}>
            {filtrados.length === 0 ? (
              <div style={{ ...S.empty, padding: "40px 20px" }}>
                <div style={S.emptyText}>Sin resultados</div>
                <div style={S.emptySub}>Ninguno coincide con "{busqueda}"</div>
              </div>
            ) : (
              <table style={S.table}>
                <thead>
                  <tr>
                    <th style={S.th}>
                      <input
                        type="checkbox"
                        checked={
                          seleccionados.length === filtrados.length &&
                          filtrados.length > 0
                        }
                        onChange={(e) =>
                          setSeleccionados(
                            e.target.checked ? filtrados.map((r) => r.id) : []
                          )
                        }
                      />
                    </th>
                    {encabezados.map((enc) => (
                      <th key={enc} style={S.th}>
                        {enc}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtrados.map((r, idx) => {
                    const est =
                      ESTADO_COLORS[r.estadoTramite] ||
                      ESTADO_COLORS["Archivo"];
                    const isEditing = editandoId === r.id;
                    return (
                      <tr
                        key={r.id}
                        style={{
                          ...S.tr,
                          background: isEditing
                            ? "#EFF6FF"
                            : idx % 2 === 0
                            ? "#FFFFFF"
                            : "#F8FAFC",
                          outline: isEditing ? "2px solid #3B82F6" : "none",
                        }}
                      >
                        <td style={{ ...S.td, ...S.tdNum }}>
                          <input
                            type="checkbox"
                            checked={seleccionados.includes(r.id)}
                            onChange={(e) =>
                              setSeleccionados((prev) =>
                                e.target.checked
                                  ? [...prev, r.id]
                                  : prev.filter((id) => id !== r.id)
                              )
                            }
                          />
                        </td>
                        {camposTabla.map((campo) => {
                          if (campo === "estadoTramite")
                            return (
                              <td key={campo} style={S.td}>
                                <span
                                  style={{
                                    display: "inline-block",
                                    padding: "3px 10px",
                                    borderRadius: 20,
                                    fontSize: 11,
                                    fontWeight: 700,
                                    whiteSpace: "nowrap",
                                    background: est.bg,
                                    color: est.color,
                                    border: `1px solid ${est.border}`,
                                  }}
                                >
                                  {r[campo] || "-"}
                                </span>
                              </td>
                            );
                          return (
                            <td key={campo} style={S.td}>
                              {r[campo] ? (
                                campo === "valor" ? (
                                  `$${r[campo]}`
                                ) : (
                                  r[campo]
                                )
                              ) : (
                                <span style={{ color: "#CBD5E1" }}>-</span>
                              )}
                            </td>
                          );
                        })}
                        <td style={{ ...S.td, whiteSpace: "nowrap" }}>
                          <button style={S.btnEdit} onClick={() => editar(r)}>
                            <Ic.Edit /> Editar
                          </button>
                          {Array.isArray(r.historial) &&
                            r.historial.length > 0 && (
                              <button
                                style={{
                                  background: "#FDF4FF",
                                  color: "#7E22CE",
                                  border: "1px solid #E9D5FF",
                                  borderRadius: 6,
                                  padding: "4px 10px",
                                  cursor: "pointer",
                                  fontSize: 11,
                                  fontWeight: 700,
                                  fontFamily: "inherit",
                                  display: "inline-flex",
                                  alignItems: "center",
                                  gap: 4,
                                  marginRight: 5,
                                }}
                                onClick={() => setHistorialVisible(r)}
                              >
                                Historial ({r.historial.length})
                              </button>
                            )}
                          <button
                            style={S.btnDel}
                            onClick={() => eliminar(r.id)}
                          >
                            <Ic.Trash />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}

      {showExport && (
        <ExportModal
          registros={registros}
          encabezados={encabezados}
          camposTabla={camposTabla}
          nombreHoja={nombreHoja}
          estados={estados}
          campoFecha={campoFecha}
          onClose={() => setShowExport(false)}
          usuario={usuario}
        />
      )}
      {showImport && (
        <ImportModal
          onClose={() => {
            setShowImport(false);
            cargar();
          }}
          onImportar={handleImportar}
          nombreHoja={nombreHoja}
          colMap={colMap}
          emptyState={emptyState}
          estadosValidos={estados}
        />
      )}

      {/* NUEVO: Modal de historial */}
      {historialVisible && (
        <div
          style={S.overlay}
          onClick={(e) =>
            e.target === e.currentTarget && setHistorialVisible(null)
          }
        >
          <div style={{ ...S.modal, maxWidth: 480 }}>
            <div style={{ ...S.modalHeader, background: "#4A1D96" }}>
              <div>
                <div style={S.modalTitle}>Historial de Modificaciones</div>
                <div style={S.modalSub}>
                  Memorando: {historialVisible.memorando || "-"}
                </div>
              </div>
              <button
                style={S.modalClose}
                onClick={() => setHistorialVisible(null)}
              >
                <Ic.Close />
              </button>
            </div>
            <div style={{ padding: 24, maxHeight: "60vh", overflowY: "auto" }}>
              {historialVisible.creadoPor && (
                <div
                  style={{
                    background: "#F0FDF4",
                    border: "1px solid #BBF7D0",
                    borderRadius: 10,
                    padding: "12px 16px",
                    marginBottom: 16,
                    fontSize: 13,
                    color: "#065F46",
                    fontWeight: 600,
                  }}
                >
                  Creado por: <strong>{historialVisible.creadoPor}</strong>
                </div>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {(historialVisible.historial || []).map((h, i) => (
                  <div
                    key={i}
                    style={{
                      background: "#FAF5FF",
                      border: "1px solid #E9D5FF",
                      borderRadius: 10,
                      padding: "12px 16px",
                      display: "flex",
                      alignItems: "center",
                      gap: 12,
                    }}
                  >
                    <div
                      style={{
                        width: 32,
                        height: 32,
                        borderRadius: "50%",
                        background: "#7E22CE",
                        color: "#fff",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 11,
                        fontWeight: 800,
                        flexShrink: 0,
                      }}
                    >
                      {h.usuario?.slice(0, 2).toUpperCase()}
                    </div>
                    <div>
                      <div
                        style={{
                          fontSize: 13,
                          fontWeight: 700,
                          color: "#1E293B",
                        }}
                      >
                        {h.usuario}
                      </div>
                      <div style={{ fontSize: 11, color: "#94A3B8" }}>
                        {h.accion} &mdash; {h.fecha}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <div style={S.modalFooter}>
              <button
                style={S.btnCancelarModal}
                onClick={() => setHistorialVisible(null)}
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function MisProcesosModal({ usuario, onClose, skPago, skGarantia }) {
  const [datos, setDatos] = useState(null);
  const esAdmin = usuario.rol === "Admin";

  useEffect(() => {
    const cargar = async () => {
      const pagos = await sharedGet(skPago, []);
      const garantias = await sharedGet(skGarantia, []);
      const usuarios = await sharedGet(SK_USUARIOS, USUARIOS_DEFAULT);

      if (esAdmin) {
        // Construir resumen por cada usuario
        const resumen = usuarios.map((u) => {
          const misPagos = pagos.filter((r) => r.creadoPor === u.nombre);
          const misGarantias = garantias.filter(
            (r) => r.creadoPor === u.nombre
          );
          return {
            nombre: u.nombre,
            rol: u.rol,
            pagos: {
              total: misPagos.length,
              atendidos: misPagos.filter((r) =>
                ESTADOS_ATENDIDOS.includes(r.estadoTramite)
              ).length,
              pendientes: misPagos.filter(
                (r) => !ESTADOS_ATENDIDOS.includes(r.estadoTramite)
              ).length,
            },
            garantias: {
              total: misGarantias.length,
              atendidos: misGarantias.filter((r) =>
                ESTADOS_ATENDIDOS.includes(r.estadoTramite)
              ).length,
              pendientes: misGarantias.filter(
                (r) => !ESTADOS_ATENDIDOS.includes(r.estadoTramite)
              ).length,
            },
          };
        });
        setDatos({ resumen });
      } else {
        const misPagos = pagos.filter((r) => r.creadoPor === usuario.nombre);
        const misGarantias = garantias.filter(
          (r) => r.creadoPor === usuario.nombre
        );
        setDatos({ pagos: misPagos, garantias: misGarantias });
      }
    };
    cargar();
  }, []);

  const contar = (lista, estado) =>
    lista.filter((r) => r.estadoTramite === estado).length;

  const atendidosDe = (lista) =>
    lista.filter((r) => ESTADOS_ATENDIDOS.includes(r.estadoTramite)).length;

  const pendientesDe = (lista) =>
    lista.filter((r) => !ESTADOS_ATENDIDOS.includes(r.estadoTramite)).length;

  return (
    <div
      style={S.overlay}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div style={{ ...S.modal, maxWidth: esAdmin ? 720 : 540 }}>
        <div style={{ ...S.modalHeader, background: "#78350F" }}>
          <div>
            <div style={S.modalTitle}>
              {esAdmin ? "Procesos de todos los trabajadores" : "Mis Procesos"}
            </div>
            <div style={S.modalSub}>
              {usuario.nombre} &mdash; {usuario.rol}
            </div>
          </div>
          <button style={S.modalClose} onClick={onClose}>
            <Ic.Close />
          </button>
        </div>

        <div style={{ padding: 24, maxHeight: "65vh", overflowY: "auto" }}>
          {!datos ? (
            <div style={{ textAlign: "center", color: "#94A3B8", padding: 32 }}>
              Cargando...
            </div>
          ) : esAdmin ? (
            // ── VISTA ADMIN: tabla resumen de todos los trabajadores ──
            <div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: "#64748B",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                  marginBottom: 12,
                }}
              >
                Resumen por trabajador
              </div>
              <div
                style={{
                  overflowX: "auto",
                  borderRadius: 10,
                  border: "1.5px solid #E2E8F0",
                }}
              >
                <table
                  style={{
                    borderCollapse: "collapse",
                    width: "100%",
                    fontSize: 12,
                  }}
                >
                  <thead>
                    <tr>
                      {[
                        "Trabajador",
                        "Rol",
                        "Pagos Total",
                        "Pagos Atend.",
                        "Pagos Pend.",
                        "Garantías Total",
                        "Garantías Atend.",
                        "Garantías Pend.",
                      ].map((h) => (
                        <th
                          key={h}
                          style={{
                            background: "#003366",
                            color: "#FFF",
                            padding: "10px 12px",
                            fontWeight: 700,
                            fontSize: 11,
                            textAlign: "left",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {datos.resumen.map((u, i) => (
                      <tr
                        key={u.nombre}
                        style={{
                          background: i % 2 === 0 ? "#FFFFFF" : "#F8FAFC",
                        }}
                      >
                        <td
                          style={{
                            padding: "10px 12px",
                            fontWeight: 700,
                            color: "#1E293B",
                            whiteSpace: "nowrap",
                          }}
                        >
                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                            }}
                          >
                            <div
                              style={{
                                width: 28,
                                height: 28,
                                borderRadius: "50%",
                                background: "#C9A84C",
                                color: "#003366",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                fontSize: 10,
                                fontWeight: 800,
                                flexShrink: 0,
                              }}
                            >
                              {u.nombre.slice(0, 2).toUpperCase()}
                            </div>
                            {u.nombre}
                          </div>
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            color: "#64748B",
                            fontSize: 11,
                          }}
                        >
                          {u.rol}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            fontWeight: 800,
                            color: "#1D4ED8",
                          }}
                        >
                          {u.pagos.total}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            fontWeight: 700,
                            color: "#065F46",
                          }}
                        >
                          {u.pagos.atendidos}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            fontWeight: 700,
                            color: "#B45309",
                          }}
                        >
                          {u.pagos.pendientes}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            fontWeight: 800,
                            color: "#1D4ED8",
                          }}
                        >
                          {u.garantias.total}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            fontWeight: 700,
                            color: "#065F46",
                          }}
                        >
                          {u.garantias.atendidos}
                        </td>
                        <td
                          style={{
                            padding: "10px 12px",
                            textAlign: "center",
                            fontWeight: 700,
                            color: "#B45309",
                          }}
                        >
                          {u.garantias.pendientes}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  {/* Fila de totales */}
                  <tfoot>
                    <tr style={{ background: "#1E293B" }}>
                      <td
                        colSpan={2}
                        style={{
                          padding: "10px 12px",
                          fontWeight: 800,
                          color: "#C9A84C",
                          fontSize: 12,
                        }}
                      >
                        TOTALES
                      </td>
                      {["pagos", "garantias"].map((tipo) =>
                        ["total", "atendidos", "pendientes"].map((campo) => (
                          <td
                            key={tipo + campo}
                            style={{
                              padding: "10px 12px",
                              textAlign: "center",
                              fontWeight: 800,
                              color: "#FFFFFF",
                            }}
                          >
                            {datos.resumen.reduce(
                              (acc, u) => acc + u[tipo][campo],
                              0
                            )}
                          </td>
                        ))
                      )}
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          ) : (
            // ── VISTA TRABAJADOR: sus propios procesos ──
            <>
              {[
                { label: "Pagos", lista: datos.pagos },
                { label: "Garantias", lista: datos.garantias },
              ].map(({ label, lista }) => (
                <div key={label} style={{ marginBottom: 24 }}>
                  <div
                    style={{
                      fontSize: 13,
                      fontWeight: 800,
                      color: "#003366",
                      marginBottom: 12,
                      textTransform: "uppercase",
                      letterSpacing: "0.5px",
                    }}
                  >
                    {label} ({lista.length} total)
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "1fr 1fr 1fr",
                      gap: 10,
                      marginBottom: 14,
                    }}
                  >
                    {[
                      {
                        label: "Ingresados",
                        val: lista.length,
                        color: "#1D4ED8",
                        bg: "#EFF6FF",
                        border: "#93C5FD",
                      },
                      {
                        label: "Atendidos",
                        val: atendidosDe(lista),
                        color: "#065F46",
                        bg: "#ECFDF5",
                        border: "#6EE7B7",
                      },
                      {
                        label: "Pendientes",
                        val: pendientesDe(lista),
                        color: "#B45309",
                        bg: "#FFFBEB",
                        border: "#FCD34D",
                      },
                    ].map((c) => (
                      <div
                        key={c.label}
                        style={{
                          background: c.bg,
                          border: `1.5px solid ${c.border}`,
                          borderRadius: 10,
                          padding: "12px 8px",
                          textAlign: "center",
                        }}
                      >
                        <div
                          style={{
                            fontSize: 22,
                            fontWeight: 800,
                            color: c.color,
                          }}
                        >
                          {c.val}
                        </div>
                        <div
                          style={{
                            fontSize: 10,
                            fontWeight: 700,
                            color: c.color,
                            marginTop: 2,
                            textTransform: "uppercase",
                          }}
                        >
                          {c.label}
                        </div>
                      </div>
                    ))}
                  </div>
                  <div
                    style={{ display: "flex", flexDirection: "column", gap: 5 }}
                  >
                    {(label === "Pagos" ? ESTADOS_PAGO : ESTADOS_GARANTIA).map(
                      (est) => {
                        const col =
                          ESTADO_COLORS[est] || ESTADO_COLORS["Archivo"];
                        const count = contar(lista, est);
                        return (
                          <div
                            key={est}
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                            }}
                          >
                            <span
                              style={{
                                width: 8,
                                height: 8,
                                borderRadius: "50%",
                                background: col.color,
                                flexShrink: 0,
                                display: "inline-block",
                              }}
                            />
                            <span
                              style={{
                                fontSize: 12,
                                fontWeight: 600,
                                color: "#475569",
                                minWidth: 110,
                              }}
                            >
                              {est}
                            </span>
                            <div
                              style={{
                                flex: 1,
                                background: "#F1F5F9",
                                borderRadius: 99,
                                height: 6,
                                overflow: "hidden",
                              }}
                            >
                              <div
                                style={{
                                  width:
                                    lista.length > 0
                                      ? `${Math.round(
                                          (count / lista.length) * 100
                                        )}%`
                                      : "0%",
                                  height: "100%",
                                  background: col.color,
                                  borderRadius: 99,
                                }}
                              />
                            </div>
                            <span
                              style={{
                                fontSize: 12,
                                fontWeight: 700,
                                color: col.color,
                                minWidth: 24,
                                textAlign: "right",
                              }}
                            >
                              {count}
                            </span>
                          </div>
                        );
                      }
                    )}
                  </div>
                </div>
              ))}
            </>
          )}
        </div>

        <div style={S.modalFooter}>
          <button style={S.btnCancelarModal} onClick={onClose}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
}

function AppInner({ usuario, setUsuario }) {
  const [tabActiva, setTabActiva] = useState("pago");
  const [showPerfil, setShowPerfil] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);
  const [showMisProcesos, setShowMisProcesos] = useState(false); // NUEVO

  const handleLogout = () => {
    if (window.confirm("Seguro que deseas cerrar sesion?")) setUsuario(null);
  };

  return (
    <div style={S.root}>
      <div style={S.header}>
        <div style={S.headerInner}>
          <svg width="38" height="38" viewBox="0 0 38 38" fill="none">
            <rect width="38" height="38" rx="8" fill="#003366" />
            <rect x="7" y="7" width="10" height="10" fill="#C9A84C" />
            <rect x="21" y="7" width="10" height="10" fill="#C9A84C" />
            <rect x="7" y="21" width="10" height="10" fill="#C9A84C" />
            <rect
              x="21"
              y="21"
              width="10"
              height="10"
              fill="rgba(201,168,76,0.4)"
            />
          </svg>
          <div>
            <div style={S.headerTitle}>
              SENAE &mdash; Bitacora de Pagos 2026
            </div>
            <div style={S.headerSub}>
              Servicio Nacional de Aduana del Ecuador
            </div>
          </div>
        </div>

        <div style={S.userBar}>
          <div style={S.userInfo}>
            <div style={S.userAvatar}>
              {usuario.nombre.slice(0, 2).toUpperCase()}
            </div>
            <div>
              <div style={S.userName}>{usuario.nombre}</div>
              <div style={S.userRol}>{usuario.rol}</div>
            </div>
          </div>
          {/* NUEVO: boton mis procesos */}
          <button
            style={{
              ...S.btnIconHeader,
              background: "rgba(201,168,76,0.2)",
              border: "1px solid rgba(201,168,76,0.4)",
              padding: "7px 12px",
              fontSize: 12,
              fontWeight: 700,
              gap: 6,
              display: "flex",
              alignItems: "center",
            }}
            onClick={() => setShowMisProcesos(true)}
            title="Mis Procesos"
          >
            Mis Procesos
          </button>

          <button
            style={S.btnIconHeader}
            onClick={() => setShowPerfil(true)}
            title="Mi Perfil"
          >
            <Ic.Settings />
          </button>
          {usuario.rol === "Admin" && (
            <button
              style={S.btnIconHeader}
              onClick={() => setShowAdmin(true)}
              title="Administrar Usuarios"
            >
              <Ic.UserPlus />
            </button>
          )}
          <button style={S.btnLogout} onClick={handleLogout}>
            <Ic.Logout /> Salir
          </button>
        </div>
      </div>

      <div style={S.tabs}>
        {[
          { key: "pago", label: "Pagos" },
          { key: "garantia", label: "Garantias" },
        ].map((t) => (
          <button
            key={t.key}
            style={{ ...S.tab, ...(tabActiva === t.key ? S.tabActive : {}) }}
            onClick={() => setTabActiva(t.key)}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div style={S.content}>
        <Hoja
          visible={tabActiva === "pago"}
          emptyState={EMPTY_PAGO}
          estados={ESTADOS_PAGO}
          camposForm={CAMPOS_FORM_PAGO}
          encabezados={ENCABEZADOS_PAGO}
          camposTabla={CAMPOS_TABLA_PAGO}
          nombreHoja="Pagos"
          campoFecha="fechaRecepcion"
          skKey={SK_PAGO}
          usuario={usuario}
          colMap={COLUMN_MAP_PAGO}
        />
        <Hoja
          visible={tabActiva === "garantia"}
          emptyState={EMPTY_GARANTIA}
          estados={ESTADOS_GARANTIA}
          camposForm={CAMPOS_FORM_GARANTIA}
          encabezados={ENCABEZADOS_GARANTIA}
          camposTabla={CAMPOS_TABLA_GARANTIA}
          nombreHoja="Garantias"
          campoFecha="fechaIngreso"
          skKey={SK_GARANTIA}
          usuario={usuario}
          colMap={COLUMN_MAP_GARANTIA}
        />
      </div>

      {showPerfil && (
        <PerfilModal
          usuario={usuario}
          onClose={() => setShowPerfil(false)}
          onActualizar={(u) => setUsuario(u)}
        />
      )}
      {showAdmin && usuario.rol === "Admin" && (
        <AdminUsuariosModal onClose={() => setShowAdmin(false)} />
      )}

      {/* NUEVO */}
      {showMisProcesos && (
        <MisProcesosModal
          usuario={usuario}
          onClose={() => setShowMisProcesos(false)}
          skPago={SK_PAGO}
          skGarantia={SK_GARANTIA}
        />
      )}
    </div>
  );
}

export default function App() {
  const [usuario, setUsuario] = useState(null);

  useEffect(() => {
    sharedGet(SK_USUARIOS, null).then(async (u) => {
      if (!u || !Array.isArray(u) || u.length === 0) {
        await sharedSet(SK_USUARIOS, USUARIOS_DEFAULT);
      }
    });
  }, []);

  if (!usuario) return <LoginScreen onLogin={setUsuario} />;
  return <AppInner usuario={usuario} setUsuario={setUsuario} />;
}

// ─── ESTILOS S (tabla, formularios, modales) ──────────────────────────────────
const S = {
  root: {
    fontFamily: "'Inter', 'Segoe UI', sans-serif",
    minHeight: "100vh",
    background: "#F1F5F9",
    color: "#1E293B",
  },

  // Header
  header: {
    background: "#003366",
    padding: "14px 24px",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 16,
    boxShadow: "0 2px 12px rgba(0,51,102,0.3)",
    position: "sticky",
    top: 0,
    zIndex: 100,
  },
  headerInner: { display: "flex", alignItems: "center", gap: 14 },
  headerTitle: {
    fontSize: 16,
    fontWeight: 800,
    color: "#FFFFFF",
    letterSpacing: "0.3px",
  },
  headerSub: { fontSize: 11, color: "rgba(255,255,255,0.55)", marginTop: 1 },
  userBar: { display: "flex", alignItems: "center", gap: 10 },
  userInfo: { display: "flex", alignItems: "center", gap: 10 },
  userAvatar: {
    width: 34,
    height: 34,
    borderRadius: "50%",
    background: "#C9A84C",
    color: "#003366",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 12,
    fontWeight: 800,
    flexShrink: 0,
  },
  userName: { fontSize: 13, fontWeight: 700, color: "#FFFFFF" },
  userRol: { fontSize: 10, color: "rgba(255,255,255,0.55)" },
  btnIconHeader: {
    background: "rgba(255,255,255,0.1)",
    border: "1px solid rgba(255,255,255,0.2)",
    borderRadius: 8,
    padding: "7px 9px",
    cursor: "pointer",
    color: "#FFFFFF",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "background 0.2s",
  },
  btnLogout: {
    background: "rgba(255,255,255,0.08)",
    border: "1px solid rgba(255,255,255,0.2)",
    borderRadius: 8,
    padding: "7px 14px",
    cursor: "pointer",
    color: "#FFFFFF",
    fontSize: 13,
    fontWeight: 600,
    display: "flex",
    alignItems: "center",
    gap: 7,
    fontFamily: "inherit",
    transition: "background 0.2s",
  },

  // Tabs
  tabs: {
    background: "#FFFFFF",
    borderBottom: "1px solid #E2E8F0",
    padding: "0 24px",
    display: "flex",
    gap: 0,
  },
  tab: {
    padding: "14px 24px",
    border: "none",
    borderBottom: "3px solid transparent",
    background: "none",
    cursor: "pointer",
    fontSize: 14,
    fontWeight: 600,
    color: "#94A3B8",
    fontFamily: "inherit",
    transition: "all 0.2s",
  },
  tabActive: { color: "#003366", borderBottom: "3px solid #C9A84C" },

  // Content
  content: { padding: "24px", maxWidth: 1800, margin: "0 auto" },

  // Action bar
  actionBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 16,
    flexWrap: "wrap",
    gap: 10,
  },
  totalBadge: {
    background: "#003366",
    color: "#FFFFFF",
    borderRadius: 20,
    padding: "4px 12px",
    fontSize: 12,
    fontWeight: 700,
  },
  btnNuevo: {
    background: "#003366",
    color: "#FFFFFF",
    border: "none",
    borderRadius: 8,
    padding: "9px 18px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
  },
  btnCancelarHeader: {
    background: "#F1F5F9",
    color: "#475569",
    border: "1.5px solid #CBD5E1",
    borderRadius: 8,
    padding: "9px 18px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
  },
  btnExportarBar: {
    background: "#ECFDF5",
    color: "#065F46",
    border: "1.5px solid #6EE7B7",
    borderRadius: 8,
    padding: "8px 14px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  btnImportarBar: {
    background: "#EFF6FF",
    color: "#1D4ED8",
    border: "1.5px solid #93C5FD",
    borderRadius: 8,
    padding: "8px 14px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
    display: "flex",
    alignItems: "center",
    gap: 6,
  },

  // Form card
  formCard: {
    background: "#FFFFFF",
    borderRadius: 14,
    border: "1.5px solid #E2E8F0",
    marginBottom: 20,
    overflow: "hidden",
    boxShadow: "0 4px 16px rgba(0,0,0,0.06)",
  },
  formHeader: {
    background: "#003366",
    padding: "14px 20px",
    display: "flex",
    alignItems: "center",
    gap: 10,
  },
  formTitle: { fontSize: 14, fontWeight: 800, color: "#FFFFFF", flex: 1 },
  editingBadge: {
    background: "#C9A84C",
    color: "#003366",
    borderRadius: 20,
    padding: "2px 10px",
    fontSize: 11,
    fontWeight: 800,
  },
  hintBadge: {
    background: "rgba(255,255,255,0.1)",
    color: "rgba(255,255,255,0.7)",
    borderRadius: 20,
    padding: "2px 10px",
    fontSize: 10,
    fontWeight: 600,
  },
  formGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: "14px 20px",
    padding: "20px",
  },
  formGroup: { display: "flex", flexDirection: "column", gap: 5 },
  label: {
    fontSize: 10,
    fontWeight: 700,
    color: "#64748B",
    letterSpacing: "0.5px",
    textTransform: "uppercase",
  },
  input: {
    border: "1.5px solid #E2E8F0",
    borderRadius: 8,
    padding: "8px 12px",
    fontSize: 13,
    color: "#1E293B",
    fontFamily: "inherit",
    outline: "none",
    transition: "border-color 0.2s",
    background: "#FAFAFA",
  },
  select: {
    border: "1.5px solid #E2E8F0",
    borderRadius: 8,
    padding: "8px 12px",
    fontSize: 13,
    color: "#1E293B",
    fontFamily: "inherit",
    background: "#FAFAFA",
    outline: "none",
  },
  formFooter: {
    padding: "16px 20px",
    borderTop: "1px solid #E2E8F0",
    display: "flex",
    justifyContent: "flex-end",
    gap: 10,
    background: "#F8FAFC",
  },
  btnCancelar: {
    background: "#F1F5F9",
    color: "#475569",
    border: "1.5px solid #CBD5E1",
    borderRadius: 8,
    padding: "9px 20px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
  },
  btnGuardar: {
    background: "#003366",
    color: "#FFFFFF",
    border: "none",
    borderRadius: 8,
    padding: "9px 22px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
    display: "flex",
    alignItems: "center",
    gap: 7,
  },

  // Search
  searchBar: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    background: "#FFFFFF",
    border: "1.5px solid #E2E8F0",
    borderRadius: 10,
    padding: "8px 14px",
    marginBottom: 14,
  },
  searchInput: {
    flex: 1,
    border: "none",
    outline: "none",
    fontSize: 14,
    color: "#1E293B",
    fontFamily: "inherit",
    background: "transparent",
  },
  searchCount: {
    background: "#EFF6FF",
    color: "#1D4ED8",
    borderRadius: 20,
    padding: "2px 10px",
    fontSize: 12,
    fontWeight: 700,
    whiteSpace: "nowrap",
  },

  // Table
  tableWrap: {
    overflowX: "auto",
    borderRadius: 12,
    border: "1.5px solid #E2E8F0",
    boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
  },
  table: {
    borderCollapse: "collapse",
    width: "100%",
    fontSize: 12,
    minWidth: 900,
  },
  th: {
    background: "#003366",
    color: "#FFFFFF",
    padding: "11px 13px",
    fontWeight: 700,
    fontSize: 11,
    textAlign: "left",
    whiteSpace: "nowrap",
    letterSpacing: "0.3px",
    borderRight: "1px solid rgba(255,255,255,0.08)",
  },
  tr: { transition: "background 0.15s" },
  td: {
    padding: "9px 13px",
    color: "#334155",
    borderBottom: "1px solid #F1F5F9",
    maxWidth: 200,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    verticalAlign: "middle",
  },
  tdNum: {
    color: "#94A3B8",
    fontWeight: 700,
    textAlign: "center",
    minWidth: 36,
  },
  btnEdit: {
    background: "#EFF6FF",
    color: "#1D4ED8",
    border: "1px solid #BFDBFE",
    borderRadius: 6,
    padding: "4px 10px",
    cursor: "pointer",
    fontSize: 11,
    fontWeight: 700,
    fontFamily: "inherit",
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
    marginRight: 5,
  },
  btnDel: {
    background: "#FFF1F2",
    color: "#BE123C",
    border: "1px solid #FDA4AF",
    borderRadius: 6,
    padding: "4px 8px",
    cursor: "pointer",
    fontSize: 11,
    fontFamily: "inherit",
    display: "inline-flex",
    alignItems: "center",
  },

  // Empty
  empty: {
    textAlign: "center",
    padding: "60px 20px",
    background: "#FFFFFF",
    borderRadius: 12,
    border: "1.5px solid #E2E8F0",
  },
  emptyIcon: { marginBottom: 16, opacity: 0.4 },
  emptyText: {
    fontSize: 16,
    fontWeight: 700,
    color: "#94A3B8",
    marginBottom: 6,
  },
  emptySub: { fontSize: 13, color: "#CBD5E1" },

  // Modal base
  overlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.45)",
    zIndex: 1000,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 20,
  },
  modal: {
    background: "#FFFFFF",
    borderRadius: 16,
    width: "100%",
    maxHeight: "90vh",
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
    boxShadow: "0 20px 60px rgba(0,0,0,0.25)",
  },
  modalHeader: {
    background: "#003366",
    padding: "18px 24px",
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 12,
  },
  modalTitle: { fontSize: 16, fontWeight: 800, color: "#FFFFFF" },
  modalSub: { fontSize: 12, color: "rgba(255,255,255,0.6)", marginTop: 3 },
  modalClose: {
    background: "rgba(255,255,255,0.1)",
    border: "none",
    borderRadius: 8,
    padding: 8,
    cursor: "pointer",
    color: "#FFFFFF",
    display: "flex",
    alignItems: "center",
    flexShrink: 0,
  },
  modalBody: { padding: "24px", overflowY: "auto" },
  modalFooter: {
    padding: "14px 24px",
    borderTop: "1px solid #E2E8F0",
    display: "flex",
    justifyContent: "flex-end",
    gap: 10,
    background: "#F8FAFC",
  },
  btnCancelarModal: {
    background: "#F1F5F9",
    color: "#475569",
    border: "1.5px solid #CBD5E1",
    borderRadius: 8,
    padding: "9px 20px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
  },
  btnExportar: {
    background: "#065F46",
    color: "#FFFFFF",
    border: "none",
    borderRadius: 8,
    padding: "9px 20px",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 700,
    fontFamily: "inherit",
    display: "flex",
    alignItems: "center",
    gap: 7,
  },

  // Export modal internals
  modoLabel: {
    fontSize: 11,
    fontWeight: 700,
    color: "#64748B",
    letterSpacing: "0.5px",
    textTransform: "uppercase",
    marginBottom: 10,
  },
  modoBtn: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "12px 14px",
    border: "2px solid #E2E8F0",
    borderRadius: 10,
    cursor: "pointer",
    background: "#FFFFFF",
    fontFamily: "inherit",
    textAlign: "left",
    transition: "all 0.2s",
  },
  modoBtnActive: { border: "2px solid #003366", background: "#EFF6FF" },
  modoBtnTitle: { fontSize: 13, fontWeight: 700, color: "#1E293B" },
  modoBtnSub: { fontSize: 11, color: "#64748B", marginTop: 2 },
  filtroPanel: {
    background: "#F8FAFC",
    border: "1.5px solid #E2E8F0",
    borderRadius: 12,
    padding: 16,
  },
  filtroLabel: {
    fontSize: 11,
    fontWeight: 700,
    color: "#64748B",
    textTransform: "uppercase",
    letterSpacing: "0.5px",
  },
  estadoGrid: { display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 },
  estadoBtn: {
    display: "flex",
    alignItems: "center",
    gap: 7,
    padding: "7px 12px",
    borderRadius: 8,
    cursor: "pointer",
    fontFamily: "inherit",
    transition: "all 0.2s",
  },
  estadoBtnDot: {
    width: 8,
    height: 8,
    borderRadius: "50%",
    flexShrink: 0,
    display: "inline-block",
  },
  estadoBtnNombre: { fontSize: 12, fontWeight: 700 },
  estadoBtnCount: {
    borderRadius: 20,
    padding: "1px 8px",
    fontSize: 11,
    fontWeight: 800,
    marginLeft: 4,
  },
  previewBox: {
    background: "#EFF6FF",
    border: "1px solid #BFDBFE",
    borderRadius: 8,
    padding: "10px 14px",
    display: "flex",
    alignItems: "center",
    gap: 8,
    fontSize: 12,
    color: "#1D4ED8",
  },
  chipBtn: {
    border: "2px solid",
    borderRadius: 8,
    padding: "6px 14px",
    cursor: "pointer",
    fontSize: 12,
    fontWeight: 700,
    fontFamily: "inherit",
  },
};

// ─── ESTILOS L (Login screen) ──────────────────────────────────────────────────
const L = {
  root: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    position: "relative",
    overflow: "hidden",
    background: "#0F172A",
    fontFamily: "'Inter', 'Segoe UI', sans-serif",
  },
  bg: {
    position: "absolute",
    inset: 0,
    background:
      "radial-gradient(ellipse at 30% 40%, rgba(0,51,102,0.8) 0%, transparent 60%), radial-gradient(ellipse at 80% 80%, rgba(201,168,76,0.15) 0%, transparent 50%)",
    zIndex: 0,
  },
  card: {
    position: "relative",
    zIndex: 1,
    background: "rgba(255,255,255,0.97)",
    borderRadius: 20,
    padding: "36px 36px 28px",
    width: "100%",
    maxWidth: 420,
    boxShadow: "0 32px 80px rgba(0,0,0,0.4), 0 0 0 1px rgba(255,255,255,0.1)",
  },
  logoArea: {
    display: "flex",
    alignItems: "center",
    gap: 14,
    marginBottom: 24,
  },
  logoTitle: {
    fontSize: 24,
    fontWeight: 900,
    color: "#003366",
    letterSpacing: "1px",
  },
  logoSub: { fontSize: 12, color: "#64748B", fontWeight: 600, marginTop: 2 },
  divider: {
    height: 1,
    background: "linear-gradient(to right, #003366, #C9A84C, transparent)",
    marginBottom: 24,
    opacity: 0.4,
  },
  field: { display: "flex", flexDirection: "column", gap: 6 },
  label: {
    fontSize: 10,
    fontWeight: 700,
    color: "#64748B",
    letterSpacing: "0.8px",
    textTransform: "uppercase",
  },
  inputWrap: { position: "relative", display: "flex", alignItems: "center" },
  inputIcon: {
    position: "absolute",
    left: 12,
    color: "#94A3B8",
    display: "flex",
    alignItems: "center",
    zIndex: 1,
  },
  input: {
    width: "100%",
    border: "1.5px solid #E2E8F0",
    borderRadius: 10,
    padding: "11px 12px 11px 42px",
    fontSize: 14,
    color: "#1E293B",
    fontFamily: "inherit",
    outline: "none",
    background: "#F8FAFC",
    boxSizing: "border-box",
    transition: "border-color 0.2s",
  },
  eyeBtn: {
    position: "absolute",
    right: 12,
    background: "none",
    border: "none",
    cursor: "pointer",
    color: "#94A3B8",
    display: "flex",
    alignItems: "center",
    padding: 0,
  },
  errorBox: {
    marginTop: 14,
    padding: "10px 14px",
    borderRadius: 8,
    background: "#FFF1F2",
    border: "1px solid #FDA4AF",
    color: "#BE123C",
    fontSize: 13,
    fontWeight: 600,
  },
  btnLogin: {
    marginTop: 20,
    width: "100%",
    background: "#003366",
    color: "#FFFFFF",
    border: "none",
    borderRadius: 10,
    padding: "13px",
    fontSize: 14,
    fontWeight: 800,
    cursor: "pointer",
    fontFamily: "inherit",
    letterSpacing: "0.3px",
    transition: "opacity 0.2s",
  },
  infoBox: {
    marginTop: 16,
    padding: "10px 14px",
    borderRadius: 8,
    background: "#EFF6FF",
    border: "1px solid #BFDBFE",
    color: "#1D4ED8",
    fontSize: 12,
    fontWeight: 500,
  },
};
