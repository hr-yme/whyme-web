import React, { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  UsersRound, CalendarClock, LogOut, RefreshCw, Download,
  UploadCloud, Plus, X, Pencil, AlertTriangle, CheckCircle2, XCircle, Clock3,
  Search, Lock, ListChecks, LayoutGrid, CalendarDays, FileSpreadsheet,
  FileCheck2, FileClock, Eye, EyeOff,
} from "lucide-react";

/* ============================================================================
   PALETA OFICIAL YME (rigorosa — usar exatamente estes 4 tons)
============================================================================ */

const COLORS = {
  navy: "#151E33",  // Fundo principal, header, modais
  pink: "#FF6EF6",  // Ações primárias, item ativo, focos/hover
  mint: "#E7F6F4",  // Cartões, tabelas, superfícies secundárias
  white: "#FCFCFC", // Títulos, texto sobre fundo escuro, ícones
};

function hexToRgba(hex, alpha) {
  const h = hex.replace("#", "");
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

/* Folha de estilos global com a paleta oficial — usada para estados
   (hover/focus) que as classes utilitárias estáticas não cobrem. */
const GLOBAL_CSS = `
  .yme-btn-primary { background-color: ${COLORS.pink}; color: ${COLORS.navy}; transition: filter .15s ease; }
  .yme-btn-primary:hover { filter: brightness(0.93); }

  .yme-btn-outline-dark { background-color: transparent; color: ${COLORS.white}; border: 1px solid ${hexToRgba(COLORS.mint, 0.25)}; transition: background-color .15s ease; }
  .yme-btn-outline-dark:hover { background-color: ${hexToRgba(COLORS.mint, 0.08)}; }

  .yme-btn-outline-light { background-color: transparent; color: ${COLORS.navy}; border: 1px solid ${hexToRgba(COLORS.navy, 0.2)}; transition: background-color .15s ease; }
  .yme-btn-outline-light:hover { background-color: ${hexToRgba(COLORS.navy, 0.06)}; }

  .yme-input { background-color: ${COLORS.white}; color: ${COLORS.navy}; border: 1px solid ${hexToRgba(COLORS.navy, 0.18)}; }
  .yme-input::placeholder { color: ${hexToRgba(COLORS.navy, 0.4)}; }
  .yme-input:focus { outline: none; border-color: ${COLORS.pink}; box-shadow: 0 0 0 3px ${hexToRgba(COLORS.pink, 0.25)}; }

  .yme-nav-item { position: relative; color: #cbd5e1; transition: color .15s ease, background-color .15s ease; }
  .yme-nav-item:hover { color: ${COLORS.pink}; }
  .yme-nav-item.active { color: ${COLORS.pink}; background-color: ${hexToRgba(COLORS.mint, 0.06)}; }
  .yme-nav-item.active::after { content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 3px; background-color: ${COLORS.pink}; }

  .yme-table-row:hover { background-color: ${hexToRgba(COLORS.navy, 0.045)}; }
  .yme-link { color: ${COLORS.navy}; text-decoration: underline; text-decoration-color: ${hexToRgba(COLORS.navy, 0.35)}; }
  .yme-link:hover { color: ${COLORS.pink}; text-decoration-color: ${COLORS.pink}; }
`;

/* ============================================================================
   CONSTANTS
============================================================================ */

const ACCESS_KEY = "YME2026";

const DAYS = ["Seg", "Ter", "Qua", "Qui", "Sex"];
const TIMES = ["09:00", "10:30", "14:00", "15:30", "17:00"];
const SLOTS = DAYS.flatMap((d) => TIMES.map((t) => `${d} ${t}`));

// Estrutura organizacional fixa da YME
const ORG = [
  { dept: "Brand Strategy", diretor: "Gustavo Dias", supervisor: "Tiago Barros", supervisorTitle: "COO", rh: ["Lara Costa"] },
  { dept: "Digital Development", diretor: "Tânia Silva", supervisor: "Inês Costa", supervisorTitle: "CMO", rh: ["Joana Furtado", "Helena Castro"] },
  { dept: "Human Resources", diretor: "Mariana Lopes", supervisor: "Beatriz Garcia", supervisorTitle: "CEO", rh: ["Andreia Freitas"] },
  { dept: "Legal & Finance", diretor: "Joana Pereira", supervisor: "Tiago Barros", supervisorTitle: "COO", rh: ["Carlota Baptista"] },
  { dept: "Quality Management", diretor: "Tiago Costa", supervisor: "Beatriz Garcia", supervisorTitle: "CEO", rh: ["Ana Pereira", "Lia Fortes"] },
  { dept: "Sales & Commercial", diretor: "Tomás Costa", supervisor: "Inês Costa", supervisorTitle: "CMO", rh: ["Catarina Lamego"] },
];
const DEPARTMENTS = ORG.map((o) => o.dept);

const PHASE_LABEL = { fase1: "Fase 1", fase2: "Fase 2", fase3: "Fase 3" };

/* Cores exatas dos 6 Departamentos, conforme o Excel Mestre oficial da YME. */
const DEPT_BADGE_CLASSES = {
  "Brand Strategy": "bg-amber-100 text-amber-800 border-amber-300",
  "Digital Development": "bg-slate-200 text-slate-700 border-slate-300",
  "Human Resources": "bg-emerald-100 text-emerald-800 border-emerald-300",
  "Legal & Finance": "bg-purple-100 text-purple-800 border-purple-300",
  "Quality Management": "bg-rose-100 text-rose-800 border-rose-300",
  "Sales & Commercial": "bg-sky-100 text-sky-800 border-sky-300",
};
const DEPT_BADGE_FALLBACK = "bg-slate-200 text-slate-700 border-slate-300";
function deptBadgeClass(dept) {
  return DEPT_BADGE_CLASSES[dept] || DEPT_BADGE_FALLBACK;
}

/* Badges de estado — cores intuitivas, com o "Agendado" corrigido para
   Laranja/Âmbar vivo (deixou de ser azul claro, que tinha pouco contraste). */
const STATUS_BADGE_CLASSES = {
  "Agendado": "bg-orange-100 text-orange-800 border-orange-300 font-semibold",
  "Aprovado": "bg-emerald-100 text-emerald-800 border-emerald-300 font-medium",
  "Rejeitado": "bg-rose-100 text-rose-800 border-rose-300 font-medium",
  "Sem Horário Comum": "bg-rose-100 text-rose-800 border-rose-300 font-medium",
  "Pendente": "bg-slate-200 text-slate-600 border-slate-300 font-medium",
};
const STATUS_BADGE_FALLBACK = "bg-slate-100 text-slate-400 border-slate-200 font-medium";
function statusBadgeClass(status) {
  return STATUS_BADGE_CLASSES[status] || STATUS_BADGE_FALLBACK;
}

/* Blocos de métricas (StatCard) — variantes usando apenas a paleta oficial */
const STAT_TONES = {
  neutral: { tile: COLORS.mint, icon: COLORS.navy },
  brand: { tile: COLORS.pink, icon: COLORS.navy },
  alert: { tile: COLORS.navy, icon: COLORS.pink },
  critical: { tile: COLORS.navy, icon: COLORS.white },
};

/* ============================================================================
   HELPERS
============================================================================ */

function hashStr(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
  return h;
}
function mulberry32(seed) {
  let a = seed;
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function pickAvailability(seedName, ratio = 0.5) {
  const rng = mulberry32(hashStr(seedName));
  const withKeys = SLOTS.map((s) => ({ s, k: rng() }));
  withKeys.sort((a, b) => a.k - b.k);
  const n = Math.max(4, Math.round(SLOTS.length * ratio));
  return withKeys.slice(0, n).map((x) => x.s).sort((a, b) => SLOTS.indexOf(a) - SLOTS.indexOf(b));
}
function uid(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 9)}`;
}
function slugify(s) {
  return s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, ".").replace(/^\.|\.$/g, "");
}
function downloadCSV(filename, rows) {
  const csv = rows.map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        resolve(XLSX.read(data, { type: "array" }));
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
function get(row, ...keys) {
  for (const k of keys) {
    for (const rk of Object.keys(row)) {
      if (rk.trim().toLowerCase() === k.toLowerCase()) return row[rk];
    }
  }
  return "";
}
function matchDept(raw) {
  const norm = String(raw || "").trim().toLowerCase();
  return DEPARTMENTS.find((d) => d.toLowerCase() === norm) || null;
}

/* ============================================================================
   SINCRONIZAÇÃO EM TEMPO REAL — EXCEL MESTRE (GOOGLE SHEETS)
   Lê diretamente da folha de cálculo mestre da YME, aba a aba, pelo nome
   exato de cada aba. Só as abas listadas abaixo são consultadas — todas
   as restantes (Capa, Utilização, Dashboard Candidatos, Organização
   Entrevistas RH, Organização Dinâmicas, Organização Entrevista Final, etc.)
   são sempre ignoradas.
============================================================================ */

const SYNC_URL_STORAGE_KEY = "yme_master_sheet_url";
const SYNC_POLL_INTERVAL_MS = 2 * 60 * 1000; // 2 em 2 minutos

// A. Abas gerais de base de dados e disponibilidades (nomes exatos das abas com aspas simples para a API do Google).
const SYNC_SHEET_NAMES = {
  departamentos: "'Base Dados Departamentos'",
  candidatos: "'Base Dados Candidatos'",
  avaliacaoCV: "'Avaliação CV e Questões Abertas'",
  dispEntrevistasRH: "'Disponibilidade Entrevistas RH'",
  dispDinamicas: "'Disponibilidade Dinâmicas'",
  dispEntrevistaFinal: "'Disponibilidade Entrevista Final'",
};

// B. Colunas de posição fixa (A=coluna 1) consultadas em cada aba de departamento.
const SYNC_DEPT_COLUMNS = {
  softSkills: "L",   // Passou Entrevista Soft Skills/RH -> avança Fase 2 (Dinâmicas)
  dinamicas: "AA",   // Passou Dinâmicas de Grupo -> avança Fase 3 (Desafio Final)
  final: "AZ",       // Passou Desafio Final/Hard Skills -> SELECIONADO / ENTROU NA YME
  talentPoolA: "AB", // Selecionado para Talent Pool (variante de coluna 1)
  talentPoolB: "BA", // Selecionado para Talent Pool (variante de coluna 2)
};
// Coluna Q da Avaliação de CV: candidato passou para a Fase 2 (Entrevista Soft Skills/RH).
const SYNC_CV_PASS_COLUMN = "Q";

/* ----------------------------------------------------------------------
   Autenticação Google (OAuth 2.0 / Google Identity Services) + leitura
   segura via Google Sheets API v4. A folha mestre é PRIVADA (RGPD): nunca
   é pedida por link público nem por endpoint CSV/gviz — só é lida com o
   Token de Acesso da conta Google autorizada e autenticada do utilizador,
   através da API oficial "sheets.googleapis.com".
------------------------------------------------------------------------- */

// Client ID OAuth 2.0 (tipo "Web application") criado na Google Cloud
// Console do projeto da YME, com o domínio desta WebApp autorizado em
// "Authorized JavaScript origins". SUBSTITUIR pelo Client ID real antes de
// publicar — sem isto a autenticação não arranca.
// Pode também ser definido via variável de ambiente na build
// (REACT_APP_GOOGLE_CLIENT_ID), que tem prioridade sobre o valor fixo
// abaixo. Propositadamente não usamos aqui nenhuma sintaxe de "meta" de
// módulo ES para ler variáveis de ambiente do Vite: em ambientes que
// corram este ficheiro como script comum (ex.: pré-visualização de
// artefactos) isso provoca um erro de parsing fatal, mesmo dentro de uma
// verificação defensiva. Se o teu setup usa Vite e precisas de ler
// VITE_GOOGLE_CLIENT_ID, faz essa leitura fora deste ficheiro (por
// exemplo, injetando o valor num ficheiro de configuração separado que
// só é processado pelo Vite) e passa o resultado para cá.
const GOOGLE_CLIENT_ID = "1073691932169-dl8eu8rlsknece09vv6d53hacgeo2h5r.apps.googleusercontent.com";



// Âmbitos pedidos: leitura da folha (spreadsheets.readonly) + identidade
// básica (openid/email/profile), só para mostrar "Sessão Ativa: conta@yme.pt"
// na interface. Nunca é pedido acesso de escrita à folha.
const GOOGLE_SHEETS_SCOPES = "openid email profile https://www.googleapis.com/auth/spreadsheets.readonly";

function loadGoogleIdentityScript() {
  return new Promise((resolve, reject) => {
    if (typeof window === "undefined") { reject(new Error("no_window")); return; }
    if (window.google?.accounts?.oauth2) { resolve(); return; }
    const existing = document.getElementById("google-identity-services");
    if (existing) {
      existing.addEventListener("load", () => resolve());
      existing.addEventListener("error", () => reject(new Error("load_failed")));
      return;
    }
    const script = document.createElement("script");
    script.id = "google-identity-services";
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.defer = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("load_failed"));
    document.head.appendChild(script);
  });
}

// Hook de autenticação Google (OAuth 2.0 implicit/token flow via Google
// Identity Services). Devolve o estado da sessão (accessToken, email,
// prontidão) e duas ações: requestToken (login/renovação) e signOut.
function useGoogleAuth() {
  const [auth, setAuth] = useState({
    ready: false,
    authenticating: false,
    accessToken: null,
    tokenExpiresAt: null,
    email: null,
    error: null,
  });
  const tokenClientRef = useRef(null);

  useEffect(() => {
    let cancelled = false;
    loadGoogleIdentityScript()
      .then(() => {
        if (cancelled) return;
        if (!GOOGLE_CLIENT_ID || GOOGLE_CLIENT_ID.startsWith("SUBSTITUIR_")) {
          setAuth((a) => ({
            ...a,
            ready: false,
            error: "GOOGLE_CLIENT_ID não configurado. Define o Client ID OAuth da YME na constante GOOGLE_CLIENT_ID (ou na variável de ambiente REACT_APP_GOOGLE_CLIENT_ID).",
          }));
          return;
        }
        tokenClientRef.current = window.google.accounts.oauth2.initTokenClient({
          client_id: GOOGLE_CLIENT_ID,
          scope: GOOGLE_SHEETS_SCOPES,
          callback: () => {}, // é substituído a cada pedido em requestToken()
        });
        setAuth((a) => ({ ...a, ready: true }));
      })
      .catch(() => {
        if (!cancelled) {
          setAuth((a) => ({ ...a, ready: false, error: "Não foi possível carregar o serviço de autenticação da Google. Confirma a tua ligação à internet." }));
        }
      });
    return () => { cancelled = true; };
  }, []);

  // Pede (ou renova) o token de acesso. { silent: true } tenta renovar sem
  // mostrar o ecrã de consentimento (usado para renovação automática em
  // segundo plano); se a sessão da Google já não for válida, falha e quem
  // chama deve pedir novo login explícito ao utilizador.
  const requestToken = ({ silent = false } = {}) => {
    return new Promise((resolve, reject) => {
      if (!tokenClientRef.current) {
        reject(new Error("O serviço de autenticação Google ainda não está pronto. Tenta novamente em instantes."));
        return;
      }
      tokenClientRef.current.callback = async (resp) => {
        if (resp.error) {
          setAuth((a) => ({ ...a, authenticating: false, error: "Autenticação Google recusada ou cancelada." }));
          reject(new Error(resp.error));
          return;
        }
        const expiresAt = Date.now() + Number(resp.expires_in || 3600) * 1000;
        let email = null;
        try {
          const profileRes = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
            headers: { Authorization: `Bearer ${resp.access_token}` },
          });
          if (profileRes.ok) {
            const profile = await profileRes.json();
            email = profile.email || null;
          }
        } catch {
          // Não crítico — falha em obter o email só afeta o texto exibido,
          // nunca a sincronização em si.
        }
        setAuth({ ready: true, authenticating: false, accessToken: resp.access_token, tokenExpiresAt: expiresAt, email, error: null });
        resolve({ accessToken: resp.access_token, email });
      };
      setAuth((a) => ({ ...a, authenticating: !silent, error: null }));
      tokenClientRef.current.requestAccessToken({ prompt: silent ? "" : "consent" });
    });
  };

  const signOut = () => {
    setAuth((a) => {
      if (a.accessToken && window.google?.accounts?.oauth2?.revoke) {
        window.google.accounts.oauth2.revoke(a.accessToken, () => {});
      }
      return { ...a, accessToken: null, tokenExpiresAt: null, email: null, error: null };
    });
  };

  return { auth, requestToken, signOut };
}

function extractSheetId(url) {
  // Aceita qualquer variante do link do Google Sheets, ou o próprio ID puro:
  //  - .../spreadsheets/d/{ID}/edit?usp=sharing
  //  - .../spreadsheets/d/{ID}/edit#gid=123
  //  - .../spreadsheets/d/{ID}  (sem sufixo)
  //  - {ID}  (ID isolado, sem URL)
  // Limpa primeiro qualquer sufixo de query-string/fragmento
  // (?usp=sharing, #gid=123, etc.) antes de procurar o ID, para nunca
  // deixar "vazar" lixo de URL para dentro do ID capturado.
  const clean = String(url || "").trim().split(/[?#]/)[0];
  const m = clean.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/) || clean.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  // Se não é um URL reconhecível, aceita o próprio valor como ID direto
  // (útil quando o utilizador cola apenas o ID do documento).
  return /^[a-zA-Z0-9-_]{20,}$/.test(clean) ? clean : null;
}

// Conversão de letra de coluna (A, L, AA, AZ, BA, ...) para índice de array (0-based).
function colLetterToIndex(letter) {
  let n = 0;
  const s = String(letter || "").trim().toUpperCase();
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

function isPositiveMark(val) {
  const v = String(val ?? "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (!v) return false;
  return ["x", "sim", "true", "1", "aprovado", "selecionado", "apto", "avanca", "yes", "✓", "v"].includes(v);
}
// "pending" = célula vazia (ainda sem decisão); "positive"/"negative" = célula preenchida.
function cellStatus(val) {
  const s = String(val ?? "").trim();
  if (!s) return "pending";
  return isPositiveMark(s) ? "positive" : "negative";
}

// Converte a grelha de valores devolvida pela Google Sheets API
// (values: [[...linha1], [...linha2], ...]) para { header, rows }, no
// mesmo formato usado por toda a lógica de mapeamento abaixo. Cada linha
// guarda tanto o objeto indexado pelo cabeçalho (colunas por nome, ex.
// "Nome", "Email") como o array bruto da linha (colunas por letra, ex.
// "Coluna L", "Coluna AZ").
function parseApiValues(values) {
  const grid = values || [];
  const header = (grid[0] || []).map((h) => String(h ?? ""));
  const rows = grid
    .slice(1)
    .filter((r) => r.some((c) => String(c ?? "").trim() !== ""))
    .map((raw) => {
      const obj = {};
      header.forEach((h, i) => { if (h) obj[h] = raw[i] ?? ""; });
      return { obj, raw };
    });
  return { header, rows };
}

// Lê uma aba da folha privada através da Google Sheets API v4
// (spreadsheets.values.get), autenticado com o accessToken da sessão do
// utilizador. encodeURIComponent no nome da aba é essencial: nomes como
// "Base Dados Departamentos" têm espaços e, sem isto, o range fica
// inválido/mal interpretado pela API.
async function fetchSheetTabApi(accessToken, sheetId, sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(sheetName)}`;
  let res;
  try {
    res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    });
  } catch {
    throw new Error("network_error");
  }

  if (res.status === 401) throw new Error("token_expired");
  if (res.status === 403) throw new Error("permission_denied");
  if (res.status === 400 || res.status === 404) throw new Error("sheet_not_found");
  if (!res.ok) throw new Error(`http_${res.status}`);

  let data;
  try {
    data = await res.json();
  } catch {
    throw new Error("read_error");
  }

  return parseApiValues(data.values);
}

// Converte os códigos de erro internos da API (token_expired,
// permission_denied, sheet_not_found, network_error, http_4xx/5xx) em
// mensagens compreensíveis para o utilizador final.
function translateSheetApiError(code, sheetName) {
  if (code === "token_expired") {
    return new Error(`A sessão Google expirou ao ler a aba "${sheetName}".`);
  }
  if (code === "permission_denied") {
    return new Error(`A tua conta Google não tem permissão para ler a aba "${sheetName}". Confirma que a folha privada foi partilhada com a tua conta @yme.pt (com acesso de Leitor).`);
  }
  if (code === "sheet_not_found") {
    return new Error(`Não foi encontrada nenhuma aba chamada exatamente "${sheetName}" nessa folha.`);
  }
  if (code === "network_error") {
    return new Error(`Não foi possível contactar a Google Sheets API para a aba "${sheetName}". Confirma a tua ligação à internet.`);
  }
  if (code === "read_error") {
    return new Error(`Não foi possível ler a resposta da aba "${sheetName}".`);
  }
  if (code?.startsWith("http_")) {
    const status = code.replace("http_", "");
    return new Error(`A aba "${sheetName}" respondeu com erro HTTP ${status}. Confirma o ID da folha e as permissões de partilha da tua conta.`);
  }
  return new Error(`Não foi possível ler a aba "${sheetName}" (${code}).`);
}

// Extrai disponibilidade de uma linha, aceitando 2 formatos comuns:
// (a) colunas-grelha em que o próprio slot é o cabeçalho ("Seg 09:00" = x/sim)
// (b) uma coluna de texto livre "Disponibilidade" com slots separados por | ; ,
function extractAvailabilityFromRow(header, row) {
  const slots = new Set();
  header.forEach((h) => {
    const clean = String(h || "").trim();
    if (SLOTS.includes(clean) && isPositiveMark(row.obj[h])) slots.add(clean);
  });
  const free = String(get(row.obj, "disponibilidade", "horarios", "horários", "slots") || "");
  free.split(/[|;,]/).map((s) => s.trim()).forEach((s) => { if (SLOTS.includes(s)) slots.add(s); });
  return Array.from(slots).sort((a, b) => SLOTS.indexOf(a) - SLOTS.indexOf(b));
}

// Aplica os dados brutos das abas do Excel Mestre ao estado de members/candidates
// da aplicação, seguindo o mapeamento estrito de abas e colunas da YME.
function applySyncedSheetsToState(raw, prevMembers, prevCandidates) {
  /* ---- A1. Base Dados Departamentos -> membros (Diretor/Supervisor/RH) ---- */
  let members = prevMembers.map((m) => ({ ...m }));
  const upsertMember = (name, role, dept) => {
    name = String(name || "").trim();
    if (!name) return;
    const idx = members.findIndex((m) => m.name.toLowerCase() === name.toLowerCase());
    if (idx >= 0) {
      const depts = new Set(members[idx].departments || []);
      if (dept) depts.add(dept);
      members[idx] = { ...members[idx], role: role || members[idx].role, departments: Array.from(depts) };
    } else {
      members.push({ id: uid("sync"), name, role: role || "RH", title: role || "Membro", departments: dept ? [dept] : [], availability: [] });
    }
  };
  (raw.departamentos?.rows || []).forEach((row) => {
    const dept = matchDept(get(row.obj, "departamento", "department"));
    const diretor = get(row.obj, "diretor", "diretor(a)");
    const supervisor = get(row.obj, "supervisor");
    const rh = String(get(row.obj, "rh", "membro rh", "membros rh", "membro de rh") || "");
    if (diretor) upsertMember(diretor, "Diretor", dept);
    if (supervisor) upsertMember(supervisor, "Supervisor", dept);
    rh.split(/[,;|]/).map((s) => s.trim()).filter(Boolean).forEach((n) => upsertMember(n, "RH", dept));
  });

  /* ---- A4/A5/A6. Disponibilidades de RH/Diretores/Supervisores -> membros ---- */
  [raw.dispEntrevistasRH, raw.dispDinamicas, raw.dispEntrevistaFinal].forEach((tab) => {
    (tab?.rows || []).forEach((row) => {
      const name = String(get(row.obj, "nome", "name") || "").trim();
      if (!name) return;
      const slots = extractAvailabilityFromRow(tab.header, row);
      if (!slots.length) return;
      const idx = members.findIndex((m) => m.name.toLowerCase() === name.toLowerCase());
      if (idx >= 0) {
        members[idx] = { ...members[idx], availability: Array.from(new Set([...(members[idx].availability || []), ...slots])) };
      } else {
        members.push({ id: uid("sync"), name, role: "RH", title: "Membro", departments: [], availability: slots });
      }
    });
  });

  /* ---- A2. Base Dados Candidatos -> dados base de cada candidato ---- */
  let candidates = prevCandidates.map((c) => ({ ...c }));
  const upsertCandidate = (patch) => {
    const idx = candidates.findIndex((c) => c.name.toLowerCase() === patch.name.toLowerCase());
    if (idx >= 0) candidates[idx] = { ...candidates[idx], ...patch };
    else candidates.push({
      id: uid("sync"), phase0Status: "Pendente", phase1Status: "—", phase2Status: "—",
      formsSubmitted: { fase1: false, fase2: false, fase3: false },
      availability: { fase1: [], fase2: [], fase3: [] },
      ...patch,
    });
  };
  (raw.candidatos?.rows || []).forEach((row) => {
    const name = String(get(row.obj, "nome completo", "nome", "name") || "").trim();
    if (!name) return;
    const dept1 = matchDept(get(row.obj, "primeira opcao", "primeira opção", "1a opcao", "1ª opção", "departamento"));
    const dept2 = matchDept(get(row.obj, "segunda opcao", "segunda opção", "2a opcao", "2ª opção"));
    upsertCandidate({
      name,
      email: String(get(row.obj, "email", "contacto", "email/contacto") || "").trim(),
      department: dept1 || DEPARTMENTS[0],
      segundaOpcaoDepartamento: dept2 || null,
      curso: String(get(row.obj, "curso") || "").trim(),
      anoLetivo: String(get(row.obj, "ano letivo", "ano") || "").trim(),
      erasmus: String(get(row.obj, "erasmus") || "").trim(),
    });
  });

  /* ---- A3. Avaliação CV e Questões Abertas -> Coluna Q (avança p/ Fase 2 = Entrevista RH) ---- */
  (raw.avaliacaoCV?.rows || []).forEach((row) => {
    const name = String(get(row.obj, "nome", "nome completo", "name") || "").trim();
    if (!name) return;
    const status = cellStatus(row.raw[colLetterToIndex(SYNC_CV_PASS_COLUMN)]);
    if (status === "pending") return;
    const patch = { name, phase0Status: status === "positive" ? "Aprovado" : "Rejeitado" };
    const idx = candidates.findIndex((c) => c.name.toLowerCase() === name.toLowerCase());
    if (idx >= 0) candidates[idx] = { ...candidates[idx], ...patch };
    else upsertCandidate({ ...patch, department: DEPARTMENTS[0] });
  });

  /* ---- B. Abas por Departamento -> progresso por fase + lógica de rejeição ---- */
  DEPARTMENTS.forEach((dept) => {
    const tab = raw.deptTabs?.[dept];
    if (!tab) return;
    tab.rows.forEach((row) => {
      const name = String(get(row.obj, "nome", "nome completo", "name") || "").trim();
      if (!name) return;
      const idx = candidates.findIndex((c) => c.name.toLowerCase() === name.toLowerCase());
      if (idx < 0) return; // candidato tem de constar em "Base Dados Candidatos"
      const cand = candidates[idx];
      const at = (letter) => row.raw[colLetterToIndex(letter)];

      const softSkills = cellStatus(at(SYNC_DEPT_COLUMNS.softSkills));
      const dinamicas = cellStatus(at(SYNC_DEPT_COLUMNS.dinamicas));
      const final = cellStatus(at(SYNC_DEPT_COLUMNS.final));
      const talentPoolMark = isPositiveMark(at(SYNC_DEPT_COLUMNS.talentPoolA)) || isPositiveMark(at(SYNC_DEPT_COLUMNS.talentPoolB));
      const anyColumnFilled = [
        SYNC_DEPT_COLUMNS.softSkills, SYNC_DEPT_COLUMNS.dinamicas, SYNC_DEPT_COLUMNS.final,
        SYNC_DEPT_COLUMNS.talentPoolA, SYNC_DEPT_COLUMNS.talentPoolB,
      ].some((letter) => String(at(letter) ?? "").trim() !== "");
      const anyPositive = softSkills === "positive" || dinamicas === "positive" || final === "positive" || talentPoolMark;

      let phase1Status = cand.phase1Status;
      if (softSkills === "positive") phase1Status = "Aprovado";
      else if (softSkills === "negative") phase1Status = "Rejeitado";

      let phase2Status = cand.phase2Status;
      if (dinamicas === "positive") phase2Status = "Aprovado";
      else if (dinamicas === "negative") phase2Status = "Rejeitado";

      let finalResult = cand.finalResult || null;
      if (final === "positive") finalResult = "Selecionado — Entrou na YME";
      else if (final === "negative") finalResult = talentPoolMark ? "Talent Pool" : "Rejeitado";

      // Lógica de rejeição automática: só atua quando há efetivamente alguma
      // marcação nas colunas de transição (não sobrepõe candidatos ainda por decidir).
      let phase0Status = cand.phase0Status;
      if (anyColumnFilled && !anyPositive && cand.phase0Status === "Aprovado") {
        phase0Status = "Rejeitado";
      }

      candidates[idx] = {
        ...cand, department: dept, phase0Status, phase1Status, phase2Status,
        finalResult, talentPool: talentPoolMark || cand.talentPool || false,
      };
    });
  });

  return { members, candidates };
}

// Sentinela lançada quando alguma aba falhou por token de acesso expirado/
// inválido (HTTP 401). Quem chama (App) apanha este erro especificamente
// para tentar uma renovação silenciosa da sessão Google antes de desistir.
const SYNC_SESSION_EXPIRED = "SESSAO_EXPIRADA";
const SYNC_SESSION_REQUIRED = "SESSAO_NECESSARIA";

// Orquestra a leitura de todas as abas mapeadas do Excel Mestre — agora
// através da Google Sheets API v4, autenticada com o accessToken da sessão
// Google do utilizador — e devolve o novo estado de members/candidates da
// aplicação. Ignora sempre quaisquer abas fora de SYNC_SHEET_NAMES /
// DEPARTMENTS. Nunca faz fetch() a um link público ou endpoint CSV — a
// folha permanece 100% privada, partilhada apenas com contas Google
// autorizadas da YME.
async function syncMasterSheet({ accessToken, sheetUrl, prevMembers, prevCandidates }) {
  const sheetId = extractSheetId(sheetUrl);
  if (!sheetId) {
    throw new Error('Não consegui encontrar o ID da folha. Cola o link de partilha completo do Google Sheets (ex: ".../spreadsheets/d/ID_DA_FOLHA/edit") ou apenas o ID do documento.');
  }
  if (!accessToken) {
    throw new Error(SYNC_SESSION_REQUIRED);
  }

  let tokenExpired = false;
  const readTab = async (key, sheetName) => {
    try {
      const data = await fetchSheetTabApi(accessToken, sheetId, sheetName);
      return [key, data];
    } catch (err) {
      if (err.message === "token_expired") tokenExpired = true;
      return [key, { error: translateSheetApiError(err.message, sheetName).message }];
    }
  };

  let generalResults, deptResults;
  try {
    [generalResults, deptResults] = await Promise.all([
      Promise.all(Object.entries(SYNC_SHEET_NAMES).map(([key, sheetName]) => readTab(key, sheetName))),
      Promise.all(DEPARTMENTS.map((dept) => readTab(dept, dept))),
    ]);
  } catch (fatalErr) {
    // Salvaguarda: qualquer falha inesperada (ex. fetch() indisponível no
    // ambiente) é convertida numa mensagem compreensível em vez de rebentar.
    throw new Error(`Falha inesperada ao sincronizar: ${fatalErr.message || fatalErr}`);
  }

  // Se qualquer aba falhou por sessão expirada, é preferível tentar renovar
  // o token e repetir tudo, em vez de avançar com dados parciais/incoerentes.
  if (tokenExpired) {
    throw new Error(SYNC_SESSION_EXPIRED);
  }

  const raw = { deptTabs: {} };
  const errors = [];
  generalResults.forEach(([key, data]) => { if (data.error) errors.push(data.error); else raw[key] = data; });
  deptResults.forEach(([dept, data]) => { if (data.error) errors.push(data.error); else raw.deptTabs[dept] = data; });

  if (!raw.candidatos && !raw.departamentos) {
    // Nenhuma aba geral foi lida — quase sempre é permissões de partilha
    // da conta Google autenticada, não do link em si (a folha é privada).
    const uniqueErrors = Array.from(new Set(errors));
    throw new Error(
      uniqueErrors[0] ||
      'Não foi possível ler nenhuma aba da folha privada. Confirma que a folha foi partilhada com a tua conta Google (@yme.pt) com acesso de Leitor, e que o ID/link está correto.'
    );
  }

  const { members, candidates } = applySyncedSheetsToState(raw, prevMembers, prevCandidates);
  return { members, candidates, errors };
}

/* ============================================================================
   MOCK DATA — MEMBROS (fixos) + 100+ CANDIDATOS
============================================================================ */

function buildMockMembers() {
  const members = [];
  const supervisorsSeen = new Map();
  ORG.forEach((o) => {
    members.push({ id: uid("dir"), name: o.diretor, role: "Diretor", title: "Diretor(a)", departments: [o.dept], availability: pickAvailability(o.diretor, 0.55) });
    o.rh.forEach((name) => {
      members.push({ id: uid("rh"), name, role: "RH", title: "Membro RH", departments: [o.dept], availability: pickAvailability(name, 0.6) });
    });
    if (supervisorsSeen.has(o.supervisor)) {
      supervisorsSeen.get(o.supervisor).departments.push(o.dept);
    } else {
      const m = { id: uid("sup"), name: o.supervisor, role: "Supervisor", title: o.supervisorTitle, departments: [o.dept], availability: pickAvailability(o.supervisor, 0.45) };
      supervisorsSeen.set(o.supervisor, m);
      members.push(m);
    }
  });
  return members;
}

const FIRST_NAMES = ["Ana", "Beatriz", "Carolina", "Diana", "Eduarda", "Filipa", "Gabriela", "Helena", "Inês", "Joana", "Leonor", "Madalena", "Mariana", "Matilde", "Rita", "Sara", "Sofia", "Vera", "Alexandre", "Bernardo", "Carlos", "Diogo", "Eduardo", "Filipe", "Gonçalo", "Hugo", "Ivo", "João", "Leonardo", "Miguel", "Nuno", "Pedro", "Ricardo", "Rodrigo", "Samuel", "Simão", "Tiago", "Tomás", "Vasco", "Xavier"];
const LAST_NAMES = ["Silva", "Santos", "Ferreira", "Pereira", "Oliveira", "Costa", "Rodrigues", "Martins", "Sousa", "Fernandes", "Gonçalves", "Gomes", "Lopes", "Marques", "Alves", "Almeida", "Ribeiro", "Carvalho", "Teixeira", "Moreira", "Correia", "Mendes", "Nunes", "Soares", "Vieira", "Monteiro", "Cardoso", "Rocha", "Neves", "Coelho", "Cunha", "Pinto", "Ramos", "Reis", "Simões", "Antunes", "Matos", "Fonseca", "Machado", "Araújo"];

function generateNamePool(n) {
  const rng = mulberry32(hashStr("yme-candidatos-2026"));
  const used = new Set();
  const out = [];
  while (out.length < n) {
    const fn = FIRST_NAMES[Math.floor(rng() * FIRST_NAMES.length)];
    const ln = LAST_NAMES[Math.floor(rng() * LAST_NAMES.length)];
    const full = `${fn} ${ln}`;
    if (used.has(full)) continue;
    used.add(full);
    out.push(full);
  }
  return out;
}

function buildMockCandidates(total = 108) {
  const names = generateNamePool(total);
  return names.map((name, i) => {
    const department = DEPARTMENTS[i % DEPARTMENTS.length];
    const phase0Status = i % 13 === 0 ? "Rejeitado" : i % 8 === 0 ? "Pendente" : "Aprovado";

    const f1 = phase0Status === "Aprovado" && i % 11 !== 0;
    const phase1Status = phase0Status !== "Aprovado" ? "—" : !f1 ? "—" : i % 17 === 0 ? "Rejeitado" : i % 9 === 0 ? "Pendente" : "Aprovado";

    const f2 = phase1Status === "Aprovado" && i % 10 !== 0;
    const phase2Status = phase1Status !== "Aprovado" ? "—" : !f2 ? "—" : i % 19 === 0 ? "Rejeitado" : i % 12 === 0 ? "Pendente" : "Aprovado";

    const f3 = phase2Status === "Aprovado" && i % 9 !== 0;

    return {
      id: uid("cand"), name, department,
      email: `${slugify(name)}@candidatos.yme.pt`,
      cvLink: `https://forms.yme.pt/questionario/${slugify(name.split(" ")[0])}`,
      phase0Status, phase1Status, phase2Status,
      formsSubmitted: { fase1: f1, fase2: f2, fase3: f3 },
      availability: {
        fase1: f1 ? pickAvailability(name + "-f1", 0.5) : [],
        fase2: f2 ? pickAvailability(name + "-f2", 0.5) : [],
        fase3: f3 ? pickAvailability(name + "-f3", 0.5) : [],
      },
    };
  });
}

/* ============================================================================
   SCHEDULING ALGORITHMS  (lógica inalterada)
============================================================================ */

function generateInterviewPhase(pool, members, existingBookings, availField, staffKeys) {
  // staffKeys: array like ["diretorId","rhId"] or ["diretorId","rhId","supervisorId"]
  const busy = {};
  const kept = existingBookings.filter((b) => b.manual && pool.some((c) => c.id === b.candidateId));
  kept.forEach((b) => {
    if (!b.slot) return;
    staffKeys.forEach((k) => {
      if (!b[k]) return;
      busy[b[k]] = busy[b[k]] || new Set();
      busy[b[k]].add(b.slot);
    });
  });
  const bookings = [...kept];
  pool.forEach((c) => {
    if (kept.some((b) => b.candidateId === c.id)) return;
    const diretor = members.find((m) => m.role === "Diretor" && m.departments.includes(c.department));
    const supervisor = staffKeys.includes("supervisorId") ? members.find((m) => m.role === "Supervisor" && m.departments.includes(c.department)) : null;
    const rhList = members.filter((m) => m.role === "RH" && m.departments.includes(c.department));

    let found = null;
    for (const slot of c.availability[availField]) {
      if (!diretor || !diretor.availability.includes(slot) || busy[diretor.id]?.has(slot)) continue;
      if (staffKeys.includes("supervisorId")) {
        if (!supervisor || !supervisor.availability.includes(slot) || busy[supervisor.id]?.has(slot)) continue;
      }
      const rh = rhList.find((r) => r.availability.includes(slot) && !busy[r.id]?.has(slot));
      if (rh) { found = { slot, diretor, rh, supervisor }; break; }
    }

    if (found) {
      const record = { id: uid("bk"), candidateId: c.id, slot: found.slot, diretorId: found.diretor.id, rhId: found.rh.id, status: "Agendado", manual: false };
      if (staffKeys.includes("supervisorId")) record.supervisorId = found.supervisor.id;
      staffKeys.forEach((k) => { const mid = record[k]; if (mid) { busy[mid] = busy[mid] || new Set(); busy[mid].add(found.slot); } });
      bookings.push(record);
    } else {
      const record = { id: uid("bk"), candidateId: c.id, slot: null, diretorId: diretor?.id || null, rhId: null, status: "Sem Horário Comum", manual: false };
      if (staffKeys.includes("supervisorId")) record.supervisorId = supervisor?.id || null;
      bookings.push(record);
    }
  });
  return bookings;
}

function generatePhase2(pool, members) {
  const buckets = {};
  DEPARTMENTS.forEach((d) => (buckets[d] = pool.filter((c) => c.department === d).slice()));
  const groupsOfCandidates = [];
  let current = [];
  let remaining = true;
  while (remaining) {
    remaining = false;
    for (const d of DEPARTMENTS) {
      if (buckets[d].length) {
        current.push(buckets[d].shift());
        remaining = true;
        if (current.length === 6) { groupsOfCandidates.push(current); current = []; }
      }
    }
  }
  if (current.length) groupsOfCandidates.push(current);

  const busy = {};
  return groupsOfCandidates.map((group, idx) => {
    const freq = {};
    group.forEach((c) => c.availability.fase2.forEach((s) => (freq[s] = (freq[s] || 0) + 1)));
    let bestSlot = null, bestCount = -1;
    SLOTS.forEach((s) => {
      const cnt = freq[s] || 0;
      if (cnt > bestCount) { bestCount = cnt; bestSlot = cnt > 0 ? s : bestSlot; }
    });

    const warnings = [];
    const missingCandidates = bestSlot ? group.filter((c) => !c.availability.fase2.includes(bestSlot)) : group;
    if (!bestSlot) warnings.push("Nenhum horário comum encontrado entre os candidatos do grupo.");
    else if (missingCandidates.length) warnings.push(`${missingCandidates.length} candidato(s) indisponível(eis) no horário escolhido: ${missingCandidates.map((c) => c.name).join(", ")}.`);

    const deptCounts = {};
    group.forEach((c) => (deptCounts[c.department] = (deptCounts[c.department] || 0) + 1));
    Object.entries(deptCounts).forEach(([d, n]) => { if (n > 2) warnings.push(`${n} candidatos do mesmo departamento (${d}) na mesma sessão — máximo recomendado: 2.`); });

    let supervisor = null;
    if (bestSlot) {
      supervisor = members.find((m) => m.role === "Supervisor" && m.availability.includes(bestSlot) && !busy[m.id]?.has(bestSlot));
      if (supervisor) { busy[supervisor.id] = busy[supervisor.id] || new Set(); busy[supervisor.id].add(bestSlot); }
      else warnings.push("Nenhum Supervisor (CEO/COO/CMO) disponível neste horário.");
    }

    const depts = [...new Set(group.map((c) => c.department))];
    const directorIds = [];
    if (bestSlot) {
      depts.forEach((d) => {
        const dir = members.find((m) => m.role === "Diretor" && m.departments.includes(d));
        if (dir && dir.availability.includes(bestSlot) && !busy[dir.id]?.has(bestSlot)) {
          directorIds.push(dir.id); busy[dir.id] = busy[dir.id] || new Set(); busy[dir.id].add(bestSlot);
        } else {
          warnings.push(`Diretor(a) de ${d} indisponível — presença prioritária, mas não bloqueante.`);
        }
      });
    }

    let rhIds = [];
    if (bestSlot) {
      const freeRH = members
        .filter((m) => m.role === "RH" && m.availability.includes(bestSlot) && !busy[m.id]?.has(bestSlot))
        .sort((a, b) => {
          const am = depts.some((d) => a.departments.includes(d)) ? 1 : 0;
          const bm = depts.some((d) => b.departments.includes(d)) ? 1 : 0;
          return bm - am;
        });
      rhIds = freeRH.slice(0, 3).map((r) => r.id);
      rhIds.forEach((id) => { busy[id] = busy[id] || new Set(); busy[id].add(bestSlot); });
      if (rhIds.length < 2) warnings.push("Menos de 2 membros de RH disponíveis para esta sessão.");
    }

    return {
      id: uid("p2"),
      name: `Grupo ${String.fromCharCode(65 + idx)}`,
      candidateIds: group.map((c) => c.id),
      slot: bestSlot,
      supervisorId: supervisor?.id || null,
      directorIds, rhIds, warnings,
    };
  });
}

const LOGO_DARK_DATA_URI = "data:image/png;base64,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";
const LOGO_LIGHT_DATA_URI = "data:image/png;base64,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";
/* ============================================================================
   SMALL SHARED UI
============================================================================ */

function Badge({ children, className = "", style }) {
  return (
    <span
      className={`inline-flex items-center px-2 py-0.5 rounded-md text-xs font-medium border ${className}`}
      style={style}
    >
      {children}
    </span>
  );
}
function DeptBadge({ dept }) {
  return <Badge className={deptBadgeClass(dept)}>{dept}</Badge>;
}
function StatusBadge({ status }) {
  return <Badge className={statusBadgeClass(status)}>{status}</Badge>;
}
function StatCard({ label, value, icon: Icon, tone = "neutral" }) {
  const t = STAT_TONES[tone] || STAT_TONES.neutral;
  return (
    <div className="rounded-xl border p-4 flex items-center gap-3" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
      <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ backgroundColor: t.tile, color: t.icon }}>
        <Icon size={18} />
      </div>
      <div>
        <p className="text-2xl font-bold leading-none font-mono" style={{ color: COLORS.navy }}>{value}</p>
        <p className="text-xs mt-1" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{label}</p>
      </div>
    </div>
  );
}
function Modal({ title, onClose, children, wide = false }) {
  return (
    <div className="fixed inset-0 flex items-center justify-center z-50 p-4" style={{ backgroundColor: hexToRgba(COLORS.navy, 0.75) }} onClick={onClose}>
      <div
        className={`rounded-2xl shadow-2xl w-full ${wide ? "max-w-2xl" : "max-w-md"} max-h-[85vh] overflow-y-auto`}
        style={{ backgroundColor: COLORS.mint }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 sticky top-0 rounded-t-2xl" style={{ backgroundColor: COLORS.mint, borderBottom: `1px solid ${hexToRgba(COLORS.navy, 0.12)}` }}>
          <h3 className="font-semibold" style={{ color: COLORS.navy }}>{title}</h3>
          <button onClick={onClose} style={{ color: hexToRgba(COLORS.navy, 0.5) }}><X size={18} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}
function Field({ label, children }) {
  return (
    <label className="block mb-3">
      <span className="block text-xs font-medium mb-1" style={{ color: hexToRgba(COLORS.navy, 0.65) }}>{label}</span>
      {children}
    </label>
  );
}
const inputCls = "yme-input w-full rounded-lg px-3 py-2 text-sm";

function usePagination(items, pageSize, resetKey) {
  const [page, setPage] = useState(1);
  useEffect(() => { setPage(1); }, [resetKey]);
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));
  const clamped = Math.min(page, totalPages);
  const pageItems = items.slice((clamped - 1) * pageSize, clamped * pageSize);
  return { page: clamped, setPage, totalPages, pageItems };
}
function PaginationBar({ page, setPage, totalPages }) {
  if (totalPages <= 1) return null;
  return (
    <div className="flex items-center justify-between px-4 py-3 text-xs" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}`, color: hexToRgba(COLORS.navy, 0.55) }}>
      <span>Página {page} de {totalPages}</span>
      <div className="flex gap-1.5">
        <button disabled={page === 1} onClick={() => setPage((p) => Math.max(1, p - 1))} className="yme-btn-outline-light px-2.5 py-1 rounded disabled:opacity-40">Anterior</button>
        <button disabled={page === totalPages} onClick={() => setPage((p) => Math.min(totalPages, p + 1))} className="yme-btn-outline-light px-2.5 py-1 rounded disabled:opacity-40">Seguinte</button>
      </div>
    </div>
  );
}

/* ============================================================================
   LOGÓTIPO OFICIAL YME
============================================================================ */

/* Logótipo YME — duas variantes de imagem (mesma marca, cor adaptada ao fundo):
   - "dark"  → letras Azul Escuro #151E33 + sublinhado Rosa — usada sobre fundos claros (Login).
   - "light" → letras Brancas #FCFCFC + sublinhado Rosa — usada sobre o fundo escuro #151E33 da Navbar. */
const LOGO_SRC = { dark: LOGO_DARK_DATA_URI, light: LOGO_LIGHT_DATA_URI };

function Logo({ size = "default", variant = "dark" }) {
  const [imgError, setImgError] = useState(false);
  const h = size === "large" ? "h-12" : "h-8";

  if (!imgError) {
    return (
      <img
        src={LOGO_SRC[variant]}
        alt="YME"
        className={`${h} w-auto object-contain`}
        onError={() => setImgError(true)}
      />
    );
  }
  // Fallback tipográfico, caso a imagem não carregue.
  const isDark = variant === "dark";
  return (
    <div className="flex flex-col leading-none select-none">
      <span className={`font-bold tracking-wide ${size === "large" ? "text-4xl" : "text-xl"}`} style={{ color: isDark ? COLORS.navy : COLORS.white }}>yme</span>
      <span className="h-[3px] w-full rounded-full mt-1.5" style={{ backgroundColor: COLORS.pink }} />
    </div>
  );
}

/* ============================================================================
   LOGIN
============================================================================ */

function LoginScreen({ onSuccess }) {
  const [key, setKey] = useState("");
  const [show, setShow] = useState(false);
  const [error, setError] = useState(false);

  const submit = () => {
    const normalized = key.trim().toLowerCase();
    if (normalized === ACCESS_KEY.toLowerCase()) {
      onSuccess();
    } else {
      setError(true);
      setTimeout(() => setError(false), 1600);
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submit();
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ backgroundColor: COLORS.mint }}>
      <style>{GLOBAL_CSS}</style>
      <div className={`w-full max-w-sm rounded-2xl p-8 shadow-2xl border ${error ? "animate-pulse" : ""}`} style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.08) }}>
        <div className="flex justify-center mb-6">
          <Logo size="large" variant="dark" />
        </div>
        <h1 className="text-lg font-bold text-center tracking-wide uppercase" style={{ color: COLORS.navy }}>Recrutamento</h1>
        <p className="text-sm mt-1 mb-6 text-center" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>Introduz a tua chave de acesso para entrar na plataforma.</p>
        <Field label="Chave de acesso">
          <div className="relative">
            <Lock size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: hexToRgba(COLORS.navy, 0.45) }} />
            <input
              autoFocus
              type={show ? "text" : "password"}
              value={key}
              onChange={(e) => setKey(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="YME2026"
              className={`${inputCls} pl-9 pr-9`}
            />
            <button
              type="button"
              onClick={() => setShow((s) => !s)}
              className="absolute right-3 top-1/2 -translate-y-1/2"
              style={{ color: hexToRgba(COLORS.navy, 0.45) }}
              tabIndex={-1}
            >
              {show ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          </div>
        </Field>
        {error && <p className="text-xs mb-3" style={{ color: "#c0227a" }}>Chave incorreta. Tenta novamente.</p>}
        <button type="button" onClick={submit} className="yme-btn-primary w-full rounded-lg py-2.5 text-sm font-semibold tracking-wide">
          ENTRAR
        </button>
        <p className="text-[11px] mt-4 text-center" style={{ color: hexToRgba(COLORS.navy, 0.4) }}>Chave de demonstração: YME2026</p>
      </div>
    </div>
  );
}

/* ============================================================================
   NAVEGAÇÃO SUPERIOR
============================================================================ */

function TopNav({ page, setPage, onLogout, counts }) {
  const items = [
    { id: "dashboard", label: "DASHBOARD", badge: counts.fase0 },
    { id: "import", label: "IMPORTAÇÃO DE DADOS" },
    { id: "fase1", label: "FASE 1 · SOFT SKILLS", badge: counts.fase1 },
    { id: "fase2", label: "FASE 2 · DINÂMICAS", badge: counts.fase2 },
    { id: "fase3", label: "FASE 3 · HARD SKILLS", badge: counts.fase3 },
  ];
  return (
    <header className="sticky top-0 z-40" style={{ backgroundColor: COLORS.navy, borderBottom: `1px solid ${hexToRgba(COLORS.mint, 0.12)}` }}>
      <div className="max-w-[1400px] mx-auto px-6 flex items-center h-16 gap-6">
        <div className="shrink-0">
          <Logo variant="light" />
        </div>
        <nav className="flex items-center h-full gap-1 flex-1 overflow-x-auto">
          {items.map((it) => {
            const active = page === it.id;
            return (
              <button
                key={it.id}
                onClick={() => setPage(it.id)}
                className={`yme-nav-item h-full px-4 flex items-center gap-2 text-xs font-semibold tracking-wider whitespace-nowrap ${active ? "active" : ""}`}
              >
                {it.label}
                {typeof it.badge === "number" && (
                  <span
                    className="text-[10px] font-mono rounded-full px-1.5 py-0.5"
                    style={active ? { backgroundColor: hexToRgba(COLORS.pink, 0.18), color: COLORS.pink } : { backgroundColor: hexToRgba(COLORS.mint, 0.12), color: "#94a3b8" }}
                  >
                    {it.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>
        <button onClick={onLogout} className="yme-nav-item shrink-0 flex items-center gap-1.5 px-2 text-xs font-medium tracking-wide h-8 rounded">
          <LogOut size={14} /> TERMINAR SESSÃO
        </button>
      </div>
    </header>
  );
}

/* ============================================================================
   PAGE: HUB DE IMPORTAÇÃO DE DADOS
============================================================================ */

function UploadCard({ icon: Icon, title, description, hint, status, onFile, accept = ".xlsx,.xls,.csv" }) {
  const [error, setError] = useState("");
  return (
    <div className="rounded-xl border p-5" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
      <div className="flex items-start gap-3 mb-3">
        <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ backgroundColor: COLORS.navy, color: COLORS.pink }}><Icon size={17} /></div>
        <div className="min-w-0">
          <p className="text-sm font-semibold" style={{ color: COLORS.navy }}>{title}</p>
          <p className="text-xs mt-0.5" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{description}</p>
        </div>
      </div>

      {status.loaded ? (
        <div className="flex items-center gap-1.5 text-xs rounded-lg px-2.5 py-1.5 mb-3 border" style={{ backgroundColor: COLORS.white, color: COLORS.navy, borderColor: hexToRgba(COLORS.navy, 0.12) }}>
          <FileCheck2 size={13} className="shrink-0" style={{ color: COLORS.navy }} />
          <span className="truncate">Carregado — {status.filename} ({status.count} registo(s))</span>
        </div>
      ) : (
        <div className="flex items-center gap-1.5 text-xs rounded-lg px-2.5 py-1.5 mb-3 border" style={{ backgroundColor: COLORS.navy, color: COLORS.pink, borderColor: hexToRgba(COLORS.pink, 0.4) }}>
          <FileClock size={13} className="shrink-0" /> Por carregar
        </div>
      )}

      {error && <p className="text-[11px] mb-2" style={{ color: "#c0227a" }}>{error}</p>}

      <label className="yme-btn-outline-light flex items-center justify-center gap-1.5 text-xs font-medium rounded-lg px-3 py-2 cursor-pointer">
        <UploadCloud size={13} /> {status.loaded ? "Substituir ficheiro" : "Carregar ficheiro"}
        <input type="file" accept={accept} className="hidden" onChange={async (e) => {
          const file = e.target.files?.[0];
          e.target.value = "";
          if (!file) return;
          setError("");
          try { await onFile(file); } catch { setError("Não foi possível processar o ficheiro. Confirma o formato das colunas."); }
        }} />
      </label>
      <p className="text-[10px] mt-2 leading-relaxed" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>{hint}</p>
    </div>
  );
}

function SyncCard({ syncUrl, syncState, auth, onSave, onSyncNow, onAuthenticate, onSignOut }) {
  const [draft, setDraft] = useState(syncUrl);
  useEffect(() => { setDraft(syncUrl); }, [syncUrl]);

  const fmtTime = (d) => (d ? new Date(d).toLocaleString("pt-PT", { hour: "2-digit", minute: "2-digit", day: "2-digit", month: "2-digit" }) : "—");
  const isAuthed = !!auth.accessToken;

  return (
    <div className="rounded-xl border p-5 mb-4" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
      <div className="flex items-start gap-3 mb-3">
        <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ backgroundColor: COLORS.navy, color: COLORS.pink }}><RefreshCw size={17} /></div>
        <div className="min-w-0">
          <p className="text-sm font-semibold" style={{ color: COLORS.navy }}>Sincronização em Tempo Real — Excel Mestre (Google Sheets, ficheiro privado)</p>
          <p className="text-xs mt-0.5" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>
            Liga com segurança à folha de cálculo mestre privada da YME, através da tua conta Google autorizada (Google Sheets API). Sincronização automática a cada 2 minutos enquanto a sessão estiver ativa.
          </p>
        </div>
      </div>

      {/* Estado de autenticação / sessão Google */}
      <div className="flex items-center justify-between flex-wrap gap-2 mb-3 rounded-lg border px-3 py-2" style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.12) }}>
        {isAuthed ? (
          <div className="flex items-center gap-1.5 text-xs min-w-0" style={{ color: COLORS.navy }}>
            <CheckCircle2 size={13} className="shrink-0" style={{ color: "#1a8f5e" }} />
            <span className="truncate">
              Sessão Ativa: <strong>{auth.email || "conta Google"}</strong>{syncState.lastSync ? " · Sincronizado" : ""}
            </span>
          </div>
        ) : (
          <div className="flex items-center gap-1.5 text-xs" style={{ color: hexToRgba(COLORS.navy, 0.65) }}>
            <Lock size={13} className="shrink-0" /> Sem sessão ativa — autentica-te para aceder à folha privada.
          </div>
        )}
        <div className="flex items-center gap-3 shrink-0">
          {isAuthed ? (
            <button onClick={onSignOut} className="yme-link text-xs">Terminar sessão</button>
          ) : (
            <button onClick={onAuthenticate} disabled={auth.authenticating || !auth.ready}
              className="yme-btn-primary flex items-center gap-1.5 text-xs font-medium rounded-lg px-3 py-1.5 disabled:opacity-50 whitespace-nowrap">
              <Lock size={13} /> {auth.authenticating ? "A autenticar…" : "Autenticar com Google para Sincronizar Ficheiro Privado"}
            </button>
          )}
        </div>
      </div>
      {auth.error && <p className="text-[11px] mb-2" style={{ color: "#c0227a" }}>{auth.error}</p>}

      <div className="flex flex-col sm:flex-row gap-2 mb-2">
        <input
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          placeholder="Cola aqui o link ou o ID da folha mestre privada (Ficheiro > Partilhar — partilhada apenas com contas Google autorizadas @yme.pt)…"
          className={`${inputCls} flex-1`}
        />
        <button onClick={() => onSave(draft.trim())} className="yme-btn-primary text-sm px-3 py-2 rounded-lg font-medium whitespace-nowrap">
          Guardar Ligação
        </button>
      </div>

      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-1.5 text-xs" style={{ color: hexToRgba(COLORS.navy, 0.65) }}>
          {syncState.syncing ? (
            <span className="inline-flex items-center gap-1.5"><RefreshCw size={13} className="animate-spin" /> A sincronizar…</span>
          ) : syncState.error ? (
            <span className="inline-flex items-center gap-1.5" style={{ color: "#c0227a" }}><AlertTriangle size={13} /> {syncState.error}</span>
          ) : syncState.lastSync ? (
            <span className="inline-flex items-center gap-1.5"><CheckCircle2 size={13} style={{ color: COLORS.navy }} /> Última sincronização: {fmtTime(syncState.lastSync)} · {syncState.lastCount ?? 0} candidato(s)</span>
          ) : (
            <span className="inline-flex items-center gap-1.5"><Clock3 size={13} /> Ainda sem sincronização.</span>
          )}
        </div>
        <button onClick={onSyncNow} disabled={!syncUrl || !isAuthed || syncState.syncing}
          className="yme-btn-outline-light flex items-center gap-1.5 text-xs font-medium rounded-lg px-3 py-1.5 disabled:opacity-40">
          <RefreshCw size={13} className={syncState.syncing ? "animate-spin" : ""} /> Sincronizar Agora
        </button>
      </div>
      <p className="text-[10px] mt-2 leading-relaxed" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>
        A folha permanece 100% privada — não precisa de estar acessível "por link". Basta estar partilhada com as contas Google autorizadas da YME (com acesso de Leitor). Abas consultadas: {Object.values(SYNC_SHEET_NAMES).join(", ")} e as 6 abas de departamento ({DEPARTMENTS.join(", ")}). Todas as restantes abas são ignoradas.
      </p>
    </div>
  );
}

function ImportHubPage({
  members, setMembers, candidates, setCandidates, importStatus, setImportStatus,
  syncUrl, syncState, auth, onSaveSyncUrl, onSyncNow, onAuthenticate, onSignOut,
}) {
  const handleExcelMembers = async (file) => {
    const wb = await readWorkbook(file);
    let count = 0;
    setMembers((prev) => {
      const next = [...prev];
      wb.SheetNames.forEach((sheetName) => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
        rows.forEach((row) => {
          const name = String(get(row, "nome", "name")).trim();
          if (!name) return;
          let role = String(get(row, "role", "cargo")).trim();
          if (!role) {
            if (/diretor/i.test(sheetName)) role = "Diretor";
            else if (/rh|recursos/i.test(sheetName)) role = "RH";
            else if (/supervisor|c-level|clevel/i.test(sheetName)) role = "Supervisor";
          }
          const deptsRaw = String(get(row, "departamentos", "departamento")).split(/[,;|]/).map((s) => matchDept(s)).filter(Boolean);
          const availRaw = String(get(row, "disponibilidade", "horarios", "slots")).split("|").map((s) => s.trim()).filter((s) => SLOTS.includes(s));
          const idx = next.findIndex((m) => m.name.toLowerCase() === name.toLowerCase());
          count++;
          if (idx >= 0) {
            next[idx] = {
              ...next[idx],
              role: role || next[idx].role,
              departments: deptsRaw.length ? deptsRaw : next[idx].departments,
              availability: availRaw.length ? availRaw : next[idx].availability,
            };
          } else {
            next.push({ id: uid("imp"), name, role: role || "RH", title: role || "Membro", departments: deptsRaw.length ? deptsRaw : [DEPARTMENTS[0]], availability: availRaw });
          }
        });
      });
      return next;
    });
    setImportStatus((prev) => ({ ...prev, excel: { loaded: true, filename: file.name, count } }));
  };

  const handleFormsPhase = (phaseKey) => async (file) => {
    const wb = await readWorkbook(file);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
    let count = 0;
    setCandidates((prev) => {
      const next = [...prev];
      rows.forEach((row) => {
        const name = String(get(row, "nome", "name")).trim();
        if (!name) return;
        const department = matchDept(get(row, "departamento", "department"));
        const email = String(get(row, "email")).trim();
        const availability = String(get(row, "disponibilidade", "horarios", "slots")).split("|").map((s) => s.trim()).filter((s) => SLOTS.includes(s));
        const idx = next.findIndex((c) => c.name.toLowerCase() === name.toLowerCase());
        count++;
        if (idx >= 0) {
          next[idx] = {
            ...next[idx],
            department: department || next[idx].department,
            email: email || next[idx].email,
            formsSubmitted: { ...next[idx].formsSubmitted, [phaseKey]: true },
            availability: { ...next[idx].availability, [phaseKey]: availability.length ? availability : next[idx].availability[phaseKey] },
          };
        } else {
          next.push({
            id: uid("cand"), name, department: department || DEPARTMENTS[0], email, cvLink: "",
            phase0Status: "Aprovado",
            phase1Status: phaseKey === "fase1" ? "Pendente" : "—",
            phase2Status: phaseKey === "fase2" ? "Pendente" : "—",
            formsSubmitted: { fase1: phaseKey === "fase1", fase2: phaseKey === "fase2", fase3: phaseKey === "fase3" },
            availability: {
              fase1: phaseKey === "fase1" ? availability : [],
              fase2: phaseKey === "fase2" ? availability : [],
              fase3: phaseKey === "fase3" ? availability : [],
            },
          });
        }
      });
      return next;
    });
    setImportStatus((prev) => ({ ...prev, [phaseKey]: { loaded: true, filename: file.name, count } }));
  };

  const downloadTemplate = (kind) => {
    if (kind === "excel") {
      downloadCSV("modelo-excel-mestre-yme.csv", [
        ["Nome", "Role", "Departamentos", "Disponibilidade"],
        ["Gustavo Dias", "Diretor", "Brand Strategy", "Seg 09:00|Ter 10:30|Qua 14:00"],
      ]);
    } else {
      downloadCSV(`modelo-forms-${kind}.csv`, [
        ["Nome", "Departamento", "Email", "Disponibilidade"],
        ["Ana Silva", "Digital Development", "ana.silva@email.com", "Seg 09:00|Qua 14:00|Sex 17:00"],
      ]);
    }
  };

  const membersByRole = { Diretor: members.filter((m) => m.role === "Diretor").length, RH: members.filter((m) => m.role === "RH").length, Supervisor: members.filter((m) => m.role === "Supervisor").length };

  return (
    <div className="p-8">
      <div className="mb-6">
        <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>Importação de Dados</h1>
        <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>Hub central para carregar o Excel Mestre da YME e os Google Forms de candidatos de cada fase.</p>
      </div>

      <SyncCard syncUrl={syncUrl} syncState={syncState} auth={auth} onSave={onSaveSyncUrl} onSyncNow={onSyncNow}
        onAuthenticate={onAuthenticate} onSignOut={onSignOut} />

      <p className="text-xs font-semibold uppercase tracking-widest mb-2" style={{ color: "#94a3b8" }}>Importação Manual (alternativa/backup)</p>
      <div className="grid grid-cols-2 gap-4 mb-4">
        <UploadCard icon={FileSpreadsheet} title="Excel Mestre da YME" status={importStatus.excel}
          description="Ficheiro .xlsx multi-abas com dados e disponibilidades de Diretores, RH e Supervisores/C-Level."
          hint='Colunas esperadas por linha: Nome, Role, Departamentos, Disponibilidade (slots separados por "|", ex: "Seg 09:00|Ter 10:30").'
          onFile={handleExcelMembers} />
        <div className="rounded-xl border p-5 flex flex-col justify-between" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
          <div>
            <p className="text-sm font-semibold mb-1" style={{ color: COLORS.navy }}>Membros ativos carregados</p>
            <div className="grid grid-cols-3 gap-2 mt-3">
              <div className="rounded-lg border p-2.5 text-center" style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
                <p className="text-lg font-bold font-mono" style={{ color: COLORS.navy }}>{membersByRole.Diretor}</p>
                <p className="text-[10px]" style={{ color: hexToRgba(COLORS.navy, 0.55) }}>Diretores</p>
              </div>
              <div className="rounded-lg border p-2.5 text-center" style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
                <p className="text-lg font-bold font-mono" style={{ color: COLORS.navy }}>{membersByRole.RH}</p>
                <p className="text-[10px]" style={{ color: hexToRgba(COLORS.navy, 0.55) }}>RH</p>
              </div>
              <div className="rounded-lg border p-2.5 text-center" style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
                <p className="text-lg font-bold font-mono" style={{ color: COLORS.navy }}>{membersByRole.Supervisor}</p>
                <p className="text-[10px]" style={{ color: hexToRgba(COLORS.navy, 0.55) }}>Supervisores</p>
              </div>
            </div>
          </div>
          <button onClick={() => downloadTemplate("excel")} className="yme-link mt-3 text-xs self-start">Descarregar modelo CSV ↓</button>
        </div>
      </div>

      <p className="text-xs font-semibold uppercase tracking-widest mb-2 mt-6" style={{ color: "#94a3b8" }}>Google Forms de Candidatos por Fase</p>
      <div className="grid grid-cols-3 gap-4 mb-8">
        <UploadCard icon={UsersRound} title="Forms — Fase 1" status={importStatus.fase1}
          description="Candidatos e disponibilidades submetidas para as entrevistas de Soft Skills."
          hint='Colunas: Nome, Departamento, Email, Disponibilidade (slots separados por "|").'
          onFile={handleFormsPhase("fase1")} />
        <UploadCard icon={LayoutGrid} title="Forms — Fase 2" status={importStatus.fase2}
          description="Candidatos e disponibilidades submetidas para as Dinâmicas de Grupo."
          hint='Colunas: Nome, Departamento, Email, Disponibilidade (slots separados por "|").'
          onFile={handleFormsPhase("fase2")} />
        <UploadCard icon={CalendarClock} title="Forms — Fase 3" status={importStatus.fase3}
          description="Candidatos e disponibilidades submetidas para as entrevistas de Hard Skills."
          hint='Colunas: Nome, Departamento, Email, Disponibilidade (slots separados por "|").'
          onFile={handleFormsPhase("fase3")} />
      </div>
      <button onClick={() => downloadTemplate("forms")} className="yme-link text-xs -mt-6 mb-8">Descarregar modelo CSV de Forms ↓</button>

      <p className="text-xs font-semibold uppercase tracking-widest mb-2" style={{ color: "#94a3b8" }}>Estrutura Organizacional (referência)</p>
      <div className="rounded-xl border overflow-hidden" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs uppercase tracking-wide" style={{ backgroundColor: COLORS.navy, color: COLORS.white }}>
              <th className="px-4 py-3 font-medium">Departamento</th>
              <th className="px-4 py-3 font-medium">Diretor(a)</th>
              <th className="px-4 py-3 font-medium">Supervisor</th>
              <th className="px-4 py-3 font-medium">Membro(s) RH</th>
            </tr>
          </thead>
          <tbody>
            {ORG.map((o) => (
              <tr key={o.dept} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                <td className="px-4 py-3"><DeptBadge dept={o.dept} /></td>
                <td className="px-4 py-3" style={{ color: COLORS.navy }}>{o.diretor}</td>
                <td className="px-4 py-3" style={{ color: COLORS.navy }}>{o.supervisor} <span className="text-xs" style={{ color: hexToRgba(COLORS.navy, 0.55) }}>({o.supervisorTitle})</span></td>
                <td className="px-4 py-3" style={{ color: COLORS.navy }}>{o.rh.join(", ")}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================================================================
   PAGE: DASHBOARD / FASE 0
============================================================================ */

function DashboardPage({ candidates, setCandidates, onAddCandidate, goToImport }) {
  const [deptFilter, setDeptFilter] = useState("Todos");
  const [statusFilter, setStatusFilter] = useState("Todos");
  const [search, setSearch] = useState("");

  const filtered = candidates.filter((c) =>
    (deptFilter === "Todos" || c.department === deptFilter) &&
    (statusFilter === "Todos" || c.phase0Status === statusFilter) &&
    c.name.toLowerCase().includes(search.toLowerCase())
  );
  const { page, setPage, totalPages, pageItems } = usePagination(filtered, 12, `${deptFilter}-${statusFilter}-${search}`);

  const counts = {
    total: candidates.length,
    aprovados: candidates.filter((c) => c.phase0Status === "Aprovado").length,
    pendentes: candidates.filter((c) => c.phase0Status === "Pendente").length,
    rejeitados: candidates.filter((c) => c.phase0Status === "Rejeitado").length,
  };

  const setStatus = (id, status) => setCandidates((prev) => prev.map((c) => (c.id === id ? { ...c, phase0Status: status } : c)));

  return (
    <div className="p-8">
      <div className="flex items-start justify-between mb-6">
        <div>
          <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>Dashboard & Fase 0 — Questionário e CV</h1>
          <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>Visão geral de {candidates.length} candidatos inscritos, distribuídos pelos 6 departamentos.</p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={goToImport} className="yme-btn-outline-dark flex items-center gap-1.5 text-sm rounded-lg px-3 py-2">
            <UploadCloud size={14} /> Hub de Importação
          </button>
          <button onClick={onAddCandidate} className="yme-btn-primary flex items-center gap-1.5 text-sm rounded-lg px-3 py-2 font-medium">
            <Plus size={14} /> Adicionar Candidato
          </button>
        </div>
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <StatCard label="Total de candidatos" value={counts.total} icon={UsersRound} tone="neutral" />
        <StatCard label="Aprovados p/ Fase 1" value={counts.aprovados} icon={CheckCircle2} tone="brand" />
        <StatCard label="Pendentes" value={counts.pendentes} icon={Clock3} tone="alert" />
        <StatCard label="Rejeitados" value={counts.rejeitados} icon={XCircle} tone="critical" />
      </div>

      <div className="flex flex-wrap items-center gap-3 mb-4">
        <div className="relative">
          <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: hexToRgba(COLORS.navy, 0.45) }} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Procurar candidato..."
            className="yme-input pl-8 pr-3 py-2 text-sm rounded-lg w-56" />
        </div>
        <select value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)} className="yme-input text-sm rounded-lg px-2.5 py-2">
          <option>Todos</option>
          {DEPARTMENTS.map((d) => <option key={d}>{d}</option>)}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="yme-input text-sm rounded-lg px-2.5 py-2">
          {["Todos", "Aprovado", "Pendente", "Rejeitado"].map((s) => <option key={s}>{s}</option>)}
        </select>
        <span className="text-xs ml-auto" style={{ color: "#94a3b8" }}>{filtered.length} resultado(s)</span>
      </div>

      <div className="rounded-xl border overflow-hidden" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs uppercase tracking-wide" style={{ backgroundColor: COLORS.navy, color: COLORS.white }}>
              <th className="px-4 py-3 font-medium">Candidato</th>
              <th className="px-4 py-3 font-medium">Departamento</th>
              <th className="px-4 py-3 font-medium">Email</th>
              <th className="px-4 py-3 font-medium">CV / Questionário</th>
              <th className="px-4 py-3 font-medium">Estado</th>
              <th className="px-4 py-3 font-medium">Avança Fase 1</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((c) => (
              <tr key={c.id} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                <td className="px-4 py-3 font-medium" style={{ color: COLORS.navy }}>{c.name}</td>
                <td className="px-4 py-3"><DeptBadge dept={c.department} /></td>
                <td className="px-4 py-3" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{c.email}</td>
                <td className="px-4 py-3">
                  <a href={c.cvLink} target="_blank" rel="noreferrer" className="yme-link text-xs">Ver questionário ↗</a>
                </td>
                <td className="px-4 py-3">
                  <select value={c.phase0Status} onChange={(e) => setStatus(c.id, e.target.value)}
                    className={`text-xs rounded-md px-2 py-1 border font-medium ${statusBadgeClass(c.phase0Status)}`}>
                    {["Aprovado", "Pendente", "Rejeitado"].map((s) => <option key={s} style={{ color: COLORS.navy }}>{s}</option>)}
                  </select>
                </td>
                <td className="px-4 py-3">
                  {c.phase0Status === "Aprovado" ? (
                    <span className="inline-flex items-center gap-1 text-xs font-medium" style={{ color: COLORS.navy }}><CheckCircle2 size={14} style={{ color: COLORS.pink }} /> Sim</span>
                  ) : (
                    <span className="inline-flex items-center gap-1 text-xs" style={{ color: hexToRgba(COLORS.navy, 0.45) }}><XCircle size={14} /> Não</span>
                  )}
                </td>
              </tr>
            ))}
            {pageItems.length === 0 && (
              <tr><td colSpan={6} className="px-4 py-10 text-center text-sm" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Sem candidatos para os filtros selecionados.</td></tr>
            )}
          </tbody>
        </table>
        <PaginationBar page={page} setPage={setPage} totalPages={totalPages} />
      </div>
    </div>
  );
}

/* ============================================================================
   PAGE: FASE 1 & FASE 3 (entrevistas 1:1)
============================================================================ */

function InterviewPhasePage({
  title, subtitle, phaseKey, availField, formsField, prevStatusField,
  candidates, members, bookings, setBookings, onGenerate, columns, showCalendar,
}) {
  const [editing, setEditing] = useState(null);
  const [view, setView] = useState("list");

  const byId = (id) => members.find((m) => m.id === id);
  const candById = (id) => candidates.find((c) => c.id === id);

  const scheduled = bookings.filter((b) => b.status === "Agendado").length;
  const conflicts = bookings.filter((b) => b.status !== "Agendado").length;
  const missingForms = candidates.filter((c) => c[prevStatusField] === "Aprovado" && !c.formsSubmitted[formsField]).length;

  const { page, setPage, totalPages, pageItems } = usePagination(bookings, 12, bookings.length);

  const exportCSV = () => {
    const header = ["Candidato", "Departamento", ...columns.map((c) => c.label), "Horário", "Estado"];
    const rows = bookings.map((b) => {
      const cand = candById(b.candidateId);
      return [cand?.name, cand?.department, ...columns.map((c) => byId(b[c.key])?.name || "—"), b.slot || "—", b.status];
    });
    downloadCSV(`${phaseKey}-agendamentos.csv`, [header, ...rows]);
  };

  return (
    <div className="p-8">
      <div className="flex items-start justify-between mb-6">
        <div>
          <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>{title}</h1>
          <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>{subtitle}</p>
        </div>
        <div className="flex items-center gap-2">
          {showCalendar && (
            <div className="flex rounded-lg overflow-hidden text-sm" style={{ border: `1px solid ${hexToRgba(COLORS.mint, 0.25)}` }}>
              <button onClick={() => setView("list")} className="px-3 py-2 flex items-center gap-1.5" style={view === "list" ? { backgroundColor: COLORS.pink, color: COLORS.navy } : { backgroundColor: "transparent", color: "#94a3b8" }}><ListChecks size={14} /> Lista</button>
              <button onClick={() => setView("calendar")} className="px-3 py-2 flex items-center gap-1.5" style={view === "calendar" ? { backgroundColor: COLORS.pink, color: COLORS.navy } : { backgroundColor: "transparent", color: "#94a3b8" }}><CalendarDays size={14} /> Calendário</button>
            </div>
          )}
          <button onClick={exportCSV} className="yme-btn-outline-dark flex items-center gap-1.5 text-sm rounded-lg px-3 py-2">
            <Download size={14} /> Exportar CSV
          </button>
          <button onClick={onGenerate} className="yme-btn-primary flex items-center gap-1.5 text-sm rounded-lg px-3 py-2 font-medium">
            <RefreshCw size={14} /> Gerar Agendamentos Automaticamente
          </button>
        </div>
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <StatCard label="Candidatos elegíveis" value={bookings.length} icon={UsersRound} tone="neutral" />
        <StatCard label="Entrevistas agendadas" value={scheduled} icon={CheckCircle2} tone="brand" />
        <StatCard label="Sem horário comum" value={conflicts} icon={AlertTriangle} tone="critical" />
        <StatCard label={`À espera do Forms ${PHASE_LABEL[formsField]}`} value={missingForms} icon={FileClock} tone="alert" />
      </div>

      {bookings.length === 0 && (
        <div className="rounded-xl border border-dashed p-10 text-center text-sm" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.2), color: hexToRgba(COLORS.navy, 0.5) }}>
          Ainda não há candidatos aprovados na fase anterior que tenham submetido o Forms desta fase.
        </div>
      )}

      {bookings.length > 0 && view === "list" && (
        <div className="rounded-xl border overflow-hidden" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs uppercase tracking-wide" style={{ backgroundColor: COLORS.navy, color: COLORS.white }}>
                <th className="px-4 py-3 font-medium">Candidato</th>
                <th className="px-4 py-3 font-medium">Departamento</th>
                {columns.map((c) => <th key={c.key} className="px-4 py-3 font-medium">{c.label}</th>)}
                <th className="px-4 py-3 font-medium">Horário</th>
                <th className="px-4 py-3 font-medium">Estado</th>
                <th className="px-4 py-3 font-medium"></th>
              </tr>
            </thead>
            <tbody>
              {pageItems.map((b) => {
                const cand = candById(b.candidateId);
                if (!cand) return null;
                return (
                  <tr key={b.id} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                    <td className="px-4 py-3 font-medium" style={{ color: COLORS.navy }}>{cand.name}{b.manual && <span className="ml-1.5 text-[10px] font-normal" style={{ color: COLORS.pink }}>(manual)</span>}</td>
                    <td className="px-4 py-3"><DeptBadge dept={cand.department} /></td>
                    {columns.map((c) => <td key={c.key} className="px-4 py-3" style={{ color: hexToRgba(COLORS.navy, 0.75) }}>{byId(b[c.key])?.name || <span className="text-xs" style={{ color: "#c0227a" }}>Sem alocação</span>}</td>)}
                    <td className="px-4 py-3 font-mono text-xs" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{b.slot || "—"}</td>
                    <td className="px-4 py-3"><StatusBadge status={b.status} /></td>
                    <td className="px-4 py-3 text-right">
                      <button onClick={() => setEditing(b)} style={{ color: hexToRgba(COLORS.navy, 0.45) }}><Pencil size={15} /></button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          <PaginationBar page={page} setPage={setPage} totalPages={totalPages} />
        </div>
      )}

      {bookings.length > 0 && view === "calendar" && (
        <CalendarView bookings={bookings} candById={candById} byId={byId} columns={columns} />
      )}

      {editing && (
        <EditBookingModal
          booking={editing} columns={columns} candidate={candById(editing.candidateId)}
          members={members}
          onClose={() => setEditing(null)}
          onSave={(updated) => { setBookings((prev) => prev.map((b) => (b.id === updated.id ? updated : b))); setEditing(null); }}
        />
      )}
    </div>
  );
}

function CalendarView({ bookings, candById, byId, columns }) {
  return (
    <div className="rounded-xl border overflow-x-auto" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
      <table className="w-full text-xs border-collapse">
        <thead>
          <tr>
            <th className="p-2 w-20" style={{ backgroundColor: COLORS.navy, border: `1px solid ${hexToRgba(COLORS.mint, 0.15)}` }}></th>
            {DAYS.map((d) => <th key={d} className="p-2 font-medium" style={{ backgroundColor: COLORS.navy, color: COLORS.white, border: `1px solid ${hexToRgba(COLORS.mint, 0.15)}` }}>{d}</th>)}
          </tr>
        </thead>
        <tbody>
          {TIMES.map((t) => (
            <tr key={t}>
              <td className="p-2 font-mono text-right" style={{ color: hexToRgba(COLORS.navy, 0.55), border: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>{t}</td>
              {DAYS.map((d) => {
                const slot = `${d} ${t}`;
                const items = bookings.filter((b) => b.slot === slot);
                return (
                  <td key={slot} className="p-1.5 align-top min-w-[140px]" style={{ border: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                    {items.map((b) => {
                      const cand = candById(b.candidateId);
                      return (
                        <div key={b.id} className="mb-1 last:mb-0 rounded-md px-2 py-1 border" style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.12) }}>
                          <p className="font-medium" style={{ color: COLORS.navy }}>{cand?.name}</p>
                          <p style={{ color: hexToRgba(COLORS.navy, 0.55) }}>{columns.map((c) => byId(b[c.key])?.name).filter(Boolean).join(" · ")}</p>
                        </div>
                      );
                    })}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function EditBookingModal({ booking, columns, candidate, members, onClose, onSave }) {
  const [form, setForm] = useState({ ...booking });
  const membersFor = (roleKey) => {
    if (roleKey === "diretorId") return members.filter((m) => m.role === "Diretor");
    if (roleKey === "rhId") return members.filter((m) => m.role === "RH");
    if (roleKey === "supervisorId") return members.filter((m) => m.role === "Supervisor");
    return [];
  };
  const save = () => onSave({ ...form, manual: true, status: form.slot ? "Agendado" : "Sem Horário Comum" });
  return (
    <Modal title={`Editar agendamento — ${candidate?.name}`} onClose={onClose} wide>
      <div className="grid grid-cols-2 gap-3">
        {columns.map((c) => (
          <Field key={c.key} label={c.label}>
            <select value={form[c.key] || ""} onChange={(e) => setForm({ ...form, [c.key]: e.target.value })} className={inputCls}>
              <option value="">— Sem alocação —</option>
              {membersFor(c.key).map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
            </select>
          </Field>
        ))}
      </div>
      <Field label="Horário">
        <select value={form.slot || ""} onChange={(e) => setForm({ ...form, slot: e.target.value })} className={inputCls}>
          <option value="">— Por definir —</option>
          {SLOTS.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
      <p className="text-xs mb-4" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Nota: a edição manual não valida automaticamente a disponibilidade de cada interveniente — confirma antes de guardar.</p>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="yme-btn-outline-light text-sm px-4 py-2 rounded-lg">Cancelar</button>
        <button onClick={save} className="yme-btn-primary text-sm px-4 py-2 rounded-lg font-medium">Guardar alteração</button>
      </div>
    </Modal>
  );
}

/* ============================================================================
   PAGE: FASE 2 — DINÂMICAS DE GRUPO
============================================================================ */

function Phase2Page({ candidates, members, groups, setGroups, onGenerate }) {
  const candById = (id) => candidates.find((c) => c.id === id);
  const byId = (id) => members.find((m) => m.id === id);
  const missingForms = candidates.filter((c) => c.phase1Status === "Aprovado" && !c.formsSubmitted.fase2).length;

  const exportCSV = () => {
    const header = ["Grupo", "Horário", "Candidatos", "Supervisor", "Diretores presentes", "RH presentes", "Avisos"];
    const rows = groups.map((g) => [
      g.name, g.slot || "—",
      g.candidateIds.map((id) => candById(id)?.name).join(" | "),
      byId(g.supervisorId)?.name || "—",
      g.directorIds.map((id) => byId(id)?.name).join(" | ") || "—",
      g.rhIds.map((id) => byId(id)?.name).join(" | ") || "—",
      g.warnings.join(" | ") || "Sem avisos",
    ]);
    downloadCSV("fase2-dinamicas-grupo.csv", [header, ...rows]);
  };

  const moveCandidate = (candId, fromGroupId, toGroupId) => {
    if (fromGroupId === toGroupId) return;
    setGroups((prev) => {
      const next = prev.map((g) => ({ ...g, candidateIds: [...g.candidateIds] }));
      const from = next.find((g) => g.id === fromGroupId);
      const to = next.find((g) => g.id === toGroupId);
      from.candidateIds = from.candidateIds.filter((id) => id !== candId);
      to.candidateIds.push(candId);
      [from, to].forEach((g) => {
        const groupCands = g.candidateIds.map((id) => candById(id)).filter(Boolean);
        const deptCounts = {};
        groupCands.forEach((c) => (deptCounts[c.department] = (deptCounts[c.department] || 0) + 1));
        const warnings = [];
        const missing = g.slot ? groupCands.filter((c) => !c.availability.fase2.includes(g.slot)) : [];
        if (missing.length) warnings.push(`${missing.length} candidato(s) indisponível(eis) no horário escolhido: ${missing.map((c) => c.name).join(", ")}.`);
        Object.entries(deptCounts).forEach(([d, n]) => { if (n > 2) warnings.push(`${n} candidatos do mesmo departamento (${d}) na mesma sessão — máximo recomendado: 2.`); });
        g.warnings = warnings;
      });
      return next;
    });
  };

  const totalWarnings = groups.reduce((acc, g) => acc + g.warnings.length, 0);

  return (
    <div className="p-8">
      <div className="flex items-start justify-between mb-6">
        <div>
          <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>Fase 2 — Dinâmicas de Grupo</h1>
          <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>Sessões de ~6 candidatos, máx. 2 por departamento, com Supervisor, Diretores e 2-3 RH.</p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={exportCSV} className="yme-btn-outline-dark flex items-center gap-1.5 text-sm rounded-lg px-3 py-2">
            <Download size={14} /> Exportar CSV
          </button>
          <button onClick={onGenerate} className="yme-btn-primary flex items-center gap-1.5 text-sm rounded-lg px-3 py-2 font-medium">
            <RefreshCw size={14} /> Gerar Agendamentos Automaticamente
          </button>
        </div>
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <StatCard label="Grupos formados" value={groups.length} icon={LayoutGrid} tone="neutral" />
        <StatCard label="Candidatos alocados" value={groups.reduce((a, g) => a + g.candidateIds.length, 0)} icon={UsersRound} tone="brand" />
        <StatCard label="Avisos de restrição" value={totalWarnings} icon={AlertTriangle} tone={totalWarnings ? "critical" : "alert"} />
        <StatCard label="À espera do Forms Fase 2" value={missingForms} icon={FileClock} tone="alert" />
      </div>

      {groups.length === 0 && (
        <div className="rounded-xl border border-dashed p-10 text-center text-sm" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.2), color: hexToRgba(COLORS.navy, 0.5) }}>
          Ainda não há candidatos aprovados na Fase 1 que tenham submetido o Forms da Fase 2.
        </div>
      )}

      <div className="grid grid-cols-2 gap-4">
        {groups.map((g) => (
          <div key={g.id} className="rounded-xl border p-4" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
            <div className="flex items-center justify-between mb-3">
              <h3 className="font-semibold" style={{ color: COLORS.navy }}>{g.name}</h3>
              <Badge style={{ backgroundColor: COLORS.white, color: COLORS.navy, borderColor: hexToRgba(COLORS.navy, 0.12) }}>
                <span className="font-mono">{g.slot || "sem horário"}</span>
              </Badge>
            </div>

            <p className="text-xs font-medium uppercase tracking-wide mb-1.5" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Candidatos ({g.candidateIds.length})</p>
            <div className="space-y-1.5 mb-3">
              {g.candidateIds.map((id) => {
                const c = candById(id);
                if (!c) return null;
                return (
                  <div key={id} className="flex items-center justify-between text-sm rounded-lg px-2.5 py-1.5 border" style={{ backgroundColor: COLORS.white, borderColor: hexToRgba(COLORS.navy, 0.08) }}>
                    <div className="flex items-center gap-2">
                      <span style={{ color: COLORS.navy }}>{c.name}</span>
                      <DeptBadge dept={c.department} />
                    </div>
                    <select
                      className="text-[11px] rounded px-1 py-0.5 border"
                      style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.15), color: hexToRgba(COLORS.navy, 0.7) }}
                      value={g.id}
                      onChange={(e) => moveCandidate(id, g.id, e.target.value)}
                    >
                      {groups.map((gg) => <option key={gg.id} value={gg.id}>{gg.name}</option>)}
                    </select>
                  </div>
                );
              })}
            </div>

            <div className="grid grid-cols-3 gap-2 text-xs mb-3">
              <div>
                <p className="mb-1" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Supervisor</p>
                <p className="font-medium" style={{ color: COLORS.navy }}>{byId(g.supervisorId)?.name || "—"}</p>
              </div>
              <div>
                <p className="mb-1" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Diretores</p>
                <p className="font-medium" style={{ color: COLORS.navy }}>{g.directorIds.map((id) => byId(id)?.name).join(", ") || "—"}</p>
              </div>
              <div>
                <p className="mb-1" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>RH</p>
                <p className="font-medium" style={{ color: COLORS.navy }}>{g.rhIds.map((id) => byId(id)?.name).join(", ") || "—"}</p>
              </div>
            </div>

            {g.warnings.length > 0 && (
              <div className="rounded-lg p-2.5 space-y-1 border" style={{ backgroundColor: COLORS.navy, borderColor: hexToRgba(COLORS.pink, 0.4) }}>
                {g.warnings.map((w, i) => (
                  <p key={i} className="text-[11px] flex items-start gap-1.5" style={{ color: COLORS.pink }}><AlertTriangle size={12} className="mt-0.5 shrink-0" />{w}</p>
                ))}
              </div>
            )}
            {g.warnings.length === 0 && (
              <p className="text-[11px] flex items-center gap-1.5" style={{ color: COLORS.navy }}><CheckCircle2 size={12} style={{ color: COLORS.pink }} /> Sem restrições violadas.</p>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================================
   ADD CANDIDATE MODAL
============================================================================ */

function AddCandidateModal({ onClose, onSave }) {
  const [form, setForm] = useState({ name: "", department: DEPARTMENTS[0], email: "", cvLink: "" });
  const save = () => {
    if (!form.name.trim()) return;
    onSave({
      id: uid("cand"), ...form,
      phase0Status: "Pendente", phase1Status: "—", phase2Status: "—",
      formsSubmitted: { fase1: false, fase2: false, fase3: false },
      availability: { fase1: [], fase2: [], fase3: [] },
    });
  };
  return (
    <Modal title="Adicionar Candidato" onClose={onClose} wide>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Nome completo"><input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
        <Field label="Departamento">
          <select className={inputCls} value={form.department} onChange={(e) => setForm({ ...form, department: e.target.value })}>
            {DEPARTMENTS.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="Email"><input className={inputCls} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
        <Field label="Link do CV / Questionário"><input className={inputCls} value={form.cvLink} onChange={(e) => setForm({ ...form, cvLink: e.target.value })} /></Field>
      </div>
      <p className="text-xs mb-3" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>A disponibilidade deste candidato para cada fase é preenchida através do respetivo Forms, no Hub de Importação.</p>
      <div className="flex justify-end gap-2 mt-2">
        <button onClick={onClose} className="yme-btn-outline-light text-sm px-4 py-2 rounded-lg">Cancelar</button>
        <button onClick={save} className="yme-btn-primary text-sm px-4 py-2 rounded-lg font-medium">Guardar candidato</button>
      </div>
    </Modal>
  );
}

/* ============================================================================
   APP
============================================================================ */

export default function App() {
  const [authed, setAuthed] = useState(false);
  const [page, setPage] = useState("dashboard");
  const [showAddCandidate, setShowAddCandidate] = useState(false);

  const [members, setMembers] = useState(() => buildMockMembers());
  const [candidates, setCandidates] = useState(() => buildMockCandidates(108));
  const [importStatus, setImportStatus] = useState({
    excel: { loaded: true, filename: "Excel_Mestre_YME_mock.xlsx", count: 17 },
    fase1: { loaded: true, filename: "Forms_Fase1_mock.csv", count: 0 },
    fase2: { loaded: true, filename: "Forms_Fase2_mock.csv", count: 0 },
    fase3: { loaded: true, filename: "Forms_Fase3_mock.csv", count: 0 },
  });

  useEffect(() => {
    setImportStatus((prev) => ({
      ...prev,
      fase1: { ...prev.fase1, count: candidates.filter((c) => c.formsSubmitted.fase1).length },
      fase2: { ...prev.fase2, count: candidates.filter((c) => c.formsSubmitted.fase2).length },
      fase3: { ...prev.fase3, count: candidates.filter((c) => c.formsSubmitted.fase3).length },
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ---- Sincronização em tempo real com o Excel Mestre (Google Sheets, privado) ---- */
  const { auth, requestToken, signOut } = useGoogleAuth();

  const [syncUrl, setSyncUrl] = useState(() => {
    try { return localStorage.getItem(SYNC_URL_STORAGE_KEY) || ""; } catch { return ""; }
  });
  const [syncState, setSyncState] = useState({ syncing: false, error: null, lastSync: null, lastCount: null });

  // Refs para o polling em segundo plano ter sempre acesso ao estado mais recente,
  // evitando "closures" desatualizadas dentro do setInterval.
  const membersRef = useRef(members);
  const candidatesRef = useRef(candidates);
  useEffect(() => { membersRef.current = members; }, [members]);
  useEffect(() => { candidatesRef.current = candidates; }, [candidates]);

  // tokenOverride permite passar diretamente um token acabado de obter
  // (login inicial ou renovação silenciosa), evitando depender do estado
  // React ainda não propagado no momento da chamada.
  const runSync = async (url, tokenOverride) => {
    const accessToken = tokenOverride || auth.accessToken;
    if (!url) return;
    if (!accessToken) {
      setSyncState((s) => ({ ...s, error: "Autentica-te com a Google para sincronizar a folha privada." }));
      return;
    }
    setSyncState((s) => ({ ...s, syncing: true, error: null }));
    try {
      const { members: nextMembers, candidates: nextCandidates, errors } =
        await syncMasterSheet({ accessToken, sheetUrl: url, prevMembers: membersRef.current, prevCandidates: candidatesRef.current });
      setMembers(nextMembers);
      setCandidates(nextCandidates);
      setImportStatus((prev) => ({
        ...prev,
        excel: { loaded: true, filename: "Sincronização em tempo real (Google Sheets, ficheiro privado)", count: nextCandidates.length },
      }));
      const uniqueErrors = Array.from(new Set(errors));
      setSyncState({
        syncing: false,
        error: uniqueErrors.length ? `${uniqueErrors.length} aba(s) com problemas: ${uniqueErrors.slice(0, 2).join(" ")}` : null,
        lastSync: new Date(),
        lastCount: nextCandidates.length,
      });
    } catch (err) {
      // Sessão expirada/necessária: tenta uma renovação silenciosa do token
      // (sem novo ecrã de consentimento) e repete a sincronização; só pede
      // login explícito ao utilizador se a renovação silenciosa falhar.
      if (err.message === "SESSAO_EXPIRADA" || err.message === "SESSAO_NECESSARIA") {
        try {
          const { accessToken: newToken } = await requestToken({ silent: true });
          await runSync(url, newToken);
        } catch {
          setSyncState((s) => ({ ...s, syncing: false, error: "A sessão Google expirou. Clica em \"Autenticar com Google\" para continuar a sincronizar." }));
        }
        return;
      }
      setSyncState((s) => ({ ...s, syncing: false, error: err.message || "Falha ao sincronizar." }));
    }
  };

  const handleAuthenticate = async () => {
    try {
      const { accessToken } = await requestToken({ silent: false });
      if (syncUrl) runSync(syncUrl, accessToken);
    } catch {
      setSyncState((s) => ({ ...s, error: "Não foi possível autenticar com a Google. Tenta novamente." }));
    }
  };

  const handleSaveSyncUrl = (url) => {
    setSyncUrl(url);
    try { localStorage.setItem(SYNC_URL_STORAGE_KEY, url); } catch { /* localStorage indisponível */ }
    if (url && auth.accessToken) runSync(url, auth.accessToken);
  };

  useEffect(() => {
    if (!syncUrl || !auth.accessToken) return;
    runSync(syncUrl, auth.accessToken);
    const interval = setInterval(() => runSync(syncUrl, auth.accessToken), SYNC_POLL_INTERVAL_MS);
    return () => clearInterval(interval);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [syncUrl, auth.accessToken]);

  const phase1Pool = useMemo(() => candidates.filter((c) => c.phase0Status === "Aprovado" && c.formsSubmitted.fase1), [candidates]);
  const phase2Pool = useMemo(() => candidates.filter((c) => c.phase1Status === "Aprovado" && c.formsSubmitted.fase2), [candidates]);
  const phase3Pool = useMemo(() => candidates.filter((c) => c.phase2Status === "Aprovado" && c.formsSubmitted.fase3), [candidates]);

  const [phase1Bookings, setPhase1Bookings] = useState(() => generateInterviewPhase(phase1Pool, members, [], "fase1", ["diretorId", "rhId"]));
  const [phase2Groups, setPhase2Groups] = useState(() => generatePhase2(phase2Pool, members));
  const [phase3Bookings, setPhase3Bookings] = useState(() => generateInterviewPhase(phase3Pool, members, [], "fase3", ["diretorId", "rhId", "supervisorId"]));

  if (!authed) return <LoginScreen onSuccess={() => setAuthed(true)} />;

  const counts = {
    fase0: candidates.length,
    fase1: phase1Bookings.length,
    fase2: phase2Groups.reduce((a, g) => a + g.candidateIds.length, 0),
    fase3: phase3Bookings.length,
  };

  return (
    <div className="min-h-screen flex flex-col font-sans" style={{ backgroundColor: COLORS.navy }}>
      <style>{GLOBAL_CSS}</style>
      <TopNav page={page} setPage={setPage} onLogout={() => setAuthed(false)} counts={counts} />
      <main className="flex-1 min-w-0 max-w-[1400px] w-full mx-auto">
        {page === "import" && (
          <ImportHubPage members={members} setMembers={setMembers} candidates={candidates} setCandidates={setCandidates}
            importStatus={importStatus} setImportStatus={setImportStatus}
            syncUrl={syncUrl} syncState={syncState} auth={auth}
            onSaveSyncUrl={handleSaveSyncUrl} onSyncNow={() => runSync(syncUrl, auth.accessToken)}
            onAuthenticate={handleAuthenticate} onSignOut={signOut} />
        )}
        {page === "dashboard" && (
          <DashboardPage candidates={candidates} setCandidates={setCandidates} onAddCandidate={() => setShowAddCandidate(true)} goToImport={() => setPage("import")} />
        )}
        {page === "fase1" && (
          <InterviewPhasePage
            title="Fase 1 — Entrevista de Soft Skills"
            subtitle="Candidato + Diretor do Departamento + 1 Membro RH — cruzamento de disponibilidades (Forms Fase 1 ∩ Excel Mestre)."
            phaseKey="fase1" availField="fase1" formsField="fase1" prevStatusField="phase0Status"
            candidates={candidates} members={members}
            bookings={phase1Bookings} setBookings={setPhase1Bookings}
            onGenerate={() => setPhase1Bookings(generateInterviewPhase(phase1Pool, members, phase1Bookings, "fase1", ["diretorId", "rhId"]))}
            columns={[{ key: "diretorId", label: "Diretor(a)" }, { key: "rhId", label: "RH" }]}
            showCalendar={false}
          />
        )}
        {page === "fase2" && (
          <Phase2Page
            candidates={candidates} members={members}
            groups={phase2Groups} setGroups={setPhase2Groups}
            onGenerate={() => setPhase2Groups(generatePhase2(phase2Pool, members))}
          />
        )}
        {page === "fase3" && (
          <InterviewPhasePage
            title="Fase 3 — Entrevista de Hard Skills"
            subtitle="Candidato + Diretor + 1 Membro RH + Supervisor do Departamento — cruzamento exato entre 4 intervenientes (Forms Fase 3 ∩ Excel Mestre)."
            phaseKey="fase3" availField="fase3" formsField="fase3" prevStatusField="phase2Status"
            candidates={candidates} members={members}
            bookings={phase3Bookings} setBookings={setPhase3Bookings}
            onGenerate={() => setPhase3Bookings(generateInterviewPhase(phase3Pool, members, phase3Bookings, "fase3", ["diretorId", "rhId", "supervisorId"]))}
            columns={[{ key: "diretorId", label: "Diretor(a)" }, { key: "rhId", label: "RH" }, { key: "supervisorId", label: "Supervisor" }]}
            showCalendar={true}
          />
        )}
      </main>

      {showAddCandidate && (
        <AddCandidateModal onClose={() => setShowAddCandidate(false)}
          onSave={(c) => { setCandidates((prev) => [...prev, c]); setShowAddCandidate(false); }} />
      )}
    </div>
  );
}
