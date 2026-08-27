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
// Etiqueta de build/versão do código — muda sempre que este ficheiro é
// atualizado (visível no canto do cabeçalho, ver TopNav). Serve para
// confirmar a olho, sem depender de memória/promessas, se a app está
// mesmo a correr a versão mais recente do código depois de um deploy —
// que foi a causa real da última ronda de "os bugs persistem": as
// correções já estavam no ficheiro entregue, mas a app em ecrã ainda
// estava a correr uma versão anterior.
const APP_BUILD = "build-2026-08-19-rh-diag";

const DAYS = ["Seg", "Ter", "Qua", "Qui", "Sex"];
const TIMES = ["09:00", "10:30", "14:00", "15:30", "17:00"];
const SLOTS = DAYS.flatMap((d) => TIMES.map((t) => `${d} ${t}`));

// Duração PADRÃO (fallback) de um slot de entrevista, em minutos — usada
// quando não há fase específica em jogo (ex. disponibilidade de
// Diretor/Supervisor/RH lida do Excel Mestre, que continua a ser um único
// "x/sim" por horário oficial, sem sub-divisão em blocos de 30 min).
// CORREÇÃO: os candidatos, no Forms, NÃO marcam os 5 horários oficiais
// diretamente — marcam uma grelha de checkboxes por BLOCO de 30 min (ex.
// "9h-9h30", "9h30-10h", "11h30-12h", ...). Um único bloco marcado só cobre
// a duração da Fase 2 (Soft Skills, 30 min); a Fase 4 (Hard Skills) precisa
// de 2 blocos SEGUIDOS (60 min) e a Fase 3 (Dinâmicas de Grupo) de 3 blocos
// SEGUIDOS (90 min) — ver PHASE_DURATION_MIN, passado a extractAvailability-
// FromRow()/parseAvailabilityCell() para o upload de cada Forms de fase.
const SLOT_DURATION_MIN = 30;
// Duração exigida (minutos) para considerar um candidato disponível num
// horário oficial, por fase de entrevista (phaseKey interno -> minutos):
// fase1 = Fase 2 (Soft Skills, 30 min), fase2 = Fase 3 (Dinâmicas de Grupo,
// 90 min), fase3 = Fase 4 (Hard Skills, 60 min).
const PHASE_DURATION_MIN = { fase1: 30, fase2: 90, fase3: 60 };
// Mapa slot canónico -> { day, startMin }, calculado uma única vez. O
// "endMin" de cada slot já NÃO é fixo aqui — passou a ser calculado no
// momento (startMin + duração exigida pela fase em causa), porque essa
// duração agora varia consoante a fase (ver PHASE_DURATION_MIN e
// expandRangesToSlots).
const SLOT_INFO = {};
SLOTS.forEach((slot) => {
  const [day, time] = slot.split(" ");
  const [hh, mm] = time.split(":").map(Number);
  const startMin = hh * 60 + mm;
  SLOT_INFO[slot] = { day, startMin };
});

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
// Opção "sem filtro" do seletor de Departamento nas fases individuais
// (Soft Skills / Hard Skills — não nas Dinâmicas de Grupo, que continuam a
// juntar candidatos de vários departamentos por desenho).
const ALL_DEPARTMENTS_OPTION = "Todos os Departamentos";

// Constrói os membros base (Diretor/Supervisor/RH) diretamente da estrutura
// organizacional fixa (ORG) — GARANTE que a associação Departamento ->
// Diretor/Supervisor/RH está sempre presente, mesmo antes de qualquer
// sincronização com o Excel Mestre ou se essa sincronização falhar a
// encontrar algum nome.
//
// CORREÇÃO PARA "Sem alocação": antes, a lista de membros só existia depois
// de uma sincronização bem-sucedida (a app arranca sempre com 0 membros —
// ver comentário mais abaixo). Isto significa que QUALQUER problema de
// sincronização (folha ainda não ligada, sessão a precisar de novo login,
// nome de RH escrito de forma diferente na aba, aba ainda não lida) deixava
// a coluna RH permanentemente "Sem alocação" para departamentos inteiros —
// não porque faltasse RH no departamento, mas porque a app simplesmente
// ainda não sabia que esse RH existia. Com esta base, a associação
// Diretor/Supervisor/RH de cada departamento está sempre lá desde o
// arranque; a sincronização com o Excel Mestre só ACRESCENTA disponibi-
// lidade (e, se necessário, atualiza/corrige nomes) por cima desta base —
// o merge é feito por nome normalizado (findMemberIndex/normKey), pelo que
// nunca cria duplicados.
function buildOrgBaselineMembers() {
  const members = [];
  const upsert = (name, role, dept) => {
    name = String(name || "").trim();
    if (!name) return;
    const idx = members.findIndex((m) => normKey(m.name) === normKey(name));
    if (idx >= 0) {
      const depts = new Set(members[idx].departments);
      depts.add(dept);
      members[idx] = { ...members[idx], departments: Array.from(depts) };
    } else {
      members.push({ id: uid("org"), name, role, title: role, departments: [dept], availability: [] });
    }
  };
  ORG.forEach((o) => {
    upsert(o.diretor, "Diretor", o.dept);
    upsert(o.supervisor, "Supervisor", o.dept);
    o.rh.forEach((n) => upsert(n, "RH", o.dept));
  });
  return members;
}

const PHASE_LABEL = { fase1: "Fase 2", fase2: "Fase 3", fase3: "Fase 4" };
// Estado de disponibilidade do candidato para a fase em curso — controlável
// manualmente via dropdown em cada separador de fase, ou atualizado
// automaticamente para "recebida" quando o respetivo Forms é importado.
const AVAILABILITY_STATES = ["nao_enviada", "pendente", "recebida"];
const AVAILABILITY_LABEL = { nao_enviada: "Não Enviada", pendente: "Pendente", recebida: "Recebida" };
const AVAILABILITY_TONE = {
  nao_enviada: { bg: "#e2e8f0", fg: "#475569" },
  pendente: { bg: "#fde68a", fg: "#92400e" },
  recebida: { bg: "#bbf7d0", fg: "#065f46" },
};
function AvailabilitySelect({ value, onChange }) {
  const tone = AVAILABILITY_TONE[value] || AVAILABILITY_TONE.nao_enviada;
  return (
    <select
      value={value || "nao_enviada"}
      onChange={(e) => onChange(e.target.value)}
      className="text-xs rounded-md px-2 py-1 border font-medium"
      style={{ backgroundColor: tone.bg, color: tone.fg, borderColor: "transparent" }}
    >
      {AVAILABILITY_STATES.map((s) => <option key={s} value={s} style={{ color: "#1f2937" }}>{AVAILABILITY_LABEL[s]}</option>)}
    </select>
  );
}

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
// Normaliza texto para comparação de cabeçalhos/valores: remove acentos,
// baixa para minúsculas, apara espaços e colapsa espaços internos múltiplos.
// Essencial para que pequenas variações no Excel Mestre (maiúsculas,
// acentuação, espaços a mais) nunca façam uma coluna "desaparecer".
function normKey(s) {
  return String(s ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .trim().toLowerCase().replace(/\s+/g, " ");
}
function get(row, ...keys) {
  for (const k of keys) {
    const nk = normKey(k);
    for (const rk of Object.keys(row)) {
      if (normKey(rk) === nk) return row[rk];
    }
  }
  return "";
}

// Limpa o CONTEÚDO BRUTO de uma célula (nome, email, disponibilidade, etc.)
// antes de qualquer trim/parse: remove caracteres invisíveis comuns em
// exports de Google/Microsoft Forms — zero-width space/joiner (\u200B-
// \u200D), BOM (\uFEFF) — troca espaço inseparável (NBSP, \u00A0) por um
// espaço normal, normaliza quebras de linha "\r\n"/"\r" para "\n", e só
// depois apara os espaços nas pontas. CORREÇÃO: estes caracteres são
// invisíveis no Excel/Sheets mas sobrevivem à exportação — uma célula que
// parece vazia ("Segunda 09:00\r\n") ou um "N/A" com um NBSP a seguir
// passavam incólumes por um simples `.trim()` e faziam o parser falhar
// silenciosamente em linhas que, a olho nu, pareciam perfeitamente normais.
function cleanCellText(raw) {
  return String(raw ?? "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\r\n?/g, "\n")
    .trim();
}

// Respostas explícitas de "não tenho disponibilidade" — o candidato
// respondeu ao Forms, mas para dizer que não tem nenhum horário possível
// (ou a pergunta não se aplica). Isto é uma resposta VÁLIDA, não um erro de
// leitura: o candidato deve continuar a ser aceite no sistema, só que com
// disponibilidade vazia ([]) nesse dia/campo — nunca contabilizado como
// "linha sem horário reconhecido" (ver isNoAvailabilityResponse).
const NO_AVAILABILITY_PHRASES = new Set([
  "nao tenho disponibilidade", "sem disponibilidade", "indisponivel",
  "nao disponivel", "n a", "na", "nenhum", "nenhuma",
  "nenhum dos horarios", "nenhum dos horarios disponiveis",
  "nenhuma das opcoes", "nenhuma opcao", "none", "n d", "nd", "-",
]);
// Devolve true se `raw` for uma célula em branco OU um dos textos acima —
// ou seja, uma resposta que deve ser aceite com disponibilidade vazia, sem
// disparar aviso de "horário não reconhecido".
function isNoAvailabilityResponse(raw) {
  const cleaned = cleanCellText(raw);
  if (!cleaned) return true;
  const norm = normKey(cleaned).replace(/[^a-z0-9]+/g, " ").trim();
  return NO_AVAILABILITY_PHRASES.has(norm);
}

// Junta uma lista de valores em texto no formato "1, 2 e 3" (PT), usado
// para listar números de linha / nomes de candidatos em avisos.
function joinWithE(items) {
  const arr = items.map(String);
  if (arr.length <= 1) return arr.join("");
  return `${arr.slice(0, -1).join(", ")} e ${arr[arr.length - 1]}`;
}

// Máximo de linhas listadas explicitamente no aviso de "sem horários
// selecionados" — evita uma mensagem gigante quando muitas linhas têm o
// mesmo problema; o resto fica resumido num "+N linha(s)".
const MAX_WARNING_ROWS = 8;
// Constrói a mensagem de aviso pedida: identifica exatamente QUAIS linhas/
// candidatos ficaram sem nenhum horário reconhecido (nunca apenas uma
// contagem), distinguindo isso de candidatos que simplesmente não
// submeteram disponibilidade (esses nem entram nesta lista — ver
// hasUnrecognizedContent em extractAvailabilityFromRow). `rows` é uma lista
// de { rowNumber, name }.
function buildUnrecognizedRowsWarning(rows) {
  if (!rows.length) return null;
  const shown = rows.slice(0, MAX_WARNING_ROWS);
  const rowNumbers = joinWithE(shown.map((r) => r.rowNumber));
  const names = shown.map((r) => r.name).join(", ");
  const extra = rows.length > MAX_WARNING_ROWS ? ` (+ ${rows.length - MAX_WARNING_ROWS} linha(s) adicional(is))` : "";
  const plural = rows.length > 1;
  return `Linha${plural ? "s" : ""} ${rowNumbers} (${names})${extra} sem horários selecionados — confirma o formato dessas células (coluna por slot, "Dia Hora" numa coluna por dia, ou texto livre "Seg 09:00 | Ter 10:30").`;
}

// Constrói o aviso INFORMATIVO (não é erro) pedido para candidatos
// legitimamente sem disponibilidade nesta fase: célula em branco,
// "Nenhum dos horários"/"N/A"/etc., OU blocos de 30 min reconhecidos mas
// que — mesmo fundidos — não chegam à duração exigida por esta fase (ex.:
// só marcou 1 bloco de 30 min numa fase que precisa de 60/90). Nestes
// casos o candidato É importado normalmente, com 0 slots — este aviso
// serve só para o RH perceber, de relance, que não foi um erro de leitura.
function buildNoAvailabilityInfo(names) {
  if (!names.length) return null;
  const shown = names.slice(0, MAX_WARNING_ROWS);
  const extra = names.length > MAX_WARNING_ROWS ? ` (+ ${names.length - MAX_WARNING_ROWS} candidato(s) adicional(is))` : "";
  const plural = names.length > 1;
  return `Candidato${plural ? "s" : ""} ${joinWithE(shown)} importado${plural ? "s" : ""} (sem disponibilidade assinalada)${extra}.`;
}

// Valores de erro típicos de fórmulas do Excel/Sheets (ex.: candidato ainda
// não chegou a essa fase -> a fórmula devolve #N/A) ou de células vazias.
// Estes NÃO são nomes/emails válidos e devem ser tratados como "slot ainda
// por preencher", nunca como um candidato desconhecido a reportar.
const SHEET_ERROR_VALUES = new Set(["#N/A", "#VALUE!", "#REF!", "#NAME?", "#NULL!", "#DIV/0!", "#ERROR!"]);
function isErrorOrEmptyValue(v) {
  if (v === null || v === undefined) return true;
  const s = String(v).trim();
  if (s === "") return true;
  return SHEET_ERROR_VALUES.has(s.toUpperCase());
}
function matchDept(raw) {
  const norm = normKey(raw).replace(/[^a-z0-9]+/g, " ").trim();
  if (!norm) return null;
  return DEPARTMENTS.find((d) => normKey(d).replace(/[^a-z0-9]+/g, " ").trim() === norm) || null;
}


// Encontra o índice de um membro (Diretor/Supervisor/RH) já existente pelo
// Nome, usando normKey() (trim + lowercase + sem acentos + espaços
// colapsados) em vez de uma simples comparação de .toLowerCase().
// CORREÇÃO CRÍTICA: era exatamente esta comparação frágil (só .toLowerCase(),
// sem normalizar acentos/espaços) que fazia a app criar DOIS registos
// diferentes para a mesma pessoa sempre que o nome vinha escrito de forma
// ligeiramente diferente entre a aba "Base Dados Departamentos" (que define
// Diretor/Supervisor/Membros RH por departamento) e as abas de
// disponibilidade "Disponibilidade Entrevistas RH/Dinâmicas/Entrevista
// Final" (ex.: "Tânia Silva" vs "Tania Silva " com espaço a mais, ou
// maiúsculas diferentes). Resultado: um registo do membro ficava com
// `departments` preenchido mas `availability: []`, e o outro com
// `availability` preenchida mas `departments: []` — daí a coluna RH
// aparecer como "Sem alocação" (o registo "certo" nunca tinha disponibilidade
// para entrar no cruzamento) e o Horário ficar vazio / "Sem Horário Comum"
// mesmo quando os dados existiam no Excel. Isto espelha o mesmo bug (e a
// mesma correção) já aplicado a matchCandidateIndex() para candidatos — só
// que nunca tinha sido replicado para membros.
function findMemberIndex(members, name) {
  const n = normKey(name);
  if (!n) return -1;
  return members.findIndex((m) => normKey(m.name) === n);
}

// Encontra o índice de um candidato já existente. Chave primária: EMAIL
// (único e constante — nunca varia entre abas). Só recorre ao Nome como
// segunda chave quando o email está em branco ou não casa, e mesmo assim
// só depois de sanitizado (.trim(), .toLowerCase(), remoção de acentos
// via normalize("NFD")) através de normKey(). Isto substitui as
// comparações antigas `c.name.toLowerCase() === name.toLowerCase()`, que
// falhavam sempre que o nome vinha escrito de forma ligeiramente
// diferente entre abas (ex.: "João Complicado" vs "Joao Complicado " com
// espaço a mais) — a causa raiz de aprovações da Coluna Q não chegarem
// aos candidatos de alguns departamentos (Quality Management, Legal &
// Finance, etc.), ficando presas apenas em Human Resources.
function matchCandidateIndex(candidates, name, email) {
  const nEmail = normKey(email);
  if (nEmail) {
    const i = candidates.findIndex((c) => c.email && normKey(c.email) === nEmail);
    if (i >= 0) return i;
  }
  const nName = normKey(name); // sanitizado: trim + lowercase + sem acentos
  if (nName) {
    const i = candidates.findIndex((c) => normKey(c.name) === nName);
    if (i >= 0) return i;
  }
  return -1;
}

// Deriva em que fase um candidato foi eliminado, olhando para o primeiro
// estado "Rejeitado" na sequência Fase 1 -> Fase 2 -> Fase 3 -> Fase 4
// (um candidato só pode ser eliminado uma vez, no primeiro corte que falhar).
function getEliminationPhase(c) {
  if (c.phase0Status === "Rejeitado") return "fase1";
  if (c.phase1Status === "Rejeitado") return "fase2";
  if (c.phase2Status === "Rejeitado") return "fase3";
  if (c.finalResult === "Rejeitado") return "fase4";
  return null;
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
  softSkills: "L",   // Passou Entrevista Soft Skills/RH -> avança Fase 3 (Dinâmicas)
  dinamicas: "AA",   // Passou Dinâmicas de Grupo -> avança Fase 4 (Hard Skills)
  final: "AZ",       // Passou Desafio Final/Hard Skills -> SELECIONADO / ENTROU NA YME
  talentPoolA: "AB", // Selecionado para Talent Pool (variante de coluna 1)
  talentPoolB: "BA", // Selecionado para Talent Pool (variante de coluna 2)
};
// Coluna Q da Avaliação de CV (Fase 1): candidato passou para a Fase 2 (Entrevista Soft Skills/RH).
const SYNC_CV_PASS_COLUMN = "Q";
// Coluna R da Avaliação de CV (Fase 1): candidato NÃO passou (Rejeitado já na Fase 1).
const SYNC_CV_FAIL_COLUMN = "R";
// Coluna A (Departamento) e B (Nome) da aba "Avaliação CV e Questões
// Abertas" — Coluna C é "Ano do Curso" e nunca deve ser lida como nome.
const SYNC_CV_DEPT_COLUMN = "A";
const SYNC_CV_NAME_COLUMN = "B";

// "Palavras-chave" de cabeçalho usadas para localizar automaticamente a
// linha de cabeçalho real de cada aba (aceita que haja 1-4 linhas de
// título/logo acima do cabeçalho, situação comum em Excels Mestre com
// formatação). Sem isto, se o cabeçalho não estiver exatamente na
// primeira linha da aba, TODAS as linhas seriam lidas como dados e
// nenhum candidato seria reconhecido (nome/email a null em todas).
const SYNC_HEADER_HINTS = {
  departamentos: ["departamento", "diretor"],
  candidatos: ["nome", "nome completo", "email"],
  avaliacaoCV: ["nome", "nome completo"],
  dispEntrevistasRH: ["nome"],
  dispDinamicas: ["nome"],
  dispEntrevistaFinal: ["nome"],
};
const SYNC_DEPT_HEADER_HINTS = ["nome", "nome completo"];

// Índice de cabeçalho FIXO (0-based) para abas cuja estrutura real foi
// confirmada manualmente e não deve depender de deteção automática por
// palavras-chave — evita falsos negativos caso o texto exato do
// cabeçalho não bata certo com nenhuma das SYNC_HEADER_HINTS. Estrutura
// confirmada da aba "Avaliação CV e Questões Abertas": cabeçalho na
// LINHA 13 (índice 12), dados a começar estritamente na LINHA 14.
const SYNC_FIXED_HEADER_IDX = { avaliacaoCV: 12 };

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
          scope: "https://www.googleapis.com/auth/spreadsheets",
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
  if (val === true) return true;
  const v = String(val ?? "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (!v) return false;
  return ["x", "sim", "true", "verdadeiro", "1", "aprovado", "selecionado", "apto", "avanca", "yes", "✓", "v", "☑", "☒"].includes(v);
}
// Extensão de isPositiveMark() específica para MARCAÇÃO DE DISPONIBILIDADE
// em colunas-grelha (uma coluna = um slot exato, ex. "Seg 09:00"): além dos
// valores positivos já reconhecidos globalmente (x, sim, true, 1, ✓, etc.),
// aceita também "Disponível"/"disponivel" e "Check", e ainda o caso em que
// o candidato escreveu a PRÓPRIA HORA na célula em vez de um "x" (ex. a
// célula da coluna "Seg 09:00" contém literalmente "09:00" ou "9h"). Estes
// dois casos extra são específicos de respostas de disponibilidade — por
// isso ficam isolados aqui, em vez de serem acrescentados a isPositiveMark()
// globalmente (que também é usada para aprovações de fase, onde não fazem
// sentido).
function isAvailabilityPositiveMark(val) {
  const cleaned = cleanCellText(val);
  if (isPositiveMark(cleaned)) return true;
  const norm = normKey(cleaned).replace(/[^a-z0-9]/g, "");
  if (["disponivel", "check"].includes(norm)) return true;
  const { inicio } = parseTimeToMinutes(cleaned);
  return inicio !== null;
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
// headerHints: palavras-chave (normalizadas) usadas para localizar a
// linha de cabeçalho real dentro das primeiras 20 linhas da aba — cobre
// o caso (confirmado no Excel Mestre da YME) de haver várias linhas de
// título/logótipo decorativo (ex: "// Recrutamento YME //") acima do
// cabeçalho real (ex: cabeçalho na linha 6 = índice 5). Sem isto, se o
// cabeçalho não estiver na linha 1, nenhuma linha seria reconhecida
// como candidato (nome/email sempre vazios).
function parseApiValues(values, headerHints = [], fixedHeaderIdx = null) {
  const grid = values || [];
  let headerIdx = 0;
  if (fixedHeaderIdx !== null && fixedHeaderIdx !== undefined) {
    // Estrutura desta aba já foi confirmada manualmente — usa o índice
    // exato em vez de tentar adivinhar pela deteção de palavras-chave.
    headerIdx = fixedHeaderIdx;
  } else {
    const hints = headerHints.map(normKey);
    if (hints.length) {
      const found = grid.slice(0, 20).findIndex((row) => {
        if (!row || !row.length) return false;
        const norm = row.map((c) => normKey(c));
        return hints.some((h) => norm.includes(h));
      });
      if (found >= 0) headerIdx = found;
    }
  }
  let header = (grid[headerIdx] || []).map((h) => String(h ?? ""));
  let dataStartIdx = headerIdx + 1;
  // CABEÇALHO EM DUAS LINHAS (dia + hora) — ver mergeTwoRowHeader. Ler só
  // UMA linha (a que headerHints encontrou, ex. por conter "Nome") dava
  // cabeçalhos sem dia nenhum sempre que o dia estivesse na linha vizinha
  // ("Dia 10 - Quarta" numa linha, "9-9:30" noutra) — daí Diretores/RH
  // aparecerem sempre "sem horários no Excel Mestre" apesar das células
  // estarem preenchidas. Tenta, por esta ordem, até encontrar o par:
  //  1) dia ACIMA da linha de cabeçalho encontrada (layout mais comum:
  //     "Nome" está na mesma linha que as horas, o dia fica na linha de
  //     cima, em células fundidas);
  //  2) a PRÓPRIA linha de cabeçalho já É a linha de dias (ex.: "Nome"
  //     está colado aos rótulos de dia) e as horas vêm na linha seguinte;
  //  3) fallback totalmente independente de headerHints/"Nome": procura em
  //     toda a folha o primeiro par de linhas consecutivas dia+hora — cobre
  //     o caso de "Nome" não aparecer em nenhuma das duas linhas de
  //     cabeçalho (só nas linhas de dados a seguir).
  // NOTA: testa sempre as 3 hipóteses (não há pré-condição "só tenta se o
  // cabeçalho ainda não tiver dia nenhum") — essa pré-condição bloqueava
  // exatamente o caso 2: quando "Nome" está colado à linha de DIAS, essa
  // linha já "tem dia" por si só, mas continua sem NENHUMA hora, por isso a
  // fusão é sempre necessária mesmo assim. A segurança contra falsos
  // positivos (não confundir uma linha de dados normal com a linha de
  // horas) vem da exigência de DENSIDADE em looksLikeHourLabelRow, não de
  // adivinhar se o cabeçalho "já parece suficiente".
  if (headerIdx > 0 && looksLikeDayLabelRow(grid[headerIdx - 1]) && looksLikeHourLabelRow(header)) {
    header = mergeTwoRowHeader(grid[headerIdx - 1], header);
    dataStartIdx = headerIdx + 1;
  } else if (looksLikeDayLabelRow(header) && looksLikeHourLabelRow(grid[headerIdx + 1] || [])) {
    header = mergeTwoRowHeader(header, grid[headerIdx + 1]);
    dataStartIdx = headerIdx + 2;
  } else {
    const pair = findDayHourHeaderRows(grid);
    if (pair) {
      header = mergeTwoRowHeader(grid[pair.dayIdx], grid[pair.hourIdx]);
      dataStartIdx = pair.hourIdx + 1;
    }
  }
  const rows = grid
    .slice(dataStartIdx)
    .filter((r) => r.some((c) => String(c ?? "").trim() !== ""))
    .map((raw) => {
      const obj = {};
      header.forEach((h, i) => { if (h) obj[h] = raw[i] ?? ""; });
      return { obj, raw };
    });
  return { header, rows, headerIdx };
}

// Lê uma aba da folha privada através da Google Sheets API v4
// (spreadsheets.values.get), autenticado com o accessToken da sessão do
// utilizador. encodeURIComponent no nome da aba é essencial: nomes como
// "Base Dados Departamentos" têm espaços e, sem isto, o range fica
// inválido/mal interpretado pela API.
async function fetchSheetTabApi(accessToken, sheetId, sheetName, headerHints = [], fixedHeaderIdx = null) {
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

  return parseApiValues(data.values, headerHints, fixedHeaderIdx);
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

// Normaliza um "token" de dia da semana para uma das 5 chaves canónicas de
// ============================================================================
// FUNÇÃO PEDIDA: parseTimeToMinutes — conversão única de QUALQUER formato de
// hora para minutos desde a meia-noite. Esta é agora a ÚNICA função em todo
// o ficheiro que interpreta dígitos de hora; tudo o resto (Excel Mestre e
// Forms dos candidatos) passa por aqui antes de qualquer comparação.
//
// CORREÇÃO CRÍTICA: os 54 "Sem Horário Comum" vinham de comparar STRINGS
// diretamente — "10h30 - 11h00" (Forms, candidatos) nunca é === "10:30 -
// 11:00" (Excel Mestre, RH/Diretores/Supervisor), mesmo sendo o MESMO
// horário. A correção não é tentar prever todas as grafias possíveis: é
// nunca mais comparar texto. Tudo passa a ser convertido para minutos
// (números) logo na leitura, e a interseção de disponibilidades passa a ser
// aritmética sobre esses números, não comparação de strings.
//
// Aceita: "10h30 - 11h00", "10:30-11:00", "10:30", "10h30", "10.30",
// "10:30 às 11:00", "10:30–11:00" (travessão), "9h-9h30", "9h- 9h30",
// "9h00 - 9h30", "11h30 -12h00", "11h30 -12h", "17h-17h30",
// "2:00 PM", "10:30 AM", "2 PM", "14:00", "14h30", "14.30", etc., e ainda,
// em modo solto (allowBareHour:true — ver TIME_TOKEN_RE_LOOSE), horas sem
// minutos e sem separador nenhum: "9h", "17h", "9", "17", "9-17", "9 às 17",
// "09-10:30".
// Exemplo pedido: parseTimeToMinutes("10h30 - 11h00") -> { inicio: 630, fim: 660 }
//                 parseTimeToMinutes("10:30 - 11:00") -> { inicio: 630, fim: 660 }
//                 parseTimeToMinutes("2:00 PM")        -> { inicio: 840, fim: null }  (== 14:00)
//                 parseTimeToMinutes("10:30 AM")       -> { inicio: 630, fim: null }  (== 10:30)
//                 parseTimeToMinutes("9-17", { allowBareHour: true }) -> { inicio: 540, fim: 1020 }

// Regex universal de horas — a ÚNICA usada em todo o ficheiro para
// reconhecer texto de horas, partilhada por parseTimeToMinutes() (para
// extrair minutos) e normalizeSlotString() (para saber onde a hora começa
// dentro de um cabeçalho). Tem 2 ramos, tentados por esta ordem em cada
// posição:
//  1) 12h com AM/PM — "2:00 PM", "10:30 AM", "2 PM", "2:00PM", "10h30 a.m."
//     (grupo `ampm` só existe se um sufixo AM/PM for encontrado a seguir)
//  2) 24h com separador ':', 'h' ou '.' — "14:00", "14h30", "14.30", "9h"
//     (minutos aqui são OPCIONAIS: "9h"/"12h" valem "9h00"/"12h00")
const TIME_TOKEN_RE = /(?<h12>\d{1,2})(?:\s*[:.h]\s*(?<m12>\d{2}))?\s*(?<ampm>[ap]\.?\s?m\.?)\b|(?<h24>\d{1,2})\s*[:h.]\s*(?<m24>\d{2})?/gi;

// Variante "solta" de TIME_TOKEN_RE: acrescenta um 3º ramo que aceita um
// número de 1-2 dígitos SOZINHO — sem `:`, `h` nem `.` a seguir — como hora
// inteira (minutos = 00). É o que faltava para "9-17", "9 às 17" ou "09-10:30"
// (o "09" aí não tem separador nenhum a seguir): a versão estrita exigia
// sempre um separador só para reconhecer a HORA, e por isso rejeitava estes
// casos por completo (`parseTimeToMinutes` devolvia inicio:null, daí o
// "nenhum horário é compatível" nestes formatos).
//
// Só é usada para interpretar VALORES de disponibilidade já isolados (texto
// livre de uma célula, ou o token depois de retirado o dia) — NUNCA para
// reconhecer cabeçalhos de coluna. Cabeçalhos do Forms costumam trazer a
// DATA por extenso ("Quarta-feira, dia 10 de junho") e vários desses
// números de dia (9, 10, 14, 15, 17) coincidem com as horas oficiais da
// grelha (09:00, 10:30, 14:00, 15:30, 17:00) — se o ramo solto fosse usado
// também aí, "dia 9 de junho" seria lido como a hora "09:00", o cabeçalho
// passaria a ser tratado como slot exato e a coluna inteira desse dia
// deixaria de ser lida como texto livre (perda de disponibilidade real, um
// bug pior que o atual). Por isso `normalizeSlotString`/`normalizeTimeToken`
// continuam, por omissão, a usar a versão ESTRITA (TIME_TOKEN_RE); só
// `parseAvailabilityRanges` (valores, não cabeçalhos) ativa o modo solto.
// Exclui ainda números colados a "/" em qualquer um dos lados, para não
// apanhar datas dd/mm escritas por engano dentro do próprio valor.
const TIME_TOKEN_RE_LOOSE = /(?<h12>\d{1,2})(?:\s*[:.h]\s*(?<m12>\d{2}))?\s*(?<ampm>[ap]\.?\s?m\.?)\b|(?<h24>\d{1,2})\s*[:h.]\s*(?<m24>\d{2})?|(?<![\d/])(?<h24b>\d{1,2})(?!\d)(?!\s*\/)/gi;

function parseTimeToMinutes(str, { allowBareHour = false } = {}) {
  const s = String(str || "");
  const re = allowBareHour ? TIME_TOKEN_RE_LOOSE : TIME_TOKEN_RE;
  const matches = [];
  for (const m of s.matchAll(re)) {
    if (m.groups.ampm !== undefined) {
      // 12h -> 24h: 12 AM = 00h; 12 PM = 12h; 1-11 AM ficam iguais;
      // 1-11 PM somam 12h. "2:00 PM" == "14:00", "10:30 AM" == "10:30".
      let hour = Number(m.groups.h12) % 12;
      if (/^p/i.test(m.groups.ampm)) hour += 12;
      matches.push(hour * 60 + Number(m.groups.m12 || 0));
    } else if (m.groups.h24 !== undefined) {
      matches.push(Number(m.groups.h24) * 60 + Number(m.groups.m24 || 0));
    } else if (m.groups.h24b !== undefined) {
      // Hora solta sem separador nenhum ("9", "17") — minutos assumidos 00.
      matches.push(Number(m.groups.h24b) * 60);
    }
  }
  if (!matches.length) return { inicio: null, fim: null };
  // Primeiro valor = início; último valor = fim (um único horário pontual
  // devolve fim=null — ver parseAvailabilityRanges abaixo, que nesse caso
  // assume a duração padrão de 1 slot).
  return { inicio: matches[0], fim: matches.length > 1 ? matches[matches.length - 1] : null };
}
// ============================================================================

// Deteta um "token" de dia da semana para uma das 5 chaves canónicas de
// DAYS ("Seg".."Sex"), aceitando abreviações, nomes completos, com/sem
// acentos, com/sem "-feira", maiúsculas/minúsculas, pontuação, etc. —
// e também os nomes em INGLÊS ("Mon"/"Monday" .. "Fri"/"Friday"), porque
// alguns exports do Forms/Sheets vêm com o idioma da conta Google em
// inglês. Funciona por PREFIXO de 3 letras depois de normalizado (normKey
// já remove acentos/maiúsculas), o que cobre "Seg", "seg.", "Segunda",
// "segunda-feira", "SEGUNDA FEIRA", "Mon", "Monday", etc. — todos colapsam
// no mesmo prefixo de 3 letras ("seg"/"mon").
const DAY_PREFIXES = {
  seg: "Seg", ter: "Ter", qua: "Qua", qui: "Qui", sex: "Sex",
  mon: "Seg", tue: "Ter", wed: "Qua", thu: "Qui", fri: "Sex",
};
function normalizeDayToken(raw) {
  const n = normKey(raw).replace(/[^a-z]/g, "");
  if (!n) return null;
  return DAY_PREFIXES[n.slice(0, 3)] || null;
}

// Normaliza um "token" de horário para o formato canónico "HH:MM" usado em
// TIMES — agora um wrapper fino sobre parseTimeToMinutes(), para que exista
// UMA SÓ função a interpretar dígitos de hora em todo o ficheiro. Devolve
// null se não encontrar nenhuma hora reconhecível.
// `allowBareHour` por omissão fica a false: esta função é chamada por
// normalizeSlotString() sobre CABEÇALHOS de coluna, onde números soltos
// costumam ser datas ("dia 10 de junho") e não horas — ver nota grande em
// TIME_TOKEN_RE_LOOSE acima. Passa allowBareHour:true apenas quando `raw`
// for a seguro de ser só texto de horário (nunca um cabeçalho com data).
function normalizeTimeToken(raw, { allowBareHour = false } = {}) {
  const { inicio } = parseTimeToMinutes(raw, { allowBareHour });
  if (inicio === null) return null;
  const hh = String(Math.floor(inicio / 60)).padStart(2, "0");
  const mm = String(inicio % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

// Mapa hora oficial (string de TIMES) -> minutos desde a meia-noite,
// construído uma única vez, para poder calcular qual a hora oficial MAIS
// PRÓXIMA de uma hora extraída de texto livre (ver nearestOfficialTime).
const OFFICIAL_TIME_MINUTES = TIMES.map((t) => {
  const [hh, mm] = t.split(":").map(Number);
  return hh * 60 + mm;
});
// Tolerância (minutos) para aceitar um cabeçalho de coluna cuja hora não
// bate CERTINHA com nenhum dos 5 horários oficiais, mas anda perto o
// suficiente para ser inequivocamente o mesmo slot (pequenas variações de
// escrita, arredondamentos do Forms, etc.).
// CORREÇÃO: o bug reportado ("X de Y linhas sem nenhum horário
// reconhecido") vinha de aqui se exigir IGUALDADE EXATA de string com um
// dos 5 horários oficiais — qualquer desvio, por mínimo que fosse, fazia a
// coluna inteira ser ignorada. Agora usa-se a hora oficial mais próxima,
// dentro desta margem.
const SLOT_MATCH_TOLERANCE_MIN = 45;

// Devolve o horário oficial (uma das strings de TIMES) mais próximo de
// `minutes` (minutos desde a meia-noite), ou null se mesmo o mais próximo
// ficar a mais de SLOT_MATCH_TOLERANCE_MIN minutos de distância — nesse
// caso não é seguro assumir que se trata do mesmo slot oficial.
function nearestOfficialTime(minutes) {
  if (minutes === null || minutes === undefined) return null;
  let best = null;
  let bestDiff = Infinity;
  TIMES.forEach((t, i) => {
    const diff = Math.abs(OFFICIAL_TIME_MINUTES[i] - minutes);
    if (diff < bestDiff) { bestDiff = diff; best = t; }
  });
  return bestDiff <= SLOT_MATCH_TOLERANCE_MIN ? best : null;
}

// Normaliza uma string livre de "Dia + Hora" (célula de cabeçalho de
// grelha, ou um item de uma lista separada por | ; ,) para um dos slots
// canónicos de SLOTS ("Seg 09:00", etc.), ou null se não for possível
// reconhecer com confiança um dia E uma hora suficientemente próxima de um
// dos 5 horários oficiais da grelha (TIMES, com tolerância — ver
// nearestOfficialTime). Usada para o caso "ponto exato" (cabeçalho = o
// próprio slot).
function normalizeSlotString(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const { inicio } = parseTimeToMinutes(s);
  const time = nearestOfficialTime(inicio);
  if (!time) return null;
  const timeIdx = s.search(TIME_TOKEN_RE);
  const dayPart = timeIdx > 0 ? s.slice(0, timeIdx) : s;
  const day = normalizeDayToken(dayPart);
  if (!day) return null;
  const slot = `${day} ${time}`;
  return SLOTS.includes(slot) ? slot : null;
}

// Deteta o DIA da semana referido num cabeçalho de coluna, mesmo quando o
// cabeçalho tem texto extra à volta — CORREÇÃO PARA EXPORTS DO GOOGLE FORMS
// COM VÁRIAS SECÇÕES: quando o Forms tem uma secção por departamento e cada
// secção repete a mesma pergunta de disponibilidade, o Excel/Sheets gera
// automaticamente um cabeçalho por secção com sufixo para os manter únicos
// — ex. "Quarta-feira, dia 10 de junho", "Quarta-feira, dia 10 de junho 2",
// "...  3", até 6. A verificação ANTERIOR exigia que o cabeçalho fosse
// EXATAMENTE "Quarta"/"Quarta-feira" (nada mais), pelo que NENHUMA destas
// colunas (nem a primeira, nem as seguintes) era reconhecida como coluna de
// disponibilidade — todas eram ignoradas, daí os 0 agendamentos. Agora
// procura-se o nome do dia (com "-feira", ou o nome completo em inglês) em
// QUALQUER parte do texto do cabeçalho, o que reconhece a coluna seja qual
// for o sufixo/data anexado E o idioma da conta Google usada no Forms.
const DAY_WORD_PATTERNS = [
  { re: /segunda[-\s]?feira/i, day: "Seg" },
  { re: /ter[cç]a[-\s]?feira/i, day: "Ter" },
  { re: /quarta[-\s]?feira/i, day: "Qua" },
  { re: /quinta[-\s]?feira/i, day: "Qui" },
  { re: /sexta[-\s]?feira/i, day: "Sex" },
  // Nomes completos SEM o sufixo "-feira" — cobre o Excel Mestre, onde a
  // linha de cabeçalho de dia costuma vir como "Dia 10 - Quarta", "Dia 11 -
  // Quinta", etc. (nunca "Quarta-feira"). \b...\b evita apanhar estas
  // palavras dentro de outras mais compridas.
  { re: /\bsegunda\b/i, day: "Seg" },
  { re: /\bter[cç]a\b/i, day: "Ter" },
  { re: /\bquarta\b/i, day: "Qua" },
  { re: /\bquinta\b/i, day: "Qui" },
  { re: /\bsexta\b/i, day: "Sex" },
  { re: /\bmonday\b/i, day: "Seg" },
  { re: /\btuesday\b/i, day: "Ter" },
  { re: /\bwednesday\b/i, day: "Qua" },
  { re: /\bthursday\b/i, day: "Qui" },
  { re: /\bfriday\b/i, day: "Sex" },
];
// Whitelist EXATA (fallback, sem "-feira"/nome completo) para cabeçalhos
// que são só a abreviação/nome do dia e mais nada (ex. "Seg", "Quarta",
// "Mon", "Tuesday"). Continua a exigir igualdade exata para nunca confundir
// "Sexo"/"Segmento" com "Sex"/"Seg" só por partilharem prefixo.
const DAY_ONLY_WORDS = {
  seg: "Seg", segunda: "Seg", segundafeira: "Seg",
  ter: "Ter", terca: "Ter", tercafeira: "Ter",
  qua: "Qua", quarta: "Qua", quartafeira: "Qua",
  qui: "Qui", quinta: "Qui", quintafeira: "Qui",
  sex: "Sex", sexta: "Sex", sextafeira: "Sex",
  mon: "Seg", monday: "Seg",
  tue: "Ter", tues: "Ter", tuesday: "Ter",
  wed: "Qua", weds: "Qua", wednesday: "Qua",
  thu: "Qui", thur: "Qui", thurs: "Qui", thursday: "Qui",
  fri: "Sex", friday: "Sex",
};
function normalizeDayOnlyHeader(raw) {
  const s = String(raw || "");
  for (const { re, day } of DAY_WORD_PATTERNS) {
    if (re.test(s)) return day;
  }
  const n = normKey(s).replace(/[^a-z]/g, "");
  if (DAY_ONLY_WORDS[n]) return DAY_ONLY_WORDS[n];
  // Fallback por NÚMERO DO DIA (ver bloco DAY_NUMBER_TO_WEEKDAY abaixo):
  // cobre cabeçalhos do Excel Mestre que só trazem a data ("10/06",
  // "Dia 10"), sem nome de dia da semana nenhum.
  const num = extractDayNumber(s);
  return num !== null ? (DAY_NUMBER_TO_WEEKDAY[num] || null) : null;
}

// ----------------------------------------------------------------------
// CABEÇALHO EM DUAS LINHAS (Excel Mestre: dia numa linha, hora na
// seguinte) — ex. linha "Dia 10 - Quarta" / "Dia 11 - Quinta" por cima da
// linha "9-9:30" / "9:30-10" / "10-10:30" / ... CORREÇÃO: ler só a linha
// das horas (a única que `parseApiValues` via considerar até agora) dava
// cabeçalhos SEM nenhum dia ("9-9:30"), que normalizeDayOnlyHeader nunca
// reconhece — daí Diretores/RH aparecerem sempre "sem horários no Excel
// Mestre" apesar de as células estarem preenchidas. As 3 funções abaixo
// combinam as duas linhas numa só, ANTES dessa combinação chegar a
// extractAvailabilityFromRow.
// ----------------------------------------------------------------------

// Deteta se uma linha da grelha se parece com uma linha de RÓTULOS DE DIA
// — basta 1 célula reconhecível por normalizeDayOnlyHeader (com "-feira",
// nome completo PT/EN, ou "Dia N"/"Dia N - <dia>").
function looksLikeDayLabelRow(row) {
  return (row || []).some((v) => {
    const s = String(v ?? "").trim();
    return s && normalizeDayOnlyHeader(s) !== null;
  });
}

// Propaga (forward-fill) o último valor não vazio de uma linha para as
// células vazias seguintes — necessário porque, no Excel/Sheets, o rótulo
// do dia normalmente só aparece na PRIMEIRA das várias colunas de meia
// hora desse dia (célula fundida/"merged cell" na folha original); as
// colunas seguintes do mesmo dia chegam vazias da API/leitura.
function forwardFillRow(row) {
  const filled = [];
  let last = "";
  (row || []).forEach((v) => {
    const s = String(v ?? "").trim();
    if (s) last = s;
    filled.push(last);
  });
  return filled;
}

// Combina uma linha de DIAS (ex. "Dia 10 - Quarta", já com forward-fill)
// com uma linha de HORAS (ex. "9-9:30") num único array de cabeçalhos —
// um por coluna. IMPORTANTE: o dia é primeiro RESOLVIDO para a sua forma
// curta canónica ("Qua") via normalizeDayOnlyHeader, em vez de manter o
// texto original ("Dia 10 - Quarta") colado à hora — isto evita que o "10"
// de "Dia 10" seja mais tarde confundido com uma hora solta quando o
// cabeçalho combinado for reprocessado (ex. "Dia 10 - Quarta 9-9:30"
// poderia, em teoria, ler "10" como hora; "Qua 9-9:30" não tem esse
// problema, porque já não sobra nenhum dígito do dia do mês).
function mergeTwoRowHeader(dayRow, hourRow) {
  const days = forwardFillRow(dayRow);
  return hourRow.map((h, i) => {
    const hourText = String(h ?? "").trim();
    const dayLabel = days[i] || "";
    if (!dayLabel) return hourText;
    const resolvedDay = normalizeDayOnlyHeader(dayLabel);
    if (!resolvedDay) return hourText ? `${dayLabel} ${hourText}`.trim() : dayLabel;
    return hourText ? `${resolvedDay} ${hourText}` : resolvedDay;
  });
}

// Deteta se uma linha da grelha se parece com uma linha de RÓTULOS DE HORA
// — exige não só pelo menos 3 células reconhecidas como um INTERVALO de
// horas genuíno (2 valores, início < fim, em modo solto — aceita "9-9:30"
// sem separador nenhum), mas também que essas células sejam a GRANDE
// MAIORIA das células não vazias da linha (≥70%). A densidade + o mínimo
// de 3 é o que distingue uma verdadeira linha de cabeçalho de horas (onde
// praticamente todas as colunas são um intervalo, tipicamente uma grelha
// de várias dezenas de blocos de 30 min) de uma linha de DADOS normal que,
// por acaso, tenha 1-2 células parecidas com horas (ex. um candidato que
// respondeu com um intervalo de texto livre em 2 colunas) misturadas com
// nome/email/etc.
function looksLikeHourLabelRow(row) {
  const cells = (row || []).map((v) => String(v ?? "").trim()).filter(Boolean);
  if (cells.length < 3) return false;
  const rangeCount = cells.filter((s) => {
    const { inicio, fim } = parseTimeToMinutes(s, { allowBareHour: true });
    return inicio !== null && fim !== null && fim > inicio;
  }).length;
  return rangeCount >= 3 && rangeCount / cells.length >= 0.7;
}

// Procura, em toda a folha (até `searchLimit` linhas), o primeiro par de
// linhas CONSECUTIVAS dia+hora — usado como último recurso quando nem a
// linha de cabeçalho encontrada por headerHints, nem a linha imediatamente
// acima dela, formam esse par (ex.: o texto "Nome" não aparece em nenhuma
// das duas linhas de cabeçalho, só nas linhas de dados a seguir).
function findDayHourHeaderRows(grid, searchLimit = 20) {
  const limit = Math.min((grid || []).length - 1, searchLimit);
  for (let i = 0; i < limit; i++) {
    if (looksLikeDayLabelRow(grid[i]) && looksLikeHourLabelRow(grid[i + 1])) {
      return { dayIdx: i, hourIdx: i + 1 };
    }
  }
  return null;
}


// MAPEAMENTO DE DIAS POR NÚMERO DO MÊS
// ----------------------------------------------------------------------
// O Forms dos candidatos identifica cada dia por NOME + DATA no mesmo
// cabeçalho: "Quarta-feira, dia 10 de junho". O Excel Mestre (disponibi-
// lidade de Diretor/Supervisor/RH), pelo contrário, identifica o mesmo dia
// só pela DATA, sem nome de dia da semana: "10/06" ou "Dia 10". Como o
// modelo interno da app organiza tudo por dia da semana (Seg..Sex, ver
// SLOTS), é preciso descobrir a que dia da semana corresponde cada data —
// e essa correspondência só está disponível nos cabeçalhos do Forms
// (que trazem os dois juntos). DAY_NUMBER_TO_WEEKDAY guarda esse mapa
// (ex.: { 10: "Qua", 11: "Qui", 12: "Sex" }), construído automaticamente
// à medida que os cabeçalhos do Forms são lidos, e é depois consultado
// para resolver as colunas do Excel Mestre que só têm a data.
let DAY_NUMBER_TO_WEEKDAY = {};

// Extrai o dia do mês (1-31) de um texto, tentando por ordem:
//  1) "dia 10", "Dia 10"
//  2) "10 de julho", "19 de junho" — dia por extenso, com nome do mês a
//     seguir (não exige a palavra "dia" antes, ao contrário do padrão 1)
//  3) "19/06", "19/07/2026", "19-06" — data numérica dd/mm(/aaaa)
// Ignora números fora do intervalo válido de um dia do mês, para nunca
// confundir com o sufixo de coluna duplicada do Forms (ex. o " 5" final de
// "...de junho 5" não é precedido de "dia"/"de <mês>" nem seguido de
// "/mês", por isso nunca é apanhado por nenhum destes padrões).
function extractDayNumber(raw) {
  const s = String(raw || "");
  let m = s.match(/dia\s*(\d{1,2})\b/i);
  if (!m) m = s.match(/\b(\d{1,2})\s*de\s*[a-zçãáéíóúâêôõ]+/i);
  if (!m) m = s.match(/\b(\d{1,2})\s*[\/\-]\s*\d{1,2}(?:\s*[\/\-]\s*\d{2,4})?\b/);
  if (!m) return null;
  const n = Number(m[1]);
  return n >= 1 && n <= 31 ? n : null;
}

// Regista no mapa DAY_NUMBER_TO_WEEKDAY a correspondência dia-do-mês -> dia
// da semana, sempre que um texto trouxer os dois em conjunto (é o caso dos
// cabeçalhos do Forms). Chamada para TODOS os cabeçalhos antes de os
// processar (ver extractAvailabilityFromRow), para que o mapa esteja
// atualizado independentemente da ordem das colunas na folha.
function registerDayNumberMapping(text) {
  const s = String(text || "");
  let weekday = null;
  for (const { re, day } of DAY_WORD_PATTERNS) {
    if (re.test(s)) { weekday = day; break; }
  }
  if (!weekday) return;
  const num = extractDayNumber(s);
  if (num !== null) DAY_NUMBER_TO_WEEKDAY[num] = weekday;
}

// Converte uma célula/token já isolado (um único dia) em intervalo(s)
// {day, startMin, endMin}, usando SEMPRE parseTimeToMinutes() para os
// números — nunca comparação de texto. Um único horário sem intervalo
// (`fim === null`) é tratado como um ponto que cobre exatamente 1 slot
// (startMin..startMin+SLOT_DURATION_MIN).
function parseAvailabilityRanges(raw, dayHint) {
  const s = String(raw || "").trim();
  if (!s) return [];
  const firstDigit = s.search(/\d/);
  const dayText = firstDigit > 0 ? s.slice(0, firstDigit) : s;
  const day = dayHint || normalizeDayToken(dayText);
  if (!day) return [];
  // allowBareHour:true — aqui `s` é sempre um VALOR de disponibilidade
  // (célula, ou item já separado por | ; , dentro dela), nunca um
  // cabeçalho de coluna, por isso é seguro aceitar horas soltas sem
  // separador ("9", "17", "9-17", "09-10:30"), que é o que faltava para
  // vários formatos do Excel Mestre e dos Forms (ver TIME_TOKEN_RE_LOOSE).
  const { inicio, fim } = parseTimeToMinutes(s, { allowBareHour: true });
  if (inicio === null) return [];
  return [{ day, startMin: inicio, endMin: fim !== null ? fim : inicio + SLOT_DURATION_MIN }];
}

// Junta intervalos {day, startMin, endMin} CONTÍGUOS ou sobrepostos do
// mesmo dia num único intervalo maior — é isto que permite reconhecer que
// vários blocos de 30 min SEPARADOS, selecionados pelo candidato no Forms
// (ex. os tokens "11h30-12h" e "12h-12h30", cada um o seu próprio item na
// célula), cobrem juntos 1 hora contínua, mesmo vindo de tokens
// independentes que — sozinhos — só dariam para a duração da Fase 2 (Soft
// Skills, 30 min). Sem esta fusão, um candidato que selecionasse
// exatamente os blocos certos para Hard Skills/Dinâmicas nunca seria
// reconhecido como disponível para essas fases (era este o motivo exato
// das linhas 41/42 falharem).
function mergeRanges(ranges) {
  const byDay = {};
  ranges.forEach(({ day, startMin, endMin }) => {
    (byDay[day] || (byDay[day] = [])).push({ startMin, endMin });
  });
  const merged = [];
  Object.keys(byDay).forEach((day) => {
    const list = byDay[day].slice().sort((a, b) => a.startMin - b.startMin);
    let current = null;
    list.forEach((r) => {
      if (!current) { current = { ...r }; return; }
      if (r.startMin <= current.endMin) {
        // adjacente ou sobreposto — funde no intervalo corrente
        current.endMin = Math.max(current.endMin, r.endMin);
      } else {
        merged.push({ day, ...current });
        current = { ...r };
      }
    });
    if (current) merged.push({ day, ...current });
  });
  return merged;
}

// Converte intervalo(s) {day, startMin, endMin} (já fundidos — ver
// mergeRanges) nos slots oficiais da grelha (SLOTS) que ficam TOTALMENTE
// cobertos, exigindo `durationMin` minutos a partir do início oficial do
// slot — ou seja, alguém que respondeu "disponível das 09:00 às 12:00"
// fica corretamente marcado como disponível para os slots oficiais
// "Seg 09:00" E "Seg 10:30" quando durationMin=30 (ambos cabem em
// 09:00-12:00), mesmo nunca tendo escrito literalmente "09:00" nem "10:30"
// como pontos exatos. `durationMin` por omissão é SLOT_DURATION_MIN (30),
// mas o upload de Forms por fase passa a duração exigida por essa fase
// (ver PHASE_DURATION_MIN) — um candidato só conta como disponível para a
// Fase 4 (Hard Skills, 60 min) se tiver 60 min seguidos a partir do início
// do slot oficial, não apenas 30.
function expandRangesToSlots(ranges, durationMin = SLOT_DURATION_MIN) {
  const result = new Set();
  ranges.forEach(({ day, startMin, endMin }) => {
    SLOTS.forEach((slot) => {
      const info = SLOT_INFO[slot];
      const requiredEnd = info.startMin + durationMin;
      if (info.day === day && startMin <= info.startMin && requiredEnd <= endMin) result.add(slot);
    });
  });
  return Array.from(result);
}

// Divide o conteúdo de UMA célula em itens (| ; , ou quebra de linha),
// interpreta cada item como um intervalo {day, startMin, endMin} — SEM
// ainda os converter em slots — funde os intervalos contíguos/sobrepostos
// do mesmo dia (mergeRanges) e só DEPOIS os converte nos slots oficiais que
// ficam totalmente cobertos, usando `durationMin` (ver expandRangesToSlots).
// Usado tanto para colunas-dia isoladas como para o campo de texto livre
// "Disponibilidade". `dayHint`, quando passado (caso das colunas-dia), fixa
// o dia para todos os itens da célula; caso contrário cada item tem de
// indicar o seu próprio dia (campo de texto livre com vários dias na mesma
// célula).
//
// Devolve { slots, parsedAnything }: `parsedAnything` diz se PELO MENOS UM
// item da célula foi reconhecido como um dia+hora válido — mesmo que,
// depois de fundido, não chegue à duração exigida pela fase. Isto permite
// ao chamador (extractAvailabilityFromRow) distinguir "o candidato marcou
// blocos reais, só não chegam para esta fase mais longa" (não é erro) de
// "não percebi nada deste texto" (erro de formato genuíno) — ver requisito
// 3: "Nenhum dos horários" e afins nunca chegam aqui (já são filtrados
// antes por isNoAvailabilityResponse), por isso um resultado vazio aqui
// nunca é, por si só, motivo para tratar a linha como inválida.
function parseAvailabilityCell(cellText, dayHint, durationMin = SLOT_DURATION_MIN) {
  const ranges = [];
  let parsedAnything = false;
  String(cellText || "").split(/[|;,\n]/).forEach((token) => {
    const found = parseAvailabilityRanges(token, dayHint);
    if (found.length) {
      parsedAnything = true;
      ranges.push(...found);
    }
  });
  return { slots: expandRangesToSlots(mergeRanges(ranges), durationMin), parsedAnything };
}

// Extrai disponibilidade de uma linha, aceitando 3 formatos comuns no Excel
// Mestre / exports de Forms:
// (a) colunas-grelha em que o próprio cabeçalho é um slot exato (ex.:
//     "Seg 09:00", "Segunda-feira, 09:00 - 09:30") marcado com um valor
//     positivo (x/sim/verdadeiro/1/☑/disponível/check/a própria hora — ver
//     isAvailabilityPositiveMark)
// (b) colunas-grelha em que o cabeçalho contém o nome do DIA — incluindo
//     cabeçalhos duplicados com sufixo de secção do Forms, ex.
//     "Quarta-feira, dia 10 de junho", "...  2", "...  3" — e a célula tem
//     um ou mais horários/intervalos em texto livre (ex. "09:00-12:00" ou
//     "09:00, 10:30"). TODAS as colunas cujo cabeçalho contenha o mesmo dia
//     são lidas e unificadas no mesmo array — cada candidato só preenche a
//     secção do SEU departamento, pelo que as restantes ficam vazias e são
//     ignoradas, sem perder a que estiver preenchida.
// (c)/(d) uma coluna de texto livre ("Disponibilidade", ou qualquer outra
//     cujo cabeçalho não indique dia/hora) com um ou mais dias e
//     intervalos, separados por | ; , ou quebras de linha, ex. "Seg 09:00
//     | Ter 10:30", "Qui 14h-15h30".
// `rowObj` é um objeto simples { "Nome da Coluna": valor, ... } — tanto faz
// vir da leitura do Google Sheets (row.obj) como do xlsx.utils.sheet_to_json
// (linha já vem nesse formato), por isso esta função serve para os dois
// caminhos de importação (sincronização automática e upload manual de Forms).
//
// Devolve { slots, hasUnrecognizedContent }, em vez de só o array de slots:
// `hasUnrecognizedContent` distingue duas situações que ANTES eram tratadas
// da mesma forma (disponibilidade vazia == "linha sem horário reconhecido",
// e portanto um aviso/erro) mas que na prática são bem diferentes:
//   - o candidato respondeu explicitamente "Não tenho disponibilidade",
//     "N/A", "Nenhum(a)" ou deixou a célula em branco -> resposta VÁLIDA,
//     disponibilidade fica [] nesse campo, `hasUnrecognizedContent` NÃO é
//     marcado (nada de errado a reportar);
//   - o candidato escreveu qualquer outra coisa numa coluna reconhecida de
//     disponibilidade (dia isolado ou campo de texto livre) e essa coisa
//     não bateu com nenhum formato suportado -> aí sim é um problema real
//     de formato, e `hasUnrecognizedContent` fica true, para o chamador
//     poder avisar exatamente quais as linhas afetadas.
function extractAvailabilityFromRow(header, rowObj, durationMin = SLOT_DURATION_MIN) {
  // Pré-passo: regista no mapa DAY_NUMBER_TO_WEEKDAY qualquer par
  // "dia da semana + data" encontrado nos cabeçalhos ANTES de extrair
  // disponibilidade — assim, mesmo que esta chamada seja sobre a aba do
  // Excel Mestre (que só tem datas) e o Forms dos candidatos só tenha sido
  // lido depois (ou antes, não importa a ordem dentro desta função), as
  // colunas "10/06"/"Dia 10" já conseguem ser resolvidas para "Qua" assim
  // que pelo menos um cabeçalho em QUALQUER aba já processada nesta sessão
  // trouxer os dois juntos.
  header.forEach(registerDayNumberMapping);

  const slots = new Set();
  let hasUnrecognizedContent = false;
  // Intervalos {day,startMin,endMin} vindos de colunas "dia+intervalo no
  // cabeçalho, célula é só Disponível/Indisponível" (ver caso (a2) abaixo)
  // — acumulados à parte para serem fundidos numa só passagem no fim,
  // exatamente como os blocos de 30 min do Forms.
  const dayRangeRanges = [];
  header.forEach((h) => {
    // (a) cabeçalho é um slot exato ("Seg 09:00") marcado x/sim/true/1/☑/
    //     disponível/check/a própria hora
    const exactSlot = normalizeSlotString(h);
    if (exactSlot) {
      if (isAvailabilityPositiveMark(rowObj[h])) slots.add(exactSlot);
      // célula vazia ou negativa aqui é uma resposta normal de checkbox
      // (candidato não marcou ESTE slot específico) — não é conteúdo por
      // reconhecer, por isso nunca marca hasUnrecognizedContent.
      return;
    }
    // (a2) cabeçalho combina DIA + INTERVALO de horas (ex. "Qua 9-9:30",
    // resultado de mergeTwoRowHeader combinando "Dia 10 - Quarta" + "9-
    // 9:30"), e a célula é um marcador booleano "Disponível"/"Indisponível"
    // — não texto livre com a hora (é essa a diferença para o caso (b)
    // abaixo). Só entra aqui se o cabeçalho tiver um dia E um INTERVALO
    // genuíno (2 horas, início < fim) — um único número solto (ex. o "10"
    // de "dia 10 de junho") nunca ativa este caso, continua a cair no (b).
    const day = normalizeDayOnlyHeader(h);
    const range = day ? parseTimeToMinutes(h, { allowBareHour: true }) : null;
    if (day && range && range.inicio !== null && range.fim !== null && range.fim > range.inicio) {
      if (isAvailabilityPositiveMark(rowObj[h])) {
        dayRangeRanges.push({ day, startMin: range.inicio, endMin: range.fim });
      }
      // "Indisponível"/vazio/etc. aqui é só "não disponível NESTE bloco de
      // meia-hora" — normal, nunca é erro de formato.
      return;
    }
    // (b) cabeçalho contém o nome do dia (com ou sem sufixo de secção); a
    // célula tem o(s) horário(s)/intervalo(s) em texto livre — incluindo a
    // grelha de blocos de 30 min do Forms, já fundida e comparada com a
    // duração exigida por `durationMin` (ver parseAvailabilityCell).
    const dayOnly = day;
    if (dayOnly) {
      const cell = cleanCellText(rowObj[h]);
      if (isNoAvailabilityResponse(cell)) return; // em branco, "N/A"/"Nenhum"/"Nenhum dos horários" — aceite, sem disponibilidade nesse dia
      const { slots: found, parsedAnything } = parseAvailabilityCell(cell, dayOnly, durationMin);
      found.forEach((s) => slots.add(s));
      // Só é erro de formato se NADA na célula foi entendido como dia+hora.
      // Se foi entendido mas os blocos (fundidos) não chegam à duração
      // exigida por esta fase (ex.: só marcou 1 bloco de 30 min mas a fase
      // precisa de 60/90), isso é uma disponibilidade legítima de ZERO
      // slots PARA ESTA FASE — não um erro de parsing.
      if (!parsedAnything) hasUnrecognizedContent = true;
      return;
    }
    // (d) FALLBACK — o cabeçalho não indica nem dia nem hora (pergunta de
    // texto livre do Forms cujo título não segue nenhum padrão reconhecido,
    // ex. "Quais os teus horários disponíveis?"): em vez de descartar a
    // coluna, tenta interpretar o CONTEÚDO da célula como uma lista de
    // slots "Dia Hora" separados por | ; , ou quebra de linha (ex.: "Seg
    // 09:00 | Ter 10:30", "Qui 14h-15h30") — sem depender do nome da
    // coluna. Cada item só produz slot se começar por um dia reconhecível,
    // pelo que colunas verdadeiramente não relacionadas (nome, email, "como
    // conheceu a YME?", etc.) não geram falsos positivos — por isso, ao
    // contrário do caso (b), uma coluna deste tipo sem nenhum slot NÃO é
    // marcada como hasUnrecognizedContent (não sabemos se era suposto ser
    // uma coluna de disponibilidade); só o campo livre "disponibilidade"
    // explícito, tratado a seguir, é que conta para esse aviso.
    const cell = cleanCellText(rowObj[h]);
    if (cell && !isNoAvailabilityResponse(cell)) {
      parseAvailabilityCell(cell, undefined, durationMin).slots.forEach((s) => slots.add(s));
    }
  });
  // Funde e expande os intervalos "dia+intervalo no cabeçalho" recolhidos
  // no caso (a2) — mesma lógica de agregação usada para os blocos de 30
  // min do Forms (mergeRanges + expandRangesToSlots), mas com a duração
  // PADRÃO (SLOT_DURATION_MIN): esta é disponibilidade de Diretor/RH, não
  // ligada a nenhuma fase específica, por isso não usa `durationMin`.
  expandRangesToSlots(mergeRanges(dayRangeRanges), SLOT_DURATION_MIN).forEach((s) => slots.add(s));
  const free = cleanCellText(get(rowObj, "disponibilidade", "horarios", "horários", "slots"));
  if (!isNoAvailabilityResponse(free)) {
    const { slots: found, parsedAnything } = parseAvailabilityCell(free, undefined, durationMin);
    found.forEach((s) => slots.add(s));
    if (!parsedAnything) hasUnrecognizedContent = true;
  }
  return {
    slots: Array.from(slots).sort((a, b) => SLOTS.indexOf(a) - SLOTS.indexOf(b)),
    hasUnrecognizedContent,
  };
}

// Aplica os dados brutos das abas do Excel Mestre ao estado de members/candidates
// da aplicação, seguindo o mapeamento estrito de abas e colunas da YME.
function applySyncedSheetsToState(raw, prevMembers, prevCandidates) {
  const warnings = [];
  /* ---- A1. Base Dados Departamentos -> membros (Diretor/Supervisor/RH) ---- */
  let members = prevMembers.map((m) => ({ ...m }));
  const upsertMember = (name, role, dept) => {
    name = String(name || "").trim();
    if (!name) return;
    const idx = findMemberIndex(members, name);
    if (idx >= 0) {
      const depts = new Set(members[idx].departments || []);
      if (dept) depts.add(dept);
      members[idx] = { ...members[idx], role: role || members[idx].role, departments: Array.from(depts) };
    } else {
      members.push({ id: uid("sync"), name, role: role || "RH", title: role || "Membro", departments: dept ? [dept] : [], availability: [] });
    }
  };
  let deptRowsSeen = 0;
  let deptRowsWithAnyMember = 0;
  (raw.departamentos?.rows || []).forEach((row) => {
    deptRowsSeen++;
    const dept = matchDept(get(row.obj, "departamento", "department"));
    const diretor = get(row.obj, "diretor", "diretor(a)");
    const supervisor = get(row.obj, "supervisor");
    // O Excel Mestre da YME guarda os membros de RH em DUAS colunas
    // distintas ("Membro RH 1" / "Membro RH 2"), em vez de uma única
    // coluna com lista separada por vírgulas — por isso lemos as duas
    // colunas de forma independente. Mantemos a coluna única "RH" como
    // alternativa de compatibilidade (formato antigo/manual).
    const rh1 = String(get(row.obj, "membro rh 1", "membro de rh 1", "rh 1", "rh1") || "").trim();
    const rh2 = String(get(row.obj, "membro rh 2", "membro de rh 2", "rh 2", "rh2") || "").trim();
    const rhCols = [rh1, rh2].filter(Boolean);
    const rhLegacy = String(get(row.obj, "rh", "membro rh", "membros rh", "membro de rh") || "")
      .split(/[,;|]/).map((s) => s.trim()).filter(Boolean);
    const rhNames = rhCols.length ? rhCols : rhLegacy;
    if (diretor) upsertMember(diretor, "Diretor", dept);
    if (supervisor) upsertMember(supervisor, "Supervisor", dept);
    rhNames.forEach((n) => upsertMember(n, "RH", dept));
    if (diretor || supervisor || rhNames.length) deptRowsWithAnyMember++;
  });
  if (raw.departamentos && deptRowsSeen > 0 && deptRowsWithAnyMember === 0) {
    warnings.push(`A aba "Base Dados Departamentos" tem ${deptRowsSeen} linha(s) de dados (cabeçalho detetado na linha ${(raw.departamentos.headerIdx ?? 0) + 1}) mas nenhuma tem Diretor, Supervisor ou Membro RH reconhecidos. Confirma os textos exatos dos cabeçalhos "Diretor", "Supervisor", "Membro RH 1" e "Membro RH 2".`);
  }

  /* ---- A4/A5/A6. Disponibilidades de RH/Diretores/Supervisores -> membros ---- */
  [raw.dispEntrevistasRH, raw.dispDinamicas, raw.dispEntrevistaFinal].forEach((tab) => {
    (tab?.rows || []).forEach((row) => {
      const name = String(get(row.obj, "nome", "name") || "").trim();
      if (!name) return;
      const { slots } = extractAvailabilityFromRow(tab.header, row.obj);
      if (!slots.length) return;
      const idx = findMemberIndex(members, name);
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
    const idx = matchCandidateIndex(candidates, patch.name, patch.email);
    if (idx >= 0) candidates[idx] = { ...candidates[idx], ...patch };
    else candidates.push({
      id: uid("sync"), phase0Status: "Pendente", phase1Status: "—", phase2Status: "—",
      email: "", telefone: "",
      formsSubmitted: { fase1: false, fase2: false, fase3: false },
      availabilityStatus: { fase1: "nao_enviada", fase2: "nao_enviada", fase3: "nao_enviada" },
      availability: { fase1: [], fase2: [], fase3: [] },
      ...patch,
    });
  };
  // Converte o texto livre da coluna "Estado"/"Fase Atual" (quando existe
  // na própria aba "Base Dados Candidatos") num phase0Status inicial.
  // É só um valor de partida: a Coluna Q da aba "Avaliação CV e Questões
  // Abertas" (processada a seguir, A3) continua a ter a última palavra
  // sempre que estiver preenchida.
  function estadoToPhase0(raw) {
    const n = normKey(raw);
    if (!n) return null;
    if (["aprovado", "aprovada", "avanca", "apto", "selecionado", "admitido"].some((x) => n.includes(x))) return "Aprovado";
    if (["rejeitado", "reprovado", "eliminado", "excluido", "nao avanca", "chumbado"].some((x) => n.includes(x))) return "Rejeitado";
    if (["pendente", "em avaliacao", "por avaliar", "em analise", "a aguardar"].some((x) => n.includes(x))) return "Pendente";
    return null;
  }

  // Filtragem silenciosa: linhas sem Nome preenchido são consideradas
  // "espaço em branco reservado para futuros candidatos" (comum no Excel
  // Mestre da YME) e são simplesmente ignoradas — sem qualquer aviso na
  // UI. A app cresce automaticamente com o Excel: hoje 124 candidatos,
  // amanhã 200+, sem tocar em código nem poluir o ecrã com alertas.
  const candidatosValidos = (raw.candidatos?.rows || []).filter((row) => {
    const name = String(get(
      row.obj, "nome completo", "nome", "name", "nome do candidato", "candidato", "full name"
    ) || "").trim();
    return name.length > 0;
  });

  candidatosValidos.forEach((row) => {
    const name = String(get(
      row.obj, "nome completo", "nome", "name", "nome do candidato", "candidato", "full name"
    ) || "").trim();
    const dept1 = matchDept(get(row.obj, "primeira opcao", "primeira opção", "1a opcao", "1ª opção", "departamento", "departamento/cargo", "cargo", "cargo pretendido"));
    const dept2 = matchDept(get(row.obj, "segunda opcao", "segunda opção", "2a opcao", "2ª opção"));
    const estadoRaw = get(row.obj, "estado", "fase atual", "fase", "estado atual", "situacao", "situação");
    const phase0FromEstado = estadoToPhase0(estadoRaw);
    // Coluna J: "Veio da Talent Pool?" (checkbox TRUE/FALSE/VERDADEIRO/☑).
    // Regra de negócio (prioridade 1 — Fast-Track): se TRUE, o candidato
    // salta a Fase 1 (Avaliação de CV) e a Fase 2 (Entrevista Soft Skills)
    // por completo e entra diretamente elegível para a Fase 3 (Dinâmicas).
    const veioTalentPool = isPositiveMark(get(row.obj, "veio da talent pool", "veio da talent pool?", "talent pool", "veio de talent pool", "veio da talentpool"));
    // Fallback: linha sem coluna de departamento reconhecida (célula vazia
    // ou valor que não corresponde a nenhum dos 6 departamentos) nunca
    // rebenta a sincronização — fica apenas marcada para confirmação manual.
    upsertCandidate({
      name,
      email: String(get(row.obj, "email", "contacto", "email/contacto") || "").trim(),
      telefone: String(get(row.obj, "telefone", "telemóvel", "telemovel", "contacto", "contacto telefónico", "contacto telefonico", "phone", "nº telemóvel", "n.º telemóvel") || "").trim(),
      department: dept1 || DEPARTMENTS[0],
      departmentPorConfirmar: !dept1,
      segundaOpcaoDepartamento: dept2 || null,
      curso: String(get(row.obj, "curso") || "").trim(),
      anoLetivo: String(get(row.obj, "ano letivo", "ano") || "").trim(),
      erasmus: String(get(row.obj, "erasmus") || "").trim(),
      veioTalentPool,
      // Fast-track: Fases 1 e 2 ficam automaticamente "Aprovado"/isentas
      // (CV e Soft Skills), avançando direto para a Fase 3 (Dinâmicas).
      ...(veioTalentPool ? { phase0Status: "Aprovado", phase1Status: "Aprovado" } : {}),
      ...(!veioTalentPool && phase0FromEstado ? { phase0Status: phase0FromEstado } : {}),
    });
  });

  // Diagnóstico: só avisa quando a aba foi lida mas ABSOLUTAMENTE nenhuma
  // linha tinha Nome preenchido — sinal de que o cabeçalho não foi
  // encontrado (problema real de configuração), nunca pelo simples facto
  // de existirem linhas em branco reservadas para candidatos futuros.
  if (raw.candidatos && raw.candidatos.rows.length > 0 && candidatosValidos.length === 0) {
    warnings.push(`A aba "Base Dados Candidatos" tem ${raw.candidatos.rows.length} linha(s) de dados (cabeçalho detetado na linha ${(raw.candidatos.headerIdx ?? 0) + 1}) mas nenhuma tem uma coluna de Nome reconhecida (procurei por "Nome", "Nome Completo", "Nome do Candidato"...). Confirma o texto exato do cabeçalho dessa coluna na folha.`);
  }

  /* ---- A3. Avaliação CV e Questões Abertas (Fase 1) -> Colunas Q/R (avança p/ Fase 2 = Entrevista Soft Skills) ----
     Estrutura confirmada desta aba (não tem coluna de Email): cabeçalho
     na linha 13, dados a partir da linha 14 (já garantido por
     SYNC_FIXED_HEADER_IDX.avaliacaoCV), Coluna A = Departamento,
     Coluna B = Nome do Candidato, Coluna C = Ano do Curso (NUNCA usada
     como nome), Coluna Q = "Passou?", Coluna R = "Não Passou".
     Lê-se SEMPRE por posição de coluna (nunca por texto de cabeçalho) e
     o matching com "Base Dados Candidatos" é feito só por Nome
     sanitizado (normKey: trim + lowercase + sem acentos + espaços
     colapsados) — não há Email nesta aba. Isto percorre TODAS as linhas
     sem qualquer interrupção por departamento: a lógica é idêntica para
     Human Resources, Quality Management, Legal & Finance, Brand
     Strategy, etc. */
  const CV_INVALID_NAME_MARKERS = ["licenciatura", "mestrado", "doutoramento", "pós-graduação", "pos-graduacao"];
  const unmatchedCV = [];
  (raw.avaliacaoCV?.rows || []).forEach((row) => {
    const rawName = row.raw[colLetterToIndex(SYNC_CV_NAME_COLUMN)];
    if (isErrorOrEmptyValue(rawName)) return; // linha vazia ou erro de fórmula (#N/A, etc.) -> ignora silenciosamente
    const name = String(rawName).trim();
    // Salvaguarda extra: se por algum motivo a célula da Coluna B trouxer
    // um grau académico (ex. valor da Coluna C "Ano do Curso" desalinhado
    // por uma linha em branco/mesclada na folha), a linha é ignorada em
    // vez de criar um candidato falso chamado "Licenciatura", etc.
    if (CV_INVALID_NAME_MARKERS.some((marker) => normKey(name).includes(marker))) return;
    const rawDept = row.raw[colLetterToIndex(SYNC_CV_DEPT_COLUMN)];
    const dept = isErrorOrEmptyValue(rawDept) ? null : matchDept(rawDept);
    const idxExisting = matchCandidateIndex(candidates, name, "");
    // Prioridade 1 (Fast-Track Talent Pool) já decidiu o estado deste
    // candidato em A2 — as Colunas Q/R (fluxo regular) nunca a sobrepõem.
    if (idxExisting >= 0 && candidates[idxExisting].veioTalentPool) return;
    // Regra de negócio (prioridade 2 — Fluxo Regular): Coluna Q ("Passou?")
    // marcada -> Aprovado na Fase 1 / avança Fase 2; Coluna R ("Não
    // Passou") marcada -> Rejeitado; nenhuma marcada -> mantém-se
    // Pendente (não se sobrepõe nada). Válido para qualquer departamento.
    const passou = isPositiveMark(row.raw[colLetterToIndex(SYNC_CV_PASS_COLUMN)]);
    const naoPassou = isPositiveMark(row.raw[colLetterToIndex(SYNC_CV_FAIL_COLUMN)]);
    if (!passou && !naoPassou) return; // nenhuma coluna marcada -> permanece Pendente
    const patch = { name, phase0Status: passou ? "Aprovado" : "Rejeitado" };
    if (idxExisting >= 0) {
      candidates[idxExisting] = {
        ...candidates[idxExisting],
        ...patch,
        // Se conseguimos ler um departamento válido nesta linha, usa-o
        // (mais fiável do que o que já estava, que pode vir de um
        // registo antigo/errado em "Base Dados Candidatos").
        ...(dept ? { department: dept } : {}),
      };
    } else {
      // Não foi possível casar por nome com nenhum candidato de "Base
      // Dados Candidatos" — em vez de o atirar silenciosamente para um
      // departamento arbitrário (o bug original: tudo o que não casasse
      // ficava sempre em Human Resources), usa-se o departamento real
      // lido da Coluna A desta própria aba sempre que possível; só cai
      // no departamento por omissão, marcado "por confirmar", quando a
      // Coluna A também não permite identificar o departamento.
      upsertCandidate({
        ...patch,
        department: dept || DEPARTMENTS[0],
        departmentPorConfirmar: !dept,
      });
      unmatchedCV.push(name);
    }
  });
  if (unmatchedCV.length) {
    warnings.push(`${unmatchedCV.length} candidato(s) da aba "Avaliação CV e Questões Abertas" não corresponderam a nenhum candidato de "Base Dados Candidatos" por nome: ${unmatchedCV.join("; ")} — foram criados marcados como "por confirmar". Confirma se o nome está escrito de forma idêntica nas duas abas.`);
  }

  /* ---- B. Abas por Departamento -> progresso por fase + lógica de rejeição ---- */
  const unmatchedDept = [];
  DEPARTMENTS.forEach((dept) => {
    const tab = raw.deptTabs?.[dept];
    if (!tab) return;
    tab.rows.forEach((row) => {
      const rawName = get(row.obj, "nome", "nome completo", "name");
      const rawEmail = get(row.obj, "email");
      // Fórmulas dinamizadas do Sheets devolvem #N/A (ou #VALUE!/#REF!) para
      // candidatos que ainda não chegaram a esta fase — não são candidatos
      // desconhecidos, são simplesmente slots por preencher. Ignora-se a
      // linha silenciosamente, sem gerar aviso nem tentar fazer match.
      if (isErrorOrEmptyValue(rawName) && isErrorOrEmptyValue(rawEmail)) return;
      const name = isErrorOrEmptyValue(rawName) ? "" : String(rawName).trim();
      if (!name) return;
      const email = isErrorOrEmptyValue(rawEmail) ? "" : String(rawEmail).trim();
      const idx = matchCandidateIndex(candidates, name, email);
      if (idx < 0) {
        unmatchedDept.push(`${email ? `${name} (${email})` : name} — aba "${dept}"`);
        return; // candidato tem de constar em "Base Dados Candidatos"
      }
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
      if (!cand.veioTalentPool) {
        if (softSkills === "positive") phase1Status = "Aprovado";
        else if (softSkills === "negative") phase1Status = "Rejeitado";
      }

      let phase2Status = cand.phase2Status;
      if (dinamicas === "positive") phase2Status = "Aprovado";
      else if (dinamicas === "negative") phase2Status = "Rejeitado";

      let finalResult = cand.finalResult || null;
      if (final === "positive") finalResult = "Selecionado — Entrou na YME";
      else if (final === "negative") finalResult = talentPoolMark ? "Talent Pool" : "Rejeitado";

      // Lógica de rejeição automática: só atua quando há efetivamente alguma
      // marcação nas colunas de transição (não sobrepõe candidatos ainda por
      // decidir), e nunca reverte o estado isento de um Fast-Track Talent Pool.
      let phase0Status = cand.phase0Status;
      if (!cand.veioTalentPool && anyColumnFilled && !anyPositive && cand.phase0Status === "Aprovado") {
        phase0Status = "Rejeitado";
      }

      candidates[idx] = {
        ...cand, department: dept, phase0Status, phase1Status, phase2Status,
        finalResult, talentPool: talentPoolMark || cand.talentPool || false,
      };
    });
  });
  if (unmatchedDept.length > 0) {
    warnings.push(`${unmatchedDept.length} linha(s) nas abas de departamento não corresponderam a nenhum candidato de "Base Dados Candidatos" por email nem por nome: ${unmatchedDept.join("; ")} — o respetivo progresso de fase não foi atualizado.`);
  }

  return { members, candidates, warnings };
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
  const readTab = async (key, sheetName, headerHints = [], fixedHeaderIdx = null) => {
    try {
      const data = await fetchSheetTabApi(accessToken, sheetId, sheetName, headerHints, fixedHeaderIdx);
      return [key, data];
    } catch (err) {
      if (err.message === "token_expired") tokenExpired = true;
      return [key, { error: translateSheetApiError(err.message, sheetName).message }];
    }
  };

  let generalResults, deptResults;
  try {
    [generalResults, deptResults] = await Promise.all([
      Promise.all(Object.entries(SYNC_SHEET_NAMES).map(([key, sheetName]) => readTab(key, sheetName, SYNC_HEADER_HINTS[key], SYNC_FIXED_HEADER_IDX[key]))),
      Promise.all(DEPARTMENTS.map((dept) => readTab(dept, `'${dept}'`, SYNC_DEPT_HEADER_HINTS))),
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

  const { members, candidates, warnings } = applySyncedSheetsToState(raw, prevMembers, prevCandidates);
  return { members, candidates, errors: [...errors, ...warnings] };
}

/* ============================================================================
   ESTADO INICIAL — SEM MOCK DATA
   A aplicação arranca sempre vazia (0 membros, 0 candidatos). O Excel
   Mestre (Google Sheets) é a ÚNICA fonte da verdade: os dados reais só
   entram em estado assim que o utilizador autentica com a Google e liga
   a folha ("Guardar Ligação"), o que despoleta uma sincronização
   automática (ver runSync/useEffect mais abaixo), ou quando clica em
   "Sincronizar Agora". As antigas funções buildMockMembers()/
   buildMockCandidates() foram removidas — nunca mais geram candidatos
   fictícios no arranque da app.
============================================================================ */

/* ============================================================================
   SCHEDULING ALGORITHMS  (lógica inalterada)
============================================================================ */

// Compara departamentos com tolerância — não exige igualdade exata mesmo
// depois de normalizado. Cobre 3 situações reais do Excel Mestre:
//  1) "Geral" ou SEM departamento definido -> conta como disponível para
//     TODOS os departamentos (Diretor/RH partilhado, ou aba ainda por
//     preencher corretamente);
//  2) abreviatura por iniciais (ex. "QM" -> "Quality Management", "RH" não
//     se aplica aqui pois é o próprio role, mas "BS" -> "Brand Strategy");
//  3) nome parcial/prefixo (ex. "Quality" dentro de "Quality Management").
// `.toLowerCase().trim()` já está embutido em deptKey() (via normKey), que
// também remove acentos e colapsa espaços a mais.
const CATCH_ALL_DEPT_KEYS = new Set([
  "geral", "todos", "all", "any", "todososdepartamentos",
  "qualquerdepartamento", "semdepartamento", "n a", "na",
]);
function deptAbbrev(name) {
  return normKey(name).replace(/[^a-z ]/g, "").split(" ").filter(Boolean).map((w) => w[0]).join("");
}
function deptMatches(memberDeptRaw, targetDept) {
  const memberKey = deptKey(memberDeptRaw);
  if (!memberKey || CATCH_ALL_DEPT_KEYS.has(memberKey.replace(/\s+/g, ""))) return true;
  const targetKey = deptKey(targetDept);
  if (memberKey === targetKey) return true;
  // Resolve o texto do membro para um dos 6 departamentos oficiais (mesma
  // tolerância já usada no resto da app — matchDept) e compara os canónicos.
  const resolved = matchDept(memberDeptRaw);
  if (resolved && deptKey(resolved) === targetKey) return true;
  // Abreviatura por iniciais (ex. "QM" -> "Quality Management").
  if (memberKey.replace(/\s+/g, "") === deptAbbrev(targetDept)) return true;
  // Nome parcial/prefixo (ex. "Quality" dentro de "Quality Management") —
  // só a partir de 3 letras, para não confundir siglas curtas com o
  // prefixo de outro departamento.
  if (memberKey.length >= 3 && (targetKey.startsWith(memberKey) || memberKey.startsWith(targetKey))) return true;
  return false;
}
// Compara departamentos ignorando maiúsculas/minúsculas, acentos e espaços
// a mais — em vez de igualdade estrita de string (===/Array.includes).
// Dois departamentos "iguais" mas escritos por fontes diferentes (Excel
// Mestre vs Google Sheets vs edição manual) podem ter, por exemplo,
// unicode de acentuação diferente ou um espaço a mais sem serem
// visivelmente distintos — e Array.includes falha nesse caso sem aviso
// nenhum. deptKey()/memberHasDept() usam a mesma normalização de normKey()
// já usada no resto do ficheiro (matchDept, findMemberIndex, etc.).
function deptKey(d) {
  return normKey(d).replace(/[^a-z0-9]+/g, " ").trim();
}
function memberHasDept(member, dept) {
  const depts = member.departments || [];
  // Sem NENHUM departamento definido -> tratado como "Geral", disponível
  // para todos (ver CORREÇÃO DA TOLERÂNCIA DE DEPARTAMENTOS acima).
  if (!depts.length) return true;
  return depts.some((d) => deptMatches(d, dept));
}
// Lista de RH atribuídos a um departamento — SEM olhar a horários. Esta é
// a "verdade" da coluna RH: sempre que não vier vazia, um candidato desse
// departamento NUNCA deve mostrar "Sem alocação", seja qual for o
// resultado do cruzamento de horários feito depois.
function rhForDepartment(members, dept) {
  return members.filter((m) => m.role === "RH" && memberHasDept(m, dept));
}

// Converte um slot (canónico "Seg 09:00", ou qualquer outra representação
// reconhecida por parseTimeToMinutes) em { day, startMin } — SEMPRE por
// aritmética de minutos, nunca por igualdade de string. CORREÇÃO CRÍTICA:
// comparar strings de hora diretamente ("9h-9h30" !== "09:00 - 09:00")
// é frágil a qualquer diferença de formatação entre fontes (Forms vs Excel
// Mestre vs edição manual) — mesmo pequenas variações invisíveis fazem o
// cruzamento falhar por completo. slotToMinutes()/availabilityMinuteSet()/
// hasSlot() garantem que a comparação de disponibilidade entre Candidato,
// Diretor e RH é sempre feita em minutos desde a meia-noite.
function slotToMinutes(slot) {
  const s = String(slot || "");
  const spaceIdx = s.indexOf(" ");
  if (spaceIdx < 0) return null;
  const day = s.slice(0, spaceIdx);
  const { inicio } = parseTimeToMinutes(s.slice(spaceIdx + 1));
  return inicio === null ? null : { day, startMin: inicio };
}
// Constrói um Set de chaves "Dia|minutos" a partir de uma lista de slots de
// disponibilidade — usado para consultas O(1) por aritmética, em vez de
// `.includes(slot)` (igualdade de string).
function availabilityMinuteSet(slots) {
  const set = new Set();
  (slots || []).forEach((s) => {
    const info = slotToMinutes(s);
    if (info) set.add(`${info.day}|${info.startMin}`);
  });
  return set;
}
function hasSlot(minuteSet, slot) {
  const info = slotToMinutes(slot);
  return info ? minuteSet.has(`${info.day}|${info.startMin}`) : false;
}

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

  // Cache de disponibilidade em minutos por membro (id -> Set "Dia|min"),
  // calculada uma única vez por membro em vez de re-parsear a cada
  // candidato — ver slotToMinutes/availabilityMinuteSet acima.
  const minuteSetCache = new Map();
  function minutesOf(member) {
    if (!member) return null;
    if (!minuteSetCache.has(member.id)) minuteSetCache.set(member.id, availabilityMinuteSet(member.availability));
    return minuteSetCache.get(member.id);
  }

  // PASSO 2 — DISTRIBUIÇÃO EQUITATIVA (ROUND-ROBIN) POR DEPARTAMENTO.
  // Um índice rotativo por departamento (não uma contagem de carga): avança
  // 1 posição por CADA candidato desse departamento, alternando estritamente
  // RH A -> RH B -> RH A -> ... independentemente de haver ou não slot
  // comum — é o que garante Candidato 1 -> RH A, Candidato 2 -> RH B,
  // Candidato 3 -> RH A, tal como pedido.
  const roundRobinIndex = {};
  function nextRoundRobinRH(dept, rhList) {
    if (!rhList.length) return null;
    const key = deptKey(dept);
    const i = roundRobinIndex[key] || 0;
    roundRobinIndex[key] = (i + 1) % rhList.length;
    return rhList[i % rhList.length];
  }

  const bookings = [...kept];
  pool.forEach((c) => {
    if (kept.some((b) => b.candidateId === c.id)) return;
    // Diretor por departamento — com FALLBACK GLOBAL: se não houver
    // Diretor explicitamente associado a este departamento no Excel
    // Mestre (mesmo já com a tolerância de deptMatches), usa qualquer
    // Diretor disponível na organização como validador, em vez de deixar
    // `diretor` por preencher e bloquear TODOS os candidatos desse
    // departamento. `diretorIsFallback` fica registado para o diagnóstico
    // poder distinguir os dois casos.
    const diretorStrict = members.find((m) => m.role === "Diretor" && memberHasDept(m, c.department));
    const diretor = diretorStrict || members.find((m) => m.role === "Diretor") || null;
    const diretorIsFallback = !diretorStrict && !!diretor;
    const supervisor = staffKeys.includes("supervisorId") ? members.find((m) => m.role === "Supervisor" && memberHasDept(m, c.department)) : null;

    // PASSO 1 — ATRIBUIÇÃO DIRETA POR DEPARTAMENTO, independente do
    // horário: primeiro decide-se QUEM é o RH responsável (rhForDepartment
    // + round-robin), só depois é que se tenta cruzar horários. A coluna
    // RH usa sempre `assignedRH`, mesmo que o cruzamento abaixo não
    // encontre nenhum slot comum. Mesmo FALLBACK GLOBAL que o Diretor: se
    // não houver nenhum RH explicitamente associado a este departamento,
    // usa a equipa de RH inteira da organização como candidatos válidos,
    // em vez de bloquear o agendamento por completo.
    let rhList = rhForDepartment(members, c.department);
    const rhIsFallback = !rhList.length;
    if (rhIsFallback) rhList = members.filter((m) => m.role === "RH");
    const assignedRH = nextRoundRobinRH(c.department, rhList);

    let found = null;
    for (const slot of c.availability[availField]) {
      if (!diretor || !hasSlot(minutesOf(diretor), slot) || busy[diretor.id]?.has(slot)) continue;
      if (staffKeys.includes("supervisorId")) {
        if (!supervisor || !hasSlot(minutesOf(supervisor), slot) || busy[supervisor.id]?.has(slot)) continue;
      }
      // PASSO 3 — FALLBACK DE DISPONIBILIDADE: tenta primeiro o RH
      // atribuído pelo round-robin (`assignedRH`); só se ele não tiver
      // este slot livre (ou estiver ocupado com outro candidato) é que se
      // testam os restantes membros de RH do mesmo departamento (ou de
      // toda a organização, se `rhIsFallback`), por ordem, antes de
      // desistir deste slot e passar ao seguinte.
      const rhCandidates = [assignedRH, ...rhList.filter((r) => r.id !== assignedRH?.id)].filter(Boolean);
      const rh = rhCandidates.find((r) => hasSlot(minutesOf(r), slot) && !busy[r.id]?.has(slot));
      if (rh) { found = { slot, diretor, rh, supervisor }; break; }
    }

    if (found) {
      const record = { id: uid("bk"), candidateId: c.id, slot: found.slot, diretorId: found.diretor.id, rhId: found.rh.id, status: "Agendado", manual: false };
      if (staffKeys.includes("supervisorId")) record.supervisorId = found.supervisor.id;
      staffKeys.forEach((k) => { const mid = record[k]; if (mid) { busy[mid] = busy[mid] || new Set(); busy[mid].add(found.slot); } });
      bookings.push(record);
    } else {
      // "Sem alocação" só deve aparecer quando NÃO EXISTE sequer 1 RH em
      // toda a organização (rhList vazia -> assignedRH null, o que só
      // acontece se `members` não tiver NENHUM role="RH"). Havendo
      // qualquer RH (do departamento ou do fallback global), a coluna
      // mostra sempre `assignedRH` — só o Horário fica por preencher e o
      // Estado mantém "Sem Horário Comum".
      const record = { id: uid("bk"), candidateId: c.id, slot: null, diretorId: diretor?.id || null, rhId: assignedRH?.id || null, status: "Sem Horário Comum", manual: false };
      if (staffKeys.includes("supervisorId")) record.supervisorId = supervisor?.id || null;
      // DIAGNÓSTICO (requisito 3): em vez de só "Sem Horário Comum" sem
      // mais nenhuma pista, guarda no próprio registo qual foi exatamente
      // o motivo — mostrado como tooltip no Estado (ver StatusBadge). A
      // ordem dos testes segue a cadeia de dependências reais do
      // cruzamento: primeiro se sequer existe Diretor/RH em TODA a
      // organização, depois se algum deles tem disponibilidade nenhuma
      // registada, e só por fim (o caso mais comum) se não há interseção
      // de horários entre quem já existe — incluindo aviso quando foi
      // preciso recorrer ao fallback global (departamento sem Diretor/RH
      // próprio no Excel Mestre).
      let reason;
      if (!diretor) {
        reason = `Nenhum(a) Diretor(a) registado no Excel Mestre (nenhum membro com role "Diretor").`;
      } else if (!minutesOf(diretor).size) {
        reason = diretorIsFallback
          ? `Nenhum(a) Diretor(a) associado ao departamento "${c.department}" — usou-se ${diretor.name} (Diretor de outro departamento) como fallback, mas também sem horários registados no Excel Mestre.`
          : `Diretor(a) ${diretor.name} sem horários registados no Excel Mestre.`;
      } else if (!rhList.length) {
        reason = `Nenhum Membro de RH registado no Excel Mestre (nenhum membro com role "RH").`;
      } else if (!rhList.some((r) => minutesOf(r).size)) {
        reason = rhIsFallback
          ? `Nenhum Membro de RH associado ao departamento "${c.department}" — usou-se a equipa de RH de outros departamentos como fallback, mas também sem horários registados no Excel Mestre.`
          : `Equipa de RH de "${c.department}" sem horários registados no Excel Mestre.`;
      } else {
        reason = `Sem interseção entre os horários de ${c.name} e a equipa (Diretor(a)/RH) ${diretorIsFallback || rhIsFallback ? "(via fallback global)" : `de "${c.department}"`}.`;
      }
      record.reason = reason;
      bookings.push(record);
    }
  });
  return bookings;
}

// Regenera o agendamento de UMA fase individual apenas para um
// DEPARTAMENTO (ou para todos, com ALL_DEPARTMENTS_OPTION) sem apagar os
// agendamentos já existentes dos restantes departamentos — CORREÇÃO: como
// generateInterviewPhase() só devolve registos para quem está no `pool`
// que recebe, gerar diretamente com um pool filtrado por departamento e
// usar esse resultado para SUBSTITUIR o estado inteiro apagaria os
// agendamentos de todos os outros departamentos (que não fariam parte
// desse pool mais pequeno). Esta função isola a regeneração ao
// departamento escolhido: filtra o pool E os `existingBookings` (para que
// os agendamentos MANUAIS desse departamento sejam preservados, tal como
// generateInterviewPhase já faz), gera só para esse subconjunto, e depois
// funde o resultado com os agendamentos dos OUTROS departamentos que já
// existiam, deixando-os intocados.
function regenerateForDepartment(pool, members, existingBookings, availField, staffKeys, department) {
  const targetPool = department === ALL_DEPARTMENTS_OPTION ? pool : pool.filter((c) => c.department === department);
  const targetIds = new Set(targetPool.map((c) => c.id));
  const regenerated = generateInterviewPhase(
    targetPool, members, existingBookings.filter((b) => targetIds.has(b.candidateId)), availField, staffKeys
  );
  const untouched = existingBookings.filter((b) => !targetIds.has(b.candidateId));
  return [...untouched, ...regenerated];
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
        const dir = members.find((m) => m.role === "Diretor" && memberHasDept(m, d));
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
          const am = depts.some((d) => memberHasDept(a, d)) ? 1 : 0;
          const bm = depts.some((d) => memberHasDept(b, d)) ? 1 : 0;
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

function Badge({ children, className = "", style, title }) {
  return (
    <span
      className={`inline-flex items-center px-2 py-0.5 rounded-md text-xs font-medium border ${className}`}
      style={style}
      title={title}
    >
      {children}
    </span>
  );
}
function DeptBadge({ dept }) {
  return <Badge className={deptBadgeClass(dept)}>{dept}</Badge>;
}
function StatusBadge({ status, title }) {
  return <Badge className={statusBadgeClass(status)} title={title}>{status}</Badge>;
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
    { id: "dashboard", label: "DASHBOARD · FASE 1 · AVALIAÇÃO CV", badge: counts.fase0 },
    { id: "import", label: "IMPORTAÇÃO DE DADOS" },
    { id: "fase1", label: "FASE 2 · SOFT SKILLS", badge: counts.fase1 },
    { id: "fase2", label: "FASE 3 · DINÂMICAS", badge: counts.fase2 },
    { id: "fase3", label: "FASE 4 · HARD SKILLS", badge: counts.fase3 },
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
      <div className="max-w-[1400px] mx-auto px-6 pb-1 -mt-1 text-right">
        <span className="text-[10px] font-mono" style={{ color: hexToRgba(COLORS.mint, 0.3) }} title="Confirma que este valor corresponde ao build mais recente enviado">{APP_BUILD}</span>
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
      {!error && status.warning && <p className="text-[11px] mb-2" style={{ color: "#c0227a" }}>{status.warning}</p>}
      {!error && status.info && <p className="text-[11px] mb-2" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{status.info}</p>}

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
        // Lê a folha em modo GRELHA BRUTA (array de arrays, não já um
        // objeto por linha) e passa pelo mesmo parseApiValues() usado na
        // sincronização com o Google Sheets — CORREÇÃO: é isto que permite
        // detetar e combinar um cabeçalho em DUAS linhas (dia numa linha,
        // ex. "Dia 10 - Quarta", hora na seguinte, ex. "9-9:30"), que o
        // antigo `XLSX.utils.sheet_to_json(ws, {defval:""})` (modo objeto,
        // só considera 1 linha de cabeçalho) nunca conseguia reconhecer —
        // cada coluna ficava só com a hora solta, sem dia nenhum associado.
        const rawGrid = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
        const { rows } = parseApiValues(rawGrid, ["nome", "name"]);
        rows.forEach(({ obj: row }) => {
          const name = cleanCellText(get(row, "nome", "name"));
          if (!name) return;
          let role = cleanCellText(get(row, "role", "cargo"));
          if (!role) {
            if (/diretor/i.test(sheetName)) role = "Diretor";
            else if (/rh|recursos/i.test(sheetName)) role = "RH";
            else if (/supervisor|c-level|clevel/i.test(sheetName)) role = "Supervisor";
          }
          const deptsRaw = cleanCellText(get(row, "departamentos", "departamento")).split(/[,;|]/).map((s) => matchDept(s)).filter(Boolean);
          // Mesma correção: usa o parser tolerante (grelha OU texto livre,
          // com normalização de dia/hora) em vez de exigir texto livre no
          // formato exato "Seg 09:00|Ter 14:00".
          const { slots: availRaw } = extractAvailabilityFromRow(Object.keys(row), row);
          const idx = findMemberIndex(next, name);
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
      // DIAGNÓSTICO (requisito 3 do pedido): lista no consola (F12) todos
      // os avaliadores (Diretor/Supervisor/RH) e a disponibilidade que
      // ficou associada a cada um, logo depois de processar este ficheiro
      // — para confirmar visualmente se Gustavo Dias, Mariana Lopes, Joana
      // Pereira, etc. ficaram com horários extraídos do Excel Mestre.
      const evaluatorsList = next
        .filter((m) => ["Diretor", "Supervisor", "RH"].includes(m.role))
        .map((m) => ({ nome: m.name, role: m.role, departamentos: m.departments, horarios: m.availability }));
      console.log("Avaliadores Mapeados:", evaluatorsList);
      return next;
    });
    setImportStatus((prev) => ({ ...prev, excel: { loaded: true, filename: file.name, count } }));
  };

  const handleFormsPhase = (phaseKey) => async (file) => {
    const wb = await readWorkbook(file);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
    let count = 0;
    // Linhas onde HAVIA texto de disponibilidade mas nenhum slot foi
    // reconhecido — diferente de candidatos que não submeteram nada ou
    // responderam explicitamente "não tenho disponibilidade" (esses são
    // aceites normalmente, com disponibilidade [], sem entrar aqui — ver
    // isNoAvailabilityResponse/hasUnrecognizedContent). `rowNumber` conta a
    // partir de 2 porque a linha 1 da folha é o cabeçalho.
    const unrecognizedRows = [];
    // Requisito 3: candidatos aceites normalmente mas com 0 slots nesta
    // fase — célula em branco, "Nenhum dos horários" e afins, OU blocos de
    // 30 min reconhecidos que não chegam à duração exigida pela fase. Não é
    // erro, só informação para o RH ("Candidato X importado (sem
    // disponibilidade assinalada)").
    const noAvailabilityNames = [];
    setCandidates((prev) => {
      const next = [...prev];
      rows.forEach((row, i) => {
        const rowNumber = i + 2;
        // Requisito 1: cada célula lida (nome/email/departamento) passa por
        // cleanCellText — remove \r, \n e caracteres invisíveis antes do
        // trim, para que uma resposta com uma quebra de linha a mais nunca
        // pareça diferente de uma sem ela.
        const name = cleanCellText(get(row, "nome", "name"));
        if (!name) return;
        const department = matchDept(get(row, "departamento", "department"));
        const email = cleanCellText(get(row, "email"));
        // CORREÇÃO: antes só se lia uma única coluna de texto livre
        // ("Disponibilidade"/"Horários"/"Slots") com formato fixo. Exports
        // reais do Google/Microsoft Forms costumam vir em formato de
        // GRELHA (uma coluna por slot, ex. "Segunda-feira, 09:00 - 09:30"),
        // exatamente como já era suportado para a disponibilidade de
        // Diretores/Supervisores/RH — extractAvailabilityFromRow() agora é
        // partilhada por ambos os caminhos, tolera variações de dia/hora
        // (ver normalizeSlotString/normalizeDayOnlyHeader) e distingue
        // "sem disponibilidade" de "formato não reconhecido" (ver
        // hasUnrecognizedContent, requisito 1 e 3 do pedido).
        const { slots: availability, hasUnrecognizedContent } = extractAvailabilityFromRow(
          Object.keys(row), row, PHASE_DURATION_MIN[phaseKey] ?? SLOT_DURATION_MIN
        );
        if (!availability.length && hasUnrecognizedContent) {
          unrecognizedRows.push({ rowNumber, name });
        } else if (!availability.length) {
          noAvailabilityNames.push(name);
        }
        const idx = matchCandidateIndex(next, name, email);
        count++;
        if (idx >= 0) {
          next[idx] = {
            ...next[idx],
            department: department || next[idx].department,
            email: email || next[idx].email,
            formsSubmitted: { ...next[idx].formsSubmitted, [phaseKey]: true },
            availabilityStatus: { ...next[idx].availabilityStatus, [phaseKey]: "recebida" },
            // Requisito 1: um candidato que respondeu "N/A"/"Nenhum"/deixou
            // em branco é aceite normalmente com [] — essa resposta é
            // válida e substitui qualquer valor anterior. Já um formato
            // realmente não reconhecido (hasUnrecognizedContent) preserva o
            // valor anterior em vez de o apagar, para uma reimportação com
            // um ficheiro com problemas não destruir disponibilidade boa já
            // lida antes.
            availability: {
              ...next[idx].availability,
              [phaseKey]: (!availability.length && hasUnrecognizedContent) ? next[idx].availability[phaseKey] : availability,
            },
          };
        } else {
          next.push({
            id: uid("cand"), name, department: department || DEPARTMENTS[0], email, cvLink: "",
            phase0Status: "Aprovado",
            phase1Status: phaseKey === "fase1" ? "Pendente" : "—",
            phase2Status: phaseKey === "fase2" ? "Pendente" : "—",
            formsSubmitted: { fase1: phaseKey === "fase1", fase2: phaseKey === "fase2", fase3: phaseKey === "fase3" },
            availabilityStatus: {
              fase1: phaseKey === "fase1" ? "recebida" : "nao_enviada",
              fase2: phaseKey === "fase2" ? "recebida" : "nao_enviada",
              fase3: phaseKey === "fase3" ? "recebida" : "nao_enviada",
            },
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
    setImportStatus((prev) => ({
      ...prev,
      [phaseKey]: {
        loaded: true, filename: file.name, count,
        // Requisito 3: em vez de uma contagem genérica, identifica
        // exatamente quais linhas/candidatos ficaram sem nenhum horário
        // reconhecido — e só entram aqui os casos de formato realmente não
        // suportado, nunca os candidatos que legitimamente não submeteram
        // disponibilidade (em branco, "Não tenho disponibilidade"/"N/A"/
        // "Nenhum dos horários", ou blocos de 30 min que não chegam à
        // duração desta fase — esses geram o aviso informativo `info`).
        warning: buildUnrecognizedRowsWarning(unrecognizedRows),
        info: buildNoAvailabilityInfo(noAvailabilityNames),
      },
    }));
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
        <UploadCard icon={UsersRound} title="Forms — Fase 2" status={importStatus.fase1}
          description="Candidatos e disponibilidades submetidas para as entrevistas de Soft Skills."
          hint='Colunas: Nome, Departamento, Email, Disponibilidade (slots separados por "|").'
          onFile={handleFormsPhase("fase1")} />
        <UploadCard icon={LayoutGrid} title="Forms — Fase 3" status={importStatus.fase2}
          description="Candidatos e disponibilidades submetidas para as Dinâmicas de Grupo."
          hint='Colunas: Nome, Departamento, Email, Disponibilidade (slots separados por "|").'
          onFile={handleFormsPhase("fase2")} />
        <UploadCard icon={CalendarClock} title="Forms — Fase 4" status={importStatus.fase3}
          description="Candidatos e disponibilidades submetidas para as entrevistas de Hard Skills."
          hint='Colunas: Nome, Departamento, Email, Disponibilidade (slots separados por "|").'
          onFile={handleFormsPhase("fase3")} />
      </div>
      <button onClick={() => downloadTemplate("forms")} className="yme-link text-xs -mt-6 mb-8">Descarregar modelo CSV de Forms ↓</button>

      <p className="text-xs font-semibold uppercase tracking-widest mb-2" style={{ color: "#94a3b8" }}>Diagnóstico ao vivo: Diretor/Supervisor/RH por departamento</p>
      <p className="text-xs mb-2" style={{ color: hexToRgba(COLORS.white, 0.55) }}>
        Ao contrário da tabela "referência" que existia aqui antes (estática, só mostrava o `ORG` fixo no código), esta lê diretamente o estado `members` que a app está a usar neste preciso momento — os mesmos dados que alimentam as colunas RH/Diretor/Horário das páginas de agendamento. Se uma célula aparecer a vermelho aqui, é um problema de DADOS (sincronização ainda não correu, nome escrito de forma diferente) — corrige na folha/import. Se aparecer tudo verde aqui mas as tabelas de agendamento continuarem a mostrar "Sem alocação", o browser está a mostrar uma versão desatualizada da app — recarrega/limpa cache.
      </p>
      <div className="rounded-xl border overflow-hidden mb-8" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
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
            {DEPARTMENTS.map((dept) => {
              const diretor = members.find((m) => m.role === "Diretor" && memberHasDept(m, dept));
              const supervisor = members.find((m) => m.role === "Supervisor" && memberHasDept(m, dept));
              const rh = rhForDepartment(members, dept);
              const missing = <span className="text-xs font-medium" style={{ color: "#c0227a" }}>— não mapeado —</span>;
              return (
                <tr key={dept} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                  <td className="px-4 py-3"><DeptBadge dept={dept} /></td>
                  <td className="px-4 py-3" style={{ color: COLORS.navy }}>{diretor?.name || missing}</td>
                  <td className="px-4 py-3" style={{ color: COLORS.navy }}>{supervisor?.name || missing}</td>
                  <td className="px-4 py-3" style={{ color: COLORS.navy }}>{rh.length ? rh.map((r) => r.name).join(", ") : missing}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <p className="text-xs font-semibold uppercase tracking-widest mb-2" style={{ color: "#94a3b8" }}>Estrutura Organizacional (referência fixa no código)</p>
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
  const [eliminationFilter, setEliminationFilter] = useState("Todos");
  const [origemFilter, setOrigemFilter] = useState("Todos");
  const [search, setSearch] = useState("");

  const filtered = candidates.filter((c) =>
    (deptFilter === "Todos" || c.department === deptFilter) &&
    (statusFilter === "Todos" || c.phase0Status === statusFilter) &&
    (eliminationFilter === "Todos" || getEliminationPhase(c) === eliminationFilter) &&
    (origemFilter === "Todos" || (origemFilter === "Talent Pool" ? !!c.veioTalentPool : !c.veioTalentPool)) &&
    c.name.toLowerCase().includes(search.toLowerCase())
  );
  const { page, setPage, totalPages, pageItems } = usePagination(filtered, 12, `${deptFilter}-${statusFilter}-${eliminationFilter}-${origemFilter}-${search}`);

  const filtersActive = deptFilter !== "Todos" || statusFilter !== "Todos" || eliminationFilter !== "Todos" || origemFilter !== "Todos" || search !== "";
  // Os contadores de topo refletem sempre o conjunto atualmente filtrado
  // (candidates completo quando não há filtros ativos, subconjunto
  // filtrado assim que qualquer filtro é aplicado).
  const counts = {
    total: filtered.length,
    aprovados: filtered.filter((c) => c.phase0Status === "Aprovado").length,
    pendentes: filtered.filter((c) => c.phase0Status === "Pendente").length,
    rejeitados: filtered.filter((c) => c.phase0Status === "Rejeitado").length,
  };

  const setStatus = (id, status) => setCandidates((prev) => prev.map((c) => (c.id === id ? { ...c, phase0Status: status } : c)));

  return (
    <div className="p-8">
      <div className="flex items-start justify-between mb-6">
        <div>
          <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>Dashboard & Fase 1 — Questionário e CV</h1>
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
        <StatCard label={filtersActive ? "Candidatos (filtrado)" : "Total de candidatos"} value={counts.total} icon={UsersRound} tone="neutral" />
        <StatCard label="Aprovados p/ Fase 2" value={counts.aprovados} icon={CheckCircle2} tone="brand" />
        <StatCard label="Pendentes" value={counts.pendentes} icon={Clock3} tone="alert" />
        <StatCard label="Rejeitados" value={counts.rejeitados} icon={XCircle} tone="critical" />
      </div>

      <div className="flex flex-wrap items-center gap-3 mb-2">
        <div className="relative">
          <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: hexToRgba(COLORS.navy, 0.45) }} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Procurar candidato..."
            className="yme-input pl-8 pr-3 py-2 text-sm rounded-lg w-56" />
        </div>
        <select value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)} className="yme-input text-sm rounded-lg px-2.5 py-2">
          <option value="Todos">Todos os departamentos</option>
          {DEPARTMENTS.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="yme-input text-sm rounded-lg px-2.5 py-2">
          <option value="Todos">Todos os estados</option>
          <option value="Aprovado">Passaram à Fase Atual</option>
          <option value="Pendente">Pendentes de Avaliação</option>
          <option value="Rejeitado">Eliminados / Rejeitados</option>
        </select>
        <select value={eliminationFilter} onChange={(e) => setEliminationFilter(e.target.value)} className="yme-input text-sm rounded-lg px-2.5 py-2">
          <option value="Todos">Eliminado em — qualquer fase</option>
          <option value="fase1">Eliminado na Fase 1 (Triagem CV)</option>
          <option value="fase2">Eliminado na Fase 2 (Soft Skills)</option>
          <option value="fase3">Eliminado na Fase 3 (Dinâmicas)</option>
          <option value="fase4">Eliminado na Fase 4 (Hard Skills)</option>
        </select>
        <select value={origemFilter} onChange={(e) => setOrigemFilter(e.target.value)} className="yme-input text-sm rounded-lg px-2.5 py-2">
          <option value="Todos">Todas as origens</option>
          <option value="Talent Pool">Talent Pool</option>
          <option value="Regular">Processo Regular</option>
        </select>
        {filtersActive && (
          <button
            onClick={() => { setDeptFilter("Todos"); setStatusFilter("Todos"); setEliminationFilter("Todos"); setOrigemFilter("Todos"); setSearch(""); }}
            className="text-xs font-medium underline" style={{ color: hexToRgba(COLORS.white, 0.7) }}
          >
            Limpar filtros
          </button>
        )}
        <span className="text-xs ml-auto" style={{ color: "#94a3b8" }}>{filtered.length} resultado(s)</span>
      </div>

      {candidates.length === 0 && (
        <div className="flex items-center gap-3 rounded-xl border px-4 py-3 mb-4" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.15) }}>
          <FileClock size={16} style={{ color: COLORS.navy }} className="shrink-0" />
          <p className="text-sm" style={{ color: COLORS.navy }}>
            Ainda sem candidatos. Liga o Excel Mestre da YME (Google Sheets) no Hub de Importação para sincronizar os dados reais — não existem dados fictícios pré-carregados.
          </p>
          <button onClick={goToImport} className="yme-btn-primary text-xs font-medium rounded-lg px-3 py-1.5 whitespace-nowrap ml-auto">
            Ir para Hub de Importação
          </button>
        </div>
      )}

      <div className="rounded-xl border overflow-hidden" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs uppercase tracking-wide" style={{ backgroundColor: COLORS.navy, color: COLORS.white }}>
              <th className="px-4 py-3 font-medium">Candidato</th>
              <th className="px-4 py-3 font-medium">Departamento</th>
              <th className="px-4 py-3 font-medium">Email</th>
              <th className="px-4 py-3 font-medium">Telefone</th>
              <th className="px-4 py-3 font-medium">CV / Questionário</th>
              <th className="px-4 py-3 font-medium">Estado</th>
              <th className="px-4 py-3 font-medium">Avança Fase 2</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((c) => (
              <tr key={c.id} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                <td className="px-4 py-3 font-medium" style={{ color: COLORS.navy }}>
                  {c.name}
                  {c.veioTalentPool && (
                    <span className="ml-1.5 text-[10px] font-normal align-middle px-1.5 py-0.5 rounded" title="Entrou via Talent Pool — salta CV e Entrevista Soft Skills" style={{ color: COLORS.navy, backgroundColor: hexToRgba(COLORS.pink, 0.18) }}>Talent Pool</span>
                  )}
                  {c.departmentPorConfirmar && (
                    <span className="ml-1.5 text-[10px] font-normal align-middle" title="Departamento não reconhecido na folha — por confirmar" style={{ color: "#c0227a" }}>(dept. por confirmar)</span>
                  )}
                </td>
                <td className="px-4 py-3"><DeptBadge dept={c.department} /></td>
                <td className="px-4 py-3" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{c.email || "—"}</td>
                <td className="px-4 py-3" style={{ color: hexToRgba(COLORS.navy, 0.6) }}>{c.telefone || "—"}</td>
                <td className="px-4 py-3">
                  {c.cvLink ? (
                    <a href={c.cvLink} target="_blank" rel="noreferrer" className="yme-link text-xs">Ver questionário ↗</a>
                  ) : (
                    <span className="text-xs" style={{ color: hexToRgba(COLORS.navy, 0.4) }}>—</span>
                  )}
                </td>
                <td className="px-4 py-3">
                  <select value={c.phase0Status} onChange={(e) => setStatus(c.id, e.target.value)}
                    className={`text-xs rounded-md px-2 py-1 border font-medium ${statusBadgeClass(c.phase0Status)}`}>
                    {["Aprovado", "Pendente", "Rejeitado"].map((s) => <option key={s} style={{ color: COLORS.navy }}>{s}</option>)}
                  </select>
                </td>
                <td className="px-4 py-3">
                  {c.phase0Status === "Aprovado" ? (
                    <span className="inline-flex items-center gap-1 text-xs font-medium" style={{ color: COLORS.navy }}>
                      <CheckCircle2 size={14} style={{ color: COLORS.pink }} />
                      {c.veioTalentPool ? "Isento (Talent Pool)" : "Sim"}
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1 text-xs" style={{ color: hexToRgba(COLORS.navy, 0.45) }}><XCircle size={14} /> Não</span>
                  )}
                </td>
              </tr>
            ))}
            {pageItems.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-10 text-center text-sm" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Sem candidatos para os filtros selecionados.</td></tr>
            )}
          </tbody>
        </table>
        <PaginationBar page={page} setPage={setPage} totalPages={totalPages} />
      </div>
    </div>
  );
}

/* ============================================================================
   PAGE: FASE 2 & FASE 4 (entrevistas 1:1)
============================================================================ */

function InterviewPhasePage({
  title, subtitle, phaseKey, availField, formsField, prevStatusField,
  candidates, setCandidates, members, bookings, setBookings, onGenerate, columns, showCalendar,
  excludeTalentPool = false,
}) {
  const [editing, setEditing] = useState(null);
  const [view, setView] = useState("list");
  // Requisito 1 (agendamento isolado por departamento): seletor local a
  // esta fase — cada separador (Soft Skills, Hard Skills) tem o seu
  // próprio filtro, começando sempre em "Todos os Departamentos".
  const [departmentFilter, setDepartmentFilter] = useState(ALL_DEPARTMENTS_OPTION);

  const byId = (id) => members.find((m) => m.id === id);
  const candById = (id) => candidates.find((c) => c.id === id);

  // Departamentos DETETADOS nos ficheiros carregados (não a lista fixa de
  // 6 departamentos "possíveis") — só aparecem no dropdown os que
  // realmente têm candidatos nesta fase.
  const availableDepartments = useMemo(
    () => Array.from(new Set(candidates.map((c) => c.department).filter(Boolean))).sort((a, b) => a.localeCompare(b)),
    [candidates]
  );

  // REMOÇÃO DA TRAVA DE VISIBILIDADE: já não se filtra por
  // `c[prevStatusField] === "Aprovado"` — todos os candidatos do
  // departamento (exceto Fast-Track da Talent Pool na Fase 2, que segue
  // rota própria) aparecem sempre na tabela, independentemente de ainda
  // não terem submetido o Forms desta fase ou da validação da fase
  // anterior. CORREÇÃO DA MATEMÁTICA DO FUNIL: continua sem exigir
  // "Aprovado", mas passa a excluir quem já ficou "Rejeitado" na fase
  // anterior (`prevStatusField`) — esses nunca chegam a esta fase, e por
  // isso não devem contar como "elegível" nem aparecer como "À espera do
  // Forms". Com 74 candidatos, 11 Rejeitados na Fase 1 e 9 da Talent Pool,
  // isto dá exatamente os 54 elegíveis da Fase 2.
  // Requisito 2: quando um departamento está selecionado, "elegível"
  // passa a significar "elegível E desse departamento" — filtra tanto os
  // candidatos como (indiretamente, via generateInterviewPhase já ser
  // department-aware) os avaliadores cruzados.
  const eligible = candidates.filter((c) =>
    (!excludeTalentPool || !c.veioTalentPool) &&
    c[prevStatusField] !== "Rejeitado" &&
    (departmentFilter === ALL_DEPARTMENTS_OPTION || c.department === departmentFilter)
  );
  // "À espera do Forms" só faz sentido dentro do próprio conjunto de
  // elegíveis — antes contava TODOS os candidatos (incluindo Rejeitados/
  // Talent Pool), o que inflacionava este número mesmo com 100% das
  // respostas dos elegíveis já recebidas.
  const missingForms = eligible.filter((c) => !c.formsSubmitted[formsField]).length;
  const availabilityConfirmed = eligible.filter((c) => (c.availabilityStatus?.[formsField] || "nao_enviada") === "recebida").length;

  // Requisito 3: a grelha/resumo final refletem só o departamento
  // selecionado — `visibleBookings` filtra os agendamentos pelo
  // departamento do respetivo candidato. Com "Todos os Departamentos"
  // continua a mostrar tudo (cada linha já tem o seu próprio DeptBadge,
  // dando a separação clara pedida sem precisar de agrupar por secções).
  const visibleBookings = departmentFilter === ALL_DEPARTMENTS_OPTION
    ? bookings
    : bookings.filter((b) => candById(b.candidateId)?.department === departmentFilter);

  const scheduled = visibleBookings.filter((b) => b.status === "Agendado").length;
  const conflicts = visibleBookings.filter((b) => b.status !== "Agendado").length;

  // DIAGNÓSTICO AGREGADO (requisito 3): quando NENHUMA entrevista fica
  // agendada, mostra logo no topo um resumo dos motivos mais comuns em vez
  // de obrigar o RH a passar o rato linha a linha — cada `b.reason`
  // (calculado em generateInterviewPhase) já identifica exatamente quem
  // falhou no cruzamento (Diretor/RH sem departamento associado, sem
  // horários no Excel Mestre, ou sem interseção de facto). Calculado sobre
  // `visibleBookings`, para o resumo bater com o departamento selecionado.
  const failureReasonCounts = {};
  visibleBookings.forEach((b) => { if (b.reason) failureReasonCounts[b.reason] = (failureReasonCounts[b.reason] || 0) + 1; });
  const topFailureReasons = Object.entries(failureReasonCounts).sort((a, b) => b[1] - a[1]).slice(0, 3);

  const setAvailability = (candId, value) => {
    setCandidates((prev) => prev.map((c) => (c.id === candId ? { ...c, availabilityStatus: { ...c.availabilityStatus, [formsField]: value } } : c)));
  };

  const { page, setPage, totalPages, pageItems } = usePagination(visibleBookings, 12, visibleBookings.length);

  const exportCSV = () => {
    const header = ["Candidato", "Departamento", ...columns.map((c) => c.label), "Horário", "Estado"];
    const rows = visibleBookings.map((b) => {
      const cand = candById(b.candidateId);
      return [cand?.name, cand?.department, ...columns.map((c) => byId(b[c.key])?.name || "—"), b.slot || "—", b.status];
    });
    const deptSuffix = departmentFilter === ALL_DEPARTMENTS_OPTION ? "" : `-${slugify(departmentFilter)}`;
    downloadCSV(`${phaseKey}-agendamentos${deptSuffix}.csv`, [header, ...rows]);
  };

  return (
    <div className="p-8">
      <div className="flex items-start justify-between mb-6">
        <div>
          <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>{title}</h1>
          <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>{subtitle}</p>
        </div>
        <div className="flex items-center gap-2">
          {/* Requisito 1: dropdown de Departamento — lista automaticamente
              os departamentos detetados em `candidates`, com "Todos os
              Departamentos" como opção por defeito. */}
          <select
            value={departmentFilter}
            onChange={(e) => setDepartmentFilter(e.target.value)}
            className="yme-input text-sm rounded-lg px-3 py-2"
            aria-label="Filtrar por departamento"
          >
            <option value={ALL_DEPARTMENTS_OPTION}>{ALL_DEPARTMENTS_OPTION}</option>
            {availableDepartments.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
          {showCalendar && (
            <div className="flex rounded-lg overflow-hidden text-sm" style={{ border: `1px solid ${hexToRgba(COLORS.mint, 0.25)}` }}>
              <button onClick={() => setView("list")} className="px-3 py-2 flex items-center gap-1.5" style={view === "list" ? { backgroundColor: COLORS.pink, color: COLORS.navy } : { backgroundColor: "transparent", color: "#94a3b8" }}><ListChecks size={14} /> Lista</button>
              <button onClick={() => setView("calendar")} className="px-3 py-2 flex items-center gap-1.5" style={view === "calendar" ? { backgroundColor: COLORS.pink, color: COLORS.navy } : { backgroundColor: "transparent", color: "#94a3b8" }}><CalendarDays size={14} /> Calendário</button>
            </div>
          )}
          <button onClick={exportCSV} className="yme-btn-outline-dark flex items-center gap-1.5 text-sm rounded-lg px-3 py-2">
            <Download size={14} /> Exportar CSV
          </button>
          <button onClick={() => onGenerate(departmentFilter)} className="yme-btn-primary flex items-center gap-1.5 text-sm rounded-lg px-3 py-2 font-medium">
            <RefreshCw size={14} /> Gerar Agendamentos Automaticamente
          </button>
        </div>
      </div>

      {visibleBookings.length > 0 && scheduled === 0 && (
        <div className="flex items-start gap-2.5 rounded-xl border px-4 py-3 mb-6" style={{ backgroundColor: hexToRgba("#c0227a", 0.08), borderColor: hexToRgba("#c0227a", 0.3) }}>
          <AlertTriangle size={16} className="shrink-0 mt-0.5" style={{ color: "#c0227a" }} />
          <div className="text-xs leading-relaxed" style={{ color: COLORS.navy }}>
            <p className="font-semibold mb-1">Nenhum horário comum encontrado para nenhum candidato {departmentFilter === ALL_DEPARTMENTS_OPTION ? "desta fase" : `de "${departmentFilter}"`}.</p>
            {topFailureReasons.length > 0 ? (
              <ul className="space-y-0.5">
                {topFailureReasons.map(([reason, n]) => (
                  <li key={reason}>• {reason}{n > 1 ? ` (${n} candidato${n > 1 ? "s" : ""})` : ""}</li>
                ))}
              </ul>
            ) : (
              <p>Passa o rato sobre o Estado de cada linha para veres o motivo específico.</p>
            )}
          </div>
        </div>
      )}

      <div className="flex flex-wrap items-center gap-2 mb-4">
        <span className="inline-flex items-center gap-1.5 text-xs font-medium rounded-full px-3 py-1.5" style={{ backgroundColor: COLORS.mint, color: COLORS.navy }}>
          <UsersRound size={13} /> {title.split("—")[0].trim()}: {eligible.length} Candidato(s)
          {departmentFilter !== ALL_DEPARTMENTS_OPTION && ` · ${departmentFilter}`}
        </span>
        <span className="inline-flex items-center gap-1.5 text-xs font-medium rounded-full px-3 py-1.5" style={{ backgroundColor: availabilityConfirmed === eligible.length && eligible.length > 0 ? "#bbf7d0" : COLORS.mint, color: COLORS.navy }}>
          <FileClock size={13} /> {availabilityConfirmed}/{eligible.length} Disponibilidades Recebidas
        </span>
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <StatCard label="Candidatos elegíveis" value={eligible.length} icon={UsersRound} tone="neutral" />
        <StatCard label="Entrevistas agendadas" value={scheduled} icon={CheckCircle2} tone="brand" />
        <StatCard label="Sem horário comum" value={conflicts} icon={AlertTriangle} tone="critical" />
        <StatCard label={`À espera do Forms ${PHASE_LABEL[formsField]}`} value={missingForms} icon={FileClock} tone="alert" />
      </div>

      <p className="text-xs font-semibold uppercase tracking-widest mb-2" style={{ color: "#94a3b8" }}>Candidatos nesta fase e disponibilidade</p>
      <div className="rounded-xl border overflow-hidden mb-8" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs uppercase tracking-wide" style={{ backgroundColor: COLORS.navy, color: COLORS.white }}>
              <th className="px-4 py-3 font-medium">Candidato</th>
              <th className="px-4 py-3 font-medium">Departamento</th>
              <th className="px-4 py-3 font-medium">Disponibilidade</th>
            </tr>
          </thead>
          <tbody>
            {eligible.map((c) => (
              <tr key={c.id} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                <td className="px-4 py-3 font-medium" style={{ color: COLORS.navy }}>
                  {c.name}
                  {c.veioTalentPool && (
                    <span className="ml-1.5 text-[10px] font-normal align-middle px-1.5 py-0.5 rounded" style={{ color: COLORS.navy, backgroundColor: hexToRgba(COLORS.pink, 0.18) }}>Talent Pool</span>
                  )}
                </td>
                <td className="px-4 py-3"><DeptBadge dept={c.department} /></td>
                <td className="px-4 py-3">
                  <AvailabilitySelect value={c.availabilityStatus?.[formsField]} onChange={(v) => setAvailability(c.id, v)} />
                </td>
              </tr>
            ))}
            {eligible.length === 0 && (
              <tr><td colSpan={3} className="px-4 py-8 text-center text-sm" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Ainda não há candidatos neste departamento.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {visibleBookings.length === 0 && (
        <div className="rounded-xl border border-dashed p-10 text-center text-sm" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.2), color: hexToRgba(COLORS.navy, 0.5) }}>
          Ainda não há candidatos para agendar {departmentFilter === ALL_DEPARTMENTS_OPTION ? "nesta fase" : `em "${departmentFilter}"`}.
        </div>
      )}

      {visibleBookings.length > 0 && view === "list" && (
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
                    <td className="px-4 py-3"><StatusBadge status={b.status} title={b.reason} /></td>
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

      {visibleBookings.length > 0 && view === "calendar" && (
        <CalendarView bookings={visibleBookings} candById={candById} byId={byId} columns={columns} />
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
   PAGE: FASE 3 — DINÂMICAS DE GRUPO
============================================================================ */

function Phase2Page({ candidates, setCandidates, members, groups, setGroups, onGenerate }) {
  const candById = (id) => candidates.find((c) => c.id === id);
  const byId = (id) => members.find((m) => m.id === id);
  const missingForms = candidates.filter((c) => c.phase1Status === "Aprovado" && !c.formsSubmitted.fase2).length;
  // Lista de elegíveis para a Fase 3: candidatos com phase1Status
  // "Aprovado" — inclui tanto quem passou a Fase 2 (Entrevista Soft
  // Skills) como os candidatos Fast-Track vindos da Talent Pool (que
  // entram já com phase1Status "Aprovado", isentos das Fases 1 e 2).
  // Não depende de já terem submetido o Forms — isso é acompanhado pela
  // coluna "Disponibilidade" abaixo, editável manualmente.
  const eligible = candidates.filter((c) => c.phase1Status === "Aprovado");
  const talentPoolAtivos = candidates.filter((c) => c.veioTalentPool && c.phase1Status === "Aprovado").length;
  const availabilityConfirmed = eligible.filter((c) => (c.availabilityStatus?.fase2 || "nao_enviada") === "recebida").length;

  const setAvailability = (candId, value) => {
    setCandidates((prev) => prev.map((c) => (c.id === candId ? { ...c, availabilityStatus: { ...c.availabilityStatus, fase2: value } } : c)));
  };

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
    downloadCSV("fase3-dinamicas-grupo.csv", [header, ...rows]);
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
          <h1 className="text-xl font-bold tracking-wide uppercase" style={{ color: COLORS.white }}>Fase 3 — Dinâmicas de Grupo</h1>
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

      <div className="flex flex-wrap items-center gap-2 mb-4">
        <span className="inline-flex items-center gap-1.5 text-xs font-medium rounded-full px-3 py-1.5" style={{ backgroundColor: COLORS.mint, color: COLORS.navy }}>
          <UsersRound size={13} /> Fase 3: {eligible.length} Candidato(s)
        </span>
        <span className="inline-flex items-center gap-1.5 text-xs font-medium rounded-full px-3 py-1.5" style={{ backgroundColor: availabilityConfirmed === eligible.length && eligible.length > 0 ? "#bbf7d0" : COLORS.mint, color: COLORS.navy }}>
          <FileClock size={13} /> {availabilityConfirmed}/{eligible.length} Disponibilidades Recebidas
        </span>
      </div>

      <div className="grid grid-cols-5 gap-4 mb-6">
        <StatCard label="Grupos formados" value={groups.length} icon={LayoutGrid} tone="neutral" />
        <StatCard label="Candidatos alocados" value={groups.reduce((a, g) => a + g.candidateIds.length, 0)} icon={UsersRound} tone="brand" />
        <StatCard label="Vindos da Talent Pool" value={talentPoolAtivos} icon={CheckCircle2} tone="brand" />
        <StatCard label="Avisos de restrição" value={totalWarnings} icon={AlertTriangle} tone={totalWarnings ? "critical" : "alert"} />
        <StatCard label="À espera do Forms Fase 3" value={missingForms} icon={FileClock} tone="alert" />
      </div>

      <p className="text-xs font-semibold uppercase tracking-widest mb-2" style={{ color: "#94a3b8" }}>Candidatos nesta fase e disponibilidade</p>
      <div className="rounded-xl border overflow-hidden mb-8" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.1) }}>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs uppercase tracking-wide" style={{ backgroundColor: COLORS.navy, color: COLORS.white }}>
              <th className="px-4 py-3 font-medium">Candidato</th>
              <th className="px-4 py-3 font-medium">Departamento</th>
              <th className="px-4 py-3 font-medium">Disponibilidade</th>
            </tr>
          </thead>
          <tbody>
            {eligible.map((c) => (
              <tr key={c.id} className="yme-table-row" style={{ borderTop: `1px solid ${hexToRgba(COLORS.navy, 0.1)}` }}>
                <td className="px-4 py-3 font-medium" style={{ color: COLORS.navy }}>
                  {c.name}
                  {c.veioTalentPool && (
                    <span className="ml-1.5 text-[10px] font-normal align-middle px-1.5 py-0.5 rounded" style={{ color: COLORS.navy, backgroundColor: hexToRgba(COLORS.pink, 0.18) }}>Talent Pool</span>
                  )}
                </td>
                <td className="px-4 py-3"><DeptBadge dept={c.department} /></td>
                <td className="px-4 py-3">
                  <AvailabilitySelect value={c.availabilityStatus?.fase2} onChange={(v) => setAvailability(c.id, v)} />
                </td>
              </tr>
            ))}
            {eligible.length === 0 && (
              <tr><td colSpan={3} className="px-4 py-8 text-center text-sm" style={{ color: hexToRgba(COLORS.navy, 0.5) }}>Ainda não há candidatos aprovados na Fase 2 nem vindos da Talent Pool.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {groups.length === 0 && (
        <div className="rounded-xl border border-dashed p-10 text-center text-sm" style={{ backgroundColor: COLORS.mint, borderColor: hexToRgba(COLORS.navy, 0.2), color: hexToRgba(COLORS.navy, 0.5) }}>
          Ainda não há candidatos aprovados na Fase 2 que tenham submetido o Forms da Fase 3.
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
                      {c.veioTalentPool && (
                        <span className="text-[10px] font-medium px-1.5 py-0.5 rounded" style={{ backgroundColor: hexToRgba(COLORS.pink, 0.18), color: COLORS.navy }}>Talent Pool</span>
                      )}
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
  const [form, setForm] = useState({ name: "", department: DEPARTMENTS[0], email: "", telefone: "", cvLink: "" });
  const save = () => {
    if (!form.name.trim()) return;
    onSave({
      id: uid("cand"), ...form,
      phase0Status: "Pendente", phase1Status: "—", phase2Status: "—",
      formsSubmitted: { fase1: false, fase2: false, fase3: false },
      availabilityStatus: { fase1: "nao_enviada", fase2: "nao_enviada", fase3: "nao_enviada" },
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
        <Field label="Telefone"><input className={inputCls} value={form.telefone} onChange={(e) => setForm({ ...form, telefone: e.target.value })} /></Field>
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

  // members arranca com a base fixa da organização (Diretor/Supervisor/RH
  // por departamento — ver buildOrgBaselineMembers), para que a coluna RH
  // nunca apareça "Sem alocação" só por falta de sincronização. candidates
  // continua vazio: só ganha conteúdo real através da sincronização com o
  // Excel Mestre (Google Sheets) — ver runSync/useEffect mais abaixo — ou,
  // como alternativa manual/backup, via upload de ficheiro no Hub de
  // Importação. A sincronização/importação de membros faz sempre merge por
  // nome com esta base (nunca substitui), acrescentando disponibilidade.
  const [members, setMembers] = useState(() => buildOrgBaselineMembers());
  const [candidates, setCandidates] = useState(() => []);
  const [importStatus, setImportStatus] = useState({
    excel: { loaded: false, filename: "", count: 0 },
    fase1: { loaded: false, filename: "", count: 0 },
    fase2: { loaded: false, filename: "", count: 0 },
    fase3: { loaded: false, filename: "", count: 0 },
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
      // DIAGNÓSTICO (requisito 3 do pedido): mesma lista no consola (F12)
      // que no upload manual de Excel — para confirmar, também no caminho
      // de sincronização automática com o Google Sheets, se Gustavo Dias,
      // Mariana Lopes, Joana Pereira, etc. ficaram com horários extraídos.
      console.log("Avaliadores Mapeados:", nextMembers
        .filter((m) => ["Diretor", "Supervisor", "RH"].includes(m.role))
        .map((m) => ({ nome: m.name, role: m.role, departamentos: m.departments, horarios: m.availability })));
      setImportStatus((prev) => ({
        ...prev,
        excel: { loaded: true, filename: "Sincronização em tempo real (Google Sheets, ficheiro privado)", count: nextCandidates.length },
      }));
      const uniqueErrors = Array.from(new Set(errors));
      setSyncState({
        syncing: false,
        error: uniqueErrors.length ? `${uniqueErrors.length} aviso(s) de sincronização: ${uniqueErrors.slice(0, 2).join(" ")}` : null,
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

  // As pools abaixo definem quem CONTA/aparece em cada separador de fase.
  // Regra única e sem exceções por departamento: basta o estado da fase
  // anterior ser "Aprovado" (+ não vir isento via Talent Pool, na Fase 2).
  // IMPORTANTE: já não depende de `formsSubmitted` — isso era o bug: o
  // Forms de disponibilidade é importado à parte (upload manual "Forms —
  // Fase X") e o matching por nome/email desse ficheiro só estava a
  // resolver para candidatos de Human Resources, pelo que os restantes
  // departamentos (aprovados nas colunas Q/R da folha, e corretamente
  // contados no Dashboard) ficavam de fora da pool aqui usada para gerar
  // o separador FASE 2 · SOFT SKILLS — daí os "17 Aprovados" no Dashboard
  // vs. apenas "8 candidatos, só de HR" no separador. A ausência de Forms
  // preenchido não é motivo de exclusão: generateInterviewPhase já trata
  // esse caso de forma explícita, marcando a entrevista como "Sem Horário
  // Comum" em vez de omitir o candidato.
  // REMOÇÃO DA TRAVA DE VISIBILIDADE: phase1Pool/phase3Pool deixam de
  // filtrar por `phase0Status`/`phase2Status === "Aprovado"` — esse filtro
  // escondia da tabela (e por isso também do algoritmo de agendamento, que
  // só gera registo para quem está no pool) candidatos que ainda não
  // tinham validação da etapa anterior ou Forms "Pendente". Agora TODOS os
  // candidatos do departamento entram no pool e recebem sempre um registo
  // de agendamento — Diretor, RH (via Round-Robin) e a tentativa de
  // cruzamento de horário correm para todos, sem exceção. veioTalentPool
  // continua de fora da Fase 2 porque é uma rota diferente por desenho (o
  // Fast-Track salta a Fase 2 e só entra a partir da Fase 3) — não é uma
  // validação de etapa a esconder candidatos, é o próprio fluxo desses
  // candidatos.
  // CORREÇÃO DA MATEMÁTICA DO FUNIL: além da Talent Pool, exclui também
  // quem já foi Rejeitado na Fase 1 (phase0Status) — esses nunca chegam à
  // Fase 2. Com 74 candidatos, 11 Rejeitados e 9 da Talent Pool, sobram
  // exatamente os 54 elegíveis pedidos.
  const phase1Pool = useMemo(
    () => candidates.filter((c) => !c.veioTalentPool && c.phase0Status !== "Rejeitado"),
    [candidates]
  );
  const phase2Pool = useMemo(() => candidates.filter((c) => c.phase1Status === "Aprovado"), [candidates]);
  const phase3Pool = useMemo(() => candidates, [candidates]);

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
            title="Fase 2 — Entrevista de Soft Skills"
            subtitle="Candidato + Diretor(a) + 1 Membro RH do Departamento — cruzamento entre 3 intervenientes (Forms Fase 2 ∩ Excel Mestre)."
            phaseKey="fase1" availField="fase1" formsField="fase1" prevStatusField="phase0Status"
            candidates={candidates} setCandidates={setCandidates} members={members}
            bookings={phase1Bookings} setBookings={setPhase1Bookings}
            onGenerate={(dept) => setPhase1Bookings(regenerateForDepartment(phase1Pool, members, phase1Bookings, "fase1", ["diretorId", "rhId"], dept))}
            columns={[{ key: "diretorId", label: "Diretor(a)" }, { key: "rhId", label: "RH" }]}
            showCalendar={false}
            excludeTalentPool
          />
        )}
        {page === "fase2" && (
          <Phase2Page
            candidates={candidates} setCandidates={setCandidates} members={members}
            groups={phase2Groups} setGroups={setPhase2Groups}
            onGenerate={() => setPhase2Groups(generatePhase2(phase2Pool, members))}
          />
        )}
        {page === "fase3" && (
          <InterviewPhasePage
            title="Fase 4 — Entrevista de Hard Skills"
            subtitle="Candidato + Diretor + 1 Membro RH + Supervisor do Departamento — cruzamento exato entre 4 intervenientes (Forms Fase 4 ∩ Excel Mestre)."
            phaseKey="fase3" availField="fase3" formsField="fase3" prevStatusField="phase2Status"
            candidates={candidates} setCandidates={setCandidates} members={members}
            bookings={phase3Bookings} setBookings={setPhase3Bookings}
            onGenerate={(dept) => setPhase3Bookings(regenerateForDepartment(phase3Pool, members, phase3Bookings, "fase3", ["diretorId", "rhId", "supervisorId"], dept))}
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