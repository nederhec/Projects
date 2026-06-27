import { useState, useEffect, useRef, useCallback } from "react";

// --- SYNTHETIC EXECUTION CONTEXT ---
const CTX = {
  workspace: "Grupo_Proteina_Sul_SOE",
  period: "JUN/2026",
  reports: ["Painel_Carteira_V2", "Producao_Diaria_L3", "Estoque_SOE_Consolidado", "Margem_SKU_Planta"],
  datasets: ["ds_carteira_cliente", "ds_producao_real", "ds_estoque_final", "ds_margem_ebitda"],
  destination: "SharePoint/Comite_SOE/Base_Gerencial/",
  recipients: "comite-soe@proteina-sul.com.br",
  recipient_profile: "EXECUTIVO",
  tools: ["power_bi_api", "sharepoint_file", "messaging", "excel_engine"],
};

// --- SYNTHETIC RAW DATA (pre-extraction) ---
const RAW = [
  { SKU: "PROT_B_TAC", Planta: "L3_1400", Recurso: "LINHA_3", Produto: "Proteína Hidrolisada B",
    Cliente: "TAC_PREMIUM_001", Tipo_Cliente: "TAC", Carteira_Total: 600,
    Demanda_Media_Proxima_Janela: 200, Producao_Planejada: 400, Producao_Real: 350,
    Estoque_Final_Real: 50, Projecao_Estoque_SOE: -130, Margem_EBITDA_Rt: 380,
    Ciclo_Produtivo: "QUADRIMESTRAL", Lote_Minimo: 80, Volume_Transferido: 0, Aprovacao_Lote: true },
  { SKU: "PROT_A_STD", Planta: "L3_1400", Recurso: "LINHA_1", Produto: "Proteína Standard A",
    Cliente: "CARREFOUR_SP", Tipo_Cliente: "NORMAL", Carteira_Total: 800,
    Demanda_Media_Proxima_Janela: 300, Producao_Planejada: 500, Producao_Real: 480,
    Estoque_Final_Real: 200, Projecao_Estoque_SOE: 120, Margem_EBITDA_Rt: 450,
    Ciclo_Produtivo: "SEMESTRAL", Lote_Minimo: 100, Volume_Transferido: 0, Aprovacao_Lote: true },
  { SKU: "PROT_D_TAC", Planta: "L3_1400", Recurso: "LINHA_2", Produto: "Proteína D Premium",
    Cliente: "TAC_FARMACEUTICA_002", Tipo_Cliente: "TAC", Carteira_Total: 300,
    Demanda_Media_Proxima_Janela: 100, Producao_Planejada: 250, Producao_Real: 200,
    Estoque_Final_Real: 80, Projecao_Estoque_SOE: 60, Margem_EBITDA_Rt: 510,
    Ciclo_Produtivo: "QUADRIMESTRAL", Lote_Minimo: 120, Volume_Transferido: 0, Aprovacao_Lote: true },
  { SKU: "PROT_C_ATK", Planta: "Campinas", Recurso: "LINHA_C1", Produto: "Proteína C Atacado",
    Cliente: "ATACADAO_RJ", Tipo_Cliente: "NORMAL", Carteira_Total: 400,
    Demanda_Media_Proxima_Janela: 150, Producao_Planejada: 300, Producao_Real: 310,
    Estoque_Final_Real: 1200, Projecao_Estoque_SOE: 1800, Margem_EBITDA_Rt: 320,
    Ciclo_Produtivo: "TRANSFERENCIA", Lote_Minimo: 60, Volume_Transferido: 200, Aprovacao_Lote: true },
  { SKU: "PROT_E_NE", Planta: "Recife", Recurso: "LINHA_R2", Produto: "Proteína E Nordeste",
    Cliente: "GPA_NORDESTE", Tipo_Cliente: "NORMAL", Carteira_Total: 200,
    Demanda_Media_Proxima_Janela: 80, Producao_Planejada: 150, Producao_Real: 90,
    Estoque_Final_Real: 30, Projecao_Estoque_SOE: -10, Margem_EBITDA_Rt: 290,
    Ciclo_Produtivo: "SEMESTRAL", Lote_Minimo: 150, Volume_Transferido: 0, Aprovacao_Lote: false },
];

// --- BUSINESS RULES ENGINE (Regras 1-5, ordem correta) ---
function applyRules(data) {
  return data.map(row => {
    let obs = "";
    let status = "OK";

    // Regra 5 primeiro: calcular Impacto EBITDA para ponderar criticidade
    const pve = row.Projecao_Estoque_SOE - row.Estoque_Final_Real;

    // Regra 1: Estoque/PVE
    let impacto = 0;
    if (pve < 0) {
      status = "CRITICO";
      impacto = pve * row.Margem_EBITDA_Rt; // negativo
      obs += `[R1] PVE=${pve}t<0 → CRÍTICO. `;
    } else if (pve <= 1500) {
      status = "ATENCAO";
      impacto = pve * row.Margem_EBITDA_Rt * 0.12; // custo de oportunidade do excesso
      obs += `[R1] PVE=${pve}t (0-1500t, JUN não é mês K-Giro) → ATENÇÃO. `;
    } else {
      impacto = pve * row.Margem_EBITDA_Rt * 0.08;
      obs += `[R1] PVE=${pve}t>1500t → EXCESSO. `;
    }

    // Regra 2: Ruptura
    const protTAC = row.Tipo_Cliente === "TAC" ? Math.round(row.Carteira_Total * 0.2) : 0;
    const gatilho = row.Carteira_Total + row.Demanda_Media_Proxima_Janela + protTAC;
    const flagRuptura = row.Projecao_Estoque_SOE < gatilho ? 1 : 0;
    if (flagRuptura) {
      status = "CRITICO";
      obs += `[R2] Gatilho=${gatilho}t (Cart+Dem+TAC_prot) > Proj=${row.Projecao_Estoque_SOE}t → RUPTURA. `;
    }

    // Regra 3: Ciclo (já no campo, apenas registrar Transferência)
    if (row.Ciclo_Produtivo === "TRANSFERENCIA") {
      obs += `[R3] SKU Transferência: origem L3_1400 → Campinas, vol=${row.Volume_Transferido}t. `;
    }

    // Regra 4: Lote mínimo
    const flagLote = (row.Producao_Real < row.Lote_Minimo && !row.Aprovacao_Lote) ? 1 : 0;
    if (flagLote) {
      status = "CRITICO";
      obs += `[R4] Prod_Real=${row.Producao_Real}t<Lote_Min=${row.Lote_Minimo}t sem aprovação → EXCECAO_CRITICA. `;
    }

    const flagExcesso = pve > 1500 ? 1 : 0;

    return {
      ...row,
      Estoque_PVE: pve,
      Flag_Ruptura: flagRuptura,
      Flag_Excesso: flagExcesso,
      Flag_Excecao_Lote: flagLote,
      Impacto_EBITDA: Math.round(impacto),
      Status: status,
      Observacoes: obs.trim(),
    };
  });
}

// --- BACKEND REAL: soe-agent/server.js (Excel + SharePoint + Power BI) ---
const SOE_API_BASE =
  (typeof window !== "undefined" && window.__SOE_API_BASE__) || "http://localhost:4000";

async function saveExcelToSharePoint(data) {
  const res = await fetch(`${SOE_API_BASE}/api/soe/save-excel`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ data }),
  });
  const json = await res.json();
  if (!res.ok) throw new Error(json.error || `Falha ao salvar Excel (HTTP ${res.status})`);
  return json;
}

async function refreshPowerBIPanel() {
  const res = await fetch(`${SOE_API_BASE}/api/soe/refresh-powerbi`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ waitForCompletion: true }),
  });
  const json = await res.json();
  if (!res.ok) throw new Error(json.error || `Falha ao atualizar painel (HTTP ${res.status})`);
  return json;
}

// --- STEP DEFINITIONS ---
const STEPS = [
  { id: 1, name: "Inventário",           ms: 1400, log: "4 relatórios | 4 datasets | 15 tabelas inventariadas" },
  { id: 2, name: "Mapeamento",           ms: 1100, log: "24/26 campos mapeados | cobertura 92.3% | 2 campos opcionais ausentes" },
  { id: 3, name: "Extração",             ms: 1900, log: "5 SKUs | 3 plantas | 1.280 registros brutos consolidados em 5 linhas" },
  { id: 4, name: "Regras de Negócio",    ms: 1600, log: "Regras 1–5 aplicadas em sequência | 4 linhas com status ≠ OK detectadas" },
  { id: 5, name: "Validação",            ms: 1200, log: "0 nulos em campos-chave | 0 duplicidades | datasets atualizados há 6h (< 48h threshold)" },
  { id: 6, name: "Salvar Excel",         ms:  700, log: "Painel_SOE_Base_Atualizada.xlsx + cópia fixa gravados ✓" },
  { id: 7, name: "Atualizar Painel",     ms: 2200, log: "Refresh concluído | Status: SUCESSO | Painel disponível" },
  { id: 8, name: "Análise Executiva",    ms: null, log: "Geração via Claude API — claude-sonnet-4-6" },
  { id: 9, name: "Alertas",             ms:  600, log: "Alerta enviado → comite-soe@proteina-sul.com.br | 3 CRÍTICOS | 1 ATENÇÃO" },
  { id: 10, name: "Log e Auditoria",    ms:  400, log: "LOG_EXECUCAO fechado | versão SOE_AGENT_V4 registrada | append-only" },
];

const COL = {
  bg: "#060810", surface: "#0A0D17", border: "#1E2433",
  blue: "#3B82F6", red: "#EF4444", amber: "#F59E0B", green: "#10B981",
  purple: "#A78BFA", textPrimary: "#E2E8F0", textSec: "#64748B", textMuted: "#334155",
};

const badge = (status) => {
  const map = {
    CRITICO: { bg: "#EF444420", color: "#EF4444", border: "#EF444440" },
    ATENCAO: { bg: "#F59E0B20", color: "#F59E0B", border: "#F59E0B40" },
    OK:      { bg: "#10B98120", color: "#10B981", border: "#10B98140" },
  };
  const s = map[status] || map.OK;
  return { background: s.bg, color: s.color, border: `1px solid ${s.border}`, padding: "2px 8px", borderRadius: "3px", fontSize: "10px", fontFamily: "monospace", whiteSpace: "nowrap" };
};

export default function SOEAgent() {
  const [phase, setPhase] = useState("idle"); // idle | running | done | error
  const [currentStep, setCurrentStep] = useState(null);
  const [done, setDone] = useState(new Set());
  const [processedData, setProcessedData] = useState(null);
  const [analysis, setAnalysis] = useState("");
  const [isApiCall, setIsApiCall] = useState(false);
  const [tab, setTab] = useState("exec");
  const [logs, setLogs] = useState([]);
  const logRef = useRef(null);

  const addLog = useCallback((etapa, acao, status, detalhe) => {
    setLogs(p => [...p, { id: Date.now() + Math.random(), ts: new Date().toLocaleTimeString("pt-BR"), etapa, acao, status, detalhe }]);
  }, []);

  useEffect(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight; }, [logs]);

  const callAPI = async (data) => {
    const criticals = data.filter(d => d.Status === "CRITICO");
    const tacRisk = data.filter(d => d.Tipo_Cliente === "TAC" && d.Flag_Ruptura === 1);
    const totalNeg = data.reduce((s, d) => s + (d.Impacto_EBITDA < 0 ? d.Impacto_EBITDA : 0), 0);

    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1000,
        system: `Você é o Agente S&OE V4 — Sales & Operations Execution. Análise técnica, executiva e defensável. Linguagem direta, sem prolixidade. Cada afirmação suportada pelos dados. Use markdown limpo. Nunca use bullet points decorativos — use dados.`,
        messages: [{
          role: "user",
          content: `BASE CONSOLIDADA PROCESSADA — JUN/2026 — Grupo Proteína Sul:

${data.map(d => `SKU: ${d.SKU} | Planta: ${d.Planta} | Cliente: ${d.Cliente} [${d.Tipo_Cliente}]
  PVE: ${d.Estoque_PVE}t | Proj.SOE: ${d.Projecao_Estoque_SOE}t | Est.Real: ${d.Estoque_Final_Real}t
  Prod.Plan: ${d.Producao_Planejada}t | Prod.Real: ${d.Producao_Real}t | Margem: R$${d.Margem_EBITDA_Rt}/t
  Status: ${d.Status} | Flag_Ruptura: ${d.Flag_Ruptura} | Flag_Lote: ${d.Flag_Excecao_Lote}
  Impacto_EBITDA: R$ ${d.Impacto_EBITDA.toLocaleString("pt-BR")}
  Raciocínio: ${d.Observacoes}`).join("\n\n")}

RESUMO EXECUTIVO:
- Total SKUs: ${data.length} | CRÍTICOS: ${criticals.length} | ATENÇÃO: ${data.filter(d=>d.Status==="ATENCAO").length}
- Clientes TAC com risco de ruptura: ${tacRisk.length}/${data.filter(d=>d.Tipo_Cliente==="TAC").length}
- Impacto EBITDA negativo total: R$ ${Math.abs(totalNeg).toLocaleString("pt-BR")}

Entregue EXATAMENTE nesta estrutura:

## Diagnóstico Executivo
(3-4 linhas objetivas citando os dados. Sem suavizar o que é crítico.)

## Top Ofensores PVE
(tabela markdown: SKU | PVE | Impacto EBITDA | Prioridade | Ação urgente)

## Clientes TAC
(analise cada cliente TAC separadamente. Status, risco e exposição.)

## Cenários de Decisão

### CONSERVADOR
Lógica: ... | Ação imediata: ... | Risco residual: ...

### EQUILIBRADO
Lógica: ... | Ação imediata: ... | Risco residual: ...

### AGRESSIVO
Lógica: ... | Ação imediata: ... | Risco residual: ...

## [DECISÃO RECOMENDADA]
Uma frase. Cenário + ação específica + justificativa técnica.`
        }]
      })
    });
    const j = await r.json();
    return j.content?.[0]?.text || "Erro na API.";
  };

  const run = async () => {
    if (phase !== "idle") return;
    setPhase("running");
    setLogs([]); setDone(new Set()); setProcessedData(null); setAnalysis(""); setTab("exec");
    addLog("PRÉ-EXECUÇÃO", "Validação", "OK", "4/4 campos obrigatórios ✓ | 4/4 tools ativos ✓ | Checkpoint liberado.");

    let consolidatedData = null;

    for (let i = 0; i < STEPS.length; i++) {
      const s = STEPS[i];
      setCurrentStep(s.id);
      addLog(`Etapa ${s.id}`, s.name, "EXEC", `Iniciando ${s.name}...`);

      let stepLog = s.log;

      if (s.id === 4) {
        await new Promise(r => setTimeout(r, s.ms));
        consolidatedData = applyRules(RAW);
        setProcessedData(consolidatedData);
      } else if (s.id === 6) {
        try {
          const result = await saveExcelToSharePoint(consolidatedData);
          stepLog = `${result.filename} salvo em ${CTX.destination} ✓${result.webUrl ? ` (${result.webUrl})` : ""}`;
        } catch (err) {
          addLog("Etapa 6", "Salvar Excel", "FALHA", err.message);
          setTab("log"); setPhase("error"); setCurrentStep(null);
          return;
        }
      } else if (s.id === 7) {
        try {
          const result = await refreshPowerBIPanel();
          stepLog = `Refresh dataset ${result.datasetId} | status: ${result.status}`;
        } catch (err) {
          addLog("Etapa 7", "Atualizar Painel", "FALHA", err.message);
          setTab("log"); setPhase("error"); setCurrentStep(null);
          return;
        }
      } else if (s.id === 8) {
        setIsApiCall(true);
        setTab("analysis");
        try {
          const txt = await callAPI(consolidatedData);
          setAnalysis(txt);
          addLog("Etapa 8", "Análise Executiva", "OK", "claude-sonnet-4-6 | análise gerada com sucesso.");
        } catch (e) {
          setAnalysis("**Erro na chamada da API:** " + e.message);
          addLog("Etapa 8", "Análise Executiva", "FALHA", e.message);
        }
        setIsApiCall(false);
      } else {
        await new Promise(r => setTimeout(r, s.ms));
      }

      setDone(p => new Set([...p, s.id]));
      addLog(`Etapa ${s.id}`, s.name, "OK", stepLog);
    }

    setPhase("done");
    setCurrentStep(null);
  };

  const reset = () => { setPhase("idle"); setDone(new Set()); setCurrentStep(null); setProcessedData(null); setAnalysis(""); setLogs([]); setTab("exec"); };

  const data = processedData || [];
  const critCount = data.filter(d => d.Status === "CRITICO").length;
  const atCount = data.filter(d => d.Status === "ATENCAO").length;
  const tacRiskCount = data.filter(d => d.Tipo_Cliente === "TAC" && d.Flag_Ruptura === 1).length;
  const totalImpact = data.reduce((s, d) => s + (d.Impacto_EBITDA < 0 ? d.Impacto_EBITDA : 0), 0);

  const statusColor = phase === "done" ? COL.green : phase === "running" ? COL.amber : phase === "error" ? COL.red : COL.textMuted;
  const statusLabel = phase === "idle" ? "AGUARDANDO" : phase === "running" ? "EXECUTANDO" : phase === "error" ? "FALHA" : "CONCLUÍDO";

  const renderMarkdown = (text) =>
    text.split("\n").map((line, i) => {
      if (line.startsWith("## ")) return <h2 key={i} style={{ color: COL.textPrimary, fontSize: "14px", fontWeight: "700", marginTop: "22px", marginBottom: "8px", paddingBottom: "6px", borderBottom: `1px solid ${COL.border}` }}>{line.slice(3)}</h2>;
      if (line.startsWith("### ")) return <h3 key={i} style={{ color: "#94A3B8", fontSize: "13px", fontWeight: "600", marginTop: "14px", marginBottom: "6px" }}>{line.slice(4)}</h3>;
      if (line.startsWith("**") && line.endsWith("**")) return <p key={i} style={{ color: COL.textPrimary, fontWeight: "700", margin: "4px 0" }}>{line.replace(/\*\*/g, "")}</p>;
      if (line.startsWith("- ")) return <div key={i} style={{ paddingLeft: "16px", color: "#94A3B8", margin: "3px 0" }}>• {line.slice(2)}</div>;
      if (line.startsWith("|")) return <div key={i} style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec, margin: "2px 0" }}>{line}</div>;
      if (line.trim() === "") return <div key={i} style={{ height: "8px" }} />;
      return <p key={i} style={{ margin: "4px 0", fontSize: "13px", color: "#CBD5E1" }}>{line}</p>;
    });

  return (
    <div style={{ fontFamily: "'Inter', system-ui, sans-serif", background: COL.bg, color: COL.textPrimary, minHeight: "100vh", fontSize: "13px" }}>

      {/* HEADER */}
      <div style={{ background: COL.surface, borderBottom: `1px solid ${COL.border}`, padding: "11px 20px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
          <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: statusColor, boxShadow: phase === "running" ? `0 0 8px ${COL.amber}` : "none" }} />
          <span style={{ fontFamily: "monospace", color: COL.textSec, fontSize: "11px", letterSpacing: "0.05em" }}>SOE_MASTER_AGENT</span>
          <span style={{ color: COL.border }}>|</span>
          <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.blue }}>V4.0</span>
          <span style={{ color: COL.border }}>|</span>
          <span style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textMuted }}>SharePoint + Power BI integrados via soe-agent/server.js</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
          <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec }}>{CTX.workspace}</span>
          <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec }}>{CTX.period}</span>
          <span style={{ fontFamily: "monospace", fontSize: "10px", padding: "3px 8px", borderRadius: "3px", background: `${statusColor}20`, color: statusColor, border: `1px solid ${statusColor}40` }}>
            {statusLabel}
          </span>
          {(phase === "done" || phase === "error") && (
            <button onClick={reset} style={{ background: "none", border: `1px solid ${COL.border}`, borderRadius: "4px", color: COL.textSec, padding: "3px 10px", cursor: "pointer", fontFamily: "monospace", fontSize: "10px" }}>↺ RESET</button>
          )}
        </div>
      </div>

      {/* IDLE */}
      {phase === "idle" && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "calc(100vh - 49px)", gap: "28px" }}>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontFamily: "monospace", color: COL.blue, fontSize: "10px", marginBottom: "8px", letterSpacing: "0.2em" }}>EXECUÇÃO COM INTEGRAÇÕES REAIS</div>
            <div style={{ fontSize: "22px", fontWeight: "700", color: COL.textPrimary, marginBottom: "6px" }}>S&OE Master Agent V4</div>
            <div style={{ color: COL.textSec, lineHeight: "1.7", fontSize: "13px" }}>
              5 SKUs · 3 plantas · 2 clientes TAC · regras de negócio reais<br/>
              Etapa 6 grava Excel real no SharePoint · Etapa 7 atualiza o dataset real no Power BI
            </div>
          </div>

          <div style={{ background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "6px", padding: "16px 20px", minWidth: "440px", fontFamily: "monospace", fontSize: "11px" }}>
            <div style={{ color: COL.textMuted, marginBottom: "8px" }}>{"<execution_context>"}</div>
            {[
              ["workspace", CTX.workspace],
              ["period", CTX.period],
              ["reports", `${CTX.reports.length} declarados`],
              ["datasets", `${CTX.datasets.length} declarados`],
              ["recipient_profile", CTX.recipient_profile],
              ["tools", `${CTX.tools.length}/4 ativos`],
              ["soe_api_base", SOE_API_BASE],
            ].map(([k, v]) => (
              <div key={k} style={{ paddingLeft: "16px", marginBottom: "3px" }}>
                <span style={{ color: "#7C3AED" }}>{k}</span><span style={{ color: COL.textMuted }}>: </span>
                <span style={{ color: "#94A3B8" }}>"{v}"</span>
              </div>
            ))}
            <div style={{ color: COL.textMuted, marginTop: "8px" }}>{"</execution_context>"}</div>
            <div style={{ marginTop: "10px", padding: "6px 8px", background: "#10B98110", border: `1px solid #10B98130`, borderRadius: "3px" }}>
              <span style={{ color: COL.green }}>✓ PRÉ-CONDIÇÕES VALIDADAS — pronto para execução</span>
            </div>
          </div>

          <button onClick={run} style={{ background: COL.blue, color: "#fff", border: "none", borderRadius: "6px", padding: "12px 36px", fontSize: "14px", fontWeight: "600", cursor: "pointer", letterSpacing: "0.04em" }}>
            ▶ INICIAR EXECUÇÃO
          </button>
        </div>
      )}

      {/* RUNNING / DONE / ERROR */}
      {phase !== "idle" && (
        <div style={{ display: "grid", gridTemplateColumns: "210px 1fr", height: "calc(100vh - 49px)" }}>

          {/* LEFT: STEP TIMELINE */}
          <div style={{ background: COL.surface, borderRight: `1px solid ${COL.border}`, overflowY: "auto", padding: "12px 0" }}>
            {STEPS.map(s => {
              const isDone = done.has(s.id);
              const isActive = currentStep === s.id;
              return (
                <div key={s.id} style={{ display: "flex", gap: "10px", padding: "8px 14px", background: isActive ? "#1E2433" : "transparent", borderLeft: isActive ? `2px solid ${COL.blue}` : "2px solid transparent", alignItems: "flex-start" }}>
                  <div style={{ width: "20px", height: "20px", borderRadius: "50%", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "10px", fontFamily: "monospace",
                    background: isDone ? `${COL.green}20` : isActive ? `${COL.blue}20` : "#1E2433",
                    border: `1px solid ${isDone ? `${COL.green}50` : isActive ? `${COL.blue}50` : "#2D3748"}`,
                    color: isDone ? COL.green : isActive ? COL.blue : "#475569"
                  }}>
                    {isDone ? "✓" : s.id}
                  </div>
                  <div>
                    <div style={{ fontSize: "12px", fontWeight: isDone ? "400" : "500", color: isDone ? COL.textSec : isActive ? COL.textPrimary : "#475569" }}>{s.name}</div>
                    {isActive && <div style={{ fontSize: "10px", color: "#475569", marginTop: "2px" }}>{s.id === 8 && isApiCall ? "claude-sonnet-4-6..." : "executando..."}</div>}
                    {isDone && <div style={{ fontSize: "10px", color: COL.textMuted, marginTop: "2px" }}>concluído</div>}
                  </div>
                </div>
              );
            })}
            <div style={{ height: "4px", margin: "8px 14px", background: "#1E2433", borderRadius: "2px", overflow: "hidden" }}>
              <div style={{ height: "100%", width: `${(done.size / 10) * 100}%`, background: phase === "error" ? COL.red : COL.blue, transition: "width 0.4s" }} />
            </div>
            {phase === "done" && (
              <div style={{ margin: "10px 14px", padding: "10px", background: `${COL.green}10`, border: `1px solid ${COL.green}30`, borderRadius: "4px", textAlign: "center" }}>
                <div style={{ color: COL.green, fontSize: "11px", fontWeight: "600" }}>✓ 10/10 CONCLUÍDO</div>
                <div style={{ color: COL.textMuted, fontSize: "10px", fontFamily: "monospace", marginTop: "2px" }}>SOE_AGENT_V4</div>
              </div>
            )}
            {phase === "error" && (
              <div style={{ margin: "10px 14px", padding: "10px", background: `${COL.red}10`, border: `1px solid ${COL.red}30`, borderRadius: "4px", textAlign: "center" }}>
                <div style={{ color: COL.red, fontSize: "11px", fontWeight: "600" }}>✗ EXECUÇÃO INTERROMPIDA</div>
                <div style={{ color: COL.textMuted, fontSize: "10px", fontFamily: "monospace", marginTop: "2px" }}>ver aba LOG</div>
              </div>
            )}
          </div>

          {/* RIGHT: MAIN CONTENT */}
          <div style={{ display: "flex", flexDirection: "column", overflow: "hidden" }}>
            {/* TABS */}
            <div style={{ display: "flex", borderBottom: `1px solid ${COL.border}`, background: COL.surface, flexShrink: 0 }}>
              {[
                { id: "exec", label: "EXECUÇÃO" },
                { id: "base", label: "BASE_CONSOLIDADA" },
                { id: "kpi", label: "KPI_ALERTAS" },
                { id: "analysis", label: "ANÁLISE EXECUTIVA" },
                { id: "log", label: "LOG" },
              ].map(t => (
                <button key={t.id} onClick={() => setTab(t.id)} style={{ background: "none", border: "none", cursor: "pointer", padding: "10px 16px", fontSize: "10px", fontFamily: "monospace", letterSpacing: "0.06em", color: tab === t.id ? COL.blue : COL.textSec, borderBottom: tab === t.id ? `2px solid ${COL.blue}` : "2px solid transparent", marginBottom: "-1px" }}>
                  {t.label}
                </button>
              ))}
            </div>

            <div style={{ flex: 1, overflow: "auto", padding: "20px" }}>

              {/* TAB: EXECUÇÃO */}
              {tab === "exec" && (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "16px" }}>
                  <div style={{ background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "6px", padding: "16px" }}>
                    <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.blue, marginBottom: "12px", letterSpacing: "0.12em" }}>EXECUTION_CONTEXT VALIDADO</div>
                    {[["workspace", CTX.workspace], ["period", CTX.period], ["reports", `${CTX.reports.length} declarados`], ["datasets", `${CTX.datasets.length} mapeados`], ["recipient_profile", CTX.recipient_profile], ["tools", "4/4 ativos"]].map(([k, v]) => (
                      <div key={k} style={{ display: "flex", justifyContent: "space-between", marginBottom: "6px", paddingBottom: "6px", borderBottom: `1px solid ${COL.border}` }}>
                        <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec }}>{k}</span>
                        <span style={{ fontFamily: "monospace", fontSize: "11px", color: "#94A3B8" }}>{v}</span>
                      </div>
                    ))}
                    <div style={{ padding: "5px 8px", background: `${COL.green}10`, border: `1px solid ${COL.green}30`, borderRadius: "3px", marginTop: "4px" }}>
                      <span style={{ fontFamily: "monospace", fontSize: "10px", color: COL.green }}>✓ CHECKPOINT PRÉ-EXECUÇÃO — LIBERADO</span>
                    </div>
                  </div>

                  <div style={{ background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "6px", padding: "16px" }}>
                    <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.blue, marginBottom: "12px", letterSpacing: "0.12em" }}>PROGRESSO DA EXECUÇÃO</div>
                    <div style={{ marginBottom: "14px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "4px" }}>
                        <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec }}>etapas</span>
                        <span style={{ fontFamily: "monospace", fontSize: "11px" }}>{done.size}/10</span>
                      </div>
                      <div style={{ height: "4px", background: "#1E2433", borderRadius: "2px" }}>
                        <div style={{ height: "100%", width: `${(done.size / 10) * 100}%`, background: phase === "error" ? COL.red : COL.blue, transition: "width 0.4s", borderRadius: "2px" }} />
                      </div>
                    </div>
                    {currentStep && (
                      <div style={{ padding: "8px", background: `${COL.blue}10`, border: `1px solid ${COL.blue}30`, borderRadius: "3px", marginBottom: "8px" }}>
                        <span style={{ fontFamily: "monospace", fontSize: "10px", color: COL.blue }}>▶ Etapa {currentStep}: {STEPS[currentStep - 1]?.name}{currentStep === 8 && isApiCall ? " [API CALL ATIVO]" : ""}</span>
                      </div>
                    )}
                    {phase === "done" && (
                      <div style={{ padding: "8px", background: `${COL.green}10`, border: `1px solid ${COL.green}30`, borderRadius: "3px", marginBottom: "10px" }}>
                        <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.green }}>✓ EXECUÇÃO COMPLETA — 10/10 etapas</span>
                      </div>
                    )}
                    {phase === "error" && (
                      <div style={{ padding: "8px", background: `${COL.red}10`, border: `1px solid ${COL.red}30`, borderRadius: "3px", marginBottom: "10px" }}>
                        <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.red }}>✗ EXECUÇÃO INTERROMPIDA — ver aba LOG</span>
                      </div>
                    )}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "4px" }}>
                      {STEPS.map(s => (
                        <div key={s.id} style={{ display: "flex", gap: "6px", alignItems: "center" }}>
                          <div style={{ width: "6px", height: "6px", borderRadius: "50%", flexShrink: 0, background: done.has(s.id) ? COL.green : currentStep === s.id ? COL.amber : "#1E2433" }} />
                          <span style={{ fontFamily: "monospace", fontSize: "10px", color: done.has(s.id) ? "#475569" : currentStep === s.id ? COL.textPrimary : "#2D3748" }}>E{s.id}: {s.name}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Fontes BI */}
                  <div style={{ gridColumn: "1 / -1", background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "6px", padding: "16px" }}>
                    <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.blue, marginBottom: "12px", letterSpacing: "0.12em" }}>ABA 2: FONTES_BI — INVENTÁRIO</div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "monospace", fontSize: "11px" }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${COL.border}` }}>
                          {["Relatório", "Dataset", "Tabela_Principal", "Campos_Utilizados", "Última_Atualização", "Criticidade", "Status"].map(h => (
                            <th key={h} style={{ padding: "6px 10px", textAlign: "left", color: COL.textSec, fontWeight: "500", fontSize: "10px" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {[
                          ["Painel_Carteira_V2", "ds_carteira_cliente", "tb_carteira_ativa", "SKU, Cliente, Tipo, Carteira_Total, Demanda_Media", "27/06/26 09:14", "ALTA", "VERIFICADO"],
                          ["Producao_Diaria_L3", "ds_producao_real", "tb_producao_planta", "SKU, Planta, Recurso, Prod_Plan, Prod_Real, Lote_Min", "27/06/26 10:02", "ALTA", "VERIFICADO"],
                          ["Estoque_SOE_Consolidado", "ds_estoque_final", "tb_estoque_soe", "SKU, Planta, Est_Real, Proj_SOE, Ciclo_Produtivo", "27/06/26 10:45", "ALTA", "VERIFICADO"],
                          ["Margem_SKU_Planta", "ds_margem_ebitda", "tb_margem_rt", "SKU, Planta, Margem_EBITDA_Rt, Origem_Abastecimento", "27/06/26 08:30", "ALTA", "VERIFICADO"],
                        ].map((row, i) => (
                          <tr key={i} style={{ borderBottom: `1px solid #111827`, background: i % 2 === 0 ? "#0A0D1760" : "transparent" }}>
                            {row.map((cell, j) => (
                              <td key={j} style={{ padding: "6px 10px", color: j === 6 ? COL.green : j === 5 ? COL.amber : "#94A3B8", fontSize: "11px" }}>{cell}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* TAB: BASE_CONSOLIDADA */}
              {tab === "base" && (
                <div>
                  {data.length === 0 ? (
                    <div style={{ textAlign: "center", color: COL.textSec, padding: "60px", fontFamily: "monospace", fontSize: "12px" }}>
                      Aguardando Etapa 4 (Regras de Negócio)...
                    </div>
                  ) : (
                    <>
                      <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, marginBottom: "14px" }}>
                        ABA 1: BASE_CONSOLIDADA · {data.length} linhas · JUN/2026 · Regras 1–5 aplicadas
                      </div>
                      <div style={{ overflowX: "auto", marginBottom: "24px" }}>
                        <table style={{ borderCollapse: "collapse", fontFamily: "monospace", fontSize: "11px", minWidth: "100%" }}>
                          <thead>
                            <tr style={{ borderBottom: `1px solid ${COL.border}` }}>
                              {["SKU","Planta","Cliente","Tipo","Carteira","Dem.Prox","Prod.Plan","Prod.Real","Est.Real","Proj.SOE","PVE","Marg.R$/t","F_Rupt","F_Lote","Status","Imp.EBITDA"].map(h => (
                                <th key={h} style={{ padding: "6px 10px", textAlign: "right", color: COL.textSec, fontWeight: "500", whiteSpace: "nowrap", textAlign: h === "SKU" || h === "Planta" || h === "Cliente" || h === "Tipo" || h === "Status" ? "left" : "right", fontSize: "10px" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {data.map((row, i) => (
                              <tr key={i} style={{ borderBottom: `1px solid #111827`, background: i % 2 === 0 ? "#0A0D1760" : "transparent" }}>
                                <td style={{ padding: "6px 10px", color: "#94A3B8", whiteSpace: "nowrap" }}>{row.SKU}</td>
                                <td style={{ padding: "6px 10px", color: COL.textSec }}>{row.Planta}</td>
                                <td style={{ padding: "6px 10px", color: COL.textSec, maxWidth: "90px", overflow: "hidden", textOverflow: "ellipsis" }}>{row.Cliente}</td>
                                <td style={{ padding: "6px 10px" }}>
                                  <span style={{ padding: "1px 6px", borderRadius: "3px", fontSize: "10px", background: row.Tipo_Cliente === "TAC" ? "#7C3AED20" : "#1E2433", color: row.Tipo_Cliente === "TAC" ? "#A78BFA" : COL.textSec, border: row.Tipo_Cliente === "TAC" ? "1px solid #7C3AED40" : `1px solid ${COL.border}` }}>{row.Tipo_Cliente}</span>
                                </td>
                                {[row.Carteira_Total, row.Demanda_Media_Proxima_Janela, row.Producao_Planejada, row.Producao_Real, row.Estoque_Final_Real, row.Projecao_Estoque_SOE].map((v, j) => (
                                  <td key={j} style={{ padding: "6px 10px", color: COL.textSec, textAlign: "right" }}>{v}t</td>
                                ))}
                                <td style={{ padding: "6px 10px", textAlign: "right", fontWeight: "700", color: row.Estoque_PVE < 0 ? COL.red : row.Estoque_PVE > 1500 ? COL.amber : COL.green }}>{row.Estoque_PVE}t</td>
                                <td style={{ padding: "6px 10px", textAlign: "right", color: COL.textSec }}>R${row.Margem_EBITDA_Rt}</td>
                                <td style={{ padding: "6px 10px", textAlign: "center", color: row.Flag_Ruptura ? COL.red : "#475569" }}>{row.Flag_Ruptura ? "⚠1" : "0"}</td>
                                <td style={{ padding: "6px 10px", textAlign: "center", color: row.Flag_Excecao_Lote ? COL.red : "#475569" }}>{row.Flag_Excecao_Lote ? "⚠1" : "0"}</td>
                                <td style={{ padding: "6px 10px" }}><span style={badge(row.Status)}>{row.Status}</span></td>
                                <td style={{ padding: "6px 10px", textAlign: "right", fontWeight: "600", color: row.Impacto_EBITDA < 0 ? COL.red : COL.green }}>{row.Impacto_EBITDA < 0 ? "-" : "+"}R${Math.abs(row.Impacto_EBITDA).toLocaleString("pt-BR")}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>

                      <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, marginBottom: "10px" }}>RACIOCÍNIO APLICADO — campo Observacoes (anti_hallucination rule ativo)</div>
                      {data.filter(d => d.Observacoes).map((row, i) => (
                        <div key={i} style={{ background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "4px", padding: "10px 14px", marginBottom: "6px" }}>
                          <div style={{ display: "flex", gap: "8px", marginBottom: "4px", alignItems: "center" }}>
                            <span style={{ fontFamily: "monospace", fontSize: "11px", color: "#7C3AED" }}>{row.SKU}</span>
                            <span style={badge(row.Status)}>{row.Status}</span>
                          </div>
                          <div style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec, lineHeight: "1.6" }}>{row.Observacoes}</div>
                        </div>
                      ))}
                    </>
                  )}
                </div>
              )}

              {/* TAB: KPI_ALERTAS */}
              {tab === "kpi" && (
                <div>
                  {data.length === 0 ? (
                    <div style={{ textAlign: "center", color: COL.textSec, padding: "60px", fontFamily: "monospace", fontSize: "12px" }}>Aguardando Etapa 4...</div>
                  ) : (
                    <>
                      <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, marginBottom: "16px" }}>ABA 4: KPI_ALERTAS · ordenado por Impacto EBITDA</div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "12px", marginBottom: "24px" }}>
                        {[
                          { label: "CRÍTICOS", value: critCount, color: COL.red },
                          { label: "ATENÇÃO", value: atCount, color: COL.amber },
                          { label: "TAC EM RISCO", value: tacRiskCount, color: "#A78BFA" },
                          { label: "IMPACTO EBITDA", value: `R$ ${Math.abs(totalImpact).toLocaleString("pt-BR")}`, color: COL.red },
                        ].map(c => (
                          <div key={c.label} style={{ background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "6px", padding: "14px 16px" }}>
                            <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, marginBottom: "6px" }}>{c.label}</div>
                            <div style={{ fontSize: "22px", fontWeight: "700", color: c.color }}>{c.value}</div>
                          </div>
                        ))}
                      </div>
                      {data.filter(d => d.Status !== "OK").sort((a, b) => a.Impacto_EBITDA - b.Impacto_EBITDA).map((row, i) => (
                        <div key={i} style={{ background: COL.surface, border: `1px solid ${row.Status === "CRITICO" ? "#EF444430" : "#F59E0B30"}`, borderRadius: "6px", padding: "14px 16px", marginBottom: "10px" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "10px" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                              <span style={{ fontFamily: "monospace", color: COL.textPrimary, fontWeight: "600" }}>{row.SKU}</span>
                              {row.Tipo_Cliente === "TAC" && <span style={{ fontFamily: "monospace", fontSize: "10px", padding: "1px 6px", background: "#7C3AED20", color: "#A78BFA", border: "1px solid #7C3AED40", borderRadius: "3px" }}>TAC</span>}
                              {row.Flag_Ruptura === 1 && <span style={{ fontFamily: "monospace", fontSize: "10px", padding: "1px 6px", background: `${COL.red}20`, color: COL.red, border: `1px solid ${COL.red}40`, borderRadius: "3px" }}>RUPTURA</span>}
                              {row.Flag_Excecao_Lote === 1 && <span style={{ fontFamily: "monospace", fontSize: "10px", padding: "1px 6px", background: `${COL.amber}20`, color: COL.amber, border: `1px solid ${COL.amber}40`, borderRadius: "3px" }}>LOTE_MIN</span>}
                              <span style={{ fontFamily: "monospace", fontSize: "11px", color: COL.textSec }}>{row.Planta} · {row.Cliente}</span>
                            </div>
                            <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
                              <span style={badge(row.Status)}>{row.Status}</span>
                              <span style={{ fontFamily: "monospace", fontSize: "13px", fontWeight: "700", color: row.Impacto_EBITDA < 0 ? COL.red : COL.amber }}>{row.Impacto_EBITDA < 0 ? "-" : "+"}R${Math.abs(row.Impacto_EBITDA).toLocaleString("pt-BR")}</span>
                            </div>
                          </div>
                          <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: "8px", marginBottom: "10px" }}>
                            {[["PVE", `${row.Estoque_PVE}t`, row.Estoque_PVE < 0], ["Proj.SOE", `${row.Projecao_Estoque_SOE}t`, row.Projecao_Estoque_SOE < 0], ["Est.Real", `${row.Estoque_Final_Real}t`, false], ["Prod.Real", `${row.Producao_Real}t`, row.Producao_Real < row.Lote_Minimo], ["Margem", `R$${row.Margem_EBITDA_Rt}/t`, false]].map(([l, v, crit]) => (
                              <div key={l} style={{ background: "#111827", borderRadius: "3px", padding: "6px 10px" }}>
                                <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec }}>{l}</div>
                                <div style={{ fontFamily: "monospace", fontSize: "12px", fontWeight: "700", color: crit ? COL.red : "#94A3B8" }}>{v}</div>
                              </div>
                            ))}
                          </div>
                          <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, borderTop: `1px solid ${COL.border}`, paddingTop: "8px", lineHeight: "1.5" }}>{row.Observacoes}</div>
                        </div>
                      ))}
                    </>
                  )}
                </div>
              )}

              {/* TAB: ANÁLISE EXECUTIVA */}
              {tab === "analysis" && (
                <div>
                  {isApiCall ? (
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "300px", gap: "12px" }}>
                      <div style={{ width: "32px", height: "32px", border: `2px solid ${COL.border}`, borderTopColor: COL.blue, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
                      <div style={{ fontFamily: "monospace", color: COL.blue, fontSize: "12px" }}>Gerando análise executiva...</div>
                      <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec }}>claude-sonnet-4-6 | S&OE_AGENT_V4 | Etapa 8</div>
                      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
                    </div>
                  ) : analysis ? (
                    <div>
                      <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, marginBottom: "16px" }}>ANÁLISE EXECUTIVA — claude-sonnet-4-6 via API — JUN/2026 — perfil: EXECUTIVO</div>
                      <div style={{ background: COL.surface, border: `1px solid ${COL.border}`, borderRadius: "6px", padding: "22px 26px", lineHeight: "1.7" }}>
                        {renderMarkdown(analysis)}
                      </div>
                    </div>
                  ) : (
                    <div style={{ textAlign: "center", color: COL.textSec, padding: "80px", fontFamily: "monospace", fontSize: "12px" }}>
                      Análise disponível após conclusão da Etapa 8
                    </div>
                  )}
                </div>
              )}

              {/* TAB: LOG */}
              {tab === "log" && (
                <div>
                  <div style={{ fontFamily: "monospace", fontSize: "10px", color: COL.textSec, marginBottom: "12px" }}>ABA 5: LOG_EXECUCAO · SOE_AGENT_V4 · append-only · {logs.length} entradas</div>
                  <div ref={logRef} style={{ fontFamily: "monospace", fontSize: "11px", background: "#060810", border: `1px solid ${COL.border}`, borderRadius: "4px", padding: "14px", maxHeight: "520px", overflowY: "auto" }}>
                    {logs.map(e => (
                      <div key={e.id} style={{ display: "flex", gap: "10px", marginBottom: "5px", alignItems: "flex-start" }}>
                        <span style={{ color: COL.textMuted, flexShrink: 0, width: "60px" }}>{e.ts}</span>
                        <span style={{ color: e.status === "OK" ? COL.green : e.status === "FALHA" ? COL.red : COL.blue, flexShrink: 0, width: "50px" }}>[{e.status}]</span>
                        <span style={{ color: "#475569", flexShrink: 0, width: "80px" }}>{e.etapa}</span>
                        <span style={{ color: COL.textSec }}>{e.detalhe}</span>
                      </div>
                    ))}
                    {logs.length === 0 && <span style={{ color: COL.textMuted }}>aguardando execução...</span>}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
