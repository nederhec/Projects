# Power BI Data Model – Cargos GBS
## Documentação de Esquema e Design

---

## 1. Estrutura de Tabelas

### Tabela Fato: `Cargos`
Tabela principal carregada pelo Power Query (`PowerQuery_Transform.m`).

| Campo | Tipo | Descrição |
|---|---|---|
| CÓDIGO CARGO | Inteiro | Chave primária do cargo |
| EMPRESA | Texto | Nome da empresa no grupo |
| EMPRESA (SAP) | Texto | Nome da empresa no sistema SAP |
| CARGO | Texto | Nome completo do cargo |
| GRADE | Inteiro | Nível salarial (2–13) |
| SITUAÇÃO | Texto | "ATIVO" ou "INATIVO" |
| GRUPO | Texto | "Corporativo" ou "Negócios & operações" |
| FAMÍLIA GBS | Texto | Família de cargos GBS |
| SUB FAMÍLIA | Texto | Subfamília de cargos |
| GRUPO HIERARQUICO | Texto | Nível hierárquico (GA, GS, Coordenador…) |
| NÍVEL | Texto | Nível detalhado (Analista Sr, Pl, Jr…) |
| CBO | Texto | Código Brasileiro de Ocupações |
| REGRA CONTROLE JORNADA | Texto | Tipo de apuração de horas |
| CONTROLE JORNADA | Texto | "Elegível" / "Não elegível" |
| COTA APRENDIZ | Texto | Elegibilidade à cota aprendiz |
| TIPO CONTRATO | Texto | CLT / PJ |
| SINDICATO | Texto | Preponderante / Específico |
| MERCADO SELECIONADO | Texto | Mercado de referência Mercer |
| DESCRIÇÃO SUMÁRIA | Texto | Descrição das atribuições do cargo |
| Cod. Mercer 2025 | Texto | Código de pesquisa Mercer 2025 |
| Cod. Mercer 2026 | Texto | Código de pesquisa Mercer 2026 |
| **IsActive** | Booleano | *Calculado:* TRUE se SITUAÇÃO = "ATIVO" |
| **HierarchyOrder** | Inteiro | *Calculado:* Ordem hierárquica (GA=1 … Auxiliar=8) |
| **HasMercer2026** | Booleano | *Calculado:* TRUE se Cod. Mercer 2026 preenchido |
| **GradeCategory** | Texto | *Calculado:* Low/Mid/Senior por faixa de grade |

---

### Tabela Dimensão: `DimEmpresa`
Criar manualmente ou via Power Query a partir de `Cargos`.

```m
DimEmpresa = Table.Distinct(
    Table.SelectColumns( Cargos, {"EMPRESA", "EMPRESA (SAP)", "MERCADO SELECIONADO"} )
)
```

| Campo | Tipo | Descrição |
|---|---|---|
| EMPRESA | Texto | Chave primária |
| EMPRESA (SAP) | Texto | Código SAP |
| MERCADO SELECIONADO | Texto | Mercado Mercer de referência |

---

### Tabela Dimensão: `DimFamília`
```m
DimFamília = Table.Distinct(
    Table.SelectColumns( Cargos, {"FAMÍLIA GBS", "SUB FAMÍLIA", "GRUPO"} )
)
```

| Campo | Tipo | Descrição |
|---|---|---|
| FAMÍLIA GBS | Texto | Chave primária |
| SUB FAMÍLIA | Texto | Subdivisão da família |
| GRUPO | Texto | Corporativo / Negócios & operações |

---

### Tabela Dimensão: `DimGrade`
Criar como tabela calculada no Power BI ou via Enter Data.

| Grade | GradeCategory | Descrição |
|---|---|---|
| 2 | Low (2-5) | Entrada operacional |
| 3 | Low (2-5) | Operacional |
| 4 | Low (2-5) | Operacional avançado |
| 5 | Low (2-5) | Técnico/Assistente |
| 6 | Mid (6-9) | Analista Jr/Assistente especializado |
| 7 | Mid (6-9) | Analista Pl |
| 8 | Mid (6-9) | Analista Sr |
| 9 | Mid (6-9) | Supervisor/Coordenador |
| 10 | Senior (10-13) | Coordenador Sênior |
| 11 | Senior (10-13) | Gerente |
| 12 | Senior (10-13) | Gerente Sênior |
| 13 | Senior (10-13) | Diretor/Gerente Executivo |

---

### Tabela Dimensão: `DimHierarquia`
```m
DimHierarquia = Table.FromRecords({
    [GRUPO_HIERARQUICO="GA",           HierarchyOrder=1, Descricao="Gerente de Área"],
    [GRUPO_HIERARQUICO="GS",           HierarchyOrder=2, Descricao="Gerente de Seção"],
    [GRUPO_HIERARQUICO="Coordenador",  HierarchyOrder=3, Descricao="Coordenador"],
    [GRUPO_HIERARQUICO="Supervisor",   HierarchyOrder=4, Descricao="Supervisor"],
    [GRUPO_HIERARQUICO="Especialista", HierarchyOrder=5, Descricao="Especialista"],
    [GRUPO_HIERARQUICO="Analista",     HierarchyOrder=6, Descricao="Analista"],
    [GRUPO_HIERARQUICO="Assistente",   HierarchyOrder=7, Descricao="Assistente"],
    [GRUPO_HIERARQUICO="Auxiliar",     HierarchyOrder=8, Descricao="Auxiliar"]
})
```

---

### Tabela de Datas: `DimData`
Gerada pelo Power Query (ver `PowerQuery_Transform.m`, seção `DimData`).
Usar para slicers de vigência quando dados históricos forem incorporados.

---

## 2. Relacionamentos Recomendados

```
Cargos[EMPRESA]            → DimEmpresa[EMPRESA]           (Many-to-One, ativo)
Cargos[FAMÍLIA GBS]        → DimFamília[FAMÍLIA GBS]        (Many-to-One, ativo)
Cargos[GRADE]              → DimGrade[Grade]                (Many-to-One, ativo)
Cargos[GRUPO HIERARQUICO]  → DimHierarquia[GRUPO_HIERARQUICO] (Many-to-One, ativo)
```

> **Nota:** Para dados salariais futuros, adicione `TabSalarios[CÓDIGO CARGO]` → `Cargos[CÓDIGO CARGO]` com cardinalidade Many-to-One.

---

## 3. Páginas do Relatório e Configuração de Visuais

---

### Página 1: Visão Executiva

**Objetivo:** Resumo executivo com os principais KPIs.

| Visual | Tipo | Campos / Configuração |
|---|---|---|
| KPI – Total Cargos | Card | Medida: `[Total Cargos]` |
| KPI – Cargos Ativos | Card | Medida: `[Total Cargos Ativos]`, subtítulo: `[% Cargos Ativos]` |
| KPI – Sem Mercer 2026 | Card | Medida: `[Cargos Sem Mercer 2026]` |
| KPI – Empresas | Card | Medida: `[Total Empresas]` |
| Donut – Ativo vs Inativo | Donut Chart | Legenda: `SITUAÇÃO`, Valores: `[Total Cargos]` |
| Donut – Corporativo vs Negócios | Donut Chart | Legenda: `GRUPO`, Valores: `[Total Cargos]` |
| Donut – Cobertura Mercer 2026 | Donut Chart | Legenda: `HasMercer2026`, Valores: `[Total Cargos]` |
| Tabela Resumo por Empresa | Table | Linhas: `EMPRESA`, Valores: `[Total Cargos]`, `[Total Cargos Ativos]`, `[% Cargos Ativos]`, `[Grades Distintos por Empresa]` |
| Slicer – GRUPO | Slicer | Campo: `Cargos[GRUPO]` |
| Slicer – EMPRESA | Slicer (dropdown) | Campo: `DimEmpresa[EMPRESA]` |

---

### Página 2: Análise de Grade e Remuneração

**Objetivo:** Distribuição dos cargos por grade e (futuramente) análise salarial.

| Visual | Tipo | Campos / Configuração |
|---|---|---|
| Barras Agrupadas – Cargos por Grade | Clustered Bar | Eixo Y: `GRADE`, Valores: `[Total Cargos Ativos]`, `[Total Cargos Inativos]`. Ordenar por `GRADE` crescente. |
| Barras – Cargos por GradeCategory | Stacked Bar | Eixo Y: `GradeCategory`, Valores: `[Total Cargos]` |
| Tabela – Grade Detail | Table | Linhas: `GRADE`, `GradeCategory`, Valores: `[Total Cargos]`, `[Total Cargos Ativos]`, `[% Cargos Ativos]`, `[% Cobertura Mercer 2026]` |
| Scatter Plot (futuro) | Scatter Chart | Eixo X: `MIDPOINT` (TabSalarios), Eixo Y: `[Compa-Ratio por Grade]`, Detalhes: `CARGO`. Adicionar linha de referência em Y=1.0. |
| Gauge – Meta Cobertura Mercer | Gauge | Valor: `[% Cobertura Mercer 2026]`, Máx: `[Gauge Max Mercer]`, Meta: `[Gauge Meta Mercer]` |
| Slicer – GradeCategory | Slicer | Campo: `Cargos[GradeCategory]` |

---

### Página 3: Estrutura Organizacional

**Objetivo:** Visualizar a estrutura de famílias, subfamílias e cargos por empresa.

| Visual | Tipo | Campos / Configuração |
|---|---|---|
| Treemap – Família / Subfamília | Treemap | Categoria: `FAMÍLIA GBS`, Detalhes: `SUB FAMÍLIA`, Valores: `[Total Cargos Ativos]`. Cores por `GRUPO`. |
| Matriz – Empresa × Hierarquia | Matrix | Linhas: `EMPRESA`, Colunas: `GRUPO HIERARQUICO` (ordenar por `HierarchyOrder`), Valores: `[Total Cargos Ativos]`. Habilitar subtotais. |
| Barras – Top Subfamílias | Bar Chart | Eixo: `SUB FAMÍLIA`, Valores: `[Total Cargos]`. Filtro Top N = 10. |
| Tabela – Detalhe Família | Table | Linhas: `FAMÍLIA GBS`, `SUB FAMÍLIA`, Valores: `[Total Cargos]`, `[Total Cargos Ativos]`, `[% Cargos Ativos]` |
| Slicer – FAMÍLIA GBS | Slicer | Campo: `DimFamília[FAMÍLIA GBS]` |
| Slicer – GRUPO | Slicer | Campo: `Cargos[GRUPO]` |

---

### Página 4: Benchmarking Mercer

**Objetivo:** Análise de cobertura e lacunas do benchmark Mercer 2026.

| Visual | Tipo | Campos / Configuração |
|---|---|---|
| Gauge – % Cobertura 2026 | Gauge | Valor: `[% Cobertura Mercer 2026]`, Meta: `[Gauge Meta Mercer]` = 95% |
| KPI – Cargos com Mercer 2025 | Card | Medida: `[Cargos Com Mercer 2025]` |
| KPI – Cargos com Mercer 2026 | Card | Medida: `[Cargos Com Mercer 2026]` |
| KPI – Gap 2025→2026 | Card | Medida: `[Gap Mercer 2026 vs 2025]`. Conditional formatting: vermelho se > 0. |
| Tabela – Gaps por Empresa | Table | Linhas: `EMPRESA`, Valores: `[Cargos Com Mercer 2026]`, `[Cargos Sem Mercer 2026]`, `[% Cobertura Mercer 2026]`. Ordenar por `[% Cobertura Mercer 2026]` crescente. |
| Tabela – Lista de Cargos sem Mercer 2026 | Table | Colunas: `CARGO`, `EMPRESA`, `GRADE`, `FAMÍLIA GBS`, `SITUAÇÃO`. Filtro: `HasMercer2026 = FALSE`. |
| Barras – Cobertura por Família | Bar Chart | Eixo: `FAMÍLIA GBS`, Valores: `[% Cobertura Mercer 2026]`. Linha de referência em 95%. |
| Slicer – EMPRESA | Slicer | Campo: `DimEmpresa[EMPRESA]` |
| Slicer – FAMÍLIA GBS | Slicer | Campo: `DimFamília[FAMÍLIA GBS]` |

---

### Página 5: Pirâmide de Cargos

**Objetivo:** Visualizar a distribuição hierárquica dos cargos (pirâmide organizacional).

| Visual | Tipo | Campos / Configuração |
|---|---|---|
| Funnel / Barras Horizontais | Bar Chart (horizontal) | Eixo Y: `GRUPO HIERARQUICO` (ordenar por `HierarchyOrder` crescente), Valores: `[Total Cargos Ativos]`. Habilitar rótulos de dados. |
| Tabela – Hierarquia Detalhada | Table | Linhas: `GRUPO HIERARQUICO`, Valores: `[Total Cargos Ativos]`, `[Total Cargos Inativos]`, `[Total Cargos]`, `[% Hierarquia sobre Total Ativo]` |
| Barras – Hierarquia por Empresa | Stacked Bar | Eixo X: `EMPRESA`, Legenda: `GRUPO HIERARQUICO`, Valores: `[Total Cargos Ativos]` |
| Matriz – Hierarquia × Grade | Matrix | Linhas: `GRUPO HIERARQUICO` (ordenar por HierarchyOrder), Colunas: `GRADE`, Valores: `[Total Cargos Ativos]` |
| Slicer – EMPRESA | Slicer | Campo: `DimEmpresa[EMPRESA]` |
| Slicer – FAMÍLIA GBS | Slicer | Campo: `DimFamília[FAMÍLIA GBS]` |

---

## 4. Configurações de Tema e Design

### Paleta de Cores Recomendada
```json
{
  "name": "Cargos GBS",
  "dataColors": [
    "#1F3864",
    "#2E75B6",
    "#70AD47",
    "#FFC000",
    "#BDD7EE",
    "#ED7D31",
    "#A9D18E",
    "#FFE699"
  ],
  "background": "#FFFFFF",
  "foreground": "#1F3864",
  "tableAccent": "#2E75B6"
}
```

### Boas Práticas
- Usar **filtros de página** em vez de slicers quando o filtro se aplica a todos os visuais
- Habilitar **drill-through** de Empresa → detalhe de cargos daquela empresa
- Adicionar **tooltips customizados** na pirâmide hierárquica com `DESCRIÇÃO SUMÁRIA`
- Usar **conditional formatting** nas tabelas: % cobertura Mercer < 80% em vermelho, 80-95% em amarelo, > 95% em verde

---

## 5. Próximos Passos para Expansão do Modelo

1. **Conectar dados salariais** (folha de pagamento ou faixa salarial por grade) para calcular Compa-Ratio
2. **Adicionar dimensão de tempo** (histórico de criação/inativação de cargos) usando a tabela `DimData`
3. **Integrar headcount** (número de colaboradores por cargo) para calcular taxa de utilização das posições
4. **Conectar pesquisa Mercer** com os valores de mercado (P25, P50, P75) por código Mercer para análise de posicionamento salarial
