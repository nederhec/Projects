// ============================================================
// Power Query M – Transformação de Dados: Cargos GBS
// Arquivo: PowerQuery_Transform.m
// Descrição: Carrega e transforma a planilha Excel de Cargos GBS,
//            criando tabelas limpas, colunas calculadas e tabela
//            de datas para uso futuro no modelo Power BI.
// ============================================================

// ────────────────────────────────────────────────────────────
// 1. TABELA PRINCIPAL: Cargos
// ────────────────────────────────────────────────────────────
let
    // --- Fonte de dados ---
    // Ajuste o caminho do arquivo conforme necessário
    Source = Excel.Workbook(
        File.Contents("C:\Dados\Cargos_GBS.xlsx"),
        null,
        true
    ),

    // Seleciona a primeira planilha (Sheet1 ou a aba de dados)
    Sheet = Source{[Item="Sheet1", Kind="Sheet"]}[Data],

    // Promove a primeira linha como cabeçalhos
    PromotedHeaders = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),

    // --- Limpeza de tipos ---
    TypedTable = Table.TransformColumnTypes(
        PromotedHeaders,
        {
            {"EMPRESA",              type text},
            {"EMPRESA (SAP)",        type text},
            {"CÓDIGO CARGO",         Int64.Type},
            {"CARGO",                type text},
            {"GRADE",                Int64.Type},
            {"SITUAÇÃO",             type text},
            {"GRUPO",                type text},
            {"FAMÍLIA GBS",          type text},
            {"SUB FAMÍLIA",          type text},
            {"GRUPO HIERARQUICO",    type text},
            {"NÍVEL",                type text},
            {"CBO",                  type text},
            {"REGRA CONTROLE JORNADA", type text},
            {"CONTROLE JORNADA",     type text},
            {"COTA APRENDIZ",        type text},
            {"TIPO CONTRATO",        type text},
            {"SINDICATO",            type text},
            {"MERCADO SELECIONADO",  type text},
            {"DESCRIÇÃO SUMÁRIA",    type text},
            {"Cod. Mercer 2025",     type text},
            {"Cod. Mercer 2026",     type text}
        }
    ),

    // --- Limpeza de espaços em branco em colunas-chave ---
    TrimmedColumns = Table.TransformColumns(
        TypedTable,
        {
            {"EMPRESA",           Text.Trim},
            {"SITUAÇÃO",          Text.Trim},
            {"GRUPO",             Text.Trim},
            {"FAMÍLIA GBS",       Text.Trim},
            {"GRUPO HIERARQUICO", Text.Trim},
            {"NÍVEL",             Text.Trim}
        }
    ),

    // ── COLUNA CALCULADA 1: IsActive ─────────────────────────
    // Boolean: true se SITUAÇÃO = "ATIVO"
    AddIsActive = Table.AddColumn(
        TrimmedColumns,
        "IsActive",
        each [SITUAÇÃO] = "ATIVO",
        type logical
    ),

    // ── COLUNA CALCULADA 2: HierarchyOrder ───────────────────
    // Ordem numérica para GRUPO HIERARQUICO (topo = 1, base = 8)
    AddHierarchyOrder = Table.AddColumn(
        AddIsActive,
        "HierarchyOrder",
        each
            if [GRUPO HIERARQUICO] = "GA"           then 1
            else if [GRUPO HIERARQUICO] = "GS"      then 2
            else if [GRUPO HIERARQUICO] = "Coordenador" then 3
            else if [GRUPO HIERARQUICO] = "Supervisor"  then 4
            else if [GRUPO HIERARQUICO] = "Especialista" then 5
            else if [GRUPO HIERARQUICO] = "Analista"    then 6
            else if [GRUPO HIERARQUICO] = "Assistente"  then 7
            else if [GRUPO HIERARQUICO] = "Auxiliar"    then 8
            else 99,
        Int64.Type
    ),

    // ── COLUNA CALCULADA 3: HasMercer2026 ────────────────────
    // Boolean: true se Cod. Mercer 2026 não estiver vazio/nulo
    AddHasMercer2026 = Table.AddColumn(
        AddHierarchyOrder,
        "HasMercer2026",
        each
            [Cod. Mercer 2026] <> null and
            Text.Trim([Cod. Mercer 2026]) <> "",
        type logical
    ),

    // ── COLUNA CALCULADA 4: GradeCategory ────────────────────
    // Categorias: Low (2-5), Mid (6-9), Senior (10-13)
    AddGradeCategory = Table.AddColumn(
        AddHasMercer2026,
        "GradeCategory",
        each
            if [GRADE] >= 2  and [GRADE] <= 5  then "Low (2-5)"
            else if [GRADE] >= 6  and [GRADE] <= 9  then "Mid (6-9)"
            else if [GRADE] >= 10 and [GRADE] <= 13 then "Senior (10-13)"
            else "Fora do padrão",
        type text
    ),

    // --- Ordenação final ---
    SortedTable = Table.Sort(
        AddGradeCategory,
        {{"EMPRESA", Order.Ascending}, {"GRADE", Order.Ascending}}
    )
in
    SortedTable,


// ────────────────────────────────────────────────────────────
// 2. TABELA DE DATAS (DimData)
// Para uso futuro com dados de vigência ou histórico salarial
// ────────────────────────────────────────────────────────────
DimData =
let
    StartDate = #date(2020, 1, 1),
    EndDate   = #date(2030, 12, 31),

    // Gera lista de datas
    DateList = List.Dates(
        StartDate,
        Duration.Days(EndDate - StartDate) + 1,
        #duration(1, 0, 0, 0)
    ),
    DateTable = Table.FromList(DateList, Splitter.SplitByNothing()),
    RenamedCol = Table.RenameColumns(DateTable, {{"Column1", "Date"}}),
    TypedDate  = Table.TransformColumnTypes(RenamedCol, {{"Date", type date}}),

    // Colunas derivadas de data
    AddYear    = Table.AddColumn(TypedDate,   "Year",    each Date.Year([Date]),   Int64.Type),
    AddMonth   = Table.AddColumn(AddYear,     "Month",   each Date.Month([Date]),  Int64.Type),
    AddMonthName = Table.AddColumn(AddMonth,  "MonthName", each Date.ToText([Date], "MMM", "pt-BR"), type text),
    AddQuarter = Table.AddColumn(AddMonthName,"Quarter", each "T" & Text.From(Date.QuarterOfYear([Date])), type text),
    AddDay     = Table.AddColumn(AddQuarter,  "Day",     each Date.Day([Date]),    Int64.Type),
    AddWeekday = Table.AddColumn(AddDay,      "Weekday", each Date.DayOfWeekName([Date], "pt-BR"), type text),
    AddIsWeekend = Table.AddColumn(
        AddWeekday,
        "IsWeekend",
        each Date.DayOfWeek([Date]) >= 5,
        type logical
    ),
    AddYearMonth = Table.AddColumn(
        AddIsWeekend,
        "YearMonth",
        each Text.From([Year]) & "-" & Text.PadStart(Text.From([Month]), 2, "0"),
        type text
    )
in
    AddYearMonth,


// ────────────────────────────────────────────────────────────
// 3. PARÂMETRO: Caminho do Arquivo Fonte
// Facilita a troca do caminho sem editar a query principal
// ────────────────────────────────────────────────────────────
SourceFilePath =
let
    Value = "C:\Dados\Cargos_GBS.xlsx"   // <-- Altere aqui se necessário
in
    Value
