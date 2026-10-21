// Query1 = saida do Livro Caixa no formato da aba mensal do Controle Financeiro.
// Coluna "Fornecedor/Cliente" :
//   1) De-Para (tabela tbl_Fornecedores, aba DePara_Fornecedores): se uma Palavra aparece no Historico de uma SAIDA,
//      vale o Fornecedor da tabela (unifica variacoes de grafia e cobre linhas sem padrao). Primeira linha que casar.
//   2) Padrao do Historico: "<tipo> <doc> -<parcela> - <NOME>" (pagamentos) ou "RECB DUP ... nn/nn - <NOME>" (recebimentos).
//   3) Se a contrapartida for a categoria de fornecedores e nada acima achou nome -> "Não Encontrado".
//   Fora disso fica vazio (nao chuta).
let
    CategoriaFornecedor = "FORNECEDORES DIVERSOS",   // Detalhamento/Categoria II que alimenta o Pareto de fornecedores
    NaoEncontrado       = "Não Encontrado",

    Validos = Base_LivroCaixa,
    Contas  = Src_Contas,

    // ---------- fornecedor/cliente a partir do Historico ----------
    ehLetraOuNum = (c as text) as logical =>
        Text.Length(Text.Select(c, {"0".."9", "A".."Z"})) = Text.Length(c),
    // "...00354404 -2"  /  "...-A"  /  "...-12"
    temParcela = (p as text) as logical =>
        let tok = List.Last(Text.Split(Text.TrimEnd(p), " "))
        in  Text.Length(tok) >= 2 and Text.Length(tok) <= 3
            and Text.Start(tok, 1) = "-" and ehLetraOuNum(Text.Range(tok, 1)),
    // "...572405-26/01"
    temData = (p as text) as logical =>
        let q = Text.End(Text.TrimEnd(p), 5)
        in  Text.Length(q) = 5 and Text.Middle(q, 2, 1) = "/"
            and Text.Length(Text.Select(Text.Start(q, 2) & Text.End(q, 2), {"0".."9"})) = 4,
    primeiro = (partes as list, f as function) as nullable number =>
        let idx = List.Select(List.Positions(List.RemoveLastN(partes, 1)), each f(partes{_}))
        in  if List.IsEmpty(idx) then null else idx{0},
    nomePadrao = (h as nullable text) as nullable text =>
        let
            partes = Text.Split(h ?? "", " - "),
            i      = primeiro(partes, temParcela) ?? primeiro(partes, temData),
            nome   = if i = null then null
                     else Text.Trim(Text.Combine(List.Skip(partes, i + 1), " - "))
        in
            if nome = "" then null else nome,

    DeParaForn = List.Buffer(
                     Table.ToRecords(
                         Table.SelectRows(Src_Fornecedores, each
                             Text.Trim(Text.From([Palavra] ?? "")) <> "" and Text.Trim(Text.From([Fornecedor] ?? "")) <> ""))),
    nomeDePara = (h as nullable text) as nullable text =>
        let
            H   = Text.Upper(h ?? ""),
            hit = List.First(List.Select(DeParaForn, each Text.Contains(H, Text.Upper(Text.Trim(Text.From([Palavra])))))
                             , null)
        in
            if hit = null then null else Text.Trim(Text.From(hit[Fornecedor])),
    fnForn = (h as nullable text, contra as nullable text, valor as nullable number) as nullable text =>
        let
            doDePara = if (valor ?? 0) < 0 then nomeDePara(h) else null,
            nome     = doDePara ?? nomePadrao(h)
        in
            if nome <> null then nome
            else if Text.Upper(Text.Trim(contra ?? "")) = CategoriaFornecedor then NaoEncontrado
            else null,

    Cats    = Table.Distinct(
                  Table.SelectColumns(
                      Table.AddColumn(Src_DePara, "ChaveCat",
                          each Text.Upper(Text.Trim(Text.From([Contrapartida]) ?? "")), type text),
                      {"ChaveCat", "Tipo", "Categoria"}),
                  {"ChaveCat"}),
    ComConta = Table.ExpandTableColumn(
                   Table.NestedJoin(Validos, {"ContaCodigo"}, Contas, {"Código"}, "cc", JoinKind.LeftOuter),
                   "cc", {"Conta (Controle Financeiro)", "Desembolso"}, {"ContaCF", "Desembolso"}),
    Norm     = Table.AddColumn(ComConta, "ChaveCat",
                   each Text.Upper(Text.Trim([Contrapartida] ?? "")), type text),
    ComCat   = Table.ExpandTableColumn(
                   Table.NestedJoin(Norm, {"ChaveCat"}, Cats, {"ChaveCat"}, "dp", JoinKind.LeftOuter),
                   "dp", {"Tipo", "Categoria"}, {"Tipo", "Categoria"}),
    ComValor = Table.AddColumn(ComCat, "ValorCF",
                   each if [Tipo] = "Receita" then [Valor] else if [Tipo] = "Gasto" then -[Valor] else Number.Abs([Valor]), type number),
    V1 = Table.AddColumn(ComValor, "Venc.",              each null, type nullable date),
    V2 = Table.AddColumn(V1,       "C. C.",              each null, type nullable text),
    V3 = Table.AddColumn(V2,       "Fornecedor/Cliente", each fnForn([Historico], [Contrapartida], [Valor]), type nullable text),
    Sel = Table.SelectColumns(V3, {"Data", "Venc.", "C. C.", "Fornecedor/Cliente",
                                   "Historico", "Contrapartida", "Tipo", "Categoria",
                                   "Desembolso", "ContaCF", "ValorCF"}),
    Ren = Table.RenameColumns(Sel, {
              {"Data", "Realiz."},
              {"Historico", "Descrição"},
              {"Contrapartida", "Detalhamento/Categoria II"},
              {"ContaCF", "Conta"},
              {"ValorCF", "Valor"}}),
    Ord = Table.Sort(Ren, {{"Realiz.", Order.Ascending}, {"Conta", Order.Ascending}})
in
    Table.TransformColumnTypes(Ord, {{"Realiz.", type date}})
