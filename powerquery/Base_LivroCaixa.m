let
    // ================== caminho da pasta Atual ==================
    PastaAtual = "C:\Demo\Livro caixa\Atual",
    // ============================================================
 
    // ---------- auxiliares ----------
    fnNum = (t as text) as nullable number =>
        let
            s     = Text.Trim(t),
            vazio = (s = "" or s = "-" or Text.Contains(s, "Cancelado")),
            neg   = not vazio and Text.EndsWith(s, "-"),
            s2    = if vazio then null else Text.Trim(Text.TrimEnd(s, "-")),
            s3    = if s2 = null then null else Text.Replace(Text.Replace(s2, ".", ""), ",", "."),
            v     = if s3 = null or s3 = "" then null else Number.FromText(s3, "en-US")
        in
            if v = null then null else (if neg then -v else v),
 
    fnUltimo = (lista as list) as text =>
        let nv = List.Select(lista, (x) => x <> "")
        in  if List.IsEmpty(nv) then "" else List.Last(nv),
 
    // ---------- le um relatorio .xls do sistema ----------
    fnArquivo = (conteudo as binary) as table =>
        let
            Livro    = Excel.Workbook(conteudo, false, true),
            Planilha = Livro{0}[Data],
            Cols     = Table.ColumnNames(Planilha),
            Base     = Table.SelectColumns(Planilha, {Cols{1}, Cols{3}}),
            Ren      = Table.RenameColumns(Base, {{Cols{1}, "L"}, {Cols{3}, "D"}}),
            Txt      = Table.TransformColumns(Ren, {
                            {"L", each if _ = null then "" else Text.From(_), type text},
                            {"D", each if _ = null then "" else Text.From(_), type text}}),
 
            // codigo da conta: linha do cabecalho, coluna D, termina com o periodo
            CabLinhas = Table.SelectRows(Txt, each
                            Text.Contains([D], " - ")
                            and Text.Length([D]) > 23
                            and Text.Length(Text.Select(Text.End(Text.TrimEnd([D]), 23), {"0".."9"})) = 16),
            ContaCod  = Text.End(Text.TrimEnd(Text.Split(Text.TrimEnd(CabLinhas{0}[D]), " - "){0}), 10),
 
            // bloco RESUMO: descricao completa da contrapartida (posicoes fixas)
            SemPipe  = Table.SelectRows(Txt, each not Text.Contains([L], "|") and Text.Length([L]) >= 49),
            SoCodigo = Table.SelectRows(SemPipe, each Text.Length(Text.Select(Text.Middle([L], 39, 10), {"0".."9"})) = 10),
            Res1     = Table.AddColumn(SoCodigo, "CodCP",  each Text.Middle([L], 39, 10), type text),
            Res2     = Table.AddColumn(Res1,     "DescCP", each Text.Trim(Text.Middle([L], 5, 34)), type text),
            Resumo   = Table.Distinct(
                           Table.SelectColumns(Table.SelectRows(Res2, each [DescCP] <> ""), {"CodCP", "DescCP"}),
                           {"CodCP"}),
 
            // linhas de detalhe
            Det   = Table.SelectRows(Txt, each Text.StartsWith([L], "   |")),
            Split = Table.AddColumn(Det, "P", each Text.Split([L], "|"), type list),
            Camp  = Table.SelectRows(Split, each List.Count([P]) >= 8),
 
            C1 = Table.AddColumn(Camp, "Lancto", each Text.Trim([P]{1}), type text),
            C2 = Table.AddColumn(C1, "HistRaw", each Text.Middle(Text.PadEnd([P]{2}, 43), 2, 40), type text),
            C3 = Table.AddColumn(C2, "DtTxt",   each Text.Trim([P]{3}), type text),
            C4 = Table.AddColumn(C3, "EntTxt",  each Text.Trim([P]{4}), type text),
            C5 = Table.AddColumn(C4, "SaiTxt",  each Text.Trim([P]{5}), type text),
            C6 = Table.AddColumn(C5, "CpTxt",   each Text.Trim([P]{7}), type text),
            C7 = Table.AddColumn(C6, "Marcador", each Text.Replace(Text.Trim([HistRaw]), " ", ""), type text),
 
            // fora: saldo anterior, saldo do dia, subtotal e a linha de fechamento
            SemSaldos = Table.SelectRows(C7, each
                            not Text.StartsWith([Marcador], "SaldoAnterior")
                            and not Text.StartsWith([Marcador], "Saldodo")
                            and not Text.StartsWith([Marcador], "---")
                            and not ([Lancto] = "" and [Marcador] = "" and [DtTxt] = "")),
 
            // linha quebrada: Lancamento em branco = continuacao da anterior
            ComIdx  = Table.AddIndexColumn(SemSaldos, "Idx", 0, 1, Int64.Type),
            ComGrp  = Table.AddColumn(ComIdx, "Grp", each if [Lancto] <> "" then [Idx] else null, Int64.Type),
            SemOrfa = Table.SelectRows(Table.FillDown(ComGrp, {"Grp"}), each [Grp] <> null),
 
            Agrup = Table.Group(SemOrfa, {"Grp"}, {
                        {"Historico", each Text.Trim(Text.Combine([HistRaw])), type text},
                        {"DtTxt",     each fnUltimo([DtTxt]), type text},
                        {"EntTxt",    each fnUltimo([EntTxt]), type text},
                        {"SaiTxt",    each fnUltimo([SaiTxt]), type text},
                        {"CpTxt",     each fnUltimo([CpTxt]), type text},
                        {"Cancelado", each List.AnyTrue(List.Transform([EntTxt] & [SaiTxt],
                                            (x) => Text.Contains(x, "Cancelado"))), type logical}
                    }, GroupKind.Local),
 
            T1 = Table.AddColumn(Agrup, "Data", each
                     if Text.Length([DtTxt]) = 10
                     then Date.FromText([DtTxt], [Format = "dd/MM/yyyy", Culture = "pt-BR"])
                     else null, type nullable date),
            T2 = Table.AddColumn(T1, "Entrada", each if [Cancelado] then 0 else (fnNum([EntTxt]) ?? 0), type number),
            T3 = Table.AddColumn(T2, "Saida",   each if [Cancelado] then 0 else (fnNum([SaiTxt]) ?? 0), type number),
            T4 = Table.AddColumn(T3, "Valor",   each Number.Round([Entrada] - [Saida], 2), type number),
            T5 = Table.AddColumn(T4, "CodCP2",  each
                     let c = Text.Start([CpTxt], 10)
                     in  if Text.Length(Text.Select(c, {"0".."9"})) = 10 then c else "", type text),
 
            J = Table.ExpandTableColumn(
                    Table.NestedJoin(T5, {"CodCP2"}, Resumo, {"CodCP"}, "r", JoinKind.LeftOuter),
                    "r", {"DescCP"}, {"Contrapartida"}),
            F = Table.AddColumn(J, "ContaCodigo", each ContaCod, type text)
        in
            Table.SelectColumns(F, {"ContaCodigo", "Data", "Historico", "Contrapartida", "Valor", "Cancelado", "CodCP2"}),
 
    // ---------- todos os .xls da pasta Atual ----------
    Arquivos = Folder.Files(PastaAtual),
    SoXls    = Table.SelectRows(Arquivos, each
                   Text.Lower([Extension]) = ".xls" and not Text.StartsWith([Name], "~$")),
    Lido     = Table.AddColumn(SoXls, "Dados", each fnArquivo([Content])),
    Base     = Table.ExpandTableColumn(Table.SelectColumns(Lido, {"Dados"}), "Dados",
                   {"ContaCodigo", "Data", "Historico", "Contrapartida", "Valor", "Cancelado", "CodCP2"}),
    Validos  = Table.SelectRows(Base, each [Data] <> null and [Cancelado] = false)
in
    Validos
