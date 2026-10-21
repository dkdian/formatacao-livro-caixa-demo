let
    PrefixoTransf = "5101",
    Chave  = (t as any) as text => Text.Upper(Text.Trim(if t = null then "" else Text.From(t))),
    Base   = Base_LivroCaixa,
    Cats   = Table.Buffer(Src_DePara),
    Regras = Table.Buffer(Table.SelectRows(Src_Regras, each Chave([Palavra]) <> "")),

    Conhecidas = List.Buffer(List.Transform(Cats[Contrapartida], Chave)),
    Norm  = Table.AddColumn(Base, "Contra", each
                if Chave([Contrapartida]) <> "" then Chave([Contrapartida])
                else "SEM DESCRICAO " & ([CodCP2] ?? ""), type text),
    Novas = Table.SelectRows(Norm, each not List.Contains(Conhecidas, [Contra])),

    Grp = Table.Group(Novas, {"Contra"}, {
              {"Ocorrências (2 meses)", each Table.RowCount(_), Int64.Type},
              {"Entradas", each List.Sum(List.Select([Valor], (v) => v > 0)) ?? 0, type number},
              {"Saídas",   each -(List.Sum(List.Select([Valor], (v) => v < 0)) ?? 0), type number},
              {"CodCP",    each List.First(List.Select([CodCP2], (c) => c <> null and c <> ""), ""), type text},
              {"Contas",   each Text.Combine(List.Distinct([ContaCodigo]), ", "), type text},
              {"Exemplo",  each List.First([Historico]), type text}}),
    Dir = Table.AddColumn(Grp, "Direção", each
              if [Entradas] > 0 and [Saídas] = 0 then "Entrada"
              else if [Saídas] > 0 and [Entradas] = 0 then "Saída" else "Misto", type text),

    // sufixo (texto após a última "/") -> Tipo/Categoria mais frequente já no tbl_DePara
    ComSuf = Table.SelectRows(
                 Table.AddColumn(Cats, "Suf", each
                     let k = Chave([Contrapartida])
                     in  if Text.Contains(k, "/") then Text.Trim(List.Last(Text.Split(k, "/"))) else null),
                 each [Suf] <> null and [Suf] <> "" and Chave([Tipo]) <> "" and Chave([Categoria]) <> ""),
    SufTab = Table.Buffer(Table.Group(ComSuf, {"Suf"}, {{"Top", each
                 Table.Sort(Table.Group(_, {"Tipo", "Categoria"}, {{"n", each Table.RowCount(_), Int64.Type}}),
                            {{"n", Order.Descending}}){0}}})),

    Sugere = (r as record) as record =>
        let
            k      = r[Contra],
            suf    = if Text.Contains(k, "/") then Text.Trim(List.Last(Text.Split(k, "/"))) else null,
            transf = r[CodCP] <> "" and Text.StartsWith(r[CodCP], PrefixoTransf),
            hitR   = Table.SelectRows(Regras, each Text.Contains(k, Chave([Palavra]))
                         and (Chave([Direção]) = "" or Chave([Direção]) = Chave(r[Direção]))),
            hitS   = Table.SelectRows(SufTab, each [Suf] = suf)
        in
            if transf and r[Direção] <> "Misto" then
                [Tipo = "Movimentação", Categoria = r[Direção], Origem = "código " & PrefixoTransf]
            else if Table.RowCount(hitR) > 0 then
                [Tipo = hitR{0}[Tipo], Categoria = hitR{0}[Categoria], Origem = "regra: " & Text.From(hitR{0}[Palavra])]
            else if Table.RowCount(hitS) > 0 then
                [Tipo = hitS{0}[Top][Tipo], Categoria = hitS{0}[Top][Categoria], Origem = "sufixo /" & suf]
            else
                [Tipo = null, Categoria = null, Origem = "sem sugestão"],

    Sug = Table.ExpandRecordColumn(Table.AddColumn(Dir, "S", each Sugere(_)), "S", {"Tipo", "Categoria", "Origem"}),
    St  = Table.AddColumn(Sug, "Status", each "Revisar", type text),
    Alt = Table.AddColumn(St, "Alternativas vistas", each null, type nullable text),
    Obs = Table.AddColumn(Alt, "Observação", each "sugestão: " & [Origem], type text),
    Ren = Table.RenameColumns(Obs, {{"Contra", "Contrapartida"}}),
    Fim = Table.SelectColumns(Ren, {"Contrapartida", "Tipo", "Categoria", "Ocorrências (2 meses)", "Status",
              "Alternativas vistas", "Observação", "Direção", "Entradas", "Saídas", "Contas", "CodCP", "Exemplo"})
in
    Table.Sort(Fim, {{"Ocorrências (2 meses)", Order.Descending}, {"Contrapartida", Order.Ascending}})
