# Livro caixa do ERP pronto para a planilha (demonstração)

Versão de demonstração de uma automação que uso em um cliente real. O projeto, com o problema, o resultado e um vídeo de 3 minutos, está no portfólio:
https://dkdian.github.io/projeto.html?p=formatacao-livro-caixa

**Todos os dados, nomes, valores e contas aqui são fictícios.** A demonstração mostra uma rotina de fechamento: cinco relatórios de livro caixa exportados por um ERP
(texto de largura fixa dentro de uma coluna do .xls) viram uma tabela classificada, no layout da planilha de controle.

- `powerquery/`: as consultas em linguagem M (`Base_LivroCaixa` lê e interpreta os .xls; `Query1` aplica os de-para e
  entrega as colunas da aba mensal; `Novas_Contrapartidas` lista o que ainda não tem classificação; `Src_*` leem as tabelas de apoio).
- `vba/`: macros `Saldo` e `AdicionarLinhas` da planilha de controle.
- `exemplo_relatorios_erp/`: os cinco relatórios fictícios de um mês, no mesmo layout do relatório real.
- `gerador/gen_erp.py`: gera relatórios fictícios nesse layout (Python + openpyxl).

Todos os nomes, valores e contas são inventados. A pasta lida pela consulta `Base_LivroCaixa` fica na variável `PastaAtual`.

## Como usar

```bash
pip install openpyxl
python gerador/gen_erp.py        # gera relatórios fictícios no layout do ERP (a conversão para .xls é feita pelo Excel)
```

No Excel, crie as consultas de `powerquery/` (cada arquivo `.m` é uma consulta), aponte `PastaAtual` para a pasta com os cinco relatórios do mês e use
*Dados > Atualizar Tudo*. As macros de `vba/` vão no módulo da planilha de controle.
