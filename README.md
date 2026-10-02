# Inadimplência — Produtos PJ Condor

Dashboard de carteira vencida / a vencer por agenda + cheques pré-datados.

## Atualizar
```
python build.py            # dashboard.html: versão completa, nominal, SÓ LOCAL
python build.py --publico  # index.html: sem nome/CPF/CNPJ/NF, vai pro GitHub Pages
python build.py "Dados/Histórico 2026/Outubro"
```
Gera `dashboard.html` (autônomo, abrir no navegador; precisa de internet p/ as libs via CDN). O build confere os cheques contra a aba RESUMO PORTADOR.

## Regras
- Data de referência = dia anterior à extração (inferida de vencimento + atraso).
- Agenda 161 existe nos dados mas não tem tradução — rotulada "Agenda 161 (sem descrição)".
- Cheques: atraso = referência − VENCTO (DEVOL. tem datas inválidas, ignorada); todos entram como vencido.
- `Dados/` e `dashboard.html` NÃO são versionados (nominais: nomes/CPF/CNPJ de devedores).

## Pendências
- Criar repositório privado no GitHub (como KPIs/Carteira/Governança); decidir como proteger os dados nominais (não publicar o HTML gerado em Pages público).
- Confirmar nome da agenda 161.
- Migrar base de Excel para outra fonte (previsto).
