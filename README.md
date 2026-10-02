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
- Agenda 161 = Cartão presente abastecimento (Postos); agendas 162 e 163 = Cartão Frota (um só produto; a coluna "Agenda / Portador" preserva a agenda de origem).
- Cheques: atraso = referência − VENCTO (DEVOL. tem datas inválidas, ignorada); todos entram como vencido.
- `Dados/` e `dashboard.html` NÃO são versionados (nominais: nomes/CPF/CNPJ de devedores).

## Pendências
- Criar repositório privado no GitHub (como KPIs/Carteira/Governança); decidir como proteger os dados nominais (não publicar o HTML gerado em Pages público).
- Migrar base de Excel para outra fonte (previsto).
