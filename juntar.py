from pathlib import Path
import pandas as pd
from formatar import formatar_relatorio  # NOVO

pasta = Path("entrada")
tabelas = []

for arquivo in sorted(pasta.glob("*.xlsx")):
    tabela = pd.read_excel(arquivo)
    loja = arquivo.stem.replace("vendas_", "").title()
    tabela["Loja"] = loja
    tabelas.append(tabela)
    print(f"{arquivo.name}: {len(tabela)} linhas")

tudo = pd.concat(tabelas, ignore_index=True)
tudo["Total"] = (tudo["Quantidade"] * tudo["Preço Unitário"]).round(2)

print()
print(tudo.head())
print()
print("Total geral de linhas:", len(tudo))
print(f"Faturamento total: R$ {tudo['Total'].sum():.2f}")

print()
print("=== Por loja ===")
por_loja = (
    tudo.groupby("Loja", as_index=False)["Total"].sum()
    .round(2)
    .sort_values("Total", ascending=False)
)
print(por_loja.to_string(index=False))

print()
print("=== Por produto ===")
por_produto = (
    tudo.groupby("Produto", as_index=False)
    .agg(Quantidade=("Quantidade", "sum"), Total=("Total", "sum"))
    .round(2)
    .sort_values("Total", ascending=False)
)
print(por_produto.to_string(index=False))

Path("saida").mkdir(exist_ok=True)

with pd.ExcelWriter("saida/relatorio_vendas.xlsx") as excel:
    por_loja.to_excel(excel, sheet_name="Por Loja", index=False)
    por_produto.to_excel(excel, sheet_name="Por Produto", index=False)
    tudo.to_excel(excel, sheet_name="Todas as Vendas", index=False)

print()
print("Relatório salvo em saida/relatorio_vendas.xlsx")

formatar_relatorio("saida/relatorio_vendas.xlsx")  # NOVO
print("Formatação aplicada.")                       # NOVO

