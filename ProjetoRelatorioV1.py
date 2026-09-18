import os
from datetime import datetime


def criar_relatorio_semanal():
    print("=== GERADOR DE RELATÓRIO SEMANAL ===")
    
    semana = input("Qual é o número da semana ou período (ex: Semana 10): ")
    atividades = input("Quais foram as principais atividades realizadas?\n> ")
    dificuldades = input("Houve algum desafio ou dificuldade?\n> ")
    proximos_passos = input("Quais são os planos para a próxima semana?\n> ")

    pasta_raiz = "Estágio"
    if not os.path.exists(pasta_raiz):
        os.makedirs(pasta_raiz)

    data_atual = datetime.now().strftime("%Y-%m-%d")
    nome_arquivo = f"relatorio_{semana.lower().replace(' ', '_')}_{data_atual}.txt"
    caminho_completo = os.path.join(pasta_raiz, nome_arquivo)

    conteudo = f"""==================================================
RELATÓRIO DE ESTÁGIO - {semana.upper()}
Data de geração: {datetime.now().strftime("%d/%m/%Y %H:%M")}
==================================================

1. PRINCIPAIS ATIVIDADES REALIZADAS:
{atividades}

2. DESAFIOS E DIFICULDADES:
{dificuldades}

3. PRÓXIMOS PASSOS / PLANEJAMENTO:
{proximos_passos}

==================================================
"""

    with open(caminho_completo, "w", encoding="utf-8") as arquivo:
        arquivo.write(conteudo)

    print(f"\nSucesso! Relatório salvo em: {caminho_completo}")


if __name__ == "__main__":
    criar_relatorio_semanal()
