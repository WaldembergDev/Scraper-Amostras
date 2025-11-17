from src.automacao.automacao_gerencia_lab import AutomacaoAmostras
import time

if __name__ == "__main__":
    automacoes = [
        AutomacaoAmostras.iniciar_automacao_geral(),
        AutomacaoAmostras.iniciar_automacao_pier(),
        AutomacaoAmostras.iniciar_automacao_atrasados()
        ]
    for automacao in automacoes:
        automacao
        time.sleep(40)