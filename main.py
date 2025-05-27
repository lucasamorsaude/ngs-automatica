# main.py

import dados
import gerar_powerpoint
import converter_imagens
import enviar_slack

def main():
    try:
        print("Iniciando coleta de dados de Medicina...")
        dados.coletar_dados()
        print("✓ Medicina OK")
    except Exception as e:
        print(f"Erro na coleta de dados de Medicina: {e}")
        return


    try:
        print("Gerando apresentação PowerPoint...")
        gerar_powerpoint.gerar()
        print("✓ PowerPoint OK")
    except Exception as e:
        print(f"Erro na geração do PowerPoint: {e}")
        return

    try:
        print("Convertendo slides em imagens...")
        converter_imagens.converter()
        print("✓ Conversão OK")
    except Exception as e:
        print(f"Erro na conversão de slides: {e}")
        return

    try:
        print("Enviando imagens para o Slack...")
        enviar_slack.enviar()
        print("✓ Slack OK")
    except Exception as e:
        print(f"Erro ao enviar para o Slack: {e}")
        return

    print("🎉 Processo finalizado!")

if __name__ == "__main__":
    main()
