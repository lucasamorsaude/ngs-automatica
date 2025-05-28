from PIL import Image, ImageDraw, ImageFont
import os
import pandas as pd

def formatar_para_contabil(valor_str):
    """
    Formata uma string numérica para o formato contábil brasileiro (R$ X.XXX,XX).
    Remove o prefixo 'R$' e sufixos como '%' antes da conversão,
    e trata corretamente separadores de milhares e decimais.
    """
    if not isinstance(valor_str, str):
        valor_str = str(valor_str)

    # Remove R$, espaços em branco, e tenta limpar o sufixo '%' se presente.
    valor_limpo = valor_str.replace("R$", "").strip().replace("%", "")

    # Se o valor for "N/A" ou vazio após a limpeza, retorna-o como está.
    if not valor_limpo or valor_limpo.upper() == "N/A":
        return valor_str

    try:
        # Tenta converter para float, tratando vírgula como decimal se houver.
        if ',' in valor_limpo and '.' not in valor_limpo:
            numero_float = float(valor_limpo.replace('.', '').replace(',', '.'))
        elif ',' in valor_limpo and '.' in valor_limpo:
            partes_decimais = valor_limpo.split(',')
            if len(partes_decimais[-1]) == 2:
                numero_float = float(valor_limpo.replace('.', '').replace(',', '.'))
            else:
                numero_float = float(valor_limpo.replace(',', ''))
        else:
            numero_float = float(valor_limpo)

        # Formata para o padrão brasileiro: separador de milhar '.' e decimal ','
        return "R$ {:,.2f}".format(numero_float).replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        # Se a conversão falhar (ex: texto que não é número), retorna o valor original.
        return valor_str


def ler_dados_da_planilha(file_path):
    """
    Lê os dados da planilha e os converte para o dicionário esperado pela função de imagem,
    aplicando formatação contábil onde necessário.
    """
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            print("Erro: Formato de arquivo não suportado. Use .csv ou .xlsx")
            return None

        df['Valor'] = df['Valor'].astype(str)
        dados_brutos = df.set_index('Indicador')['Valor'].to_dict()

        dados_formatados = {
            "agendamentos_valor": dados_brutos.get("NUM_AGENDAMENTOS", "N/A"),
            "qca_valor": dados_brutos.get("NUM_QCA", "N/A"),
            "consultas_valor": dados_brutos.get("NUM_CONSULTAS", "N/A"),
            "conversao_valor": dados_brutos.get("CONVERSAO", "N/A") + "%",
            "faturamento_valor": dados_brutos.get("R$_FATURAMENTO_MED", "N/A"),
            "ticket_medio_valor": dados_brutos.get("R$_TM_MEDIO_MEDICINA", "N/A"),
            "exames_lab_valor": dados_brutos.get("R$_EXAMES_LABORATORIAIS", "N/A"),
            "tm_exames_lab_valor": dados_brutos.get("TM_EXAMES_LABORATORIAIS", "N/A"), # Incluído na lista de campos
            "exames_imagem_valor": dados_brutos.get("R$_CAIXA_TOTAL", "N/A"),
            "procedimentos_valor": dados_brutos.get("NOVOS_PACIENTES", "N/A")
        }

        # Dicionário de mapeamento das chaves do dicionário 'dados_formatados'
        # para as chaves originais da planilha que contêm 'R$' e devem ser formatadas.
        campos_para_formatar_contabil = {
            "faturamento_valor": "R$_FATURAMENTO_MED",
            "ticket_medio_valor": "R$_TM_MEDIO_MEDICINA",
            "exames_lab_valor": "R$_EXAMES_LABORATORIAIS",
            "exames_imagem_valor": "R$_CAIXA_TOTAL",
            "tm_exames_lab_valor": "TM_EXAMES_LABORATORIAIS" # Agora incluído aqui
        }

        for key_formatado, key_original_planilha in campos_para_formatar_contabil.items():
            valor = dados_formatados.get(key_formatado)
            if valor is not None:
                # Chama a função de formatação contábil para o valor
                dados_formatados[key_formatado] = formatar_para_contabil(valor)

        return dados_formatados

    except Exception as e:
        print(f"Erro ao ler ou processar a planilha: {e}")
        return None

def preencher_template_medicina(data_to_fill, template_image_path="template_medicina.png", output_image_path="relatorio_medicina_preenchido.png"):
    """
    Carrega o template "Resultados Medicina" e preenche os dados dinâmicos (valores principais).
    """
    try:
        img = Image.open(template_image_path).convert("RGB")
        print(f"Template '{template_image_path}' carregado com sucesso.")
    except FileNotFoundError:
        print(f"Erro: Imagem de template não encontrada em '{template_image_path}'.")
        print("Certifique-se de que a imagem 'template_medicina.png' está no mesmo diretório do script.")

        print("Criando um template.png dummy para demonstração. Por favor, substitua-o pelo seu template real.")
        img = Image.new('RGB', (800, 600), color=(230, 240, 255))
        draw_dummy = ImageDraw.Draw(img)
        try:
            dummy_font = ImageFont.truetype("arial.ttf", 60)
        except IOError:
            dummy_font = ImageFont.load_default()
        draw_dummy.text((100, 250), "SEU TEMPLATE AQUI", fill=(0,0,0), font=dummy_font)
        img.save(template_image_path)
        img = Image.open(template_image_path).convert("RGB")

    draw = ImageDraw.Draw(img)

    try:
        font_path_bold = "arialbd.ttf"
        font_valor = ImageFont.truetype(font_path_bold, 40)
        print(f"Fonte negrito '{font_path_bold}' carregada com sucesso.")
    except IOError:
        print(f"Aviso: Fonte negrito '{font_path_bold}' não encontrada. Tentando fonte regular ou padrão.")
        try:
            font_path_regular = "arial.ttf"
            font_valor = ImageFont.truetype(font_path_regular, 40)
            print(f"Aviso: Usando fonte regular '{font_path_regular}'.")
        except IOError:
            print("Aviso: Nenhuma fonte Arial encontrada. Usando fonte padrão da Pillow.")
            font_valor = ImageFont.load_default()

    campos_para_preencher = {
        "consultas_valor": {"text": data_to_fill.get("consultas_valor", "N/A"), "position": (200, 400), "color": (0, 0, 0), "font": font_valor},
        "qca_valor": {"text": data_to_fill.get("qca_valor", "N/A"), "position": (600, 400), "color": (0, 0, 0), "font": font_valor},
        "conversao_valor": {"text": data_to_fill.get("conversao_valor", "N/A"), "position": (200, 635), "color": (0, 0, 0), "font": font_valor},
        "faturamento_valor": {"text": data_to_fill.get("faturamento_valor", "N/A"), "position": (1110, 285), "color": (0, 0, 0), "font": font_valor},
        "ticket_medio_valor": {"text": data_to_fill.get("ticket_medio_valor", "N/A"), "position": (1580, 285), "color": (0, 0, 0), "font": font_valor},
        "exames_lab_valor": {"text": data_to_fill.get("exames_lab_valor", "N/A"), "position": (1110, 550), "color": (0, 0, 0), "font": font_valor},
        "tm_exames_lab_valor": {"text": data_to_fill.get("tm_exames_lab_valor", "N/A"), "position": (1580, 550), "color": (0, 0, 0), "font": font_valor},
        "exames_imagem_valor": {"text": data_to_fill.get("exames_imagem_valor", "N/A"), "position": (1110, 810), "color": (0, 0, 0), "font": font_valor},
        "procedimentos_valor": {"text": data_to_fill.get("procedimentos_valor", "N/A"), "position": (1580, 810), "color": (0, 0, 0), "font": font_valor},
    }

    for campo, config in campos_para_preencher.items():
        text_to_draw = config["text"]
        position = config["position"]
        fill_color = config["color"]
        current_font = config["font"]

        draw.text(position, text_to_draw, fill=fill_color, font=current_font)
        print(f"Preenchido '{text_to_draw}' em {position} (Campo: {campo}).")

    img.save(output_image_path)
    print(f"Imagem final gerada e salva em: {output_image_path}")
    return output_image_path

if __name__ == "__main__":
    planilha_file = "indicadores.xlsx"
    template_image = "template_medicina.png"
    output_image = "relatorio_medicina_preenchido.png"

    dados_da_planilha = ler_dados_da_planilha(planilha_file)

    if dados_da_planilha:
        generated_image_path = preencher_template_medicina(dados_da_planilha, template_image, output_image)
        if generated_image_path:
            print(f"\nProcesso completo: Dados da planilha lidos e imagem gerada em: {generated_image_path}")
        else:
            print("\nFalha na geração da imagem.")
    else:
        print("\nNão foi possível ler os dados da planilha. Verifique o arquivo e o caminho.")