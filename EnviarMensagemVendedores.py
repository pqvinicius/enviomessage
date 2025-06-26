# =============================================================================
# IMPORTAÇÕES NECESSÁRIAS
# =============================================================================
import pywhatkit
import time
import datetime
import pandas as pd
import sys

# =============================================================================
# CARREGAMENTO DOS DADOS E TEMPLATE
# =============================================================================
try:
    with open('message.txt', 'r', encoding='utf-8') as file:
        template_mensagem = file.read()
    print("✅ Template de mensagem carregado.")
except FileNotFoundError:
    print("❌ ERRO: Arquivo 'message.txt' não encontrado. Verifique se ele está na mesma pasta do script.")
    sys.exit()

try:
    nome_aba = 'basededados'
    df_vendedores = pd.read_excel('contatosvendedores.xlsx', sheet_name=nome_aba)
    
    print(f"✅ Planilha com {len(df_vendedores)} vendedores carregada.")
    
    # =============================================================================
    # LIMPEZA E PADRONIZAÇÃO DOS DADOS
    # =============================================================================
    print("⏳ Padronizando os dados para o envio...")
    
    # Garante que a coluna 'Telefone' seja tratada como texto e adiciona o '+' se faltar.
    df_vendedores['Telefone_Formatado'] = df_vendedores['Telefone'].astype(str).apply(
        lambda telefone_str: f"+{telefone_str}" if not telefone_str.startswith('+') else telefone_str
    )
    # Extrai o primeiro nome e formata a capitalização (Ex: JOÃO -> João).
    df_vendedores['primeiro_nome'] = df_vendedores['Nome'].str.split().str[0].str.title()
    
    print("✅ Dados limpos e padronizados com sucesso.")

except FileNotFoundError:
    print("❌ ERRO: Planilha 'contatosvendedores.xlsx' não encontrada.")
    sys.exit()
except KeyError as e:
    print(f"❌ ERRO: A coluna {e} não foi encontrada na planilha ao tentar formatar. Verifique os nomes 'Telefone' e 'Nome'.")
    sys.exit()

# =============================================================================
# LÓGICA PRINCIPAL DE LEITURA E ENVIO
# =============================================================================
print("\n🚀 Iniciando processamento e envio de mensagens personalizadas...")

for indice, vendedor in df_vendedores.iterrows():
    try:
        # Extração dos dados da linha atual
        nome = vendedor['primeiro_nome']
        telefone = str(vendedor['Telefone_Formatado'])
        faturado_mes = vendedor['Faturado_mes']
        meta = vendedor['Meta']
        alcance = vendedor['Alcance']
        fat_projetado = vendedor['Fat_Projetado']
        pct_projetado = vendedor['Pct_Projetado']
        meta_diaria = vendedor['Meta_diaria']

        print(f"\n- Processando vendedor: {nome}")

        # Geração da data atual
        data_atual_str = datetime.date.today().strftime('%d/%m/%Y')
        
        # Formatação dos números
        faturado_mes_fmt = f"{faturado_mes:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        meta_fmt = f"{meta:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        fat_projetado_fmt = f"{fat_projetado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        meta_diaria_fmt = f"{meta_diaria:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Montagem da mensagem final
        mensagem_final = template_mensagem.format(
            Nome=nome,
            data_atual=data_atual_str,
            Faturado_mes=faturado_mes_fmt,
            Meta=meta_fmt,
            Alcance=f"{alcance:.2f}", 
            Fat_Projetado=fat_projetado_fmt,
            Pct_Projetado=f"{pct_projetado:.2f}",
            Meta_diaria=meta_diaria_fmt
        )

        print("  Mensagem montada. Agendando envio...")
        print(mensagem_final)

        # Envio
        horario_envio = datetime.datetime.now() + datetime.timedelta(minutes=2)
        pywhatkit.sendwhatmsg(
            phone_no=telefone,
            message=mensagem_final,
            time_hour=horario_envio.hour,
            time_min=horario_envio.minute,
            wait_time=30,
            tab_close=True,
            close_time=10
        )
        print(f"  ✅ Mensagem para {nome} agendada com sucesso!")

        time.sleep(15)


    except KeyError as e:
        print(f"❌ ERRO: A coluna {e} não foi encontrada na planilha. Verifique os nomes das colunas.")
    except Exception as e:
        print(f"❌ ERRO ao processar o vendedor {nome}: {e}")

print("\n\n📊 Processo finalizado!")