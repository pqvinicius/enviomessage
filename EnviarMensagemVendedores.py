# Importando as bibliotecas necessárias
import pywhatkit
import time
import datetime
import pandas as pd
import sys

# --- CARREGAMENTO DOS DADOS E TEMPLATE ---
try:
    with open('message.txt', 'r', encoding='utf-8') as file:
        template_mensagem = file.read()
    print("✅ Template de mensagem carregado.")
except FileNotFoundError:
    print("❌ ERRO: Arquivo 'message.txt' não encontrado.")
    sys.exit()

try:
    df_vendedores = pd.read_excel('contatosvendedores.xlsx')
    print(f"✅ Planilha com {len(df_vendedores)} vendedores carregada.")
except FileNotFoundError:
    print("❌ ERRO: Planilha 'contatosvendedores.xlsx' não encontrada.")
    sys.exit()

# --- LÓGICA PRINCIPAL DE LEITURA E ENVIO ---
print("\n🚀 Iniciando processamento e envio de mensagens personalizadas...")

# Loop que itera sobre cada linha da planilha (cada vendedor)
for indice, vendedor in df_vendedores.iterrows():
    try:
        # --- 1. EXTRAÇÃO DIRETA DE TODOS OS DADOS DA PLANILHA ---
        nome = vendedor['Nome']
        telefone = str(vendedor['Telefone'])
        faturado_mes = vendedor['Faturado_mes']
        meta = vendedor['Meta']
        alcance = vendedor['Alcance']
        fat_projetado = vendedor['Fat_Projetado']
        pct_projetado = vendedor['Pct_Projetado']
        meta_diaria = vendedor['Meta_diaria']

        print(f"\n- Processando vendedor: {nome}")

        # --- 2. GERAÇÃO DA DATA ATUAL ---
    
        data_atual_str = datetime.date.today().strftime('%d/%m/%Y')
        
        # --- 3. FORMATAÇÃO OPCIONAL PARA MELHOR VISUALIZAÇÃO ---
        faturado_mes_fmt = f"{faturado_mes:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        meta_fmt = f"{meta:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        fat_projetado_fmt = f"{fat_projetado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        meta_diaria_fmt = f"{meta_diaria:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # --- 4. MONTAGEM DA MENSAGEM FINAL ---
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

        # --- 5. ENVIO ---
        horario_envio = datetime.datetime.now() + datetime.timedelta(minutes=2)
        pywhatkit.sendwhatmsg(
            phone_no=telefone,
            message=mensagem_final,
            time_hour=horario_envio.hour,
            time_min=horario_envio.minute,
            wait_time=45,
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