import os
from datetime import datetime
import win32com.client as win32

# ========== CONFIG ==========
remetente = "vinicius.domingues@brktecnologia.com.br"  # <== troque pro seu e-mail correto
cc_padrao = [
    "iara.araujo@bayer.com", 
    "luiz.osato@bayer.com", 
    "luizpaulo.veiga@bayer.com", 
    "felipe.barbosa1@bayer.com", 
    "julia.merli@bayer.com", 
    "marcos.moura@bayer.com", 
    "vitoria.maciel@bayer.com", 
    "marcus.ramos@brktecnologia.com.br", 
    "marlon.felix@brktecnologia.com.br"
]
pasta_base = r"C:\Users\vinicius.domingues\Documents\Projeto\Base Bayer\Exp e Transito"

# Corpo padrão
mensagem = """Bom dia, time, Tudo bem?

Segue em anexo Tracking de CP com notas que estão em trânsito e fora do prazo e notas que ainda estão em expedição, podem justificar os atrasos, e informar uma previsão (caso haja) para as notas que ainda estão em expedição?

Fico no aguardo. Por gentileza, retornar até às 15:00 horas

Atte.
Vini G."""

# Gera data no formato dia.mes
hoje = datetime.now()
data_str = f"{hoje.day}.{hoje.month}"

# ========== NOMES DOS ARQUIVOS ==========
arquivos_esperados = [
    "BRAVO.xlsx",
    "LUFT DC Carazinho CP.xlsx",
    "LUFT DC Paulinia CP.xlsx",
    "TONIATO DC Ibipora CP.xlsx",
    "TONIATO DC Rio Verde CP.xlsx",
    "TONIATO DC Paulinia CP.xlsx",
    "TONIATO WH Belford Roxo CP.xlsx"
]

# ========== DESTINATÁRIOS POR GRUPO ==========
destinatarios_por_grupo = {
    "BRAVO": ["andreia.martins@bravolog.com.br", "barbara.goncalves@bravolog.com.br"],
    "LUFT CARAZINHO": ["taise.schmitt.ext@bayer.com", "roteirizacao.carazinho@luftagro.com.br", "thiago.fagundes@luftagro.com.br"],
    "LUFT PAULINIA": ["vanessa.manteiga@luftagro.com.br", "andrine.santos@luftagro.com.br", "fabio.silva@luftagro.com.br"],
    "TONIATO IBIPORA": ["milena.silva@grupotoniato.com.br", "viviane.garcia@grupotoniato.com.br"],
    "TONIATO PAULINIA": ["maria.samara@grupotoniato.com.br", "jovana.cerqueira@grupotoniato.com.br"],
    "TONIATO BELFORD ROXO": ["ronald.melo@grupotoniato.com.br", "elton.santos@grupotoniato.com.br"],
    "TONIATO RIO VERDE": ["jean.hossel@grupotoniato.com.br", "kelly.silva@grupotoniato.com.br", "samuel.pereira@grupotoniato.com.br"]
}

# ========== ENVIO ==========
try:
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Confere a conta de envio
    conta = None
    for a in namespace.Accounts:
        if a.SmtpAddress.lower() == remetente.lower():
            conta = a
            break

    if not conta:
        print(f"❌ Conta '{remetente}' não encontrada no Outlook.")
    else:
        for arquivo in arquivos_esperados:
            caminho_anexo = os.path.join(pasta_base, arquivo)

            # Verifica se o arquivo existe na pasta
            if not os.path.isfile(caminho_anexo):
                print(f"❌ Arquivo não encontrado: {arquivo}")
                continue

            # Identifica o grupo com base no nome do arquivo
            if "BRAVO" in arquivo:
                grupo = "BRAVO"
            elif "LUFT DC Carazinho CP" in arquivo:
                grupo = "LUFT CARAZINHO"
            elif "LUFT DC Paulinia CP" in arquivo:
                grupo = "LUFT PAULINIA"
            elif "TONIATO DC Ibipora CP" in arquivo:
                grupo = "TONIATO IBIPORA"
            elif "TONIATO DC Paulinia CP" in arquivo:
                grupo = "TONIATO PAULINIA"
            elif "TONIATO WH Belford Roxo CP" in arquivo:
                grupo = "TONIATO BELFORD ROXO"
            elif "TONIATO DC Rio Verde CP" in arquivo:
                grupo = "TONIATO RIO VERDE"
            else:
                print(f"❌ Grupo não encontrado para o arquivo: {arquivo}")
                continue

            # Destinatários específicos para o grupo
            destinatarios = destinatarios_por_grupo.get(grupo, [])
            if not destinatarios:
                print(f"❌ Não há destinatários definidos para o grupo {grupo}.")
                continue

            # Cria o e-mail
            mail = outlook.CreateItem(0)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, conta))  # Usa a conta certa
            mail.To = ";".join(destinatarios)
            mail.CC = ";".join(cc_padrao)  # Copia para os e-mails padrão
            mail.Subject = f"Tracking - {arquivo.replace('.xlsx', '')} {data_str}"  # Título com a data
            mail.Body = mensagem
            mail.Attachments.Add(caminho_anexo)
            mail.Send()  # Envia o e-mail diretamente
            print(f"✅ E-mail enviado para o grupo {grupo}: {arquivo}")

except Exception as e:
    print("❌ Erro no envio:", str(e))
