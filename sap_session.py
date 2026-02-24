import subprocess
import time
import win32com.client as win32
import pythoncom
import os
from dotenv import load_dotenv

load_dotenv()


def open_sap_session():
    """Abre sessão SAP e retorna objeto session."""
    try:
        # Inicializa COM para a thread (necessário quando roda via API/FastAPI)
        pythoncom.CoInitialize()

        # Carrega credenciais do .env
        user = os.getenv("SAP_USER")
        password = os.getenv("SAP_PASSWORD")
        language = os.getenv("SAP_LANGUAGE", "PT")
        system = os.getenv("SAP_SYSTEM")
        client = os.getenv("SAP_CLIENT", "400")

        print(f"[SAP] Conectando ao sistema {system} (cliente {client})...")

        # Tenta conectar ao SAP GUI existente primeiro
        try:
            sap_gui = win32.GetObject("SAPGUI")
            app = sap_gui.GetScriptingEngine

            # Se já existe conexão, usa ela
            if app.Children.Count > 0:
                connection = app.Children(0)
                if connection.Children.Count > 0:
                    session = connection.Children(0)
                    print("[SAP] ✅ Sessão existente reutilizada")
                    return session
        except Exception:
            # SAP GUI não está aberto, vamos abrir
            pass

        # Se não tem conexão, abre nova
        print(f"[SAP] Abrindo nova sessão...")
        subprocess.call(
            f"START sapshcut -user={user} -pw={password} -language={language} -system={system} -client={client}",
            shell=True,
        )

        # Aguarda SAP abrir e conectar
        print("[SAP] Aguardando inicialização (10 segundos)...")
        time.sleep(10)

        # Tenta conectar
        print("[SAP] Procurando sessão disponível...")
        for tentativa in range(1, 16):
            try:
                sap_gui = win32.GetObject("SAPGUI")
                app = sap_gui.GetScriptingEngine

                if app.Children.Count > 0:
                    connection = app.Children(0)
                    if connection.Children.Count > 0:
                        session = connection.Children(0)
                        session.findById("wnd[0]").maximize()
                        print(f"[SAP] ✅ Sessão conectada (tentativa {tentativa}/15)")
                        return session

                if tentativa % 3 == 0:  # Log a cada 3 tentativas
                    print(f"[SAP] Aguardando... (tentativa {tentativa}/15)")
                time.sleep(1)
            except Exception:
                if tentativa % 3 == 0:
                    print(f"[SAP] Aguardando... (tentativa {tentativa}/15)")
                time.sleep(1)

        print("[SAP] ❌ Tempo esgotado: sessão não ficou disponível")
        raise Exception("Sessão SAP não ficou disponível após 15 tentativas")

    except Exception as erro:
        print(f"[SAP] ❌ Erro ao conectar: {erro}")
        return None


def close_sap_session():
    """Encerra processos do SAP."""
    print("[SAP] Encerrando processos...")
    subprocess.call(
        f'taskkill /f /fi "USERNAME eq %username%" /im "saplogon.exe"', shell=True
    )
    subprocess.call(
        f'taskkill /f /fi "USERNAME eq %username%" /im "saplgpad.exe"', shell=True
    )
    print("[SAP] ✅ Processos encerrados")