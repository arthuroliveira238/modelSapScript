# importações de módulos
from sap_session import open_sap_session, close_sap_session
import os
import time
import sys


# função principal
def main():
    # Verifica ambiente antes de executar
    # if not verificar_ambiente():
    #     return
    # Inicia a sessão SAP
    session = open_sap_session()
    # Se a sessão não for iniciada, encerra o programa
    if not session:
        return

    # Função para garantir que estamos no menu principal do SAP
    # def para criar funções - nome_função(parâmetros):
    def garantir_menu_principal(session):
        # procura pelo id wnd[0]/tbar[0]/okcd e insere o comando /n para voltar ao menu principal
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        # procura pelo id wnd[0] e envia a tecla Enter (VKey 0)
        session.findById("wnd[0]").sendVKey(0)

    # Exemplo de função para realizar uma ação no SAP
    def abrir_mb51(session):
        # procura pelo id wnd[0]/tbar[0]/okcd e insere o comando se16n
        print("Abrindo transação MB51")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "mb51"
            # procura pelo id wnd[0] e envia a tecla Enter (VKey 0)
            session.findById("wnd[0]").sendVKey(0)
        except:
            print("Erro: A transação MB51 não foi aberta!")
            raise
        
        print("Transação MB51 foi aberta")

    def gerar_relatorio_excel(session):

        print("Gerando relatório excel")

        try:
            session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "B84387"

            session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "4097"

            session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "0001"

            session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "01022026"

            session.findById("wnd[0]/usr/ctxtBUDAT-LOW").setFocus()

            session.findById("wnd[0]/usr/ctxtBUDAT-LOW").caretPosition = 8

            session.findById("wnd[0]").sendVKey (0)

            session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = "20022026"

            session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus()

            session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 8

            session.findById("wnd[0]").sendVKey (0)

            session.findById("wnd[0]").sendVKey (8,4788)

            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"

            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()

            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")

            session.findById("wnd[0]").sendVKey (0)

            session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\arthur.oliveira\Documents\DESENVOLVIMENTO\arquivosexcel"

            session.findById("wnd[0]").sendVKey (0)
            time.sleep(1)
        except:
            raise ValueError("Erro ao gerar relatório do Excel")
            
        print("Gerado relatório excel")
      

    # Exemplo de chamadadas das funções
    # ==================================================================================
    garantir_menu_principal(session)
    time.sleep(5)  # espera 5 segundo para garantir que a tela foi carregada
    abrir_mb51(session)
    gerar_relatorio_excel(session)
    time.sleep(7)

    # encerrar a sessão SAP
    close_sap_session()
    print("\nProcessos SAP encerrados.")


if __name__ == "__main__":
    main()
