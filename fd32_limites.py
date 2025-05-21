# ================================================== #

# ~~ Bibliotecas.
import xlwings as xw
import win32com.client
from datetime import datetime
import os

# ================================================== #

# ~~ Coletando caminho da planilha.
path = os.path.abspath(__file__)
path = path.split(r"\fd32_limites.py")[0] + r"\fd32_limites.xlsm"

# ================================================== #

# ~~ Classe de alteração de limite.
class AlterarLimites:

    """
    Alteração de limites da FD32.
    """

# ================================================== #

    # ~~ Função.
    def alterar_limites(self) -> None:

        """
        Resumo:
        * Altera os limites de clientes no SAP.

        Parâmetros:
        * ===
        
        Retorna:
        * ===
        
        Exceções:
        * ===
        """

        # ~~ Loop.
        última_linha = self.ws.range("F" + str("999999")).end("up").row
        última_linha = última_linha + 1
        for linha in range(última_linha, 999999):

            # ~~ Coleta dados.
            cliente = self.ws.range("A" + str(linha)).value
            vencimento = self.ws.range("B" + str(linha)).value
            risco = self.ws.range("C" + str(linha)).value
            limite_total = self.ws.range("D" + str(linha)).value
            limite_segurado = self.ws.range("E" + str(linha)).value
            data_atual = datetime.now().date()
            data_atual = datetime.strftime(data_atual, "%d.%m.%Y")
            
            # ~~ Se encontrar linha vazia, encerrou lista.
            if cliente is None:
                self.session = None
                self.ws = None
                self.wb = None
                print("Lista encerrada.")
                print("==================================================")
                exit()
            
            # ~~ Converte dados e remove espaços.
            cliente = int(cliente)
            if vencimento is None:
                vencimento = ""
            else:
                try:
                    vencimento = datetime.strftime(vencimento, "%d.%m.%Y")
                except:
                    pass
            if risco is None:
                risco = ""
            else:
                risco = str(risco).strip()
            if limite_total is None:
                limite_total = ""
            else:
                limite_total = int(limite_total)
            if limite_segurado is None:
                limite_segurado = ""
            else:
                limite_segurado = int(limite_segurado)
            if not limite_total == "":
                limite_total = str(limite_total).strip()
            if not limite_segurado == "":
                limite_segurado = str(limite_segurado).strip()

            # ~~ Printa mensagem.
            print(f"Iniciando alteração da conta: {cliente}")
            print("---")

            # ~~ Para captura de erros.
            try:

                # ~~ Acessa FD32.
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NFD32"
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]/usr/ctxtRF02L-KUNNR").text = cliente
                self.session.findById("wnd[0]/usr/chkRF02L-D0105").selected = False
                self.session.findById("wnd[0]/usr/chkRF02L-D0110").selected = False
                self.session.findById("wnd[0]/usr/chkRF02L-D0120").selected = False
                self.session.findById("wnd[0]/usr/chkRF02L-D0220").selected = False
                self.session.findById("wnd[0]/usr/chkRF02L-D0210").selected = True
                self.session.findById("wnd[0]/usr/ctxtRF02L-KKBER").text = "1000"
                self.session.findById("wnd[0]").sendVKey(0)

                # ~~ Verifica se cliente está em agrupamento de contas e pula ele.
                msg_bar = self.session.findById("wnd[0]/sbar").text
                if "são atualizados na conta" in msg_bar:
                    print(f"Cliente: {cliente} está em agrupamento de contas. Continuando pro próximo.")
                    self.ws.range("F" + str(linha)).value = "AGRUPAMENTO"
                    continue
                elif "está marcado para eliminação" in msg_bar:
                    print(f"Cliente: {cliente} marcado para eliminação. Pulando.")
                    self.ws.range("F" + str(linha)).value = "ELIMINAÇÃO"
                    continue

                # ~~ Insere dados na transação.
                self.session.findById("wnd[0]/usr/txtKNKK-KLIMK").text = limite_total
                self.session.findById("wnd[0]/usr/ctxtKNKK-CTLPC").text = risco
                self.session.findById("wnd[0]/usr/ctxtKNKK-NXTRV").text = vencimento
                self.session.findById("wnd[0]/usr/ctxtKNKK-DTREV").text = data_atual
                self.session.findById("wnd[0]/usr/subARI-02:ZSAPMF2C:2000/txtKNKK-ZZLIMITE_SEG").text = limite_segurado

                # ~~ Limpa campos necessários.
                self.session.findById("wnd[0]/usr/subARI-02:ZSAPMF2C:2000/ctxtKNKK-ZLSCH").text = ""
                self.session.findById("wnd[0]/usr/subARI-02:ZSAPMF2C:2000/ctxtKNKK-ZTERM").text = ""

                # ~~ Salvamento.
                tentativa = 0
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                while tentativa != 5:
                    tentativa += 1
                    msg_bar = self.session.findById("wnd[0]/sbar").text
                    if msg_bar in ["Dados da área de controle 1000 modificado", "Não foi efetuada nenhuma modificação"]:
                        print(f"Cliente: {cliente} atualizado com sucesso.")
                        print("==================================================")
                        self.ws.range("F" + str(linha)).value = "OK"
                        break
                    else:
                        self.session.findById("wnd[0]").sendVKey(0)

                # ~~ Se houver erro ao salvar.
                if tentativa == 5:
                    print(f"Insucesso ao atualizar conta: {cliente}")
                    print("==================================================")
                    self.ws.range("F" + str(linha)).value = "ERRO"
            
            # ~~ Caso haja algum erro.
            except:
                print(f"Insucesso ao atualizar conta: {cliente}")
                print("==================================================")
                self.ws.range("F" + str(linha)).value = "ERRO"

    # ================================================== #

    # ~~ Função para instanciar objetos.
    def instanciar(self) -> None:

        """
        Resumo:
        * Faz vínculo com SAP e planilha.

        Parâmetros:
        * ===
        
        Retorna:
        * ===
        
        Exceções:
        * "Não logado no SAP.": Necessário fazer login.
        """

        # ~~ Cria instância da planilha.
        self.wb = xw.Book(path)
        self.ws = self.wb.sheets["FD32"]

        # ~~ Cria instância do SAP.
        try:
            gui = win32com.client.GetObject("SAPGUI")
            app = gui.GetScriptingEngine
            con = app.Children(0)
            self.session = con.Children(0)
        except:
            raise Exception("Não logado no SAP.")

    # ================================================== #

    # ~~ Função main.
    def main(self) -> None:

        """
        Resumo:
        * Função principal.

        Parâmetros:
        * ===
        
        Retorna:
        * ===
        
        Exceções:
        * ===
        """

        # ~~ Abertura.
        print("==================================================")

        # ~~ Instancia.
        try:
            self.instanciar()
        except Exception:
            exit()
        
        # ~~ Inicia agrupamento.
        self.alterar_limites()

# ================================================== #

# ~~ Inicia código.
if __name__ == "__main__":

    # ~~ Instancia AlterarLimites.
    alterar_limites = AlterarLimites()

    # ~~ Inicia código.
    alterar_limites.main()

# ================================================== #