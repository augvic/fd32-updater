# ================================================== #

# ~~ Bibliotecas.
import xlwings as xw
import win32com.client
import os

# ================================================== #

# ~~ Coletando caminho da planilha.
path = os.path.abspath(__file__)
path = path.split(r"\fd32_agrupamento.py")[0] + r"\fd32_agrupamento.xlsm"

# ================================================== #

# ~~ Classe de agrupamento de contas.
class AgrupamentoContas:

    """
    Agrupamento de contas na FD32.
    """

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
        * "Não logado no SAP.": Necessário realizar login.
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

    # ~~ Função de agrupar.
    def agrupar(self) -> None:

        """
        Resumo:
        * Faz agrupamento de contas na FD32.

        Parâmetros:
        * ===
        
        Retorna:
        * ===
        
        Exceções:
        * ===
        """

        # ~~ Coleta linha do último cliente agrupado.
        última_linha = self.ws.range("C" + str("999999")).end("up").row
        última_linha += 1

        # ~~ Loop.
        for linha in range(última_linha, 999999):

            # ~~ Coleta dados.
            cliente = self.ws.range("A" + str(linha)).value
            agrupar_em = self.ws.range("B" + str(linha)).value

            # ~~ Se encontrar linha vazia, encerrou lista.
            if cliente is None:
                self.session = None
                self.ws = None
                self.wb = None
                print("Lista encerrada.")
                print("==================================================")
                exit()

            # ~~ Converte para int.
            cliente = int(cliente)
            agrupar_em = int(agrupar_em)
            print(f"Iniciando agrupamento da conta: {cliente}.")
            print("---")

            # ~~ Captura erros.
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

                # ~~ Agrupa conta.
                self.session.findById("wnd[0]/mbar/menu[1]/menu[0]").select()

                # ~~ Verifica se já está agrupado na mesma conta.
                msg_bar = self.session.findById("wnd[0]/sbar").text
                if "1 clientes ainda c/referência à conta" in msg_bar:
                    self.ws.range("C" + str(linha)).value = "AGRUPADO"
                    continue

                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.session.findById("wnd[1]/usr/ctxtKNKK-KNKLI").text = agrupar_em
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                try:
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except:
                    pass

                # ~~ Salvamento.
                tentativa = 0
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                while tentativa != 5:
                    tentativa += 1
                    msg_bar = self.session.findById("wnd[0]/sbar").text
                    if msg_bar in ["Dados da área de controle 1000 modificado", "Não foi efetuada nenhuma modificação"]:
                        print(f"Conta {cliente} agrupada na conta {agrupar_em}.")
                        print("==================================================")
                        self.ws.range("C" + str(linha)).value = "AGRUPADO"
                        break
                    else:
                        self.session.findById("wnd[0]").sendVKey(0)

                # ~~ Se houver erro ao salvar.
                if tentativa == 5:
                    print(f"Erro ao agrupar a conta: {cliente} na conta {agrupar_em}.")
                    print("==================================================")
                    self.ws.range("C" + str(linha)).value = "ERRO"

            # ~~ Se encontrar qualquer erro.
            except:
                print(f"Erro ao agrupar a conta: {cliente} na conta {agrupar_em}.")
                print("==================================================")
                self.ws.range("C" + str(linha)).value = "ERRO"

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
        self.agrupar()

# ================================================== #

# ~~ Inicia código.
if __name__ == "__main__":

    # ~~ Instancia AgrupamentoContas.
    agrupamento = AgrupamentoContas()

    # ~~ Inicia função main.
    agrupamento.main()

# ================================================== #