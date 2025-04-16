import tkinter as tk
from tkinter import messagebox
import win32print
import win32ui
import datetime
import os
NOME_IMPRESSORA = "ZDesigner GK420t"
CAMINHO_ARQUIVO_LOG =os.path.join(os.path.dirname(os.path.abspath(__file__)), "registros_impressao.txt")
# Função para quebrar o texto do motivo em até 37 caracteres por linha
def quebrar_texto(texto, limite=37):
    linhas = []
    while texto:
        linha = texto[:limite]
        texto = texto[limite:]
        linhas.append(linha.strip())
    return linhas
def registrar_log(loja, setor, motivo):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(CAMINHO_ARQUIVO_LOG, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} | Loja: {loja} | Setor: {setor} | Motivo: {motivo}\n")
# Função chamada ao clicar no botão "Gerar ZPL"
def gerar_zpl():
    loja = entrada_loja.get()
    motivo = entrada_motivo.get("1.0", tk.END).strip()
    setor = entrada_setor.get()

    if not loja or not motivo or not setor:
        messagebox.showwarning("Campos obrigatórios", "Preencha todos os campos.")
        return

    linhas_motivo = quebrar_texto(motivo, 37)

    zpl = "^XA\n"
    zpl += "^FO000,00^GFA,6422,6422,38,,::::::::::::::::::::::::::U07IF,T07JFE,S01LF8,S07LFE,S0NF,R03NF8,R07NFC,R0OFE,Q01PFS07F8,Q03PFR03FFEgG01C,Q07PF9FFEN07IF80FC07E3IFEFC7FFC001FF803IFE,Q0QFDIF8L01JFC1FC0FF7IFEFE7IF007FFE07JF,Q0QFCIFEL03JFE1FE0FF7IFEFE7IF80JF0KF,P01KFC7JFCP03JFE1FE0FF7IFCFE7IFC1JF8JFE,P03JFE00JFEJFL07KF1FC0FF7IFCFE7IFE3JF8JFE,P03JFC007IFE7IF8K0FFC1FF1FC0FE7IF1FEFF3FE7FEFFCJFE,P07JF8007IFE7IF8K0FF81FF1FC0FE7F001FEFE1FEFF83FC007FC,P07JFI03JF7IF8J01FF00FF3FC0FEFF001FEFE1FEFF01FC00FF8,P0JFEI03JFO01FF00FFBFC1FEFF001FCFE1FEFE01FC01FF,P0JFCI03JF7IFCJ03FE00FFBFC1FEFF001FCFE1IFE01FC03FE,P0JFCI01JF3IFEJ03FE00FFBF81FCIFE1FCFE1IFE01FE03FC,O01JF8I01JF3IFEJ03FE00FFBF81FCIFE3FDFE3FDFE01FC07F8,O01JF8I01JFO03FC00FFBF81FCIFC3FDJF9FC01FC0FF,O01JFJ01JFO03FC00FF7F83FCIFC3FDJFBFC01FC1FE,O03JFJ01JF3JFJ03FC00FF7F83FDIFC3F9IFE3FC03FC3FC,O03JFJ01JF3JFJ03FC01FF7F83F9IFC3F9IFC3FC03FC7FC,O03IFEJ01JF3JFJ03FC01FF7F83F9FC007F9IFC3FC03FCFF8,O03IFEJ01JFO03FC01FE7F03F9FC007FBFDFE3FC07F8FF,O07IFEJ01JF3JF8I03FC03FE7F87F9FC007FBF9FE1FC07F9FE,O07IFEJ01JF3JF8I03FE07FE7F87FBFC007FBF8FF1FE0FF3FC,O07IFEJ03JF7JF8I03FE0FFC7JF3IFC7F3F8FF1KF7JF,O07IFEJ03JF3JFCI03KFC7JF3IFE7F3F8FF8JFE7JF,O07IFEJ03JFO01KF87IFE3IFE7F7F87F8JFC7JF8,O07IFCJ03JF7JFCJ0KF03IFC3IFEFF7F87F87IF87JF8,O07IFCJ03IFE7JFEJ0JFE01IF83IFEFF7F03FC3IF0KF8,O07IFCJ07IFE7JFEJ07IFC007FE03IFEFE7F01F80FFC07JF,O07IFCJ07IFE7JFEJ03IF,O07IFCJ07IFEQ07IF8,O07IFCJ0JFCLFL0IF,O07IFCJ0JFCLFL0IFgG07E6,O07IFEI01JFDLFL03FEgG07FE,O07IFEI01JF8R01FCgG0FFE,O07IFEI03JF9LF8gO067E,O07IFEI03JF3LF8V01CV01C,O03JFI07JF7LF8J0FC1IFE1F801FF803F807FF003F801FF8,O03JF800JFE7LFCI01FE3IFE3FC03FFC03FC0IF803FC03FFC,O03JFC03JFEQ01FE7IFE3FC07FFC07FC0IFC07FC07FFE,O01KF0KFCMFCI03FE7IFC7FE0IFE07FC0IFE07FC0IFE,O01QF9MFEI03FF7IF87FE1FE7C0FFC0FCFE0FFC1FE7F,O01QF9MFEI07FF01F80FFE1F81C0FFC0FC7F0FFC1F83F,P0QF3MFEI07FF03F80FFE3F8001FFC0FC3F1FFE3F83F,P0PFER0IF03F01FBE3FI01F7E1FC3F1F7E3F03F,P07OFCFE003JFI0F9F03F01F3E7FI03F7E1F83F3F7E3F03F,P03OF9FC001JF001F9F03F03F3E7FI03E7E1F83F3E7E7F03F,P01OF3FC001JF001F1F83F03F3F7EI07E7E1F83F7E7E7E03F,Q0NFER03IF83F07IF7EI07FFE1F87F7FFE7E03F,Q07MF8FF8I0JF803IF83F07IF7EI0IFE1F87EIFE7E07F,Q03MF3FF8I0JF807IF87E0JF7F030JF3F0FEIFE7E07E,R0LFC7FFJ0JFC07IF87E0JF7F079JF3F1LF7F0FE,R03KFC1FK07IFC0JF87E0JF7F8F9JF3IFDJF3F9FE,S0LF803P0FC1F87E1F81FBIF9F83F3IFBF83F3IFC,T03MFJ03IFC1FC0FC7E1F81FBIFBF03F3IF3F03F1IF8,T03MFJ01IFE1F80FCFE3F01F9IF3F03F3FFE3F03F1IF,T03MFK0IFE3F80FCFC3F01F8FFE7E03F7FF87E03F07FC,T03MFgI01FV01F,T03MF,T01LFE,U0LFE,U07KFE,U03KFC,V0KF8P03C06R02O02I01002,V07JFQ0C606M08J02O02I01003,V01IFEP018206M0CJ02O02I01003,W03FFQ018304M0CJ06O06I03002,gQ010307E3E3E7CE3C07E3C10CF1E7E3E3F1E2,gQ03030I62708C8460CI61891B8C6626I32,gQ03020C24261898C2086C619B13086C642636,gQ03020C2C26109882084CC19373084C6C2I6,gQ03060C6C66319886184F00B3C218486C2784,gQ03040C6C6431908618C800A20218C84C64,gQ018C08CCE41390CC08CC40E31208JCE624,gQ01F80F87E41F9E780FE7C0C1F20FE767E3EC,,:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::^FS"

    zpl += "^FO330,60^A0N,40,40^FD Setor de TI^FS\n"

    # Caixa do motivo
    zpl += "^FO120,190^GB500,240,3^FS\n"
    zpl += "^FO10,200^A0N,30,30^FDMotivo:^FS\n"

    y_offset = 200
    for i, linha in enumerate(linhas_motivo):
        zpl += f"^FO120,{y_offset + (i * 30)}^A0N,30,30^FD {linha}^FS\n"

    # Setor
    zpl += "^FO420,120^GB200,40,3^FS\n"
    zpl += "^FO320,130^A0N,30,30^FDSetor:^FS\n"
    zpl += f"^FO430,130^A0N,30,30^FD{setor}^FS\n"

    # Loja
    zpl += "^FO110,120^GB200,40,3^FS\n"
    zpl += "^FO10,130^A0N,30,30^FDLoja:^FS\n"
    zpl += f"^FO120,130^A0N,30,30^FD{loja}^FS\n"

    zpl += "^XZ"
    registrar_log(loja, setor, motivo)

"""try:
        printer_handle = win32print.OpenPrinter(NOME_IMPRESSORA)
        job = win32print.StartDocPrinter(printer_handle, 1, ("ZPL Print Job", None, "RAW"))
        win32print.StartPagePrinter(printer_handle)
        win32print.WritePrinter(printer_handle, zpl.encode('utf-8'))
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        win32print.ClosePrinter(printer_handle)
        
        registrar_log(loja, setor, motivo)

        messagebox.showinfo("Impressão", "Etiqueta enviada para a impressora.")
**except Exception as e:
        messagebox.showerror("Erro na impressão", f"Ocorreu um erro ao imprimir:\n{e}")
"""
# Interface com tkinter
janela = tk.Tk()
janela.title("Gerador de ZPL")

tk.Label(janela, text="Loja:").grid(row=0, column=0, sticky="e")
entrada_loja = tk.Entry(janela, width=30)
entrada_loja.grid(row=0, column=1, pady=5)

tk.Label(janela, text="Setor:").grid(row=1, column=0, sticky="e")
entrada_setor = tk.Entry(janela, width=30)
entrada_setor.grid(row=1, column=1, pady=5)

tk.Label(janela, text="Motivo:").grid(row=2, column=0, sticky="ne")
entrada_motivo = tk.Text(janela, width=30, height=4)
entrada_motivo.grid(row=2, column=1, pady=5)

botao_gerar = tk.Button(janela, text="Gerar ZPL", command=gerar_zpl)
botao_gerar.grid(row=3, column=0, pady=10)

botao_gerar = tk.Button(janela, text="Imprimir", command=gerar_zpl)
botao_gerar.grid(row=3, column=0, columnspan=2, pady=10)

janela.mainloop()
