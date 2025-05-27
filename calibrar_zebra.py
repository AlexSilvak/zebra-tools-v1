import win32print




def interpretar_status(status_code):
    
    status_flags = {
        0x00000001: "Pausada",
        0x00000002: "Erro",
        0x00000004: "Excluindo",
        0x00000008: "Aguardando",
        0x00000010: "Imprimindo",
        0x00000020: "Porta Fechada",
        0x00000040: "Sem Papel",
        0x00000080: "Uso Manual",
        0x00000100: "Offline",
        0x00000200: "Sem Tinta/Toner",
        0x00000400: "Problema com Papel",
        0x00000800: "Porta Ocupada",
        0x00001000: "Inicializando",
        0x00002000: "Warming Up",
        0x00004000: "Processando",
        0x00008000: "Impressora Compartilhada",
        0x00080000: "Baixo Toner",
    }
    status_list = [desc for bit, desc in status_flags.items() if status_code & bit]
    return status_list if status_list else ["Pronta"]

# Lista de portas USB válidas
valid_ports = ["USB005", "USB004", "USB003", "USB002", "USB001", "USB000"]

# Lista todas as impressoras
printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)

# Procura a impressora em uma das portas USB válidas
printer_name = None
for printer in printers:
    name = printer[2]
    handle = win32print.OpenPrinter(name)
    try:
        info = win32print.GetPrinter(handle, 2)
        port_name = info["pPortName"]
        if port_name in valid_ports:
            printer_name = name
            print(f"Impressora encontrada: {printer_name} na porta {port_name}")
            break
    finally:
        win32print.ClosePrinter(handle)

if not printer_name:
    print("Nenhuma impressora encontrada em portas USB válidas.")
    exit()


   

try:
    printer_handle = win32print.OpenPrinter(printer_name)
except Exception as e:
    print(f"Erro ao abrir a impressora: {e}")
    exit()

# Comandos ZPL
zpl_print_restaure_default = b"^XA^JUF^XZ"
zpl_print_calibration = b"~JC^XA^XZ"
zpl_print_clear = b"~JA"
zpl_print_teste = b"~WC"
zpl_print_clear_buffer = b"~JA"
zpl_feed_label = b"~VS"



# Menu de opções
while True:
    print(" ZDesigner")
    print("\nEscolha uma opção:")
    print("1 - Calibração")
    print("2 - Restauração de fábrica")
    print("3 - Limpeza da fila Impressão")
    print("4 - Teste de impressão")
    print("5 - Limpeza da Memoria de Impressão")
    print("6 - Escuridão")
    print("7 - Ver status da impressora")
    print("8 - Avançar etiqueta (Feed)")
    print("9 - Sair do Modo Pause")
    print("0 - Sair")
    
    try:
        option = int(input("Digite o número da opção: "))
        if option == 0:
            print("Finalizado.")
            break

        if option == 1:
            zpl_command = zpl_print_calibration
        elif option == 2:
            zpl_command = zpl_print_restaure_default
        elif option == 3:
            zpl_command = zpl_print_clear
        elif option == 4:
            zpl_command = zpl_print_teste
        elif option == 5:
            zpl_command = zpl_print_clear_buffer
        elif option == 6:
            darkness_level = int(input("Digite o nível de escuridão (0 a 15): "))
            if 0 <= darkness_level <= 15:
                zpl_command = f"~SD{darkness_level}".encode('utf-8')
            else:
                print("Valor inválido. Digite um número entre 0 e 15.")
                continue
        elif option == 8:
               zpl_command = zpl_feed_label
        elif option == 7:
            info = win32print.GetPrinter(printer_handle, 2)
            status_code = info["Status"]
            status_legivel = interpretar_status(status_code)

            print("\n📋 Status da Impressora:")
            print(f"Nome: {printer_name}")
            print(f"Porta: {info['pPortName']}")
            print(f"Driver: {info['pDriverName']}")
            print(f"Local: {info['pLocation']}")
            print(f"Status: {', '.join(status_legivel)}\n")
            continue
        else:
            print("Opção inválida. Tente novamente.")
            continue

        # Envia o comando
        job = win32print.StartDocPrinter(printer_handle, 1, ("Comando ZPL", None, "RAW"))
        win32print.StartPagePrinter(printer_handle)
        win32print.WritePrinter(printer_handle, zpl_command)
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        print("Comando enviado com sucesso.")

    except ValueError:
        print("Por favor, digite um número válido.")
    except Exception as e:
        print(f"Erro ao enviar o comando para a impressora: {e}")
