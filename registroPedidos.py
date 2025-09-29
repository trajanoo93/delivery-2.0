import time
from functionsRegistroPedidos import main as processar_pedidos

def main():
    while True:
        try:
            processar_pedidos()
            print("Checando novos pedidos...")
            time.sleep(30)  # espera 1 minuto
        except Exception as e:
            print(f"Erro inesperado: {str(e)}")
            print("Aguardando 120 segundos antes de tentar novamente...")
            time.sleep(120)

if __name__ == "__main__":
    main()
