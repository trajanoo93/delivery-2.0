from datetime import datetime, date
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import time
import json
import os
import shutil
from reportlab.lib.pagesizes import portrait
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from io import BytesIO
import qrcode
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import logging
import logging.handlers
import colorlog
from contextlib import contextmanager
import pytz

# Configuração do Logging
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_file = os.path.join(log_dir, f"pedidos_app_{datetime.now().strftime('%Y%m%d')}.log")
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

file_handler = logging.handlers.TimedRotatingFileHandler(
    log_file, when="midnight", interval=1, backupCount=30, encoding='utf-8'
)
file_handler.setLevel(logging.DEBUG)
file_formatter = logging.Formatter(
    "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
file_handler.setFormatter(file_formatter)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = colorlog.ColoredFormatter(
    "%(log_color)s%(levelname)s: %(message)s",
    log_colors={
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
    }
)
console_handler.setFormatter(console_formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Configurações do Google Sheets
credentials_file = 'impressao-belvedere-8a876abef441.json'
scopes = ['https://www.googleapis.com/auth/spreadsheets']
spreadsheet_key = '1dYwkJXVHXXx__cYMd-xUoWnDNZd58xIvKVfLHJMYF_c'
sheet_name = 'Novos Pedidos'

creds = Credentials.from_service_account_file(credentials_file, scopes=scopes)
service = build('sheets', 'v4', credentials=creds)
client = gspread.authorize(creds)
sheet = client.open_by_key(spreadsheet_key).worksheet(sheet_name)

url = "https://shop.fabapp.com/panel/stores/26682591/orders"
registered_orders_file = "registrado2.json"

delivery_mapeamento = {
    "in_home": "delivery",
    "on_site_pickup": "pickup",
}

data_atual = date.today()
data_formatada = data_atual.strftime('%d/%m')

@contextmanager
def file_lock(file_path):
    """Garante acesso exclusivo ao arquivo usando um arquivo de trava."""
    lock_file = file_path + '.lock'
    while os.path.exists(lock_file):
        time.sleep(0.1)
    try:
        with open(lock_file, 'w') as f:
            pass
        yield
    finally:
        try:
            os.remove(lock_file)
        except OSError:
            pass

def format_phone_number(phone):
    if len(phone) == 11 and phone.isdigit():
        ddd = phone[0:2]
        prefix = phone[2:3]
        first_part = phone[3:7]
        second_part = phone[7:11]
        return f"({ddd}) {prefix} {first_part}-{second_part}"
    return phone

def enviar_mensagem_whatsapp(celular, mensagem):
    if not celular.startswith("+55"):
        celular = "+55" + celular
    celular = ''.join(filter(str.isdigit, celular))

    url = "https://api.wzap.chat/v1/messages"
    payload = {"phone": celular, "message": mensagem}
    headers = {
        "Content-Type": "application/json",
        "Token": "7343607cd11509da88407ea89353ebdd8a79bdf9c3152da4025274c08c370b7b90ab0b68307d28cf"
    }
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=10)
        logger.info(f"Mensagem WhatsApp enviada com sucesso para {celular}")
        return response
    except Exception as e:
        logger.error(f"Falha ao enviar mensagem WhatsApp para {celular}: {str(e)}")
        return None

def criar_pdf_invoice_app(order_data):
    order_number = order_data['orderNumber']
    user_name = order_data['userName']
    formatted_phone = format_phone_number(order_data['userPhone'][2:])
    address_full = f"{order_data['address']['address']}, {order_data['address']['number']}, {order_data['address']['complement']}, {order_data['address']['neighborhood']}, {order_data['address']['city']} - CEP: {order_data['address']['zipCode']}"
    delivery_method = delivery_mapeamento.get(order_data.get('delivery', {}).get('method', 'indefinido'), 'indefinido')
    observacao = ""

    page_width, page_height = portrait((72 * mm, 297 * mm))
    pdf_path = os.path.join('C:/Users/ESCRITORIO/Desktop/invoices', f'Invoice_App_{order_number}.pdf')
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=(page_width, page_height),
        leftMargin=5 * mm,
        rightMargin=5 * mm,
        topMargin=5 * mm,
        bottomMargin=5 * mm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Title"], fontSize=16, spaceAfter=4)
    header_style = ParagraphStyle("Header", parent=styles["Heading3"], fontSize=11, spaceAfter=2)
    body_style = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=10, leading=11)
    address_style = ParagraphStyle("Address", parent=styles["BodyText"], fontSize=11, leading=12, fontName="Helvetica-Bold")
    section_style = ParagraphStyle("Section", parent=styles["BodyText"], fontSize=11, textColor=colors.white, backColor=colors.black, alignment=1, spaceAfter=2, spaceBefore=2, fontName="Helvetica-Bold")

    elements = []

    elements.append(Paragraph(f"Pedido App #{order_number}", title_style))
    elements.append(Spacer(1, 4 * mm))

    data = [[Paragraph("Dados do Cliente", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2)]))
    elements.append(section_table)
    elements.append(Paragraph(f"<b>Nome:</b> {user_name}", body_style))
    elements.append(Paragraph(f"<b>Telefone:</b> {formatted_phone}", body_style))
    elements.append(Spacer(1, 2 * mm))

    data = [[Paragraph("Dados da Entrega/Endereço", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2)]))
    elements.append(section_table)
    if delivery_method == "pickup":
        elements.append(Paragraph(f"<b>Tipo:</b> Retirada na Unidade (Av. Silviano Brandão, 685, Sagrada Família)", address_style))
    else:
        elements.append(Paragraph(f"<b>Tipo:</b> Entrega", body_style))
        elements.append(Paragraph(f"<b>Endereço:</b> {address_full}", address_style))
    elements.append(Spacer(1, 2 * mm))

    if observacao:
        data = [[Paragraph("Observações", section_style)]]
        section_table = Table(data, colWidths=[69 * mm])
        section_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2)]))
        elements.append(section_table)
        elements.append(Paragraph(observacao, body_style))
        elements.append(Spacer(1, 2 * mm))

    data = [[Paragraph("Confirmar Saída para Entrega", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2)]))
    elements.append(section_table)
    script_url = "https://script.google.com/macros/s/AKfycbzSXA2EPQY7snNG0Hfuksnelh_dp6EwOFLc_4vcLMFiFCZ1bpsStt0WM5lWA4pi76q3/exec"
    qr_url = f"{script_url}?action=AssignDelivery&id={order_number}"
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(qr_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")  # Corrigido: usar qr.make_image
    qr_buffer = BytesIO()
    qr_img.save(qr_buffer, format="PNG")
    qr_buffer.seek(0)
    qr_image = Image(qr_buffer, width=30 * mm, height=30 * mm)
    elements.append(qr_image)
    elements.append(Paragraph("Escaneie para registrar a saída do pedido", body_style))
    elements.append(Spacer(1, 2 * mm))

    elements.append(Paragraph("Obrigado por comprar na Ao Gosto Carnes!", body_style))

    try:
        pdf.build(elements)
        logger.info(f"Arquivo PDF do pedido app criado: {pdf_path}")
        return True
    except Exception as e:
        logger.error(f"Erro ao criar o PDF para pedido app {order_number}: {str(e)}")
        return False

def verificar_valores_dropdown(values, order_number):
    valid_payment_methods = [
        "Cartão", "Dinheiro", "Pix", "Vale Alimentação", "V.A", "Crédito Site",
        "Cartão Presente Ao Gosto Card", "Sem método de pagamento"
    ]
    valid_statuses = ["Pendente", "Agendado", "-", "Saiu para Entrega", "Entregue"]
    valid_entregadores = ["-", "Nenhum"]

    pagamento = values[5]
    status = values[10]
    entregador = values[11]

    if pagamento not in valid_payment_methods:
        logger.warning(f"Pedido {order_number}: Método de pagamento '{pagamento}' inválido. Usando 'Sem método de pagamento'.")
        values[5] = "Sem método de pagamento"

    if status not in valid_statuses:
        logger.warning(f"Pedido {order_number}: Status '{status}' inválido. Usando 'Pendente'.")
        values[10] = "Pendente"

    if entregador not in valid_entregadores:
        logger.warning(f"Pedido {order_number}: Entregador '{entregador}' inválido. Usando '-'.")
        values[11] = "-"

    return values

def open_spreadsheet(order_data, registered_orders):
    try:
        order_number = order_data['orderNumber']
        if order_number in registered_orders:
            logger.info(f"Pedido {order_number} duplicado. Não foi inserido!")
            return

        # Verificar e expandir colunas na planilha
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_key).execute()
        for worksheet in sheet_metadata['sheets']:
            if worksheet['properties']['title'] == sheet_name:
                column_count = worksheet['properties']['gridProperties'].get('columnCount', 26)
                if column_count < 35:
                    logger.warning(f"A aba '{sheet_name}' tem apenas {column_count} colunas. Expandindo para 35.")
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_key,
                        body={
                            "requests": [{
                                "updateSheetProperties": {
                                    "properties": {
                                        "sheetId": worksheet['properties']['sheetId'],
                                        "gridProperties": {"columnCount": 35}
                                    },
                                    "fields": "gridProperties.columnCount"
                                }
                            }]
                        }
                    ).execute()

        last_row = len(sheet.get_all_values()) + 1
        range_to_write = f"{sheet_name}!A{last_row}:AI{last_row}"

        shipping_tax = round(float(order_data['shippingTax']) / 100, 2)
        valor_total = round(float(order_data['amountFinal']) / 100, 2)

        payment_mapping = {
            'Cartão de Crédito': 'Cartão',
            'Cartão de Débito': 'Cartão',
            'Voucher': 'V.A'
        }

        neighborhood = order_data['address']['neighborhood']
        user_name = order_data['userName'].split(' ')[0]
        payment_method = order_data['paymentMethod']['option']['title']
        payment_method = payment_mapping.get(payment_method, 'Sem método de pagamento')
        status = order_data['status']['title']
        company = 'App'
        address = order_data['address']['address']
        numero = order_data['address']['number']
        celular = order_data['userPhone']
        zipCode = order_data['address']['zipCode']
        complement = order_data['address']['complement']
        longitude = float(order_data['address']['lng'])
        latitude = float(order_data['address']['lat'])
        delivery_method = order_data.get('delivery', {}).get('method', 'indefinido')
        delivery_method = delivery_mapeamento.get(delivery_method, delivery_method)
        cidade = order_data['address']['city']
        data_formato = order_data['createdAt']
        data_objeto = datetime.strptime(data_formato, "%Y-%m-%dT%H:%M:%S.%f%z")
        horario = data_objeto.strftime("%H:%M")
        datapedido = data_formato[:10]
        dia_mes = datapedido[8:10] + '-' + datapedido[5:7]

        tz = pytz.timezone('America/Sao_Paulo')
        datadia = data_objeto.astimezone(tz)
        dia_formatado = datadia.strftime('%d/%m')

        if dia_formatado != data_formatada:
            logger.info(f"Pedido {order_number} não é do dia atual. Não foi inserido!")
            return

        celular = celular[2:]

        productsList = ''
        if 'items' in order_data:
            items = order_data['items']
            for item in items:
                item_quantity = str(item['quantity'])
                item_product_name = item['productName']
                productsList += f"{item_product_name} (Qtd: {item_quantity}) *\n"

        order_row = [
            order_number, dia_mes, horario, neighborhood, user_name, payment_method,
            valor_total, valor_total, company, shipping_tax, '-', '-',
            address, numero, zipCode, complement, latitude, longitude, '',
            None, cidade, '', celular, '', delivery_method, '', datapedido,
            None, None, productsList, '', '', '', '', ''
        ]

        order_row = verificar_valores_dropdown(order_row, order_number)

        if len(order_row) != 35:
            logger.error(f"Pedido {order_number}: Número incorreto de valores ({len(order_row)}, esperado 35)")
            return

        logger.debug(f"Valores para pedido {order_number}: {order_row}")

        # Inserir na planilha
        try:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_key,
                range=range_to_write,
                valueInputOption="RAW",
                body={"values": [order_row]}
            ).execute()
            logger.info(f"Pedido {order_number} inserido com sucesso na aba '{sheet_name}'")
        except Exception as e:
            logger.error(f"Erro ao inserir pedido {order_number} na aba '{sheet_name}': {str(e)}")
            return

        # Marcar o pedido como registrado antes de PDF e WhatsApp
        registered_orders.add(order_number)
        save_registered_orders(registered_orders)

        # Tentar gerar PDF
        if not criar_pdf_invoice_app(order_data):
            logger.warning(f"Continuando processamento do pedido {order_number} apesar de falha no PDF")

        # Tentar enviar mensagem WhatsApp
        if delivery_method == "pickup":
            mensagem = f"""
Olá {user_name}! 👋

Seu pedido chegou aqui na Ao Gosto Carnes e já está sendo montado. 🥩📦
Pedimos um prazo de 30 minutos para montar o pedido. 😊

📍Retirada: Av. Silviano Brandão, 685, Sagrada Família. (Basta subir o portão grande de garagem, temos estacionamento!)
Ah, lembramos que os pedidos para retirada são guardados somente até o final do dia.
            """
        else:
            mensagem = f"""
Ei {user_name}! 👋

Seu pedido na Ao Gosto Carnes foi confirmado e já estamos preparando tudo. Aqui está o endereço de entrega:
📍{address}, {numero}, {complement}, {neighborhood} - {cidade} | CEP: {zipCode}

Se o endereço está correto, em breve sua caixinha laranja estará aí com você!
O prazo de entrega varia de 30 minutos a 2 horas em BH e até 3 horas em outras localidades.
Estamos empenhados em entregar o mais rápido possível! 😊

Desejamos uma excelente experiência com nossos produtos!
            """

        logger.info(f"Enviando mensagem ao cliente do pedido {order_number}")
        if not enviar_mensagem_whatsapp(celular, mensagem):
            logger.warning(f"Continuando processamento do pedido {order_number} apesar de falha no WhatsApp")

        logger.info(f"Pedido de {user_name} ({order_number}) registrado com sucesso!")
        return True

    except Exception as e:
        logger.error(f"Erro ao processar pedido {order_number}: {str(e)}")
        return False

def load_registered_orders():
    try:
        with file_lock(registered_orders_file):
            with open(registered_orders_file, 'r', encoding='utf-8') as file:
                content = file.read()
                if not content.strip():
                    logger.info(f"Arquivo {registered_orders_file} está vazio. Retornando conjunto vazio.")
                    return set()
                try:
                    # Try to parse as JSON first
                    return set(json.loads(content))
                except json.JSONDecodeError:
                    # Fallback to legacy newline-separated format
                    logger.warning(f"Arquivo {registered_orders_file} não é JSON válido. Tentando formato legado (linhas separadas).")
                    try:
                        return set(line for line in content.splitlines() if line.strip())
                    except Exception as e:
                        logger.error(f"Erro ao processar formato legado em {registered_orders_file}: {str(e)}")
                        # Create a backup and reset the file
                        backup_file = registered_orders_file + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        shutil.copy2(registered_orders_file, backup_file)
                        logger.info(f"Arquivo corrompido salvo como {backup_file}. Resetando para lista vazia.")
                        with open(registered_orders_file, 'w', encoding='utf-8') as f:
                            json.dump([], f)
                        return set()
    except FileNotFoundError:
        logger.info(f"Arquivo {registered_orders_file} não encontrado. Retornando conjunto vazio.")
        return set()

def save_registered_orders(registered_orders):
    try:
        with file_lock(registered_orders_file):
            with open(registered_orders_file, 'w', encoding='utf-8') as file:
                json.dump(list(registered_orders), file, ensure_ascii=False, indent=2)
                logger.debug(f"Arquivo {registered_orders_file} atualizado com {len(registered_orders)} pedidos.")
    except Exception as e:
        logger.error(f"Erro ao atualizar {registered_orders_file}: {str(e)}")
        raise

def tentar_executar_com_retries(funcao, *args, max_tentativas=5, intervalo_tentativas=25):
    for tentativa in range(1, max_tentativas + 1):
        try:
            return funcao(*args)
        except requests.RequestException as e:
            logger.error(f"Tentativa {tentativa} falhou: erro de conexão ({str(e)})")
            if tentativa < max_tentativas:
                logger.info(f"Aguardando {intervalo_tentativas} segundos para tentar novamente...")
                time.sleep(intervalo_tentativas)
            else:
                logger.error("Máximo de tentativas atingido. Operação abortada.")
                return None
        except Exception as e:
            logger.error(f"Erro inesperado na tentativa {tentativa}: {str(e)}")
            raise

def check_new_orders():
    registered_orders = load_registered_orders()
    num_orders_to_process = 5

    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            json_data = response.json()
            orders = json_data.get('data', [])

            if orders:
                for latest_order in orders:
                    order_id = latest_order['id']
                    if order_id not in registered_orders:
                        try:
                            link_pedido = f"https://shop.fabapp.com/panel/stores/26682591/orders/{order_id}"
                            response = requests.get(link_pedido, timeout=10)
                            if response.status_code == 200:
                                order_data = response.json()
                                success = open_spreadsheet(order_data, registered_orders)
                                if success:
                                    num_orders_to_process -= 1
                                    if num_orders_to_process <= 0:
                                        break
                            else:
                                logger.error(f"Erro ao obter os dados JSON para o pedido {order_id}: {response.status_code}")
                        except requests.RequestException as e:
                            logger.error(f"Erro de conexão ao obter os dados JSON para o pedido {order_id}: {str(e)}")
                    else:
                        logger.info(f"Pedido {order_id} já registrado. Ignorando duplicata...")
            else:
                logger.info("Nenhum novo pedido encontrado.")
        else:
            logger.error(f"Erro ao obter a lista de pedidos: {response.status_code}")
    except requests.RequestException as e:
        logger.error(f"Erro de conexão: {str(e)}")
    except Exception as e:
        logger.error(f"Erro inesperado: {str(e)}")

    save_registered_orders(registered_orders)
    time.sleep(65)

def main():
    while True:
        tentar_executar_com_retries(check_new_orders)

if __name__ == "__main__":
    main()