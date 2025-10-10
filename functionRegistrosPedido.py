import logging
import logging.handlers
from datetime import datetime, timedelta
import os
import json
from dateutil import parser
import requests
import time
import re
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import colorlog
from ast import literal_eval
import shutil
from contextlib import contextmanager
import gspread
from googleapiclient.errors import HttpError
import pytz
# =========================
# Utils
# =========================
def _col_label(n: int) -> str:
    # 1 -> A, 36 -> AJ
    label = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        label = chr(65 + rem) + label
    return label
def _get_sheet_props(service, spreadsheet_id, title):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties"
    ).execute()
    for s in meta["sheets"]:
        p = s["properties"]
        if p["title"] == title:
            grid = p.get("gridProperties", {})
            return p["sheetId"], grid.get("rowCount", 1000), grid.get("columnCount", 26)
    raise ValueError(f"Aba '{title}' não encontrada")
def write_row_with_template(
        service, spreadsheet_id, sheet_title, row_index, values,
        template_row=2, total_cols=None, text_cols=(2, 27)
):
    # Define range
    range_end_col = total_cols if total_cols is not None else len(values)
    last_col_letter = _col_label(range_end_col)
    range_str = f"{sheet_title}!A{row_index}:{last_col_letter}{row_index}"
    # Ajusta comprimento de values ao range
    if len(values) > range_end_col:
        logger.warning(f"values tem {len(values)} colunas; truncando para {range_end_col}.")
        values = values[:range_end_col]
    elif len(values) < range_end_col:
        values += [""] * (range_end_col - len(values))
    # 1) Copiar FORMATAÇÃO + VALIDAÇÃO da linha template para a linha destino
    sheet_id, _, _ = _get_sheet_props(service, spreadsheet_id, sheet_title)
    copy_src = {"sheetId": sheet_id,
                "startRowIndex": template_row - 1, "endRowIndex": template_row,
                "startColumnIndex": 0, "endColumnIndex": range_end_col}
    copy_dst = {"sheetId": sheet_id,
                "startRowIndex": row_index - 1, "endRowIndex": row_index,
                "startColumnIndex": 0, "endColumnIndex": range_end_col}
    requests = [
        {"copyPaste": {"source": copy_src, "destination": copy_dst, "pasteType": "PASTE_FORMAT"}},
        {"copyPaste": {"source": copy_src, "destination": copy_dst, "pasteType": "PASTE_DATA_VALIDATION"}},
    ]
    # 2) Forçar número/format das colunas B e AA como TEXTO na linha que vamos preencher
    for c in text_cols:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_index - 1, "endRowIndex": row_index,
                    "startColumnIndex": c - 1, "endColumnIndex": c
                },
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
                "fields": "userEnteredFormat.numberFormat"
            }
        })
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()
    # 3) Escrever valores como RAW (sem parsing do Sheets)
    body = {"values": [values]}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueInputOption="RAW",
        body=body
    ).execute()
    logger.info(f"Linha {row_index} escrita com sucesso na aba '{sheet_title}'")
def set_columns_as_text(service, spreadsheet_id, sheet_title, columns=(2, 27, 23)):
    sheet_id, row_count, _ = _get_sheet_props(service, spreadsheet_id, sheet_title)
    requests = []
    for c in columns:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": row_count,
                    "startColumnIndex": c - 1, "endColumnIndex": c
                },
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
                "fields": "userEnteredFormat.numberFormat"
            }
        })
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()
def normalize_id(order_id) -> str:
    """Normaliza IDs para string (evita duplicações por tipos diferentes)."""
    return str(order_id).strip()
def safe_parse_coupon_info(raw):
    """
    Faz parsing resiliente do campo coupon_info:
    - Se já for dict/list, retorna direto
    - Tenta json.loads
    - Tenta ast.literal_eval
    - Se tudo falhar, devolve string
    """
    if raw is None:
        return None
    if isinstance(raw, (dict, list, tuple)):
        return raw
    s = str(raw).strip()
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        pass
    try:
        return literal_eval(s)
    except Exception:
        return s
# =========================
# Configuração do Logging
# =========================
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)
log_file = os.path.join(log_dir, f"pedidos_{datetime.now().strftime('%Y%m%d')}.log")
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
console_handler.setLevel(logging.DEBUG)
console_formatter = colorlog.ColoredFormatter(
    "%(log_color)s%(levelname)s: %(message)s",
    log_colors={
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'DEBUG': 'blue',
    }
)
console_handler.setFormatter(console_formatter)
if os.name == 'nt':
    import sys
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
logger.addHandler(file_handler)
logger.addHandler(console_handler)
# =========================
# Configurações do WooCommerce
# =========================
consumer_key = 'ck_91e048e8dcc2e0dbd385f8c8b3c3c7c48062d80e'
consumer_secret = 'cs_32931adade86ee5b65eb3c62085e2e1b5b8b96e3'
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
def fetch_orders():
    """Busca os últimos 5 pedidos com status 'processing' modificados na última hora."""
    base_url = "https://aogosto.com.br/delivery/wp-json/wc/v3/orders"
    modified_after = (datetime.now() - timedelta(hours=1)).isoformat()
    params = {
        'per_page': 5,
        'order': 'desc',
        'orderby': 'modified',
        'status': 'processing',
        'modified_after': modified_after
    }
    try:
        response = requests.get(base_url, auth=(consumer_key, consumer_secret), params=params)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        logger.error(f"Erro ao buscar pedidos: {str(e)}")
        return []
def convert_data(data_original):
    try:
        logger.debug(f"Convertendo data: {data_original}")
        data_obj = parser.parse(data_original)
        return data_obj.date().strftime("%Y-%m-%d")
    except ValueError as e:
        logger.error(f"Falha ao converter data '{data_original}': {str(e)}")
        return None
def limpar_endereco(endereco):
    padrao = r'\s-\s(?:até|de|lado).*'
    return re.sub(padrao, '', endereco)
def separar_numero_endereco(endereco):
    match = re.match(r'^(\d+)\s+(.*)$', endereco)
    if match:
        numero, rua_limpa = match.groups()
        return rua_limpa, numero
    return endereco, None
def getLatLong(endereco):
    return None
def check_status(delivery_time, delivery_date):
    tz = pytz.timezone('America/Sao_Paulo')
    today = datetime.now(tz).date()
    try:
        delivery_date_dt = datetime.strptime(delivery_date, "%Y-%m-%d").date()
    except ValueError:
        logger.warning(f"Data inválida para agendamento: {delivery_date}")
        return 'Agendado'
    if delivery_date_dt == today:
        logger.debug(f"check_status: Pedido para hoje ({delivery_date}), retornando '-'")
        return '-'
    horarios = delivery_time.split(' - ')
    if len(horarios) != 2:
        logger.warning(f"Formato de horário inválido: {delivery_time}")
        return 'Agendado'
    try:
        _ = datetime.strptime(horarios[0], "%H:%M").time()
        _ = datetime.strptime(horarios[1], "%H:%M").time()
    except ValueError as e:
        logger.error(f"Erro ao parsear horário: {str(e)}")
        return 'Agendado'
    logger.debug(f"check_status: Pedido para data futura ({delivery_date}), retornando 'Agendado'")
    return 'Agendado'
def checkValidateAgendado(delivery_date):
    tz = pytz.timezone('America/Sao_Paulo')
    today = datetime.now(tz).date()
    try:
        delivery_date_dt = datetime.strptime(delivery_date, "%Y-%m-%d").date()
        return delivery_date_dt == today
    except ValueError:
        logger.warning(f"Data inválida para validação de agendamento: {delivery_date}")
        return False
def enviar_mensagem_whatsapp(celular, mensagem):
    if not celular or not celular.isdigit() or len(celular) < 10:
        logger.error(f"Falha ao enviar mensagem WhatsApp: número de celular inválido ({celular})")
        raise ValueError(f"Número de celular inválido: {celular}")
    url = "http://82.25.71.135:8080/message/sendText/central_delivery"
    payload = {
        "number": celular,
        "text": mensagem
    }
    headers = {
        "Content-Type": "application/json",
        "apikey": "3f0d87b1-0c4a-4e9c-bf14-9a07f6b7e9d3"
    }
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=10)
        response.raise_for_status()
        logger.info(f"Mensagem WhatsApp enviada com sucesso para {celular} via Evolution API")
        return response.json()
    except requests.RequestException as e:
        logger.error(f"Falha ao enviar mensagem WhatsApp para {celular}: {str(e)}")
        raise
def enviar_erro_ao_gestor(id_pedido, erro):
    """Envia mensagem ao WhatsApp do gestor com o erro e o ID do pedido."""
    gestor_numero = "5531998501560"
    mensagem = f"⚠️ Erro no processamento do pedido {id_pedido}: {str(erro)}"
    try:
        enviar_mensagem_whatsapp(gestor_numero, mensagem)
        logger.info(f"Mensagem de erro enviada ao gestor para o pedido {id_pedido}")
    except Exception as e:
        logger.error(f"Falha ao enviar mensagem de erro ao gestor para o pedido {id_pedido}: {str(e)}")
def enviar_mensagem_cliente(values, pedido, addressFull, unidade_retirada):
    id_pedido = pedido['id']
    user_name = values[4]
    if not user_name:
        logger.error(f"Falha ao enviar mensagem ao cliente do pedido {id_pedido}: nome não definido")
        raise ValueError("Nome do cliente não definido")
    celular = pedido['billing']['phone']
    if not celular:
        logger.error(f"Falha ao enviar mensagem ao cliente do pedido {id_pedido}: celular não definido")
        raise ValueError("Número de celular não definido")
    celular = ''.join(filter(str.isdigit, celular.lstrip('+55').lstrip('55')))
    if not celular.startswith("55"):
        celular = "55" + celular
    try:
        if values[24] == "pickup":
            if unidade_retirada == "Central Distribuição (Sagrada Família)":
                mensagem = f"""
Olá {user_name}! 👋
Seu pedido chegou aqui na Ao Gosto Carnes e já está sendo montado. 🥩📦
Pedimos um prazo de 30 minutos para montar o pedido. 😊
⚠️ *Mensagem automática — não responda por aqui*. Para falar com nossa equipe, utilize a Central Oficial: *(31) 3461-3297*.
📍Retirada: Av. Silviano Brandão, 685, Sagrada Família. (Basta subir o portão grande de garagem, temos estacionamento!)
Ah, lembramos que os pedidos para retirada são guardados somente até o final do dia.
                """.strip()
            elif unidade_retirada == "Unidade Barreiro":
                mensagem = f"""
Olá {user_name}! 👋
Seu pedido foi recebido pela *Ao Gosto Carnes* e já está em preparação! 🥩📦
Pedimos um prazo de 30 minutos para montá-lo. 😊
Caso tenha alguma dúvida, envie uma mensagem para a nossa Unidade do Barreiro: *(31) 99534-8704*
_Informações de Retirada:_
📆 *Data:* {values[26] or 'Não informada'}
⏰ *Horário:* {values[25] or 'Não informado'}
📍 *Local:* Av. Sinfrônio Brochado, 612 - Barreiro, Belo Horizonte
Obrigado por escolher a *Ao Gosto Carnes*!
(Mensagem Automática, favor não responder)
                """.strip()
            elif unidade_retirada == "Unidade Sion":
                mensagem = f"""
Olá {user_name}! 👋
Seu pedido foi recebido pela *Ao Gosto Carnes* e já está em preparação! 🥩📦
Pedimos um prazo de 30 minutos para montá-lo. 😊
Para falar na Unidade Sion, basta chamar nesse número: *(31) 9 8311-2919*.
_Informações de Retirada:_
📆 *Data:* {values[26] or 'Não informada'}
⏰ *Horário:* {values[25] or 'Não informado'}
📍 *Local:* Rua Haití, 354 - loja 5 - Sion, Belo Horizonte
Obrigado por escolher a *Ao Gosto Carnes*!
(Mensagem Automática, favor não responder)
                """.strip()
            else:
                mensagem = f"""
Olá {user_name}! 👋
Seu pedido chegou aqui na Ao Gosto Carnes e já está sendo montado. 🥩📦
📍Retirada: Av. Silviano Brandão, 685, Sagrada Família. (Basta subir o portão grande de garagem, temos estacionamento!)
⚠️ *Mensagem automática — não responda por aqui*.
Para falar com nossa equipe, utilize a central oficial: *(31) 3461-3297*.
                """.strip()
        else:
            mensagem = f"""
Ei {user_name}! 👋
Seu pedido na Ao Gosto Carnes foi *confirmado* e já estamos preparando tudo.
Aqui está o endereço de entrega:
📍 *{addressFull}*
⚠️ *Mensagem automática — não responda por aqui*.
Para falar com nossa equipe, utilize a central oficial: *(31) 3461-3297*.
Se o endereço está correto, em breve sua caixinha laranja estará aí com você!
⏰ O prazo de entrega varia de *30 minutos* a *2 horas* em BH e até 3 horas em outras localidades.
Estamos empenhados em entregar o mais rápido possível! 😊
Desejamos uma excelente experiência com nossos produtos!
            """.strip()
        logger.info(f"Enviando mensagem ao cliente do pedido {id_pedido}")
        enviar_mensagem_whatsapp(celular, mensagem)
    except Exception as e:
        logger.error(f"Falha ao enviar mensagem ao cliente do pedido {id_pedido}: {str(e)}")
def get_sheets_service(credentials_file):
    try:
        with open(credentials_file, 'r') as f:
            pass # Verifica se o arquivo é acessível
    except FileNotFoundError:
        logger.error(f"Arquivo de credenciais não encontrado: {credentials_file}")
        enviar_erro_ao_gestor("N/A", f"Arquivo de credenciais não encontrado: {credentials_file}")
        raise
    except PermissionError:
        logger.error(f"Sem permissão para acessar o arquivo de credenciais: {credentials_file}")
        enviar_erro_ao_gestor("N/A", f"Sem permissão para acessar o arquivo de credenciais: {credentials_file}")
        raise
    try:
        creds = Credentials.from_service_account_file(
            credentials_file, scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        return build('sheets', 'v4', credentials=creds)
    except Exception as e:
        logger.error(f"Falha ao inicializar serviço do Google Sheets: {str(e)}")
        enviar_erro_ao_gestor("N/A", str(e))
        raise
def verificar_valores_dropdown(values, id_pedido):
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
        logger.warning(
            f"Pedido {id_pedido}: Método de pagamento '{pagamento}' inválido. Usando 'Sem método de pagamento'.")
        values[5] = "Sem método de pagamento"
    if status not in valid_statuses:
        logger.warning(f"Pedido {id_pedido}: Status '{status}' inválido. Usando 'Pendente'.")
        values[10] = "Pendente"
    if entregador not in valid_entregadores:
        logger.warning(f"Pedido {id_pedido}: Entregador '{entregador}' inválido. Usando '-'.")
        values[11] = "-"
    return values
def processar_pedido_normal(values, pedido, addressFull, sheet, sheetAgendado, id_pedido, service, spreadsheet_key):
    from registroPedidosmanual import criar_pdf_invoice
    id_pedido_str = normalize_id(id_pedido)
    logger.debug(f"Processando pedido {id_pedido_str} com status={values[10]}, data_agendamento={values[26]}")
    values = verificar_valores_dropdown(values, id_pedido_str)
    # Determina a aba de destino
    if values[10] == 'Agendado' and not checkValidateAgendado(values[26]):
        logger.info(f"Pedido {id_pedido_str} agendado para o futuro, inserido na aba 'Agendados'")
        sheet_title = "Agendados"
        target_sheet = sheetAgendado
    else:
        logger.info(f"Pedido {id_pedido_str} inserido na aba 'Novos Pedidos'")
        sheet_title = "Novos Pedidos"
        target_sheet = sheet
    # Verifica se o pedido já existe na aba
    column_a = target_sheet.col_values(1)
    if id_pedido_str in column_a:
        logger.info(
            f"Pedido {id_pedido_str} já existe na aba '{sheet_title}', ignorando inserção, PDF e envio de mensagem.")
        return
    # Expande colunas da aba, se necessário
    next_row = len(target_sheet.col_values(1)) + 1
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_key).execute()
        for worksheet in sheet_metadata['sheets']:
            if worksheet['properties']['title'] == sheet_title:
                column_count = worksheet['properties']['gridProperties'].get('columnCount', 26)
                if column_count < len(values):
                    logger.warning(f"A aba '{sheet_title}' tem {column_count} colunas. Expandindo para {len(values)}.")
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_key,
                        body={
                            "requests": [{
                                "updateSheetProperties": {
                                    "properties": {
                                        "sheetId": worksheet['properties']['sheetId'],
                                        "gridProperties": {"columnCount": len(values)}
                                    },
                                    "fields": "gridProperties.columnCount"
                                }
                            }]
                        }
                    ).execute()
    except Exception as e:
        logger.error(f"Erro ao verificar/expandir colunas da aba '{sheet_title}' para pedido {id_pedido_str}: {str(e)}")
        enviar_erro_ao_gestor(id_pedido_str, str(e))
        return
    # Insere o pedido na aba
    try:
        write_row_with_template(service, spreadsheet_key, sheet_title, next_row, values, template_row=2)
        logger.info(f"Pedido {id_pedido_str} inserido com sucesso na aba '{sheet_title}', linha {next_row}")
    except HttpError as e:
        logger.error(f"Erro ao inserir pedido {id_pedido_str} na aba '{sheet_title}': {str(e)}")
        enviar_erro_ao_gestor(id_pedido_str, str(e))
        return
    # Determina o destino para o PDF
    destination = "CD Central"
    if values[10] == 'Agendado' and not checkValidateAgendado(values[26]):
        destination = "CD Central" # Agendados vão para CD Central
    else:
        destination = "CD Central" # Novos Pedidos vão para CD Central
    # Gera PDF para todos os pedidos
    try:
        logger.debug(f"Tentando gerar PDF para pedido {id_pedido_str} com destination={destination}")
        if criar_pdf_invoice(id_pedido_str, pedido, values[26], values[25], values[24], destination=destination):
            logger.info(f"PDF gerado com sucesso para pedido {id_pedido_str}")
        else:
            logger.error(f"Falha ao gerar PDF para pedido {id_pedido_str} (função retornou False)")
            enviar_erro_ao_gestor(id_pedido_str, "Falha ao gerar PDF (função retornou False)")
    except Exception as e:
        logger.error(f"Erro ao gerar PDF para pedido {id_pedido_str}: {str(e)}")
        enviar_erro_ao_gestor(id_pedido_str, str(e))
    # Envia mensagem ao cliente
    try:
        enviar_mensagem_cliente(values, pedido, addressFull, values[18])
    except Exception as e:
        logger.error(f"Erro ao enviar mensagem WhatsApp para pedido {id_pedido_str}: {str(e)}")
def getDictPedidos(values):
    return {
        'id': values[0], 'nome': values[4], 'total': values[6], 'rua': values[12],
        'numero': values[13], 'bairro': values[3], 'complemento': values[15],
        'cep': values[14], 'cidade': values[20], 'latitude': values[16],
        'longitude': values[17], 'telefone': values[22], 'observacao': values[23],
        'taxa_entrega': values[9], 'metodo_Pagamento': values[5],
        'delivery_date': values[26], 'horario': values[2], 'horarioEntrega': values[25]
    }
def load_registered_orders(registered_orders_file):
    try:
        with file_lock(registered_orders_file):
            with open(registered_orders_file, 'r', encoding='utf-8') as file:
                content = file.read().strip()
                if not content:
                    logger.info(f"Arquivo {registered_orders_file} está vazio. Retornando conjunto vazio.")
                    return set()
                try:
                    raw_list = json.loads(content)
                    normalized = {normalize_id(x) for x in raw_list}
                    if len(raw_list) != len(normalized):
                        logger.warning("IDs normalizados (havia mistura de tipos).")
                    logger.debug(f"Pedidos carregados de {registered_orders_file}: {normalized}")
                    return set(normalized)
                except json.JSONDecodeError as e:
                    logger.error(f"Erro ao decodificar JSON em {registered_orders_file}: {str(e)}")
                    logger.debug(f"Conteúdo do arquivo: {content[:1000]}...")
                    backup_file = registered_orders_file + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    shutil.copy2(registered_orders_file, backup_file)
                    logger.info(f"Arquivo corrompido salvo como {backup_file}. Resetando para lista vazia.")
                    with open(registered_orders_file, 'w', encoding='utf-8') as f:
                        json.dump([], f)
                    return set()
    except FileNotFoundError:
        logger.info(f"Arquivo {registered_orders_file} não encontrado. Retornando conjunto vazio.")
        return set()
def update_registered_orders(registered_orders, registered_orders_file):
    try:
        normalized = [normalize_id(x) for x in registered_orders]
        with file_lock(registered_orders_file):
            with open(registered_orders_file, 'w', encoding='utf-8') as file:
                json.dump(list(sorted(set(normalized))), file, ensure_ascii=False, indent=2)
                logger.debug(f"Arquivo {registered_orders_file} atualizado com {len(set(normalized))} pedidos.")
    except Exception as e:
        logger.error(f"Erro ao atualizar {registered_orders_file}: {str(e)}")
        enviar_erro_ao_gestor("N/A", str(e))
        raise
def tentar_executar_com_retries(funcao, *args, max_tentativas=5, intervalo_tentativas=25):
    for tentativa in range(1, max_tentativas + 1):
        try:
            return funcao(*args)
        except requests.RequestException as e:
            logger.error(f"Tentativa {tentativa} falhou: erro de conexão ({str(e)})")
            enviar_erro_ao_gestor("N/A", str(e))
            if tentativa < max_tentativas:
                logger.info(f"Aguardando {intervalo_tentativas} segundos para tentar novamente...")
                time.sleep(intervalo_tentativas)
            else:
                logger.error("Máximo de tentativas atingido. Operação abortada.")
                return None
        except Exception as e:
            logger.error(f"Erro inesperado na tentativa {tentativa}: {str(e)}")
            enviar_erro_ao_gestor("N/A", str(e))
            raise
def check_values(pedido):
    delivery_time = None
    delivery_date = None
    pickup_time = None
    pickup_date = None
    store_final = None
    effective_store_final = None
    payment_account_stripe = None
    payment_account_pagarme = None
    substitutions = {
        '7': 'Alline', '07': 'Alline', '74': 'Alinne', '73': 'Carlos jr',
        '77': 'Maria Eduarda', '78': 'Cássio Vinicius'
    }
    id_pedido = pedido['id']
    if len(str(id_pedido)) < 6:
        return None
    status = pedido['status']
    if status not in ['processing', 'saiu-pra-entrega', 'wc-agendado', 'lala-move']:
        logger.warning(f"Pedido {id_pedido} ignorado: status não permitido ({status})")
        return None
    for meta_data in pedido['meta_data']:
        if meta_data['key'] == 'delivery_time':
            delivery_time = meta_data['value']
        elif meta_data['key'] == 'delivery_date':
            delivery_date = meta_data['value']
        elif meta_data['key'] == 'pickup_time':
            pickup_time = meta_data['value']
        elif meta_data['key'] == 'pickup_date':
            pickup_date = meta_data['value']
        elif meta_data['key'] == '_effective_store_final':
            effective_store_final = meta_data['value']
        elif meta_data['key'] == '_payment_account_stripe':
            payment_account_stripe = meta_data['value']
        elif meta_data['key'] == '_payment_account_pagarme':
            payment_account_pagarme = meta_data['value']
        elif meta_data['key'] == '_store_final':
            store_final = meta_data['value']
    shipping_store = effective_store_final or store_final or ""
    if delivery_time and not re.match(r'^\d{2}:\d{2} - \d{2}:\d{2}$', delivery_time):
        logger.warning(f"Formato de delivery_time inválido para pedido {id_pedido}: {delivery_time}")
        delivery_time = None
    if delivery_date:
        delivery_date = convert_data(delivery_date)
        if delivery_date and not re.match(r'^\d{4}-\d{2}-\d{2}$', delivery_date):
            logger.warning(f"Formato de delivery_date inválido após conversão para pedido {id_pedido}: {delivery_date}")
            delivery_date = None
    if pickup_date:
        pickup_date = convert_data(pickup_date)
        if pickup_date and not re.match(r'^\d{4}-\d{2}-\d{2}$', pickup_date):
            logger.warning(f"Formato de pickup_date inválido após conversão para pedido {id_pedido}: {pickup_date}")
            pickup_date = None
    endereco_final = limpar_endereco(pedido['billing']['address_1'])
    endereco_limpo, numero = separar_numero_endereco(pedido['billing']['address_1'])
    if numero:
        pedido['billing']['number'] = numero
    endereco_final = limpar_endereco(endereco_limpo)
    payment_method = pedido['payment_method'].strip()
    logger.info(f"Valor bruto de payment_method para pedido {id_pedido}: '{payment_method}'")
    payment_method_mapping = {
        'cod': 'Cartão',
        'cartão_na_entrega': 'Cartão',
        'custom_729b8aa9fc227ff': 'Cartão',
        'woo_payment_on_delivery': 'Dinheiro',
        'vale_alimentação': 'Vale Alimentação',
        'voucher': 'V.A',
        'dinheiro_na_entrega': 'Dinheiro',
        'custom_e876f567c151864': 'V.A',
        'pagarme_custom_pix': 'Pix',
        'todo_incomm': 'Cartão Presente Ao Gosto Card',
        'stripe': 'Crédito Site',
        'stripe_cc': 'Crédito Site',
        'eh_stripe_pay': 'Crédito Site',
    }
    logger.info(f"Chaves disponíveis no payment_method_mapping: {list(payment_method_mapping.keys())}")
    payment_method = payment_method_mapping.get(payment_method, None)
    if payment_method is None and pedido.get('payment_method_title', '').startswith('Vale Alimentação'):
        logger.info(f"Usando fallback para payment_method_title: {pedido['payment_method_title']} -> V.A")
        payment_method = 'V.A'
    else:
        payment_method = payment_method or 'Sem método de pagamento'
    if payment_method == 'Sem método de pagamento' and pedido['payment_method']:
        logger.warning(f"Método de pagamento não mapeado para pedido {id_pedido}: {pedido['payment_method']}")
    celular = pedido['billing']['phone'].lstrip('+55').lstrip('55')
    if not celular.startswith("55"):
        celular = "55" + celular
    celular = ''.join(filter(str.isdigit, celular))
    horario = pedido['date_created'][-8:-3]
    data = pedido['date_created'][8:10] + "-" + pedido['date_created'][5:7]
    status_atualizado = check_status(
        delivery_time if delivery_time else pickup_time,
        delivery_date if delivery_date else pickup_date
    ) if (delivery_time or pickup_time) else '-'
    valor_total = float(pedido['total'].replace(',', '.'))
    taxa_entrega = float(pedido['shipping_total'])
    nome_cliente = pedido['billing']['first_name'].split(' ')[0]
    billing_company = substitutions.get(str(pedido['billing']['company']), pedido['billing']['company'] or 'Site')
    latitude, longitude = None, None
    for meta_item in pedido['meta_data']:
        if meta_item['key'] == 'billing_lat':
            latitude = float(meta_item['value'])
        elif meta_item['key'] == 'billing_long':
            longitude = float(meta_item['value'])
    if latitude is None or longitude is None:
        cordenadas = getLatLong(pedido['billing']['address_1'] + ",Minas Gerais - Brasil")
        latitude, longitude = cordenadas if cordenadas else ('latitude', 'longitude')
    delivery_type = "delivery"
    for shipping_line in pedido['shipping_lines']:
        method_title = shipping_line['method_title']
        if method_title == "Retirada na Unidade":
            delivery_type = "pickup"
        elif method_title == "Motoboy":
            delivery_type = "delivery"
        break
    if not pedido['billing']['neighborhood']:
        pedido['billing']['neighborhood'] = pedido['billing']['city']
    productList = []
    base_url = "https://aogosto.com.br/delivery/wp-json/wc/v3/products"
    for item in pedido['line_items']:
        variations = []
        product_id = item.get('product_id')
        if product_id:
            url_produto = f"{base_url}/{product_id}"
            try:
                resp = requests.get(url_produto, auth=(consumer_key, consumer_secret))
                if resp.status_code == 200:
                    product_data = resp.json()
                    product_meta_data = product_data.get('meta_data', [])
                    for meta in product_meta_data:
                        if meta.get('key') == '_weight_grams' and meta.get('value'):
                            item['meta_data'].append({
                                'key': '_weight_grams',
                                'value': meta['value']
                            })
                else:
                    logger.error(f"Falha ao buscar metadados do produto {product_id}: {resp.status_code} - {resp.text}")
                    enviar_erro_ao_gestor(id_pedido,
                                          f"Falha ao buscar metadados do produto {product_id}: {resp.status_code} - {resp.text}")
            except requests.RequestException as e:
                logger.error(f"Erro ao buscar metadados do produto {product_id}: {str(e)}")
                enviar_erro_ao_gestor(id_pedido, str(e))
        for meta in item.get('meta_data', []):
            key = meta.get('display_key', meta.get('key', ''))
            value = meta.get('display_value', meta.get('value', ''))
            if key == '_weight_grams' and value:
                variations.append(f"Peso: {value}g")
            elif key and value and key != '_weight_grams':
                variations.append(f"{key}: {value}")
        product = {
            "name": item['name'],
            "quantity": item['quantity'],
            "variations": variations
        }
        productList.append(product)
    productList_str = "\n".join(
        [f"{p['name']} (Qtd: {p['quantity']})" + (f" - {' | '.join(p['variations'])}" if p['variations'] else "") + " *"
         for p in productList]
    )
    coupon_code = ""
    coupon_value = ""
    coupon_type = ""
    if pedido.get('coupon_lines'):
        for coupon in pedido['coupon_lines']:
            coupon_code = coupon.get('code', '') or coupon_code
            discount_type = coupon.get('discount_type', '')
            discount_value = coupon.get('discount', '')
            found_info = False
            for meta in coupon.get('meta_data', []):
                if meta.get('key') == 'coupon_info':
                    ci = safe_parse_coupon_info(meta.get('value'))
                    try:
                        if isinstance(ci, (list, tuple)) and len(ci) >= 4:
                            valor = ci[3]
                            tipo = ci[2]
                            coupon_value = str(valor)
                            coupon_type = str(tipo)
                            found_info = True
                            break
                        elif isinstance(ci, dict):
                            val = ci.get('amount') or ci.get('value') or ci.get('discount') or ci.get('discount_value')
                            typ = ci.get('type') or ci.get('discount_type')
                            if val is not None:
                                coupon_value = str(val)
                                found_info = True
                            if typ is not None:
                                coupon_type = str(typ)
                                found_info = True
                            if found_info:
                                break
                        elif isinstance(ci, str) and ci:
                            pass
                    except Exception as e:
                        logger.warning(f"Falha ao interpretar coupon_info para pedido {pedido.get('id')}: {e}")
                        enviar_erro_ao_gestor(id_pedido, str(e))
            if not found_info:
                coupon_value = str(discount_value)
                coupon_type = discount_type or coupon_type
            break
    gift_card_discount = ""
    if pedido.get('fee_lines'):
        for fee in pedido['fee_lines']:
            if fee.get('name') == "Cartão Presente Ao Gosto Card":
                gift_card_discount = fee.get('total', '')
                break
    id_pedido_str = normalize_id(id_pedido)
    horario_agendamento = delivery_time if delivery_time else pickup_time
    if horario_agendamento:
        if not re.match(r'^\d{2}:\d{2} - \d{2}:\d{2}$', horario_agendamento):
            try:
                hora = datetime.strptime(horario_agendamento, "%H:%M").time()
                hora_inicio = hora.strftime("%H:%M")
                hora_fim = (datetime.combine(datetime.today(), hora) + timedelta(hours=3)).strftime("%H:%M")
                horario_agendamento = f"{hora_inicio} - {hora_fim}"
            except ValueError:
                logger.warning(f"Formato de horário inválido para pedido {id_pedido}: {horario_agendamento}")
                horario_agendamento = "12:00 - 15:00"
    else:
        horario_agendamento = "12:00 - 15:00"
    data_agendamento = delivery_date if delivery_date else pickup_date
    if data_agendamento:
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', data_agendamento):
            try:
                data_obj = parser.parse(data_agendamento)
                data_agendamento = data_obj.strftime("%Y-%m-%d")
            except ValueError:
                logger.warning(f"Formato de data inválido para pedido {id_pedido}: {data_agendamento}")
                data_agendamento = datetime.now(pytz.timezone('America/Sao_Paulo')).strftime("%Y-%m-%d")
    else:
        data_agendamento = datetime.now(pytz.timezone('America/Sao_Paulo')).strftime("%Y-%m-%d")
    values = [
        id_pedido_str,
        data, horario, pedido['billing']['neighborhood'], nome_cliente,
        payment_method, valor_total, valor_total, billing_company, taxa_entrega,
        status_atualizado, '-', endereco_final, pedido['billing'].get('number', ''),
        pedido['billing']['postcode'], pedido['billing']['address_2'], latitude, longitude,
        shipping_store, None, pedido['billing']['city'], '', celular, pedido['customer_note'],
        delivery_type, horario_agendamento,
        data_agendamento,
        None, None, productList_str
    ]
    values.extend(["", "", coupon_code, coupon_value, coupon_type, gift_card_discount, payment_account_stripe,
                   effective_store_final, payment_account_pagarme])
    return values
def adicionar_pedido_ao_google_sheets(pedido, registered_orders, sheet, sheetAgendado, sheetCDBarreiro, sheetCDSion,
                                      spreadsheet_key, client, registered_orders_file):
    from registroPedidosmanual import criar_pdf_invoice
    id_pedido = pedido['id']
    id_pedido_str = normalize_id(id_pedido)
    logger.info(f"Processando pedido {id_pedido_str}")
    if id_pedido_str in registered_orders:
        logger.info(f"Pedido {id_pedido_str} já registrado. ✅")
        return
    values = check_values(pedido)
    if not values:
        if len(str(id_pedido)) >= 6:
            logger.error(f"Falha ao processar pedido {id_pedido_str}: dados inválidos ❌")
            enviar_erro_ao_gestor(id_pedido_str, "Dados inválidos")
        return
    try:
        service = get_sheets_service(
            'C:/Users/ESCRITORIO/PycharmProjects/Delivery 2.0/impressao-belvedere-8a876abef441.json')
    except Exception as e:
        logger.error(f"Falha ao inicializar serviço do Google Sheets para pedido {id_pedido_str}: {str(e)}")
        enviar_erro_ao_gestor(id_pedido_str, str(e))
        return
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_key).execute()
        for worksheet in sheet_metadata['sheets']:
            if worksheet['properties']['title'] in [sheet.title, sheetAgendado.title, sheetCDBarreiro.title,
                                                    sheetCDSion.title]:
                column_count = worksheet['properties']['gridProperties'].get('columnCount', 26)
                if column_count < len(values):
                    logger.warning(
                        f"A aba '{worksheet['properties']['title']}' tem apenas {column_count} colunas. Expandindo para {len(values)}.")
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_key,
                        body={
                            "requests": [{
                                "updateSheetProperties": {
                                    "properties": {
                                        "sheetId": worksheet['properties']['sheetId'],
                                        "gridProperties": {"columnCount": len(values)}
                                    },
                                    "fields": "gridProperties.columnCount"
                                }
                            }]
                        }
                    ).execute()
    except Exception as e:
        logger.error(f"Erro ao verificar/expandir colunas da planilha para pedido {id_pedido_str}: {str(e)}")
        enviar_erro_ao_gestor(id_pedido_str, str(e))
        return
    values = verificar_valores_dropdown(values, id_pedido_str)
    logger.debug(f"Valores para pedido {id_pedido_str}: {values}")
    try:
        if id_pedido_str in registered_orders:
            logger.info(f"Pedido {id_pedido_str} já registrado, ignorando (dedup)")
            return
        scheduled_date = values[26]
        is_future_date = False
        if scheduled_date:
            try:
                scheduled_date_dt = datetime.strptime(scheduled_date, "%Y-%m-%d").date()
                tz = pytz.timezone('America/Sao_Paulo')
                is_future_date = scheduled_date_dt > datetime.now(tz).date()
                if is_future_date:
                    values[10] = "Agendado"
            except ValueError:
                logger.warning(f"Data de agendamento inválida para pedido {id_pedido_str}: {scheduled_date}")
                enviar_erro_ao_gestor(id_pedido_str, f"Data de agendamento inválida: {scheduled_date}")
                return
        store_final = (values[18] or "").strip().lower()
        logger.debug(f"Pedido {id_pedido_str}: store_final normalizado='{store_final}' (repr={repr(store_final)})")
        pedido_datetime = datetime.strptime(pedido['date_created'], "%Y-%m-%dT%H:%M:%S")
        pedido_hora = pedido_datetime.time()
        pedido_dia_semana = pedido_datetime.weekday()
        hora_seg_sex = datetime.strptime("18:00", "%H:%M").time()
        hora_sab = datetime.strptime("17:00", "%H:%M").time()
        hora_dom = datetime.strptime("13:00", "%H:%M").time()
        hora_fecha_cd = datetime.strptime("21:00", "%H:%M").time()

        # Definição de hora_limite com base no dia da semana
        if pedido_dia_semana < 5:  # Segunda a Sexta (0-4)
            hora_limite = hora_seg_sex
        elif pedido_dia_semana == 5:  # Sábado
            hora_limite = hora_sab
        else:  # Domingo (ou fallback)
            hora_limite = hora_dom

        addressFull = f"{pedido['billing']['address_1']}, {values[13]} / {pedido['billing']['address_2']}, {pedido['billing']['neighborhood']} - {pedido['billing']['city']} | Cep: {pedido['billing']['postcode']}"
        # Determina a aba de destino com base em store_final
        if store_final in ["unidade barreiro", "barreiro", "cd barreiro", "unidadebarreiro"]:
            target_sheet = sheetCDBarreiro
            sheet_title = "CD Barreiro"
            column_a = target_sheet.col_values(1)
            if id_pedido_str in column_a:
                logger.info(f"Pedido {id_pedido_str} já existe na aba 'CD Barreiro', ignorando inserção.")
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
            next_row = len(target_sheet.col_values(1)) + 1
            try:
                if (
                        hora_limite <= pedido_hora < hora_fecha_cd or pedido_hora >= hora_fecha_cd) and not checkValidateAgendado(
                        values[26]):
                    values[10] = "Agendado"
                    if not values[26]:
                        values[26] = (pedido_datetime + timedelta(days=1)).strftime("%Y-%m-%d")
                write_row_with_template(service, spreadsheet_key, sheet_title, next_row, values, template_row=2)
                logger.info(f"Pedido {id_pedido_str} inserido na aba 'CD Barreiro', linha {next_row}")
                try:
                    enviar_mensagem_cliente(values, pedido, addressFull, values[18])
                except Exception as e:
                    logger.error(f"Erro ao enviar mensagem WhatsApp para pedido {id_pedido_str}: {str(e)}")
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
            except Exception as e:
                logger.error(f"Erro no fluxo CD Barreiro para pedido {id_pedido_str}: {str(e)}")
                enviar_erro_ao_gestor(id_pedido_str, str(e))
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
        elif store_final in ["unidade sion", "sion", "cd sion", "unidadesion"]:
            target_sheet = sheetCDSion
            sheet_title = "CD Sion"
            column_a = target_sheet.col_values(1)
            if id_pedido_str in column_a:
                logger.info(f"Pedido {id_pedido_str} já existe na aba 'CD Sion', ignorando inserção.")
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
            next_row = len(target_sheet.col_values(1)) + 1
            try:
                if (
                        hora_limite <= pedido_hora < hora_fecha_cd or pedido_hora >= hora_fecha_cd) and not checkValidateAgendado(
                        values[26]):
                    values[10] = "Agendado"
                    if not values[26]:
                        values[26] = (pedido_datetime + timedelta(days=1)).strftime("%Y-%m-%d")
                write_row_with_template(service, spreadsheet_key, sheet_title, next_row, values, template_row=2)
                logger.info(f"Pedido {id_pedido_str} inserido na aba 'CD Sion', linha {next_row}")
                try:
                    enviar_mensagem_cliente(values, pedido, addressFull, values[18])
                except Exception as e:
                    logger.error(f"Erro ao enviar mensagem WhatsApp para pedido {id_pedido_str}: {str(e)}")
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
            except Exception as e:
                logger.error(f"Erro no fluxo CD Sion para pedido {id_pedido_str}: {str(e)}")
                enviar_erro_ao_gestor(id_pedido_str, str(e))
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
        else:
            if is_future_date or (pedido_hora >= hora_fecha_cd and not values[26]):
                if pedido_hora >= hora_fecha_cd and not is_future_date and not values[26]:
                    values[26] = (pedido_datetime + timedelta(days=1)).strftime("%Y-%m-%d")
                    values[10] = "Agendado"
                target_sheet = sheetAgendado
                sheet_title = "Agendados"
            else:
                target_sheet = sheet
                sheet_title = "Novos Pedidos"
            column_a = target_sheet.col_values(1)
            if id_pedido_str in column_a:
                logger.info(f"Pedido {id_pedido_str} já existe na aba '{sheet_title}', ignorando inserção.")
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
            next_row = len(target_sheet.col_values(1)) + 1
            try:
                write_row_with_template(service, spreadsheet_key, sheet_title, next_row, values, template_row=2)
                logger.info(f"Pedido {id_pedido_str} inserido na aba '{sheet_title}', linha {next_row}")
                if checkValidateAgendado(values[26]):
                    try:
                        logger.debug(f"Tentando gerar PDF para pedido {id_pedido_str}")
                        if criar_pdf_invoice(id_pedido_str, pedido, values[26], values[25], values[24]):
                            logger.info(f"PDF gerado com sucesso para pedido {id_pedido_str}")
                        else:
                            logger.error(f"Falha ao gerar PDF para pedido {id_pedido_str} (função retornou False)")
                            enviar_erro_ao_gestor(id_pedido_str, "Falha ao gerar PDF (função retornou False)")
                    except Exception as e:
                        logger.error(f"Erro ao gerar PDF para pedido {id_pedido_str}: {str(e)}")
                        enviar_erro_ao_gestor(id_pedido_str, str(e))
                try:
                    enviar_mensagem_cliente(values, pedido, addressFull, values[18])
                except Exception as e:
                    logger.error(f"Erro ao enviar mensagem WhatsApp para pedido {id_pedido_str}: {str(e)}")
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
            except Exception as e:
                logger.error(f"Erro ao inserir pedido {id_pedido_str} na aba '{sheet_title}': {str(e)}")
                enviar_erro_ao_gestor(id_pedido_str, str(e))
                registered_orders.add(id_pedido_str)
                update_registered_orders(registered_orders, registered_orders_file)
                return
    except Exception as e:
        logger.error(f"Falha ao processar pedido {id_pedido_str}: {str(e)}")
        enviar_erro_ao_gestor(id_pedido_str, str(e))
        registered_orders.add(id_pedido_str)
        update_registered_orders(registered_orders, registered_orders_file)

def main():
    spreadsheet_key = "1dYwkJXVHXXx__cYMd-xUoWnDNZd58xIvKVfLHJMYF_c"
    credentials_file = 'C:/Users/ESCRITORIO/PycharmProjects/Delivery 2.0/impressao-belvedere-8a876abef441.json'
    registered_orders_file = "registered_orders.json"
    try:
        service = get_sheets_service(credentials_file)
        client = gspread.authorize(Credentials.from_service_account_file(
            credentials_file, scopes=['https://www.googleapis.com/auth/spreadsheets']
        ))
        for aba in ["Novos Pedidos", "Agendados", "CD Barreiro", "CD Sion"]:
            set_columns_as_text(service, spreadsheet_key, aba, columns=(2, 27, 23))
        sheet = client.open_by_key(spreadsheet_key).worksheet("Novos Pedidos")
        sheetAgendado = client.open_by_key(spreadsheet_key).worksheet("Agendados")
        sheetCDBarreiro = client.open_by_key(spreadsheet_key).worksheet("CD Barreiro")
        sheetCDSion = client.open_by_key(spreadsheet_key).worksheet("CD Sion")
    except Exception as e:
        logger.error(f"Erro ao inicializar planilha ou serviço: {str(e)}")
        enviar_erro_ao_gestor("N/A", str(e))
        return
    registered_orders = load_registered_orders(registered_orders_file)
    pedidos = fetch_orders()
    for pedido in pedidos:
        adicionar_pedido_ao_google_sheets(
            pedido, registered_orders, sheet, sheetAgendado, sheetCDBarreiro, sheetCDSion,
            spreadsheet_key, client, registered_orders_file
        )
if __name__ == "__main__":
    main()