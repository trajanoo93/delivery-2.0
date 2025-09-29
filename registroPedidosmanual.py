import os
from reportlab.lib.pagesizes import portrait
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from datetime import datetime
import qrcode
from PIL import Image as PILImage
from io import BytesIO

def format_phone_number(phone):
    # Formata o número de telefone de 31998501560 para (31) 9 9850-1560
    if len(phone) == 11 and phone.isdigit():
        ddd = phone[0:2]
        prefix = phone[2:3]
        first_part = phone[3:7]
        second_part = phone[7:11]
        return f"({ddd}) {prefix} {first_part}-{second_part}"
    return phone # Retorna o número original se não estiver no formato esperado

def criar_pdf_invoice(id_pedido, pedido, delivery_date, delivery_time, delivery_type):
    # Inicializar variáveis
    delivery_time = delivery_time or ""
    delivery_date = delivery_date or ""
    pickup_time, pickup_date = "", ""
    delivery_method, delivery_cost = "", "0,00"

    # Extrair informações dos metadados do pedido
    for meta_data in pedido['meta_data']:
        if meta_data['key'] == 'pickup_time':
            pickup_time = meta_data['value']
        elif meta_data['key'] == 'pickup_date':
            pickup_date = meta_data['value']

    # Encontrar o método de entrega e o custo
    for shipping_line in pedido.get('shipping_lines', []):
        delivery_method = shipping_line.get('method_title', 'Não especificado')
        delivery_cost = shipping_line.get('total', '0,00')

    # Definir tamanho da página (A4 adaptado para impressora térmica)
    page_width, page_height = portrait((72 * mm, 297 * mm))

    # Criar diretório de invoices se não existir
    invoice_dir = 'C:/Users/ESCRITORIO/Desktop/invoices'
    os.makedirs(invoice_dir, exist_ok=True)

    # Criar objeto PDF
    pdf_path = os.path.join(invoice_dir, f'Invoice_{id_pedido}.pdf')
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=(page_width, page_height),
        leftMargin=5 * mm,
        rightMargin=5 * mm,
        topMargin=5 * mm,
        bottomMargin=5 * mm
    )

    # Definir estilos com ajustes para maior legibilidade
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Title"], fontSize=16, spaceAfter=4)
    order_number_style = ParagraphStyle("OrderNumber", parent=styles["Title"], fontSize=16, fontName="Helvetica-Bold",
                                       spaceAfter=4)
    header_style = ParagraphStyle("Header", parent=styles["Heading3"], fontSize=11, spaceAfter=2)
    body_style = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=10, leading=11)
    address_style = ParagraphStyle("Address", parent=styles["BodyText"], fontSize=11, leading=12,
                                  fontName="Helvetica-Bold")
    total_style = ParagraphStyle("Total", parent=styles["BodyText"], fontSize=11, textColor=colors.red, spaceAfter=4)
    urgent_style = ParagraphStyle("Urgent", parent=styles["BodyText"], fontSize=13, textColor=colors.red, alignment=1,
                                 spaceAfter=4)
    section_style = ParagraphStyle("Section", parent=styles["BodyText"], fontSize=11, textColor=colors.white,
                                  backColor=colors.black, alignment=1, spaceAfter=2, spaceBefore=2,
                                  fontName="Helvetica-Bold")
    qty_style = ParagraphStyle("Quantity", parent=styles["BodyText"], fontSize=12, fontName="Helvetica-Bold")
    date_highlight_style = ParagraphStyle("DateHighlight", parent=styles["BodyText"], fontSize=11,
                                         fontName="Helvetica-Bold")
    today_style = ParagraphStyle("Today", parent=styles["BodyText"], fontSize=10, fontName="Helvetica-Bold",
                                textColor=colors.green)
    highlight_style = ParagraphStyle("Highlight", parent=styles["BodyText"], fontSize=10, fontName="Helvetica-Bold")

    # Elementos do PDF
    elements = []

    # Cabeçalho
    logo_path = "logo_aogosto.png"
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=20 * mm, height=20 * mm)
        elements.append(logo)

    elements.append(Paragraph(f"#{id_pedido}", order_number_style)) # Número do pedido destacado
    elements.append(Spacer(1, 4 * mm))

    # Dados do Cliente (usando Table para alinhar com a tabela de produtos)
    nome_completo = f"{pedido['billing']['first_name']} {pedido['billing']['last_name']}"
    data = [[Paragraph("Dados do Cliente", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]))
    elements.append(section_table)
    elements.append(Paragraph(f"<b>Nome:</b> {nome_completo}", body_style))
    formatted_phone = format_phone_number(pedido['billing']['phone'])
    elements.append(Paragraph(f"<b>Telefone:</b> {formatted_phone}", body_style))
    if delivery_type == "pickup":
        unidade = next((meta['value'] for meta in pedido['meta_data'] if meta['key'] == '_store_final'),
                      'Central Distribuição (Sagrada Família)')
        elements.append(Paragraph(f"<b>Retirada:</b> {unidade}", address_style))
    else:
        address = f"{pedido['billing']['address_1']}, {pedido['billing']['number']}, {pedido['billing']['address_2']}, {pedido['billing']['neighborhood']}, {pedido['billing']['city']} - CEP: {pedido['billing']['postcode']}"
        elements.append(Paragraph(f"<b>Endereço:</b> {address}", address_style))
    elements.append(Spacer(1, 2 * mm))

    # Dados da Entrega (usando Table para alinhar)
    data = [[Paragraph("Dados da Entrega", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]))
    elements.append(section_table)
    if delivery_type == "pickup":
        elements.append(Paragraph(f"<b>Tipo:</b> Retirada na Unidade", body_style))
    else:
        elements.append(Paragraph(f"<b>Tipo:</b> Entrega", body_style))

    # Processar a data de entrega ou retirada
    current_date = datetime.now().strftime('%Y-%m-%d')
    scheduled_date = delivery_date or pickup_date
    is_scheduled = False
    is_today = False
    formatted_date = scheduled_date

    if scheduled_date:
        try:
            # Converter a data de YYYY-MM-DD para DD/MM/YYYY
            date_obj = datetime.strptime(scheduled_date, '%Y-%m-%d')
            formatted_date = date_obj.strftime('%d/%m/%Y')
            # Verificar se é agendado (data futura) ou para hoje
            if scheduled_date > current_date:
                is_scheduled = True
            elif scheduled_date == current_date:
                is_today = True
        except ValueError:
            formatted_date = 'Não informada'

    # Exibir a data com destaque se agendado ou com "Para Hoje" se for do dia
    if is_scheduled:
        date_text = f"<b>Data:</b> {formatted_date} <b>Agendado</b>"
        elements.append(Paragraph(date_text, date_highlight_style))
    elif is_today:
        date_text = f"<b>Data:</b> {formatted_date} <font color=green><b>(Para Hoje)</b></font>"
        elements.append(Paragraph(date_text, body_style))
    else:
        elements.append(Paragraph(f"<b>Data:</b> {formatted_date}", body_style))

    elements.append(Paragraph(f"<b>Horário:</b> {delivery_time or pickup_time or 'Não informado'}", body_style))
    elements.append(Paragraph(f"<b>Método:</b> {delivery_method}", body_style))
    if delivery_method.lower() == "go! express (em até 1 hora!)":
        elements.append(Paragraph("ENTREGA URGENTE", urgent_style))
    elements.append(Spacer(1, 2 * mm))

    # Verificar se algum item tem variações ou peso (_weight_grams)
    has_variations = False
    for item in pedido['line_items']:
        variations = [meta for meta in item.get('meta_data', []) if
                      meta.get('display_key') and meta.get('display_value')]
        weight_grams = next((meta['value'] for meta in item.get('meta_data', []) if meta['key'] == '_weight_grams'),
                            None)
        if variations or weight_grams:
            has_variations = True
            break

    # Tabela de Produtos (sem o cabeçalho "Itens do Pedido")
    if has_variations:
        # Tabela com coluna de variações
        data = [["Produto", "Variações", "Qtd", "Subtotal"]]
        col_widths = [25 * mm, 17 * mm, 13 * mm, 14 * mm] # Total: 69mm
    else:
        # Tabela sem coluna de variações
        data = [["Produto", "Qtd", "Subtotal"]]
        col_widths = [42 * mm, 13 * mm, 14 * mm] # Total: 69mm, mais espaço para "Produto"

    total = 0
    table_styles = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, 1), 8),
        ('FONTSIZE', (1 if has_variations else 0, 1), (1 if has_variations else 0, -1), 8),
        ('FONTSIZE', (2 if has_variations else 1, 1), (2 if has_variations else 1, -1), 12),
        ('FONTNAME', (2 if has_variations else 1, 1), (2 if has_variations else 1, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]

    row_index = 1 # Começa em 1 porque a linha 0 é o cabeçalho
    for item in pedido['line_items']:
        variations_list = [
            f"{meta['display_key']}: {meta['display_value']}"
            for meta in item.get('meta_data', [])
            if meta.get('display_key') and meta.get('display_value')
        ]
        weight_grams = next((meta['value'] for meta in item.get('meta_data', []) if meta['key'] == '_weight_grams'),
                            None)
        if weight_grams:
            variations_list.append(f"Aprox. {weight_grams}g")
        variations = ", ".join(variations_list)
        subtotal = float(item['total'].replace(',', '.'))
        total += subtotal

        # Verificar se o item é "Carvão" ou "Acendedor de Churrasqueira – Fogaço | R$: 2,00 (uni)" para aplicar destaque
        is_highlighted = item['name'].strip().lower() in ["carvão", "acendedor de churrasqueira – fogaço | r$: 2,00 (uni)"]
        item_style = highlight_style if is_highlighted else body_style
        # Adicionar símbolos visuais para destacar em impressoras térmicas
        item_name = f"*** {item['name']} ***" if is_highlighted else item['name']

        if has_variations:
            data.append([
                Paragraph(item_name, item_style),
                Paragraph(variations or "-", item_style),
                Paragraph(str(item['quantity']), qty_style),
                Paragraph(f"R$ {subtotal:.2f}", body_style)
            ])
        else:
            data.append([
                Paragraph(item_name, item_style),
                Paragraph(str(item['quantity']), qty_style),
                Paragraph(f"R$ {subtotal:.2f}", body_style)
            ])
        row_index += 1

    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle(table_styles))
    elements.append(table)
    elements.append(Spacer(1, 2 * mm))

    # Totais (usando Table para alinhar)
    data = [[Paragraph("Totais", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]))
    elements.append(section_table)
    taxa_entrega = float(pedido['shipping_total'].replace(',', '.'))
    elements.append(Paragraph(f"<b>Subtotal:</b> R$ {total:.2f}", body_style))
    elements.append(Paragraph(f"<b>Taxa de Entrega:</b> R$ {taxa_entrega:.2f}", body_style))

    # Descontos (cupons e fees, apenas para exibição)
    if pedido.get('coupon_lines'):
        for coupon in pedido['coupon_lines']:
            desconto = float(coupon['discount'].replace(',', '.'))
            elements.append(Paragraph(f"<b>Cupom ({coupon['code']}):</b> -R$ {desconto:.2f}", body_style))
    if pedido.get('fee_lines'):
        for fee in pedido['fee_lines']:
            valor = float(fee['total'].replace(',', '.'))
            if valor < 0:
                elements.append(Paragraph(f"<b>Desconto ({fee['name']}):</b> -R$ {abs(valor):.2f}", body_style))

    # Usar o valor total do pedido diretamente da API
    total_final = float(pedido['total'].replace(',', '.'))
    elements.append(Paragraph(f"<b>Total Final:</b> R$ {total_final:.2f}", total_style))
    elements.append(Paragraph(f"<b>Forma de Pagamento:</b> {pedido['payment_method_title']}", body_style))

    # Vendedor
    vendedor_mapping = {
        '62': 'Lorena', '062': 'Lorena', '43': 'Carol', '043': 'Carol', '7': 'Alline', '07': 'Alline',
        '52': 'Luciene', '052': 'Luciene', '74': 'Alinne', '71': 'Luiz', '071': 'Luiz', '73': 'Carlos jr',
        '77': 'Maria Eduarda'
    }
    vendedor_code = pedido['billing']['company']
    vendedor_nome = vendedor_mapping.get(vendedor_code, vendedor_code or 'Não especificado')
    elements.append(Paragraph(f"<b>Vendedor:</b> {vendedor_nome}", body_style))
    elements.append(Spacer(1, 2 * mm))

    # Observações (usando Table para alinhar)
    observacao = pedido['customer_note']
    if observacao:
        data = [[Paragraph("Observações", section_style)]]
        section_table = Table(data, colWidths=[69 * mm])
        section_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ]))
        elements.append(section_table)
        elements.append(Paragraph(observacao, body_style))
        elements.append(Spacer(1, 2 * mm))

    # QR Code (usando Table para alinhar)
    data = [[Paragraph("Confirmar Saída para Entrega", section_style)]]
    section_table = Table(data, colWidths=[69 * mm])
    section_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]))
    elements.append(section_table)
    script_url = "https://script.google.com/macros/s/AKfycbzSXA2EPQY7snNG0Hfuksnelh_dp6EwOFLc_4vcLMFiFCZ1bpsStt0WM5lWA4pi76q3/exec"
    qr_url = f"{script_url}?action=AssignDelivery&id={id_pedido}"
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(qr_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")

    # Converter imagem para stream compatível com reportlab
    qr_buffer = BytesIO()
    qr_img.save(qr_buffer, format="PNG")
    qr_buffer.seek(0)
    qr_image = Image(qr_buffer, width=30 * mm, height=30 * mm)
    elements.append(qr_image)
    elements.append(Paragraph("Escaneie para registrar a saída do pedido", body_style))
    elements.append(Spacer(1, 2 * mm))

    # Rodapé
    elements.append(Paragraph("Obrigado por comprar na Ao Gosto Carnes!", body_style))
    elements.append(Paragraph("Contato: (31) 3461-3297 | aogosto.com.br", body_style))

    # Construir PDF
    try:
        pdf.build(elements)
        print(f"Arquivo PDF do invoice criado: {pdf_path}")
        return True
    except Exception as e:
        print(f"Erro ao criar o PDF: {str(e)}")
        return False