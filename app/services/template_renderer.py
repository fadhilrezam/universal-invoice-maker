"""Generator template HTML A4 & Styling Dokumen Invoice.

Mengadopsi desain dokumen dari reference-invoicely.html:
1. 4 Variasi Gaya Header (Modern Studio, Bold Banner, Korporat Klasik, Minimalis)
2. 3 Variasi Gaya Tabel (Zebra Striped, Grid Bordered, Clean Lines)
3. Stempel Status (LUNAS, BELUM BAYAR, DRAFT, CUSTOM)
4. Full Customization via DesignConfig (CSS variables & inline attributes)
5. Multi-page A4 Pagination & Print CSS presisi (@media print { size: A4 portrait; margin: 0; })
"""
from __future__ import annotations

import html
import math
from dataclasses import asdict, dataclass, field
from datetime import date
from typing import TYPE_CHECKING, Sequence

if TYPE_CHECKING:
    from .invoice_export import InvoiceData, InvoiceLine


@dataclass
class DesignConfig:
    """Konfigurasi desain dokumen invoice A4 (diadopsi dari Invoicely)."""
    # 1. Gaya Layout Dokumen & Header
    # 'modern' (Modern Studio), 'banner' (Bold Banner), 'classic' (Korporat Klasik), 'minimal' (Minimalis)
    layout: str = "modern"
    header_align: str = "default"  # 'default' (kanan), 'swapped' (kiri), 'centered' (tengah)
    
    # 2. Gaya Tabel
    # 'clean' (Clean Lines), 'zebra' (Zebra Striped), 'bordered' (Grid Bordered)
    table_style: str = "clean"
    show_row_num: bool = True
    row_padding: int = 10  # padding baris tabel dalam px
    
    # 3. Warna Tema & Palet (Hex code)
    color_primary: str = "#2563eb"     # Aksen utama / brand (biru/merah/dsb)
    color_table_head: str = "#f8fafc"  # Background header tabel
    color_table_head_text: str = "#475569" # Teks header tabel
    color_paper: str = "#ffffff"       # Warna dasar kertas A4
    color_box_bg: str = "#f8fafc"      # Background kotak total / instruksi bayar
    color_text: str = "#1e293b"        # Teks utama
    
    # 4. Tipografi & Ukuran Kertas
    font_family: str = "font-sans"     # 'font-sans', 'font-jakarta', 'font-serif', 'font-mono'
    font_scale: int = 100              # Zoom / font scale dalam % (85 - 115)
    page_padding: int = 48             # Padding kertas A4 dalam px (24 - 64)
    border_radius: str = "rounded-lg"  # 'rounded-none', 'rounded-md', 'rounded-lg', 'rounded-xl'
    
    # 5. Logo & Gambar
    show_logo: bool = True
    logo_data_url: str = ""            # Base64 string / path URL gambar
    logo_size: int = 44                # Tinggi logo dalam px (30 - 80)
    
    # 6. Stempel / Watermark Status
    # 'none', 'LUNAS', 'BELUM BAYAR', 'DRAFT', 'CUSTOM'
    stamp_status: str = "none"
    stamp_custom_text: str = "VERIFIED"
    
    # 7. Visibilitas Elemen
    show_due_date: bool = True
    show_hero_total: bool = True
    show_payment_box: bool = True
    show_signature: bool = True
    show_footer: bool = True
    
    # 8. Label Dokumen & Kolom (Dapat dikustomisasi)
    doc_title: str = "INVOICE"
    bill_to_label: str = "Ditagihkan Kepada:"
    th_num: str = "#"
    th_desc: str = "Deskripsi Item / Jasa"
    th_qty: str = "Qty"
    th_unit: str = "Satuan"
    th_price: str = "Harga Satuan"
    th_total: str = "Jumlah"
    
    # 9. Informasi Finansial Ekstra & Catatan
    tax_rate: float = 0.0              # Persentase Pajak / PPN (0 = off)
    discount_rate: float = 0.0         # Persentase Diskon (0 = off)
    currency_symbol: str = "Rp"        # Simbol mata uang
    due_date_text: str = ""            # Tanggal jatuh tempo opsional
    bank_info: str = "BCA: 123-456-7890 a.n. Perusahaan\nMandiri: 098-765-4321"
    notes: str = "Terima kasih atas kerja sama Anda."
    signer_name: str = "Finance & Accounting"
    
    # 10. Kontrol Paginasi & Interaktivitas
    items_per_page: int = 0            # 0 = auto paginate
    content_editable: bool = False     # Aktifkan contenteditable="true"


def format_rupiah_simple(amount: float | int) -> str:
    """Format angka nominal ke representasi rupiah."""
    try:
        val = int(round(amount))
        return f"Rp {val:,}".replace(",", ".")
    except Exception:
        return f"Rp {amount}"


def paginate_invoice_lines(
    lines: Sequence[InvoiceLine],
    custom_limit: int = 0,
    row_pad: int = 10,
) -> list[list[tuple[int, InvoiceLine]]]:
    """Bagi baris invoice ke dalam halaman A4.

    Mengadaptasi logika `paginateInvoiceItems` dari reference-invoicely.html.
    Kembalikan list of pages, tiap page adalah list of (global_index, InvoiceLine).
    """
    total = len(lines)
    if total == 0:
        return [[]]

    density_factor = 0.8 if row_pad > 13 else (1.25 if row_pad < 9 else 1.0)
    first_page_cap = custom_limit if custom_limit > 0 else max(6, round(11 * density_factor))
    next_page_cap = custom_limit if custom_limit > 0 else max(10, round(20 * density_factor))

    indexed_lines = list(enumerate(lines))
    if total <= first_page_cap:
        return [indexed_lines]

    pages = []
    pages.append(indexed_lines[:first_page_cap])
    cursor = first_page_cap
    while cursor < total:
        pages.append(indexed_lines[cursor : cursor + next_page_cap])
        cursor += next_page_cap

    return pages


def render_invoice_html(
    invoice: InvoiceData,
    config: DesignConfig | None = None,
) -> str:
    """Render dokumen invoice lengkap dalam format HTML A4 multi-halaman.

    Menghasilkan HTML mandiri dengan CSS variables, 4 tema header,
    3 tema tabel, stempel status, serta @media print A4 presisi.
    """
    cfg = config or DesignConfig()
    editable_attr = 'contenteditable="true"' if cfg.content_editable else ""

    # Perhitungan Total, Diskon, Pajak
    subtotal = invoice.subtotal
    discount_amount = subtotal * (cfg.discount_rate / 100.0) if cfg.discount_rate > 0 else 0
    taxable_amount = subtotal - discount_amount
    tax_amount = taxable_amount * (cfg.tax_rate / 100.0) if cfg.tax_rate > 0 else 0
    grand_total = taxable_amount + tax_amount

    # Paginasi Halaman
    pages = paginate_invoice_lines(invoice.lines, cfg.items_per_page, cfg.row_padding)
    num_pages = len(pages)

    # Sender & Client info
    company_name = html.escape(invoice.company_name or "Perusahaan Anda")
    company_address = html.escape(invoice.company_address or "").replace("\n", "<br>")
    client_name = html.escape(invoice.client_name or "Klien")
    client_address = html.escape(invoice.client_address or "Alamat belum dicantumkan").replace("\n", "<br>")
    invoice_number = html.escape(invoice.po_number or invoice.label or "INV-001")
    invoice_date_text = html.escape(invoice.invoice_date_text or str(invoice.invoice_date))
    due_date_text = html.escape(cfg.due_date_text or invoice_date_text)
    doc_title = html.escape(cfg.doc_title or "INVOICE")
    bill_to_label = html.escape(cfg.bill_to_label or "Ditagihkan Kepada:")
    bank_info = html.escape(cfg.bank_info or "").replace("\n", "<br>")
    notes = html.escape(cfg.notes or "").replace("\n", "<br>")
    signer_name = html.escape(cfg.signer_name or company_name)

    # Generate CSS Variables
    css_vars = f"""
        --color-primary: {cfg.color_primary};
        --color-table-head: {cfg.color_table_head};
        --color-table-head-text: {cfg.color_table_head_text};
        --color-paper: {cfg.color_paper};
        --color-box-bg: {cfg.color_box_bg};
        --color-text: {cfg.color_text};
        --page-padding: {cfg.page_padding}px;
        --row-padding: {cfg.row_padding}px;
        --font-scale: {cfg.font_scale}%;
    """

    # Generate Font Family Class
    font_css = {
        "font-sans": "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
        "font-jakarta": "'Plus Jakarta Sans', Inter, sans-serif",
        "font-serif": "'Playfair Display', Georgia, serif",
        "font-mono": "'JetBrains Mono', Consolas, monospace",
    }.get(cfg.font_family, "Inter, sans-serif")

    # Render Stamp HTML
    stamp_html = ""
    if cfg.stamp_status != "none":
        stamp_text = cfg.stamp_custom_text if cfg.stamp_status == "CUSTOM" else cfg.stamp_status
        stamp_color = (
            "#059669" if cfg.stamp_status == "LUNAS"
            else ("#e11d48" if cfg.stamp_status == "BELUM BAYAR" else "#64748b")
        )
        stamp_html = f"""
        <div class="inv-status-stamp" style="
            position: absolute; top: 11rem; right: 3rem; border: 4px solid {stamp_color};
            color: {stamp_color}; font-weight: 800; font-size: 1.5rem; padding: 0.375rem 1.25rem;
            border-radius: 0.5rem; text-transform: uppercase; letter-spacing: 0.1em;
            opacity: 0.25; transform: rotate(-12deg); z-index: 10; pointer-events: none;
        ">{html.escape(stamp_text)}</div>
        """

    pages_html_list = []
    for page_idx, chunk in enumerate(pages):
        is_first_page = (page_idx == 0)
        is_last_page = (page_idx == num_pages - 1)

        # ------------------- HEADER COMPONENT -------------------
        if is_first_page:
            # Logo element
            logo_markup = ""
            if cfg.show_logo:
                if cfg.logo_data_url:
                    logo_markup = f'<img src="{cfg.logo_data_url}" alt="Logo" style="height:{cfg.logo_size}px; object-fit:contain; border-radius:4px;">'
                else:
                    initial = company_name[0].upper() if company_name else "I"
                    logo_markup = f"""
                    <div style="width:44px; height:44px; border-radius:10px; background-color:var(--color-primary); color:#ffffff;
                                display:flex; align-items:center; justify-content:center; font-weight:700; font-size:1.25rem; flex-shrink:0;">
                        {initial}
                    </div>
                    """

            # Header Style Variants: modern, banner, classic, minimal
            header_container_style = "display:flex; justify-content:space-between; align-items:flex-start; padding-bottom:1.5rem; gap:1rem;"
            if cfg.layout == "banner":
                header_container_style += f" background-color:var(--color-primary); color:#ffffff; padding:1.5rem; border-radius:0.75rem;"
            elif cfg.layout == "classic":
                header_container_style += f" border-bottom:4px double var(--color-primary);"
            elif cfg.layout == "minimal":
                header_container_style += " border-bottom:1px solid #cbd5e1;"
            else:  # modern
                header_container_style += f" border-bottom:2px solid var(--color-primary);"

            # Header Alignment Variants: default, swapped, centered
            flex_dir = "row"
            right_align = "text-align:right;"
            if cfg.header_align == "swapped":
                flex_dir = "row-reverse"
                right_align = "text-align:left;"
            elif cfg.header_align == "centered":
                flex_dir = "column"
                header_container_style += " align-items:center; text-align:center;"
                right_align = "text-align:center;"

            header_container_style += f" flex-direction:{flex_dir};"

            title_color = "#ffffff" if cfg.layout == "banner" else "var(--color-primary)"
            sender_color = "#ffffff" if cfg.layout == "banner" else "var(--color-text)"

            header_html = f"""
            <div class="inv-header-wrapper" style="{header_container_style}">
                <div style="display:flex; align-items:flex-start; gap:0.875rem; {'flex-direction:column; align-items:center;' if cfg.header_align == 'centered' else ''}">
                    {logo_markup}
                    <div>
                        <h1 {editable_attr} style="font-size:1.25rem; font-weight:700; color:{sender_color}; margin:0; letter-spacing:-0.025em;">{company_name}</h1>
                        <p {editable_attr} style="font-size:0.75rem; opacity:0.8; line-height:1.5; margin:0.25rem 0 0 0; max-width:22rem;">{company_address}</p>
                    </div>
                </div>

                <div style="{right_align}">
                    <span {editable_attr} style="font-size:1.75rem; font-weight:800; text-transform:uppercase; letter-spacing:-0.025em; color:{title_color}; display:block;">{doc_title}</span>
                    <div style="margin-top:0.5rem; font-size:0.75rem; line-height:1.4;">
                        <p style="margin:2px 0; opacity:0.8;"><span {editable_attr}>No. Dokumen:</span> <strong {editable_attr} style="font-family:'JetBrains Mono', monospace; font-size:0.875rem;">{invoice_number}</strong></p>
                        <p style="margin:2px 0; opacity:0.8;"><span {editable_attr}>Tanggal:</span> <span {editable_attr}>{invoice_date_text}</span></p>
                        {'<p style="margin:2px 0; opacity:0.8;"><span ' + editable_attr + '>Jatuh Tempo:</span> <span ' + editable_attr + '>' + due_date_text + '</span></p>' if cfg.show_due_date else ''}
                    </div>
                </div>
            </div>

            <!-- Bill To & Hero Total Box -->
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:1.5rem; padding:1.25rem 0;">
                <div>
                    <p {editable_attr} style="font-size:0.6875rem; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; opacity:0.5; margin:0 0 0.375rem 0;">{bill_to_label}</p>
                    <h3 {editable_attr} style="font-size:1rem; font-weight:700; margin:0;">{client_name}</h3>
                    <p {editable_attr} style="font-size:0.75rem; opacity:0.8; margin:0.25rem 0 0 0; line-height:1.5;">{client_address}</p>
                </div>
                <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:center;">
                    {f'''
                    <div class="inv-total-box" style="background-color:var(--color-box-bg); border:1px solid #e2e8f0; padding:0.75rem 1.25rem; text-align:right; border-radius:0.5rem;">
                        <p {editable_attr} style="font-size:0.6875rem; opacity:0.6; text-transform:uppercase; font-weight:600; margin:0;">Total Tagihan</p>
                        <p style="font-size:1.5rem; font-family:'JetBrains Mono', monospace; font-weight:700; margin:0.25rem 0 0 0; color:var(--color-text);">{format_rupiah_simple(grand_total)}</p>
                    </div>
                    ''' if cfg.show_hero_total else ''}
                </div>
            </div>
            """
        else:
            # Continuation Header for subsequent pages
            header_html = f"""
            <div class="inv-continuation-header" style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #e2e8f0; padding-bottom:0.75rem; margin-bottom:1.25rem; font-size:0.75rem;">
                <div style="display:flex; align-items:center; gap:0.5rem;">
                    <strong>{company_name}</strong>
                    <span style="opacity:0.4;">•</span>
                    <span style="opacity:0.7;">Lanjutan Rincian Item ({client_name})</span>
                </div>
                <div style="text-align:right; font-family:'JetBrains Mono', monospace;">
                    <strong>{invoice_number}</strong>
                    <span style="margin-left:0.5rem; padding:0.125rem 0.5rem; background:#f1f5f9; border-radius:0.25rem; color:#475569; font-size:0.6875rem;">Hal. {page_idx + 1} / {num_pages}</span>
                </div>
            </div>
            """

        # ------------------- TABLE COMPONENT -------------------
        # Table styles: zebra, bordered, clean
        table_border_style = ""
        th_extra_style = ""
        if cfg.layout == "classic":
            th_extra_style = f"border-top:2px solid var(--color-primary); border-bottom:2px solid var(--color-primary);"
        elif cfg.layout == "minimal":
            th_extra_style = f"background:transparent !important; color:var(--color-primary) !important; border-bottom:2px solid var(--color-primary);"

        rows_html = []
        for local_idx, (global_idx, item) in enumerate(chunk):
            row_bg = "transparent"
            if cfg.table_style == "zebra":
                row_bg = "#f8fafc" if global_idx % 2 == 1 else "transparent"

            cell_border = "border-bottom: 1px solid #e2e8f0;"
            if cfg.table_style == "bordered":
                cell_border = "border: 1px solid #cbd5e1;"

            pad_style = f"padding: {cfg.row_padding}px 12px;"

            row_markup = f"""
            <tr style="background-color:{row_bg}; {cell_border}">
                {f'<td style="{pad_style} color:#94a3b8; font-family:\'JetBrains Mono\', monospace; width:40px; text-align:center;">{global_idx + 1}</td>' if cfg.show_row_num else ''}
                <td {editable_attr} style="{pad_style} font-weight:500;">{html.escape(item.name)}</td>
                <td {editable_attr} style="{pad_style} text-align:center; font-family:'JetBrains Mono', monospace; width:60px;">{item.qty}</td>
                <td {editable_attr} style="{pad_style} text-align:right; font-family:'JetBrains Mono', monospace; width:120px;">{format_rupiah_simple(item.unit_price)}</td>
                <td style="{pad_style} text-align:right; font-family:'JetBrains Mono', monospace; font-weight:600; width:140px;">{format_rupiah_simple(item.amount)}</td>
            </tr>
            """
            rows_html.append(row_markup)

        table_html = f"""
        <table style="width:100%; border-collapse:collapse; font-size:0.75rem; text-align:left; margin-top:0.25rem;">
            <thead>
                <tr style="background-color:var(--color-table-head); color:var(--color-table-head-text); font-size:0.6875rem; text-transform:uppercase; letter-spacing:0.05em; {th_extra_style}">
                    {f'<th style="padding:10px 12px; font-weight:600; width:40px; text-align:center;">{cfg.th_num}</th>' if cfg.show_row_num else ''}
                    <th style="padding:10px 12px; font-weight:600;">{cfg.th_desc}</th>
                    <th style="padding:10px 12px; font-weight:600; text-align:center; width:60px;">{cfg.th_qty}</th>
                    <th style="padding:10px 12px; font-weight:600; text-align:right; width:120px;">{cfg.th_price}</th>
                    <th style="padding:10px 12px; font-weight:600; text-align:right; width:140px;">{cfg.th_total}</th>
                </tr>
            </thead>
            <tbody>
                {''.join(rows_html)}
            </tbody>
        </table>
        """

        # ------------------- FOOTER & SUMMARY COMPONENT -------------------
        if is_last_page:
            discount_row = ""
            if cfg.discount_rate > 0:
                discount_row = f"""
                <div style="display:flex; justify-content:space-between; padding:3px 0; opacity:0.8;">
                    <span><span {editable_attr}>Diskon</span> ({cfg.discount_rate}%)</span>
                    <span style="font-family:'JetBrains Mono', monospace; font-weight:500; color:#e11d48;">- {format_rupiah_simple(discount_amount)}</span>
                </div>
                """

            tax_row = ""
            if cfg.tax_rate > 0:
                tax_row = f"""
                <div style="display:flex; justify-content:space-between; padding:3px 0; opacity:0.8;">
                    <span><span {editable_attr}>PPN</span> ({cfg.tax_rate}%)</span>
                    <span style="font-family:'JetBrains Mono', monospace; font-weight:500;">{format_rupiah_simple(tax_amount)}</span>
                </div>
                """

            summary_and_footer_html = f"""
            <div style="margin-top:1.5rem; display:flex; justify-content:space-between; align-items:flex-start; gap:1.5rem;">
                {f'''
                <div class="inv-payment-box" style="background-color:var(--color-box-bg); padding:1rem; border:1px solid #e2e8f0; border-radius:0.5rem; max-width:22rem; width:100%; font-size:0.75rem;">
                    <p {editable_attr} style="font-weight:700; margin:0 0 0.375rem 0;">Instruksi Pembayaran</p>
                    <p {editable_attr} style="opacity:0.8; line-height:1.5; margin:0;">{bank_info}</p>
                    <p {editable_attr} style="opacity:0.7; margin:0.625rem 0 0 0; padding-top:0.5rem; border-top:1px solid #e2e8f0; font-style:italic;">{notes}</p>
                </div>
                ''' if cfg.show_payment_box else '<div style="flex:1;"></div>'}

                <div style="width:18rem; font-size:0.75rem;">
                    <div style="display:flex; justify-content:space-between; padding:3px 0; opacity:0.8;">
                        <span {editable_attr}>Subtotal</span>
                        <span style="font-family:'JetBrains Mono', monospace; font-weight:500;">{format_rupiah_simple(subtotal)}</span>
                    </div>
                    {discount_row}
                    {tax_row}
                    <div style="display:flex; justify-content:space-between; padding:0.625rem 0 0.375rem 0; border-top:2px solid var(--color-primary); font-size:0.875rem; font-weight:700; margin-top:0.25rem;">
                        <span {editable_attr}>TOTAL AKHIR</span>
                        <span style="font-family:'JetBrains Mono', monospace; font-size:1rem; color:var(--color-primary);">{format_rupiah_simple(grand_total)}</span>
                    </div>
                </div>
            </div>

            {f'''
            <div class="inv-footer-section" style="margin-top:2rem; padding-top:1rem; border-top:1px solid #e2e8f0; display:flex; justify-content:space-between; align-items:flex-end; font-size:0.75rem; opacity:0.8;">
                <div>
                    <p {editable_attr} style="margin:0;">Dokumen ini sah dan diterbitkan secara resmi oleh sistem.</p>
                    <p style="font-size:0.6875rem; opacity:0.7; margin:2px 0 0 0; font-family:'JetBrains Mono', monospace;">Halaman {page_idx + 1} dari {num_pages}</p>
                </div>
                {f'''
                <div class="inv-signature-box" style="text-align:center;">
                    <p {editable_attr} style="margin:0 0 2.5rem 0;">Hormat Kami,</p>
                    <p {editable_attr} style="font-weight:700; border-top:1px solid #cbd5e1; padding-top:0.25rem; min-width:120px; margin:0;">{signer_name}</p>
                </div>
                ''' if cfg.show_signature else ''}
            </div>
            ''' if cfg.show_footer else ''}
            """
        else:
            summary_and_footer_html = f"""
            <div style="margin-top:1.5rem; padding-top:0.75rem; border-top:1px dashed #cbd5e1; display:flex; justify-content:space-between; align-items:center; font-size:0.75rem; opacity:0.7;">
                <span>Menampilkan item {chunk[0][0] + 1} – {chunk[-1][0] + 1} dari total {len(invoice.lines)} item</span>
                <span style="font-weight:600;">Bersambung ke Halaman {page_idx + 2} →</span>
            </div>
            """

        page_container = f"""
        <div class="invoice-a4-page" style="
            position: relative;
            box-sizing: border-box;
            width: 210mm;
            min-height: 297mm;
            background-color: var(--color-paper);
            color: var(--color-text);
            padding: var(--page-padding);
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
            border-radius: 4px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            overflow: hidden;
            margin: 0 auto 2rem auto;
            page-break-after: always;
            break-after: page;
        ">
            {stamp_html}
            <div>
                {header_html}
                {table_html}
            </div>
            <div>
                {summary_and_footer_html}
            </div>
        </div>
        """
        pages_html_list.append(page_container)

    # Wrap in complete standalone HTML document
    return f"""<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{doc_title} - {invoice_number}</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&family=Playfair+Display:wght@600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            {css_vars}
        }}
        * {{
            box-sizing: border-box;
        }}
        body {{
            margin: 0;
            padding: 24px;
            background-color: #f1f5f9;
            font-family: {font_css};
            color: var(--color-text);
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }}
        .invoice-canvas {{
            display: flex;
            flex-direction: column;
            align-items: center;
            zoom: var(--font-scale);
        }}
        [contenteditable="true"]:hover {{
            outline: 1px dashed var(--color-primary);
            outline-offset: 2px;
            cursor: text;
        }}
        [contenteditable="true"]:focus {{
            outline: 2px solid var(--color-primary);
            outline-offset: 2px;
            background-color: rgba(37, 99, 235, 0.05);
        }}
        @media print {{
            body {{
                margin: 0 !important;
                padding: 0 !important;
                background: #ffffff !important;
            }}
            .no-print {{
                display: none !important;
            }}
            .invoice-canvas {{
                zoom: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
            }}
            .invoice-a4-page {{
                width: 210mm !important;
                min-height: 297mm !important;
                height: 297mm !important;
                margin: 0 !important;
                box-shadow: none !important;
                border: none !important;
                border-radius: 0 !important;
                page-break-after: always !important;
                break-after: page !important;
            }}
            .invoice-a4-page:last-child {{
                page-break-after: auto !important;
                break-after: auto !important;
            }}
            @page {{
                size: A4 portrait;
                margin: 0;
            }}
        }}
    </style>
</head>
<body>
    <div class="invoice-canvas">
        {''.join(pages_html_list)}
    </div>
</body>
</html>
"""
