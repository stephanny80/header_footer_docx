import os
import copy
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree

# Define a comprehensive namespace map used throughout the script
nsmap_comprehensive = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}


def copy_layout_and_header_footer(template_path, dest_path, output_path):
    """
    Copia o LAYOUT DA PÁGINA, ESTILOS, e o conteúdo de CABEÇALHOS/RODAPÉS
    (incluindo a configuração de "Primeira Página Diferente") de um documento
    modelo para um de destino, preservando tudo.
    """
    print("Iniciando processo de cópia de layout e conteúdo...")
    try:
        template_doc = Document(template_path)
        dest_doc = Document(dest_path)
    except Exception as e:
        print(f"Erro ao abrir os documentos: {e}")
        return

    style_pPr_cache = {}

    for i, dest_section in enumerate(dest_doc.sections):
        if i < len(template_doc.sections):
            template_section = template_doc.sections[i]
            print(f"\nProcessando Seção {i}...")

            # Copiar propriedades da seção (margens, tamanho, e a configuração de primeira página)
            _copy_section_properties(template_section, dest_section)

            # Copiar o cabeçalho/rodapé padrão (para páginas 2 em diante)
            _copy_part(template_doc, dest_doc, dest_section, template_section.header, 'default_header', style_pPr_cache)
            _copy_part(template_doc, dest_doc, dest_section, template_section.footer, 'default_footer', style_pPr_cache)

            # Copiar o cabeçalho/rodapé da primeira página, se a configuração estiver ativa no template
            # A verificação correta é pela presença do elemento XML 'titlePg'
            if template_section._sectPr.titlePg is not None:
                print("  - 'Primeira Página Diferente' detectado. Copiando conteúdo específico.")
                _copy_part(template_doc, dest_doc, dest_section, template_section.first_page_header,
                           'first_page_header', style_pPr_cache)
                _copy_part(template_doc, dest_doc, dest_section, template_section.first_page_footer,
                           'first_page_footer', style_pPr_cache)

    try:
        dest_doc.save(output_path)
        print(f"\nProcesso concluído! Documento modificado salvo em: {output_path}")
    except Exception as e:
        print(f"Erro ao salvar o documento final: {e}")


def _get_style_pPr(doc, style_id, cache):
    """Busca e armazena em cache o elemento pPr de um estilo de parágrafo."""
    if style_id in cache:
        return cache[style_id]

    try:
        style = doc.styles[style_id]
        if style.type == WD_STYLE_TYPE.PARAGRAPH:
            pPr = style.element.find('.//w:pPr', namespaces=nsmap_comprehensive)
            cache[style_id] = pPr
            return pPr
    except KeyError:
        pass
    cache[style_id] = None
    return None


def _copy_section_properties(source_section, dest_section):
    """Copia as propriedades de layout da seção (margens, tamanho, orientação, etc.)."""
    print("  - Copiando propriedades da seção (margens, tamanho da página, etc.)...")
    source_sectPr = source_section._sectPr
    dest_sectPr = dest_section._sectPr
    # Adicionado 'titlePg' para copiar a configuração "Primeira Página Diferente"
    tags_to_copy = ['pgSz', 'pgMar', 'cols', 'docGrid', 'titlePg']

    for child in list(dest_sectPr):
        if hasattr(child, 'tag') and etree.QName(child).localname in tags_to_copy:
            dest_sectPr.remove(child)

    for child in source_sectPr:
        if hasattr(child, 'tag') and etree.QName(child).localname in tags_to_copy:
            dest_sectPr.append(copy.deepcopy(child))


def _copy_part(template_doc, dest_doc, dest_section, source_element, part_type, style_cache):
    """Copia o conteúdo de uma parte, internalizando estilos para preservar a formatação da régua/tabulação."""
    if not source_element.paragraphs and not source_element.tables:
        print(f"  - A parte '{part_type}' está vazia no template. Pulando.")
        return

    print(f"  - Copiando conteúdo da parte: {part_type}...")

    if part_type == 'default_header':
        dest_element = dest_section.header
    elif part_type == 'default_footer':
        dest_element = dest_section.footer
    elif part_type == 'first_page_header':
        dest_element = dest_section.first_page_header
    elif part_type == 'first_page_footer':
        dest_element = dest_section.first_page_footer
    else:
        return

    # 1. Obter as partes de origem e destino
    source_part = source_element.part
    dest_part = dest_element.part

    # 2. Copiar relações de imagem e criar o mapa de rIds, necessário para levar as informações da imagem
    rid_map = {}
    for rel in source_part.rels.values():
        if "image" in rel.target_ref:
            rid_map[rel.rId] = dest_part.relate_to(rel.target_part, rel.reltype)

    # 3. Limpar o conteúdo do elemento de destino
    dest_xml_element = dest_element._element
    dest_xml_element.clear()

    # 4. Copiar os filhos (parágrafos, tabelas) do elemento de origem para o de destino
    for child_element in source_element._element:
        # Cria uma cópia do elemento filho (ex: um parágrafo <w:p>)
        new_child_docx_el = copy.deepcopy(child_element)

        if etree.QName(new_child_docx_el).localname == 'p':
            pPr = new_child_docx_el.find('w:pPr', namespaces=nsmap_comprehensive)
            if pPr is None:
                pPr = etree.Element(f'{{{nsmap_comprehensive["w"]}}}pPr')
                new_child_docx_el.insert(0, pPr)

            style_tag = pPr.find('w:pStyle', namespaces=nsmap_comprehensive)
            if style_tag is not None:
                style_id = style_tag.get(f'{{{nsmap_comprehensive["w"]}}}val')
                style_pPr = _get_style_pPr(template_doc, style_id, style_cache)

                if style_pPr is not None:
                    for style_prop in style_pPr:
                        prop_tag_name = etree.QName(style_prop).localname
                        if pPr.find(f'w:{prop_tag_name}', namespaces=nsmap_comprehensive) is None:
                            pPr.insert(0, copy.deepcopy(style_prop))

        # Cria uma cópia do elemento filho
        child_xml_string = etree.tostring(new_child_docx_el, encoding='unicode')

        # Substitui os rIds das imagens na string XML deste filho
        for old_rid, new_rid in rid_map.items():
            child_xml_string = child_xml_string.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')

        # Converte a string de volta para um elemento XML
        final_lxml_el = etree.fromstring(child_xml_string)

        # Anexa o novo filho (com rIds corrigidos) ao elemento de destino
        dest_xml_element.append(final_lxml_el)


# --- Função Principal ---
if __name__ == "__main__":
    TEMPLATE_FILE = "header1.docx"
    DEST_FILE = "original1.docx"
    OUTPUT_FILE = "resultado1.docx"

    # apagar arquivo destino
    if os.path.exists(OUTPUT_FILE):
        os.remove(OUTPUT_FILE)

    copy_layout_and_header_footer(TEMPLATE_FILE, DEST_FILE, OUTPUT_FILE)