import os
from docx import Document
from lxml import etree


def copy_header_footer_with_images(template_path, dest_path, output_path):
    """
    Copia cabeçalhos e rodapés de um documento modelo para um de destino,
    incluindo imagens e preservando a formatação (alinhamento, etc.).
    """
    print("Iniciando o processo de cópia de cabeçalho/rodapé com imagens...")
    try:
        template_doc = Document(template_path)
        dest_doc = Document(dest_path)
    except Exception as e:
        print(f"Erro ao abrir os documentos: {e}")
        return

    for i, dest_section in enumerate(dest_doc.sections):
        if i < len(template_doc.sections):
            template_section = template_doc.sections[i]
            print(f"\nProcessando Seção {i}...")
            _copy_part(template_doc, dest_doc, dest_section, template_section.header, 'header')
            _copy_part(template_doc, dest_doc, dest_section, template_section.footer, 'footer')

    try:
        dest_doc.save(output_path)
        print(f"\nProcesso concluído! Documento modificado salvo em: {output_path}")
    except Exception as e:
        print(f"Erro ao salvar o documento final: {e}")


# --- Função para copiar o conteúdo dos arquivos ---
def _copy_part(template_doc, dest_doc, dest_section, source_element, part_type):
    """
    Função auxiliar que copia o conteúdo de uma parte (cabeçalho ou rodapé)
    modificando o elemento XML de destino in-place para máxima compatibilidade.
    """
    if not source_element.paragraphs and not source_element.tables:
        print(f"  - {part_type.capitalize()} do template está vazio. Pulando.")
        return

    print(f"  - Copiando {part_type.capitalize()}...")

    if part_type == 'header':
        dest_element = dest_section.header
    else:
        dest_element = dest_section.footer

    # 1. Obter as partes de origem e destino
    source_part = source_element.part
    dest_part = dest_element.part

    # 2. Copiar relações de imagem e criar o mapa de rIds, necessário para levar as informações da imagem
    rid_map = {}
    for rel in source_part.rels.values():
        if "image" in rel.target_ref:
            old_rid = rel.rId
            image_part = rel.target_part
            new_rid = dest_part.relate_to(image_part, rel.reltype)
            rid_map[old_rid] = new_rid
            print(f"    - Imagem encontrada (rId: {old_rid}). Copiada e relacionada (novo rId: {new_rid}).")

    # 3. Limpar o conteúdo do elemento de destino
    dest_xml_element = dest_element._element
    dest_xml_element.clear()

    # 4. Copiar os filhos (parágrafos, tabelas) do elemento de origem para o de destino
    for child_element in source_element._element:
        # Cria uma cópia do elemento filho (ex: um parágrafo <w:p>)
        child_xml_string = etree.tostring(child_element, encoding='unicode')

        # Substitui os rIds das imagens na string XML deste filho
        for old_rid, new_rid in rid_map.items():
            child_xml_string = child_xml_string.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')

        # Converte a string de volta para um elemento XML
        new_child = etree.fromstring(child_xml_string)

        # Anexa o novo filho (com rIds corrigidos) ao elemento de destino
        dest_xml_element.append(new_child)


# --- Função Principal ---
if __name__ == "__main__":
    TEMPLATE_FILE = "1-cabecalho_e_rodape.docx"
    DEST_FILE = "2-original.docx"
    OUTPUT_FILE = "3-destino_com_cab_e_rod.docx"

    # apagar arquivo destino
    if os.path.exists(OUTPUT_FILE):
        os.remove(OUTPUT_FILE)

    copy_header_footer_with_images(TEMPLATE_FILE, DEST_FILE, OUTPUT_FILE)