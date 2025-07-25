import copy
from lxml import etree
from .style_handler import StyleHandler
from docx.image.image import Image  # Importação necessária para o objeto Image


class PartCopier:
    """Copia o conteúdo de uma parte do documento (cabeçalho/rodapé)."""

    def __init__(self, source_element, dest_element, dest_doc_part, style_handler: StyleHandler, nsmap: dict,
                 part_name: str, view):
        self._source_element = source_element
        self._dest_element = dest_element
        self._dest_doc_part = dest_doc_part
        self._style_handler = style_handler
        self._nsmap = nsmap
        self._part_name = part_name
        self._view = view

    def copy_content(self):
        """Executa o processo completo de cópia de conteúdo para esta parte."""
        if not self._source_element.paragraphs and not self._source_element.tables:
            self._view.log_action(f"Parte '{self._part_name}' está vazia no template. Pulando.")
            return

        self._view.log_action(f"Copiando conteúdo da parte: {self._part_name}...")

        rid_map = self._copy_relationships()
        dest_xml_element = self._dest_element._element
        dest_xml_element.clear()

        for child_element in self._source_element._element:
            processed_element = self._process_child_element(child_element, rid_map)
            dest_xml_element.append(processed_element)

    def _get_next_rId(self, part) -> str:
        """Gera manualmente o próximo ID de relacionamento (rId) disponível para uma parte."""
        rIds = part.rels.keys()
        if not rIds:
            return "rId1"

        max_id_num = 0
        for rId in rIds:
            if rId.startswith("rId"):
                try:
                    num = int(rId[3:])
                    if num > max_id_num:
                        max_id_num = num
                except ValueError:
                    continue
        return f"rId{max_id_num + 1}"

    def _get_or_add_image_part_by_hash(self, source_image_part):
        """
        Adiciona uma parte de imagem ao pacote de destino, evitando duplicatas
        através da verificação do hash SHA1 da imagem. Esta é a abordagem robusta.
        """
        source_image = Image.from_blob(source_image_part.blob)
        image_collection = self._dest_doc_part.package.image_parts

        # Procura por uma imagem existente com o mesmo hash
        for image_part in image_collection:
            if image_part.sha1 == source_image.sha1:
                return image_part  # Encontrou a imagem, reutiliza a parte existente

        # Se não encontrou, cria uma nova parte de imagem usando chamada interno
        return image_collection._add_image_part(source_image)

    def _copy_relationships(self) -> dict:
        """
        Copia todas as relações relevantes (imagens, hiperlinks) e retorna
        um mapa de IDs antigos para novos.
        """
        rid_map = {}
        source_part = self._source_element.part
        dest_part = self._dest_element.part

        for rId, rel in source_part.rels.items():
            if "image" in rel.reltype:
                source_image_part = rel.target_part

                # 1. Usa o chamada auxiliar para adicionar ou obter a imagem de forma segura por hash
                new_image_part = self._get_or_add_image_part_by_hash(source_image_part)

                # 2. Cria a relação local (do cabeçalho para a imagem)
                new_rId = dest_part.relate_to(new_image_part, rel.reltype)
                rid_map[rId] = new_rId

            elif "hyperlink" in rel.reltype:
                new_rId = self._get_next_rId(dest_part)
                dest_part.rels.add_relationship(rel.reltype, rel.target_ref, new_rId, is_external=True)
                rid_map[rId] = new_rId
        return rid_map

    def _process_child_element(self, child_element, rid_map: dict):
        """
        Processa um único elemento filho (parágrafo, tabela), aplicando estilos
        e corrigindo as referências de relacionamento (rIds).
        """
        new_child_element = copy.deepcopy(child_element)

        if etree.QName(new_child_element).localname == 'p':
            self._style_handler.inline_paragraph_style(new_child_element)

        lxml_element = etree.fromstring(etree.tostring(new_child_element))

        for blip_el in lxml_element.xpath('.//a:blip', namespaces=self._nsmap):
            old_rid = blip_el.get(f'{{{self._nsmap["r"]}}}embed')
            if old_rid in rid_map:
                blip_el.set(f'{{{self._nsmap["r"]}}}embed', rid_map[old_rid])

        for hlink_el in lxml_element.xpath('.//w:hyperlink', namespaces=self._nsmap):
            old_rid = hlink_el.get(f'{{{self._nsmap["r"]}}}id')
            if old_rid in rid_map:
                hlink_el.set(f'{{{self._nsmap["r"]}}}id', rid_map[old_rid])

        return lxml_element