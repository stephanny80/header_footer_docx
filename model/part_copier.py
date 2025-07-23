import copy
from lxml import etree
from .style_handler import StyleHandler


class PartCopier:
    """Copia o conteúdo de uma parte do documento (cabeçalho/rodapé)."""

    def __init__(self, source_element, dest_element, style_handler: StyleHandler, nsmap: dict, part_name: str, view):
        self._source_element = source_element
        self._dest_element = dest_element
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

        rid_map = self._copy_image_relationships()
        dest_xml_element = self._dest_element._element
        dest_xml_element.clear()

        for child_element in self._source_element._element:
            processed_element = self._process_child_element(child_element, rid_map)
            dest_xml_element.append(processed_element)

    def _copy_image_relationships(self) -> dict:
        """Copia relações de imagem e retorna um mapa de IDs antigos para novos."""
        rid_map = {}
        source_part = self._source_element.part
        dest_part = self._dest_element.part
        for rel in source_part.rels.values():
            if "image" in rel.target_ref:
                rid_map[rel.rId] = dest_part.relate_to(rel.target_part, rel.reltype)
        return rid_map

    def _process_child_element(self, child_element, rid_map: dict):
        """Processa um único elemento filho (parágrafo, tabela), aplicando estilos e corrigindo referências."""
        new_child_element = copy.deepcopy(child_element)

        if etree.QName(new_child_element).localname == 'p':
            self._style_handler.inline_paragraph_style(new_child_element)

        child_xml_string = etree.tostring(new_child_element, encoding='unicode')
        for old_rid, new_rid in rid_map.items():
            child_xml_string = child_xml_string.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')

        return etree.fromstring(child_xml_string)