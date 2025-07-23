import copy
from lxml import etree


class SectionProcessor:
    """Responsável por manipular as propriedades de uma seção do documento."""

    def __init__(self, source_section, dest_section, nsmap: dict, view):
        self._source_sectPr = source_section._sectPr
        self._dest_sectPr = dest_section._sectPr
        self._nsmap = nsmap
        self._view = view
        self._tags_to_copy = ['pgSz', 'pgMar', 'cols', 'docGrid', 'titlePg']

    def copy_properties(self):
        """Substitui as propriedades de layout da seção de destino pelas da origem."""
        self._view.log_action("Sincronizando propriedades da seção (margens, tamanho, etc.)...")

        # Remove propriedades antigas para uma substituição limpa
        for child in list(self._dest_sectPr):
            if hasattr(child, 'tag') and etree.QName(child).localname in self._tags_to_copy:
                self._dest_sectPr.remove(child)

        # Adiciona novas propriedades a partir do template
        for child in self._source_sectPr:
            if hasattr(child, 'tag') and etree.QName(child).localname in self._tags_to_copy:
                self._dest_sectPr.append(copy.deepcopy(child))

    def is_first_page_different(self) -> bool:
        """Verifica se a configuração 'Primeira Página Diferente' está ativa."""
        return self._source_sectPr.find('w:titlePg', namespaces=self._nsmap) is not None