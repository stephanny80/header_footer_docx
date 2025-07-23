import copy
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree


class StyleHandler:
    """Classe utilitária com a única responsabilidade de lidar com estilos do Word."""

    def __init__(self, template_doc, nsmap: dict):
        self._template_doc = template_doc
        self._nsmap = nsmap
        self._pPr_cache = {}  # Cache para propriedades de estilo de parágrafo

    def _fetch_style_pPr(self, style_id: str):
        """Busca (com cache) as propriedades de um estilo de parágrafo pelo seu ID."""
        if style_id in self._pPr_cache:
            return self._pPr_cache[style_id]

        try:
            style = self._template_doc.styles[style_id]
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                pPr = style.element.find('.//w:pPr', namespaces=self._nsmap)
                self._pPr_cache[style_id] = pPr
                return pPr
        except KeyError:
            pass  # Estilo não encontrado no documento

        self._pPr_cache[style_id] = None
        return None

    def inline_paragraph_style(self, p_element):
        """
        Copia as propriedades de um estilo referenciado (como paradas de tabulação)
        diretamente para o XML do parágrafo, tornando-o autossuficiente ("inlining").
        """
        pPr = p_element.find('w:pPr', namespaces=self._nsmap)
        if pPr is None:
            pPr = etree.Element(f'{{{self._nsmap["w"]}}}pPr')
            p_element.insert(0, pPr)

        style_tag = pPr.find('w:pStyle', namespaces=self._nsmap)
        if style_tag is not None:
            style_id = style_tag.get(f'{{{self._nsmap["w"]}}}val')
            style_pPr = self._fetch_style_pPr(style_id)

            if style_pPr is not None:
                # Itera sobre as propriedades do estilo (ex: <w:tabs>)
                for style_prop in style_pPr:
                    prop_tag_name = etree.QName(style_prop).localname
                    # Se o parágrafo não tem essa propriedade, herda do estilo
                    if pPr.find(f'w:{prop_tag_name}', namespaces=self._nsmap) is None:
                        pPr.insert(0, copy.deepcopy(style_prop))