from docx import Document
from view.console_view import ConsoleView
from model.style_handler import StyleHandler
from model.section_processor import SectionProcessor
from model.part_copier import PartCopier


class DocxSynchronizer:
    """
    Controlador que orquestra a sincronização de layout de um documento Word
    de template para um documento de destino.
    """
    _NSMAP = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    def __init__(self, template_path: str, dest_path: str):
        self.view = ConsoleView()
        try:
            self.template_doc = Document(template_path)
            self.dest_doc = Document(dest_path)
            self.style_handler = StyleHandler(self.template_doc, self._NSMAP)
        except Exception as e:
            self.view.display_error(f"Falha ao abrir os documentos: {e}")
            raise

    def synchronize(self):
        """Executa o processo de sincronização para todas as seções."""
        self.view.start_process()
        for i, dest_section in enumerate(self.dest_doc.sections):
            if i >= len(self.template_doc.sections):
                break

            self.view.log_processing_section(i)
            template_section = self.template_doc.sections[i]

            section_processor = SectionProcessor(template_section, dest_section, self._NSMAP, self.view)
            section_processor.copy_properties()

            self._process_section_parts(template_section, dest_section)

    def _process_section_parts(self, template_section, dest_section):
        """Processa a cópia de todas as partes relevantes de uma seção."""
        parts_to_copy = {
            'default_header': (template_section.header, dest_section.header),
            'default_footer': (template_section.footer, dest_section.footer),
        }

        section_processor = SectionProcessor(template_section, dest_section, self._NSMAP, self.view)
        if section_processor.is_first_page_different():
            self.view.log_action("'Primeira Página Diferente' detectado. Processando partes específicas.")
            parts_to_copy.update({
                'first_page_header': (template_section.first_page_header, dest_section.first_page_header),
                'first_page_footer': (template_section.first_page_footer, dest_section.first_page_footer),
            })

        for part_name, (source_el, dest_el) in parts_to_copy.items():
            copier = PartCopier(
                source_element=source_el,
                dest_element=dest_el,
                dest_doc_part=self.dest_doc.part,
                style_handler=self.style_handler,
                nsmap=self._NSMAP,
                part_name=part_name,
                view=self.view
            )
            copier.copy_content()

    def save(self, output_path: str):
        """Salva o documento de destino modificado."""
        try:
            self.dest_doc.save(output_path)
            self.view.end_process(output_path)
        except Exception as e:
            self.view.display_error(f"Falha ao salvar o documento final: {e}")
            raise