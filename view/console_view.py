class ConsoleView:
    """Responsável por toda a apresentação de informações ao usuário via console."""

    def start_process(self):
        # Inicia a exibição do processo.
        print("Iniciando processo de sincronização de layout e conteúdo...")

    def log_processing_section(self, index: int):
        # Log para indicar qual seção está sendo processada.
        print(f"\nProcessando Seção {index}...")

    def log_action(self, message: str):
        # Log para uma ação principal dentro de uma seção.
        print(f"  - {message}")

    def end_process(self, output_path: str):
        # Exibe a mensagem de conclusão.
        print(f"\nProcesso concluído! Documento modificado salvo em: {output_path}")

    def display_error(self, message: str):
        # Exibe uma mensagem de erro formatada.
        print(f"\nERRO: {message}")