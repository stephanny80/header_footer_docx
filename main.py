# main.py

import os
from controller.docx_synchronizer import DocxSynchronizer


def run_synchronization(template_file: str, source_file: str, output_file: str):
    """Função principal que encapsula a lógica de execução."""
    if not all(os.path.exists(p) for p in [template_file, source_file]):
        print(f"ERRO: Arquivo de template '{template_file}' ou de origem '{source_file}' não encontrado.")
        return

    if os.path.exists(output_file):
        os.remove(output_file)
        print(f"Arquivo de saída existente '{output_file}' removido.")

    try:
        # O Controlador gerencia o processo inteiro
        controller = DocxSynchronizer(template_file, source_file)
        controller.synchronize()
        controller.save(output_file)
    except Exception as e:
        print(f"Ocorreu um erro fatal durante a execução: {e}")


if __name__ == "__main__":
    TEMPLATE_FILE = "header1.docx"
    DEST_FILE = "original1.docx"
    OUTPUT_FILE = "resultado1.docx"

    run_synchronization(TEMPLATE_FILE, DEST_FILE, OUTPUT_FILE)