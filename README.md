# Automação de Injeção de Cabeçalho e Rodapé em Documentos Word

## Descrição do Projeto

Este projeto Python oferece uma solução para automatizar a injeção de cabeçalhos e rodapés de um documento Word (.docx) modelo em um documento Word de destino. A ferramenta é útil para padronizar a identidade visual de múltiplos documentos, garantindo que logotipos, informações de contato, numeração de páginas e outros elementos padrão de cabeçalho e rodapé sejam aplicados de forma consistente, sem a necessidade de edição manual.

A implementação lida com a complexidade do formato `.docx` (que é um arquivo ZIP contendo XMLs e mídias), extraindo e injetando as partes XML e os arquivos de imagem associados aos cabeçalhos e rodapés.

## Funcionalidades

  * **Injeção de Cabeçalhos:** Copia o conteúdo completo do cabeçalho da primeira seção do documento modelo para o documento de destino.
  * **Injeção de Rodapés:** Copia o conteúdo completo do rodapé da primeira seção do documento modelo para o documento de destino.
  * **Suporte a Imagens:** Tenta copiar imagens presentes nos cabeçalhos e rodapés do template, garantindo que os arquivos de mídia e seus relacionamentos sejam corretamente transferidos para o documento de destino.
  * **Formatação Básica:** Preserva a formatação básica de texto (negrito, itálico, sublinhado, tamanho e cor da fonte) dos parágrafos copiados.
  * **Manipulação de Arquivos DOCX:** Opera diretamente na estrutura ZIP e XML dos arquivos `.docx` para maior robustez na cópia de elementos complexos.

## Pré-requisitos

Para executar este projeto, você precisará ter o Python instalado em sua máquina (Ubuntu ou qualquer outro sistema operacional).

  * **Python 3.x:**
    ```bash
    python3 --version
    ```
    Se não tiver, instale com:
    ```bash
    sudo apt update
    sudo apt install python3 python3-pip
    ```
  * **Pip (gerenciador de pacotes do Python):**
    ```bash
    pip3 --version
    ```
  * **Bibliotecas Python:**
    Instale as bibliotecas necessárias usando `pip`:
    ```bash
    pip3 install python-docx lxml Pillow
    ```
      * `python-docx`: Para manipulação de alto nível de documentos Word.
      * `lxml`: Para manipulação eficiente de XML, usada para extrair e injetar as partes do DOCX.
      * `Pillow`: Necessária para a geração de imagens de teste no script de exemplo, se você não tiver arquivos de imagem próprios.

## Como Usar

Siga os passos abaixo para preparar e executar o script:

1.  **Prepare os Documentos Word:**
    O script espera encontrar dois arquivos na mesma pasta:

      * **`1-cabecalho_e_rodape.docx`**: Este será o documento no qual os cabeçalhos e rodapés do template serão injetados.
      * **`2-original.docx`**: Este será o seu documento modelo, contendo o cabeçalho e o rodapé que você deseja injetar. Certifique-se de que ele tenha um cabeçalho e/ou rodapé configurado com o conteúdo desejado, incluindo texto e imagens (se aplicável).

2.  **Execute o Script:**
    Abra o terminal, navegue até a pasta onde você salvou o script e os documentos Word, e execute o comando:

    ```bash
    python3 main.py
    ```

3.  **Verifique o Resultado:**
    Após a execução, um novo arquivo chamado `3-destino_com_cab_e_rod.docx` será criado na mesma pasta.
    
## Notas Importantes e Limitações

  * **Fidelidade de Formatação:** A grande força deste script é sua capacidade de preservar a formatação. Ao copiar diretamente os nós XML de parágrafos (<w:p>) e tabelas (<w:tbl>), ele mantém com sucesso propriedades como alinhamento, espaçamento, fontes e cores definidas no template.
  * **Substituição vs. Fusão:** O script *substituir* o conteúdo dos cabeçalhos e rodapés padrão (`header1.xml` e `footer1.xml`) do documento de destino pelo conteúdo do template. Se o seu caso de uso exigir a fusão ou combinação de conteúdo existente com o do template, a lógica seria mais complexa.
  * **Formatação de Header e Footer:** Cópia de configuração de formatação de régua e tabulação.
  * **Tratamento de Configuração de Primeira Página:** Tratamento diferenciado para entender e copiar a configuração de header e footer da primeira página.
  * **MVC Architecture:** Projeto utilizando arquitetura MVC, com pouca carga intrisseca e aplicação de conceitos de SOLID e clean code.

-----