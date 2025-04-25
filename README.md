# Gerar HTML a partir de Excel

Este projeto é um script em C# que lê um arquivo Excel (`Excel.xlsx`) e gera uma representação do seu conteúdo em uma tabela HTML estilizada com Bootstrap.

## Funcionalidades

* **Leitura de arquivos Excel:** Permite ler dados de arquivos `.xlsx`.
* **Geração de HTML:** Converte o conteúdo da primeira planilha do Excel em uma tabela HTML.
* **Suporte a células mescladas:** Identifica e renderiza corretamente células mescladas do Excel, utilizando os atributos `colspan` e `rowspan` no HTML.
* **Estilização com Bootstrap:** A tabela HTML é estilizada com as classes do Bootstrap para uma aparência responsiva e moderna.
* **Tratamento de múltiplos valores em células mescladas:** Se uma célula mesclada no Excel contiver uma coleção de valores, eles serão concatenados em um único texto na célula HTML.

## Pré-requisitos

* **Visual Studio:** Necessário para compilar e executar o projeto C#.
* **Pacote NuGet OfficeOpenXml:** Este pacote é utilizado para trabalhar com arquivos Excel. Ele será instalado automaticamente ao compilar o projeto no Visual Studio.
* **Arquivo Excel (`Excel.xlsx`):** O script espera encontrar um arquivo chamado `Excel.xlsx` no seguinte caminho: `C:\\Users\\danielcaldeira\\Desktop\\`. Certifique-se de que o arquivo exista neste local ou modifique o caminho no código.

## Como Usar

1.  **Clone ou faça o download deste repositório.**
2.  **Abra a solução (`.sln`) no Visual Studio.**
3.  **Verifique se o pacote NuGet `OfficeOpenXml` está instalado.** Caso contrário, restaure os pacotes NuGet.
4.  **Certifique-se de ter um arquivo Excel chamado `Excel.xlsx` no caminho `C:\\Users\\danielcaldeira\\Desktop\\` com os dados que você deseja converter.**
5.  **Execute o projeto (por exemplo, pressionando F5 no Visual Studio).**
6.  **O script irá gerar uma string HTML que representa a tabela do seu arquivo Excel.** Atualmente, essa string é criada na memória, mas não está sendo salva em um arquivo. Você pode modificar o código para salvar essa string em um arquivo `.html` se desejar.

## Observações

* O script atualmente lê apenas a primeira planilha do arquivo Excel.
* O caminho do arquivo Excel está fixo no código. Para maior flexibilidade, você pode considerar parametrizar esse caminho.
* O HTML gerado inclui links para as bibliotecas JavaScript do Bootstrap, mas elas não parecem ser utilizadas para a renderização da tabela em si. Elas podem ser removidas se não houver intenção de adicionar interatividade ao HTML posteriormente.
* A string HTML gerada passa por uma substituição de aspas simples por aspas duplas.

## Próximos Passos (Sugestões)

* **Salvar o HTML em um arquivo:** Modificar o código para salvar a string HTML gerada em um arquivo `.html`.
* **Permitir a seleção da planilha:** Adicionar a opção de especificar qual planilha do Excel deve ser convertida.
* **Tornar o caminho do arquivo configurável:** Permitir que o usuário especifique o caminho do arquivo Excel a ser lido.
* **Adicionar opções de formatação:** Incluir opções para personalizar a formatação da tabela HTML (cores, estilos, etc.).
* **Tratar diferentes tipos de dados:** Melhorar o tratamento de diferentes tipos de dados do Excel (datas, números, etc.) para garantir uma representação HTML adequada.

Sinta-se à vontade para usar e modificar este README conforme necessário para o seu projeto! 😊 Se tiver mais alguma dúvida ou precisar de ajuda com alguma modificação no código, é só me dizer!
