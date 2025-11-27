# DocsDownloader-SubC

Este projeto é uma aplicação com interface gráfica (GUI) desenvolvida em Python para facilitar a visualização de dados e o download de anexos de listas do SharePoint. Ele utiliza scripts em PowerShell em segundo plano para realizar a comunicação com o SharePoint.

## Funcionalidades

- **Interface Gráfica**: Visualização amigável dos dados do SharePoint.
- **Conexão Segura**: Utiliza `PnP.PowerShell` para autenticação e conexão.
- **Download de Anexos**: Permite selecionar itens e baixar seus anexos automaticamente.
- **Exportação de Dados**: Gera relatórios em Excel.

## Como Usar (Recomendado)

A maneira mais fácil de executar o projeto é utilizando o script de inicialização automática:

1. Localize o arquivo **`iniciar.bat`** na raiz do projeto.
2. Dê um **duplo clique** nele.
3. O script irá tentar executar a aplicação utilizando o Python instalado no sistema ou o Anaconda.

## Atualização

O arquivo **`att.bat`** é responsável por atualizar a aplicação. Ele realiza o download da versão mais recente do repositório e substitui os arquivos no diretório de destino (configurado como `C:\DocsDownloader-SubC`).

## Scripts Principais

### `src/downloadFiles.py`
A aplicação principal em Python. Responsável pela interface gráfica e por orquestrar as chamadas aos scripts de PowerShell.

### `src/downloadAttachments.ps1`
Script PowerShell invocado pela aplicação Python para realizar o download efetivo dos anexos dos itens selecionados.

### `src/exportAllColumns.ps1`
Script PowerShell para exportar dados completos das listas do SharePoint.

## Pré-requisitos

- **Python 3.x** ou **Anaconda** instalado.
- **PowerShell 5.1** ou superior.
- Módulos do PowerShell (instalados automaticamente se necessário, mas listados aqui para referência):
    - `PnP.PowerShell`
    - `ImportExcel`
- Bibliotecas Python (listadas em `requirements.txt`):
    - `pandas`
    - `openpyxl`

## Instalação Manual (Opcional)

Caso prefira não usar o `iniciar.bat`, você pode preparar o ambiente manualmente:

1. Instale as dependências Python:
   ```bash
   pip install -r requirements.txt
   ```
2. Instale os módulos do PowerShell:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser -Force
   Install-Module ImportExcel -Scope CurrentUser -Force
   ```
3. Execute a aplicação:
   ```bash
   python src/downloadFiles.py
   ```

## Notas

- Certifique-se de ter as permissões necessárias para acessar os sites e listas do SharePoint.
- A primeira execução pode solicitar autenticação no SharePoint via navegador.