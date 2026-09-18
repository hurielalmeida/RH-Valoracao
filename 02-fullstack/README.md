# RH Valoração — Full Stack V4.1

Aplicação local para operar o **Processador Universal V11** sem Google Colab. O notebook operacional do Colab permanece separado e não é alterado por este projeto.

## O que esta versão entrega

- interface web local simples;
- seleção de arquivos individuais;
- seleção de pasta inteira;
- arrastar e soltar;
- Excel `.xlsx` ou `.xlsm`;
- PDFs individuais;
- ZIP com Excel/PDF e subpastas;
- uma única base Excel obrigatória por execução;
- um ou vários PDFs na mesma execução;
- validação de arquivos e limite de 150 MB;
- processamento sequencial para evitar concorrência sobre o motor;
- diagnóstico por competência e layout;
- download do Excel atualizado;
- endpoint `/health` para diagnóstico técnico;
- proteção contra caminhos relativos perigosos e IDs de download inválidos;
- tratamento de erros para apresentação na interface;
- motor de processamento isolado em `core/`.

## Arquitetura

```text
Navegador
   ↓
frontend/index.html
   ↓ HTTP
backend/app.py (FastAPI)
   ↓
core/rh_valoracao_core.py (Processador Universal V11)
   ↓
Excel atualizado
```

O projeto foi deliberadamente separado do Colab para que a interface possa evoluir sem quebrar a ferramenta operacional já homologada.

## Como usar — quando esta versão for homologada

### Windows

1. Instale **Python 3.11 ou superior**.
2. Extraia esta pasta para um local permanente.
3. Dê duplo clique em `start_windows.bat`.
4. Na primeira execução, o programa cria um ambiente virtual e instala as dependências.
5. Abra no navegador `http://127.0.0.1:8000`.
6. Selecione **1 Excel-base** e os PDFs desejados, ou selecione uma pasta.
7. Clique em **Processar**.
8. Leia o diagnóstico apresentado.
9. Clique em **Baixar Excel atualizado**.
10. Feche a janela do terminal para encerrar o servidor.

### Linux/macOS

Execute `start_linux_mac.sh` em um terminal.

## Estrutura

- `core/` — motor V11 usado pelo aplicativo.
- `backend/` — API local FastAPI.
- `frontend/` — interface web.
- `runs/` — criada em tempo de execução para armazenar os arquivos de cada execução.
- `requirements.txt` — dependências.
- `start_windows.bat` — inicialização no Windows.
- `start_linux_mac.sh` — inicialização em Linux/macOS.
- `RELATORIO_TESTES.md` — registro de homologação técnica.

## Regras de entrada

A execução precisa conter:

- exatamente **1 Excel-base** (`.xlsx` ou `.xlsm`);
- pelo menos **1 PDF**, diretamente ou dentro de ZIP.

O sistema aceita arquivos dentro de pastas. Os caminhos relativos são preservados apenas durante a execução.

## Segurança e escopo

Esta versão é destinada a **uso local**. Não exponha a porta 8000 à internet.

Para uma futura versão multiusuário/servidor seriam necessários autenticação, autorização, limpeza automática de execuções, armazenamento apropriado, auditoria e hardening adicionais.

## Status do projeto

Esta é a versão consolidada do laboratório Full Stack. A V4.1 corrige o armazenamento temporário no Windows para evitar problemas quando o projeto estiver dentro do OneDrive e registra o traceback técnico em cada execução com erro. Antes do uso operacional definitivo, recomenda-se uma homologação no computador do usuário com uma cópia de trabalho real. O Colab V11 continua sendo a referência operacional atual.
