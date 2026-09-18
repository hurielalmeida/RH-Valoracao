# Relatório de testes — RH Valoração Full Stack V4

## Escopo

A V4 foi validada como aplicação local separada do Colab, usando o motor universal V11 incorporado em `core/`.

## Testes executados

### 1. Interface e serviço
- `GET /` → **PASS** (HTTP 200; página RH Valoração carregada).
- `GET /health` → **PASS** (serviço identificado; versão 4.0).
- JavaScript da interface → **PASS** (`node --check`).

### 2. Processamento multi-layout
Entrada: 1 Excel-base + 2 PDFs.
- Log Planning, competência 01/2025 → **5/5 encontrados**.
- Datasul, competência 12/2025 → **8/8 encontrados**.
- Excel final → **PASS**.
- Download do resultado → **PASS**.

### 3. ZIP
Um ZIP contendo Excel e os dois PDFs, em subpastas.
- Extração → **PASS**.
- Identificação de 1 Excel + 2 PDFs → **PASS**.
- Processamento → **PASS**.
- Download → **PASS**.

### 4. Estrutura de pastas
Arquivos enviados com caminhos relativos simulando seleção de pasta.
- Preservação dos caminhos durante a execução → **PASS**.
- Processamento → **PASS**.

### 5. Validações
- Extensão não aceita → **PASS**, HTTP 400.
- Dois Excels → **PASS**, HTTP 400.
- ID de download inválido → **PASS**, HTTP 400.
- ZIP inválido → **PASS**, HTTP 400.

### 6. Regressão MG Info 2024
Entrada: Excel-base histórico + 12 PDFs já homologados.
- Meses já existentes foram ignorados sem sobrescrita.
- 06/2024 → **74/74**.
- 07/2024 → **73/73**.
- 08/2024 → **74/74**.
- 09/2024 → **74/74**.
- 10/2024 → **68/68**.
- 11/2024 → **67/67**.
- 12/2024 → **57/57**.
- Resultado final → **738 linhas × 98 colunas**.
- Download → **PASS**.

## Conclusão

A V4 passou pelos testes funcionais, de validação, multi-layout, ZIP, pasta e regressão do conjunto MG Info 2024 utilizado na homologação do motor V11.

A aplicação permanece destinada a uso **local**. A próxima etapa, quando o usuário decidir iniciar a implantação, é executar a instalação/homologação no computador real seguindo o `README.md`. Não é necessário alterar o notebook Colab operacional.
