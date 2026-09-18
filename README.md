# RH Valoração

> Transformando um processo manual de leitura e estruturação de folhas de pagamento em uma solução automatizada com Python.

Projeto de automação desenvolvido a partir de uma necessidade real de transformar informações presentes em folhas de pagamento em PDF em uma base estruturada de RH em Excel.

O projeto começou como uma atividade operacional, passou por diferentes abordagens de automação e evoluiu até uma aplicação em Python com processamento de PDFs, interface web local e versão desktop para Windows.

---

## 🎯 O problema

O processo original envolvia a leitura de folhas de pagamento em PDF e a transferência manual das informações para uma estrutura de Excel já existente.

Entre os dados tratados estavam:

- Colaboradores
- Cargos
- Códigos
- Rubricas
- Valores
- Competências
- Informações cadastrais necessárias para a estrutura da base

Dependendo da quantidade de colaboradores, meses e rubricas, essa atividade poderia consumir várias horas ou até dias de trabalho.

A proposta do projeto foi automatizar principalmente a etapa de **leitura, transformação e estruturação dos dados**, mantendo a análise humana necessária para as etapas posteriores.

---

## 💡 A solução

A aplicação recebe uma base Excel e uma ou várias folhas de pagamento em PDF.

O processamento identifica a estrutura dos documentos, extrai as informações relevantes e utiliza esses dados para atualizar uma nova versão da base Excel.

Também é possível trabalhar com arquivos ZIP contendo múltiplas folhas de pagamento.

### Fluxo simplificado

```text
Folhas de pagamento
        │
        ▼
     PDF / ZIP
        │
        ▼
 Processamento em Python
        │
        ├── Identificação da estrutura
        ├── Extração dos dados
        ├── Identificação de colaboradores
        ├── Identificação de rubricas e códigos
        └── Organização dos valores
        │
        ▼
 Base Excel estruturada
🔎 Diferentes estruturas de PDF

Um dos principais desafios encontrados durante o desenvolvimento foi perceber que diferentes sistemas podem gerar folhas de pagamento com estruturas completamente diferentes.

Por isso, o projeto não utiliza uma única regra genérica para qualquer PDF.

Durante o desenvolvimento foram trabalhadas três estruturas diferentes de folhas de pagamento, com leitores específicos para cada padrão identificado.

Essa abordagem permite que novas estruturas sejam analisadas e incorporadas ao processamento conforme novas necessidades apareçam.

🛠️ Tecnologias utilizadas
Processamento e dados
Python
Pandas
PyMuPDF
OpenPyXL
Aplicação
FastAPI
HTML
CSS
JavaScript
Desktop
PyInstaller
Windows .exe
Desenvolvimento
Google Colab
Jupyter Notebook
GitHub
🚀 Evolução do projeto

O projeto não começou como uma aplicação em Python.

Ele evoluiu conforme novas necessidades e desafios apareceram:

Processo manual
      ↓
Excel
      ↓
Power Query
      ↓
Python
      ↓
Processamento de PDFs
      ↓
FastAPI + Interface Web
      ↓
Aplicação Desktop
      ↓
Executável Windows (.exe)

Cada etapa trouxe novos aprendizados sobre automação, tratamento de dados, desenvolvimento de software e integração entre diferentes tecnologias.

📂 Estrutura do repositório

O projeto está dividido em três etapas que representam sua evolução:

RH-Valoracao/
│
├── 01-colab/
│   └── Notebook do processamento em Python
│
├── 02-fullstack/
│   ├── backend/
│   ├── core/
│   ├── frontend/
│   ├── requirements.txt
│   └── arquivos de execução e documentação
│
└── 03-desktop-app/
    ├── backend/
    ├── core/
    ├── frontend/
    ├── build_windows/
    ├── desktop_launcher/
    ├── requirements/
    ├── requirements-build/
    └── RH_Valoracao.spec
01-colab

Primeira etapa do processamento em Python.

A versão utiliza o Google Colab para executar o processamento sem exigir a configuração de um ambiente Python local.

02-fullstack

Evolução do processamento para uma aplicação com separação entre:

Frontend
Backend
Núcleo de processamento

O backend utiliza FastAPI para realizar a comunicação entre a interface e o processamento.

03-desktop-app

Versão voltada para utilização local em Windows.

A aplicação reúne interface, backend e núcleo de processamento e pode ser empacotada como um executável .exe, permitindo utilizar a solução sem configurar manualmente um ambiente Python.

🧪 Testes e validação

Durante o desenvolvimento foram realizados testes com diferentes conjuntos de folhas de pagamento e estruturas de documentos.

A validação envolveu, entre outros pontos:

Quantidade de colaboradores identificados
Correspondência entre colaboradores e registros
Identificação das rubricas
Valores processados
Estrutura final da planilha
Preservação das informações existentes
Processamento de múltiplos meses
Processamento de arquivos ZIP
Execução repetida do processo

Também foram realizados testes de regressão para verificar se alterações realizadas para suportar novas estruturas não afetavam os processamentos anteriormente validados.

📚 Aprendizados

Mais do que automatizar uma tarefa, este projeto foi uma oportunidade para entender na prática como uma solução de dados pode evoluir para uma aplicação.

Durante o desenvolvimento, trabalhei com conceitos como:

Tratamento e transformação de dados
Processamento e leitura de PDFs
Python
Pandas
APIs
Backend
Frontend
Automação de processos
Processamento de arquivos
Empacotamento de aplicações
Testes e validação
Organização de projetos
Git e GitHub

O projeto também mostrou que automatizar um processo real não significa apenas escrever código.

É necessário entender o problema, identificar padrões, definir regras, testar hipóteses, validar resultados e lidar com situações inesperadas.

🤖 Uso de IA

A inteligência artificial foi utilizada como parceira de desenvolvimento ao longo do projeto.

Ela ajudou na exploração de soluções, estruturação de código, identificação de problemas e aprendizado de novas tecnologias.

Ao mesmo tempo, a evolução do projeto envolveu entender o problema, testar as soluções, validar os resultados, identificar erros e decidir quais abordagens deveriam ser mantidas ou modificadas.

🔐 Dados e privacidade

Este repositório não contém folhas de pagamento reais, dados pessoais de colaboradores ou documentos de clientes.

Os arquivos disponibilizados têm finalidade técnica e educacional, permitindo demonstrar a estrutura e a evolução da solução sem expor informações sensíveis.

📌 Próximos passos

O projeto pode continuar evoluindo principalmente em duas frentes:

Suporte a novas estruturas de folhas de pagamento
Evolução da interface e da experiência de utilização

A arquitetura também permite que novos formatos de entrada sejam estudados e incorporados ao processamento conforme novas necessidades apareçam.

👨‍💻 Sobre o projeto

Este projeto faz parte da minha jornada de aprofundamento em tecnologia, conectando minha experiência com processos e análise de dados ao estudo de Python, automação e desenvolvimento de aplicações.

A ideia é continuar transformando problemas reais em projetos práticos, explorando cada vez mais a interseção entre:

Dados + Tecnologia + Automação + Resolução de Problemas

📫 Contato

Huriel Teixeira de Almeida

LinkedIn:
https://www.linkedin.com/in/huriel-teixeira-de-almeida-a2363823b

GitHub:
https://github.com/hurielalmeida
