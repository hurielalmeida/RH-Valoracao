# RH Valoração

Automação do processo de estruturação de dados de folhas de pagamento em PDF para uma base de RH em Excel.

Este projeto nasceu de uma necessidade real: transformar informações presentes em diferentes estruturas de folhas de pagamento em uma base organizada e padronizada, reduzindo etapas manuais e tornando o processo mais rápido, consistente e reutilizável.

O projeto evoluiu de uma solução inicial baseada em Excel e Power Query para uma aplicação desenvolvida em Python, com processamento de PDFs, interface web local e, posteriormente, uma versão desktop executável para Windows.

---

## 🎯 O problema

O processo original envolvia a leitura de folhas de pagamento em PDF e a transferência manual de informações para uma estrutura de Excel já existente.

Entre as informações tratadas estavam:

- colaboradores;
- cargos;
- códigos;
- rubricas;
- valores;
- competências;
- informações cadastrais necessárias para a estrutura da base.

Dependendo da quantidade de colaboradores, meses e rubricas, essa atividade poderia consumir várias horas ou até dias de trabalho.

Além do tempo envolvido, o processo manual também aumentava a quantidade de etapas sujeitas a erros de preenchimento, cópia e organização.

A ideia do projeto foi automatizar principalmente essa etapa de **leitura, transformação e estruturação dos dados**, mantendo a análise humana necessária para as etapas posteriores.

---

## 💡 A solução

O processador recebe uma base Excel e uma ou várias folhas de pagamento em PDF.

A aplicação identifica a estrutura dos documentos, extrai as informações relevantes, interpreta os dados encontrados e utiliza essas informações para atualizar uma nova versão da base Excel.

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
