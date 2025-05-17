# Gerador Inteligente de Documentos Escolares com Google Gemini API

## 🚀 Projeto da Imersão AI - Alura + Google (Maio 2025)

Este projeto visa automatizar e otimizar a criação de diversos documentos escolares (certificados, declarações, etc.) utilizando o poder da API Google Gemini integrada ao Google Planilhas através do Google Apps Script. A ferramenta foca na **geração inteligente do conteúdo textual**, que pode ser subsequentemente utilizado por ferramentas como o **Autocrat** ou scripts customizados para produzir os documentos finais formatados (ex: PDFs) e enviá-los por e-mail aos interessados. O objetivo é reduzir o trabalho manual, garantir consistência e permitir personalização avançada.

## ✨ Funcionalidades Principais

* **Geração de Diversos Tipos de Documentos:** Suporte para uma ampla gama de documentos escolares, como:
  * Declarações de Comparecimento
  * Declarações de Atuação como Professor Voluntário
  * Certificados de Ouvinte ou Ministrante de Palestras
  * Certificados de Participação ou Organização em Minicursos, Oficinas, Exposições, Saídas Técnicas, etc.
  * (Facilmente expansível para outros modelos)
* **Textos Personalizados com IA:** Utiliza a API Google Gemini (modelo `gemini-1.5-flash-latest` por padrão, mas configurável) para gerar textos contextualmente relevantes e com linguagem natural.
* **Integração Total com Google Planilhas:** Interface amigável diretamente na planilha para entrada de dados e visualização dos resultados.
* **Sistema de "Texto Base" Inteligente:**
  * **Criação de Modelos Reutilizáveis:** Opção para salvar um texto gerado pela IA como um "modelo". Ao fazer isso, o nome do participante atual é **automaticamente substituído pelo placeholder `{{NOME_PARTICIPANTE}}`**, preparando o texto para reuso.
  * **Reutilização Consistente:** Permite usar esses modelos (com placeholders) para gerar documentos para múltiplos participantes, garantindo consistência no corpo do texto e alterando apenas os dados variáveis.
* **Código de Verificação Único:** Gera automaticamente um UUID (Identificador Único Universal) para cada documento emitido, permitindo a criação de um sistema de validação de autenticidade.
* **Tratamento Flexível de Datas:** A coluna de data aceita tanto datas únicas (`dd/mm/aaaa`) quanto períodos textuais (ex: `15/05/2025 a 16/05/2025`), que são passados para a IA interpretar. Datas únicas são formatadas para o padrão brasileiro.
* **Menu Interativo:** Funções acessíveis através de um menu customizado na planilha ("Gerador de Documentos IA").
* **Confirmação de Ações:** Diálogos de confirmação (usando `ui.prompt`) para operações em lote, evitando ações acidentais.

## Demonstração em Vídeo

Confira uma rápida demonstração de como o Gerador Inteligente de Documentos funciona!

[![Demonstração Animada do Gerador de Documentos](./videos/demo.gif)](./videos/gerador_documentos_IA.mp4)


## 🛠️ Tecnologias Utilizadas

* **Google Planilhas:** Interface de dados e controle.
* **Google Apps Script:** Lógica de automação, orquestração da API e manipulação da planilha.
* **Google Gemini API:** Para a geração de texto inteligente.
* **Git & GitHub:** Para versionamento de código e compartilhamento do projeto.
* **(Opcional, para etapa seguinte) Autocrat:** Add-on do Google Sheets para mesclagem de dados e geração de documentos.

## 📊 Configuração da Planilha Google

Para o correto funcionamento do script, sua planilha (recomendamos nomear a aba principal como "Certificados") deve conter as seguintes colunas. A ordem é importante e deve corresponder às constantes definidas no script:

1. **`Nome do Participante`** (Ex: Coluna A)
2. **`Nome do Evento/Atividade`** (Ex: Coluna B)
3. **`Modelo do Documento`** (Ex: Coluna C) - *Dica: Use Validação de Dados.*
4. **`Data do Evento`** (Ex: Coluna D)
5. **`Carga Horária`** (Ex: Coluna E)
6. **`Detalhes Adicionais`** (Ex: Coluna F)
7. **`Local Específico`** (Ex: Coluna G)
8. **`Horário Início Comparecimento`** (Ex: Coluna H)
9. **`Horário Fim Comparecimento`** (Ex: Coluna I)
10. **`Nome da Disciplina/Projeto Voluntário`** (Ex: Coluna J)
11. **`Período Voluntariado`** (Ex: Coluna K)
12. **`Título da Palestra/Minicurso`** (Ex: Coluna L)
13. **`CPF do Participante`** (Ex: Coluna M)
14. **`Salvar Como Modelo Reutilizável?`** (Ex: Coluna N): **[Caixa de Seleção]**
15. **`Texto Base do Documento`** (Ex: Coluna O)
16. **`Usar Texto Base?`** (Ex: Coluna P): **[Caixa de Seleção]**
17. **`Texto do Certificado Gerado`** (Ex: Coluna Q): **[Saída]**
18. **`Status Geração`** (Ex: Coluna R): **[Saída]**
19. **`Código de Verificação`** (Ex: Coluna S): **[Saída]**

*A primeira linha da planilha deve conter os cabeçalhos. Os dados começam na linha 2.*

## ⚙️ Configuração do Script

1. **Copiar o Código:** Em `Extensões > Apps Script`, cole o conteúdo do `Code.gs` deste repositório.
2. **Ajustar Constantes:** No início do script, revise `COLUNA_...`, `NOME_PLANILHA_PRINCIPAL` e `NOME_INSTITUICAO`.
3. **Configurar Chave API Gemini:**
  * **NUNCA coloque sua chave diretamente no código público!**
  * No editor: `Arquivo > Propriedades do Projeto > Propriedades do Script`.
  * Adicionar: Nome `GEMINI_API_KEY`, Valor `SUA_CHAVE_API_AQUI`.
  * Salvar.
4. **Salvar e Autorizar:** Salve o script. Recarregue a planilha. O menu "Gerador de Documentos IA" aparecerá. Execute uma função para autorizar.

## 🚀 Como Usar

1. **Preencher Dados:** Insira as informações na planilha.
2. **Gerar Novo Documento com IA:** Deixe "Salvar Como Modelo Reutilizável?" e "Usar Texto Base?" desmarcadas. Use o menu para gerar.
3. **Criar Novo Modelo Reutilizável:**
  * Preencha dados do primeiro participante.
  * Marque **`Salvar Como Modelo Reutilizável?`**.
  * Deixe "Usar Texto Base?" desmarcada.
  * Execute. `Texto do Certificado Gerado` recebe o texto específico; `Texto Base do Documento` recebe o mesmo texto, mas com o nome do participante **automaticamente substituído por `{{NOME_PARTICIPANTE}}`**.
4. **Usar Modelo Reutilizável:**
  * Preencha dados do novo participante.
  * Copie o modelo com `{{NOME_PARTICIPANTE}}` da coluna `Texto Base do Documento` (da linha onde o modelo foi criado) para a coluna `Texto Base do Documento` da linha atual.
  * Marque **`Usar Texto Base?`**.
  * Deixe "Salvar Como Modelo Reutilizável?" desmarcada.
  * Execute.
5. **Código de Verificação:** Gerado automaticamente.
6. **Confirmação:** Digite "SIM" quando solicitado para ações em lote.

## 🔗 Integração com Ferramentas de Geração de Documentos (Ex: Autocrat)

O principal resultado deste script são os textos nas colunas **`Texto do Certificado Gerado`** e **`Código de Verificação`**. Após estas colunas serem preenchidas:

1. **Configuração do Autocrat (ou similar):**
  * Utilize uma ferramenta de mesclagem como o [Autocrat](https://workspace.google.com/marketplace/app/autocrat/539341275670) (um complemento para Google Sheets) ou um script próprio.
  * Crie um modelo de documento no Google Docs ou Google Slides (ex: o layout do seu certificado).
  * Neste modelo, insira as tags de mesclagem (`<<Nome da Coluna>>`) correspondentes aos cabeçalhos das suas colunas. Por exemplo:
    * `<<Nome do Participante>>`
    * `<<Modelo do Documento>>`
    * `<<Data do Evento>>`
    * `<<Texto do Certificado Gerado>>` (para o corpo principal do texto)
    * `<<Código de Verificação>>`
    * E quaisquer outras colunas que você queira incluir.
2. **Geração e Envio:**
  * Configure o Autocrat para usar sua planilha como fonte de dados e mapeie as tags para as colunas.
  * Defina o formato de saída (ex: PDF), nomes de arquivo e local de salvamento (Google Drive).
  * O Autocrat também pode ser configurado para enviar os documentos gerados por e-mail (se você tiver uma coluna com os e-mails dos destinatários).

Dessa forma, este projeto de IA cuida da redação do conteúdo, e ferramentas como o Autocrat cuidam da formatação final e distribuição.

## 🧩 Placeholders para "Texto Base"

Ao usar/editar um "Texto Base do Documento", os seguintes placeholders (definidos na função `substituirPlaceholders`) serão substituídos:

* `{{NOME_PARTICIPANTE}}`
* `{{NOME_EVENTO_ATIVIDADE}}`
* `{{MODELO_DOCUMENTO}}`
* `{{DATA_EVENTO}}`
* `{{CARGA_HORARIA}}`
* `{{DETALHES_ADICIONAIS}}`
* `{{LOCAL_ESPECIFICO}}`
* `{{HORA_INICIO_COMPARECIMENTO}}`
* `{{HORA_FIM_COMPARECIMENTO}}`
* `{{DISCIPLINA_PROJETO_VOLUNTARIO}}`
* `{{PERIODO_VOLUNTARIADO}}`
* `{{TITULO_PALESTRA_MINICURSO}}`
* `{{CPF_PARTICIPANTE}}`
* `{{NOME_INSTITUICAO}}`

## 🔑 Código de Verificação

* Um UUID é gerado para cada documento e salvo na coluna "Código de Verificação".
* Inclua este código no seu modelo de certificado.
* **Ideia para verificação:** A planilha atua como registro mestre. Futuramente, um Web App pode consultar este código para validar o certificado.

## ✍️ Personalização dos Prompts da IA

A qualidade dos textos depende dos prompts na função `construirPrompt`. **Revise, teste e refine os prompts** para cada `Modelo do Documento` que utilizar. Adicione novos `case` conforme necessário.

## 📝 Observações

* **Workaround `ui.confirm`:** O script usa `ui.prompt` para confirmações (digite "SIM").
* **Logs:** Verifique os logs no editor do Apps Script para erros (`Visualizar > Execuções`).
* **Cotas da API:** O uso da API Gemini tem cotas. O script inclui um `Utilities.sleep()` para processamento em lote.

## 💡 Ideias para Melhorias Futuras (Contribua!)

* Interface de usuário mais rica com HTML Service.
* Seleção de modelo da API Gemini via interface.
* Gerenciamento centralizado de "Textos Base Mestres".
* Geração de QR Code para o código de verificação.
* Web App público para validação de certificados.

* * *

Espero que este projeto seja útil e inspirador!
*Sebastião Tadeu de Oliveira Almeida*
