# WFS Shuttle - Gestão de Escalas e Fretado

Sistema web desenvolvido em Google Apps Script para automatizar a gestão de escalas operacionais, controle de folgas e requisições de transporte fretado.

## 🚀 Funcionalidades Principais
* **Controle de Acessos:** Perfis distintos para Administradores (Visão global e configurações) e Gestores/Focais (Visão limitada ao próprio setor).
* **Importação Inteligente (Modo Planilha):** Interface que simula uma planilha para colagem em lote de escalas e folgas, com validação automática de matrículas e sanitização de dados.
* **Gestão de Prazos:** Trava automática do sistema baseada em datas de abertura e fechamento definidas pela administração.
* **Integração com APIs:** Busca automática de endereços via CEP (ViaCEP / BrasilAPI).
* **Rastreamento de Atividade:** Monitoramento de quais setores já atualizaram suas respectivas escalas no ciclo atual.

## 🛠️ Tecnologias Utilizadas
* **Backend:** Google Apps Script (V8)
* **Banco de Dados:** Google Sheets
* **Frontend:** HTML5, CSS3, JavaScript vanilla
* **Bibliotecas Externas:** SweetAlert2 (Alertas modais), SheetJS (Processamento de uploads XLSX), Material Icons.

## 📂 Estrutura do Projeto
* `Code.gs`: Regras de negócio, endpoints de comunicação e integrações com o Google Sheets.
* `Index.html`: Interface principal, dashboards e formulários (SPA - Single Page Application).
* `ModalColagem.html`: Componente isolado da interface de importação em lote (Mini-Excel).
* `appsscript.json`: Manifesto de permissões do Google Workspace.

## ⚙️ Configuração e Instalação
1. Crie um novo projeto no [Google Apps Script](https://script.google.com/).
2. Copie os arquivos deste repositório para o seu projeto.
3. No arquivo `Code.gs`, atualize as constantes `ID_PLANILHA_DEMITIDOS` e `ID_BASE_EFETIVOS` com os IDs reais das suas planilhas de banco de dados.
4. Faça o Implante (Deploy) do projeto como um **Aplicativo da Web**.
