# 📊 Resumo Fácil de Planilhas (CBSMExcell)

> **Sistema universal, acessível e inteligente de simplificação, visualização e análise de planilhas Excel (.xlsx, .xls) e arquivos CSV/TXT.**

[![Status](https://img.shields.io/badge/Deploy-Online-success?style=flat-square)](https://cbsm-excell.web.app)
[![License](https://img.shields.io/badge/License-MIT-blue?style=flat-square)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Web%20%7C%20Mobile%20%7C%20Desktop-informational?style=flat-square)](#)

---

## 🌐 Acesso Online Imediato
O sistema está hospedado de forma permanente e pode ser acessado diretamente de qualquer navegador:
* **Link Principal:** [https://cbsm-excell.web.app](https://cbsm-excell.web.app)
* **Link Alternativo:** [https://cbsm-excell.firebaseapp.com](https://cbsm-excell.firebaseapp.com)

---

## 📥 Como Baixar e Usar Localmente (Versão Offline Funcional)

Este projeto foi construído para funcionar de forma **100% autônoma e baixável**, sem necessidade de instalação de banco de dados ou configurações complexas de servidor.

### Opção 1: Uso Direto pelo Navegador (Sem Instalação)
1. Baixe o repositório como arquivo **ZIP** (clicando em `Code` > `Download ZIP` no GitHub) ou clone o repositório:
   ```bash
   git clone https://github.com/jvq3422/CBSMExcell.git
   ```
2. Descompacte a pasta no seu computador.
3. Dê um **duplo clique no arquivo `index.html`**.
4. O sistema abrirá imediatamente em qualquer navegador (Chrome, Edge, Firefox, Safari) pronto para uso!

### Opção 2: Executando via Servidor Local Node.js
Se você tiver o [Node.js](https://nodejs.org/) instalado e preferir rodar como um servidor local na porta 8080:
```bash
# 1. Clone o repositório
git clone https://github.com/jvq3422/CBSMExcell.git

# 2. Acesse a pasta do projeto
cd CBSMExcell

# 3. Inicie o servidor
npm start
# ou: node server.js
```
Abra seu navegador em [http://localhost:8080](http://localhost:8080).

---

## ✨ Principais Funcionalidades

### 1. 🔍 Compatibilidade Universal de Arquivos
* Suporta arquivos do Microsoft Excel moderno (`.xlsx`, `.xlsm`, `.xlsb`), Excel legado (`.xls`), LibreOffice (`.ods`) e arquivos de texto delimitado (`.csv`, `.tsv`, `.txt`).
* **Header Sniffer Automático:** Detecta automaticamente a linha correta de cabeçalho, mesmo se houver títulos, logotipos ou linhas vazias no topo da planilha.
* **Leitor de Múltiplas Abas:** Permite alternar facilmente entre todas as pastas/planilhas do arquivo com um clique.

### 2. 💵 Detecção Universal de Moedas e Valores (Zero Hardcode)
* **Inspeção Nativa de Células:** Lê diretamente a máscara contábil do Excel (`cell.z` / `cell.w`) preservando moedas sem distorções.
* **Inteligência Semântica:** Reconhece termos contábeis e contratuais em múltiplos idiomas (`autorizado`, `medido`, `saldo`, `preço`, `custo`, `total`, `salary`, `price`, `revenue`, etc.).
* **Preservação de Dados Gerais:** Índices, quantidades, notas acadêmicas, anos e pontuações continuam como números limpos (ex: `8,5`, `15`), sem inclusão indevida de símbolo monetário.
* **Botão Rápido de Formato:** Permite forçar ou alternar a qualquer momento entre **`[ 💵 R$ Moeda ]`** e **`[ 🔢 Número Geral ]`**.

### 3. 📈 5 Modelos de Gráficos com Escala Dinâmica
* **Gráfico de Pizza (Pie):** Exibe proporções, percentuais e soma dos registros em 100% da amostra.
* **Gráfico de Barras (Bar):** Comparação categórica com limite superior dinâmico (headroom de 18% a 20%) para que nenhuma barra encoste no topo.
* **Gráfico de Linha (Line):** Tendências contínuas com respiro visual acima do ponto mais alto.
* **Histograma (Histogram):** Frequência agrupada em faixas calculadas automaticamente.
* **Gráfico de Dispersão (Scatter):** Relações e distribuições entre linhas e valores.

### 4. 📄 Exportação Acessível de Resumos
* **Relatório em PDF Profissional:** Gera um documento A4 diagramado com métricas consolidadas, lista de destaques, texto narrativo e os gráficos capturados com **proporção geométrica estrita (círculos 100% redondos, sem distorção ou achatamento)**.
* **Quadro Interativo TXT:** Abre um modal acessível para ler, selecionar com um clique, copiar para a área de transferência ou baixar o arquivo `.txt`.
* **Leitor por Voz (Text-to-Speech):** Narra o resumo da planilha em áudio claro para pessoas com deficiência visual ou idosos.

### 5. ⚖️ Comparador Interativo de Itens
* Permite escolher quaisquer dois lançamentos da tabela para visualizar a diferença de valores, a proporção relativa em gráfico de pizza e a explicação em áudio.

### 6. ♿ Acessibilidade e Inclusão (Design Universal)
* Fontes grandes e legíveis com botões de tamanho ampliado.
* Modo **Alto Contraste** para idosos e pessoas com baixa visão.
* Layout 100% responsivo para celulares, tablets e computadores.

---

## 🛠️ Tecnologias Utilizadas
* **HTML5 / CSS3 / JavaScript (ES6+):** Código limpo, moderno e nativo, sem frameworks pesados.
* **[SheetJS (xlsx)](https://cdn.sheetjs.com/):** Leitura de planilhas nos formatos Excel e CSV.
* **[Chart.js](https://www.chartjs.org/):** Renderização de gráficos dinâmicos e responsivos.
* **[jsPDF](https://github.com/parallax/jsPDF):** Geração de relatórios em PDF no próprio navegador do usuário.
* **Web Speech API:** Síntese de voz nativa para acessibilidade.

---

## 📄 Licença
Distribuído sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para obter mais informações.
