let currentWorkbook = null;
let currentFileName = "";
let currentActiveSheet = "";
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const resultsSection = document.getElementById("resultsSection");
const narrativeText = document.getElementById("narrativeText");
const metricsGrid = document.getElementById("metricsGrid");
const topItemsList = document.getElementById("topItemsList");
const fileNameDisplay = document.getElementById("fileNameDisplay");
const sheetSelectorArea = document.getElementById("sheetSelectorArea");

let currentFontScale = 1;
document.getElementById("btnFontIncrease").addEventListener("click", () => {
    currentFontScale = Math.min(currentFontScale + 0.15, 1.8);
    document.documentElement.style.setProperty("--font-scale", `${currentFontScale}rem`);
});
document.getElementById("btnFontDecrease").addEventListener("click", () => {
    currentFontScale = Math.max(currentFontScale - 0.15, 0.85);
    document.documentElement.style.setProperty("--font-scale", `${currentFontScale}rem`);
});
document.getElementById("btnFontNormal").addEventListener("click", () => {
    currentFontScale = 1;
    document.documentElement.style.setProperty("--font-scale", "1rem");
});
document.getElementById("btnHighContrast").addEventListener("click", () => {
    document.body.classList.toggle("high-contrast");
});

// 2. Eventos de Drag & Drop Universais e Seguros
window.addEventListener("dragover", (e) => {
    e.preventDefault();
}, false);

window.addEventListener("drop", (e) => {
    e.preventDefault();
}, false);

["dragenter", "dragover"].forEach(name => {
    dropZone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add("dragover");
    });
});

["dragleave"].forEach(name => {
    dropZone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove("dragover");
    });
});

dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove("dragover");
    const files = e.dataTransfer ? e.dataTransfer.files : null;
    if (files && files.length > 0) {
        processFile(files[0]);
    }
});

// Toque ou clique no quadro para abrir o seletor de arquivos (facilidade em celular e desktop)
dropZone.addEventListener("click", (e) => {
    if (e.target !== fileInput && e.target.tagName !== "LABEL") {
        fileInput.click();
    }
});

fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
        processFile(e.target.files[0]);
    }
});

// 3. Conversor Numérico Universal (Moeda Brasileira, Americana, Contábil e Pura)
function parseUniversalNumber(val) {
    if (typeof val === "number") return isNaN(val) ? null : val;
    if (!val) return null;
    if (typeof val !== "string") val = String(val);

    let clean = val.trim();
    if (!clean) return null;

    // Detecta formato contábil negativo: (1.250,00) ou (1,250.00)
    let isNegative = false;
    if (clean.startsWith("(") && clean.endsWith(")")) {
        isNegative = true;
        clean = clean.slice(1, -1).trim();
    } else if (clean.startsWith("-")) {
        isNegative = true;
        clean = clean.slice(1).trim();
    }

    // Remove símbolos de moeda e texto irrelevante
    clean = clean.replace(/^(R\$|\$|€|£|BRL|USD|EUR)\s*/i, "").trim();
    clean = clean.replace(/(\s*(R\$|\$|€|£|BRL|USD|EUR))$/i, "").trim();

    // Se tem tanto ponto quanto vírgula:
    // Ex: "1.250,50" -> ponto é milhar, vírgula é decimal (Brasil)
    // Ex: "1,250.50" -> vírgula é milhar, ponto é decimal (EUA)
    if (clean.includes(".") && clean.includes(",")) {
        if (clean.lastIndexOf(",") > clean.lastIndexOf(".")) {
            // Formato BR: 1.250,50 -> remove ponto, troca vírgula por ponto
            clean = clean.replace(/\./g, "").replace(",", ".");
        } else {
            // Formato EUA: 1,250.50 -> remove vírgula
            clean = clean.replace(/,/g, "");
        }
    } else if (clean.includes(",")) {
        // Apenas vírgula: se tiver mais de uma, é milhar; se tiver uma, é decimal
        const parts = clean.split(",");
        if (parts.length === 2) {
            clean = clean.replace(",", ".");
        } else {
            clean = clean.replace(/,/g, "");
        }
    } else if (clean.includes(".")) {
        // Apenas ponto: se tiver mais de um ponto (ex: 1.000.000), é milhar
        const parts = clean.split(".");
        if (parts.length > 2) {
            clean = clean.replace(/\./g, "");
        }
    }

    const num = parseFloat(clean);
    if (isNaN(num)) return null;
    return isNegative ? -Math.abs(num) : num;
}

// 3.1. Detecção Inteligente de Moeda vs Número Geral (Índices, Pontuações, Estatísticas)
let userForcedCurrencyMode = null; // null = automático, 'currency_brl' = forçar R$, 'number' = forçar número geral
let currentSheetColumnFormats = {}; // Mapeamento de formatos nativos de células extraídos diretamente do Excel

function detectColumnType(colName, rows) {
    if (userForcedCurrencyMode === "currency_brl") {
        return { isCurrency: true, symbol: "R$" };
    }
    if (userForcedCurrencyMode === "number") {
        return { isCurrency: false, symbol: "" };
    }

    // 1. Verifica se a coluna possui formatação monetária gravada no próprio arquivo Excel (cell.z / cell.w)
    if (currentSheetColumnFormats && currentSheetColumnFormats[colName]) {
        return {
            isCurrency: true,
            symbol: currentSheetColumnFormats[colName]
        };
    }

    const nameLower = String(colName || "").toLowerCase();
    let currencySymbol = null;
    let foundExplicit = false;

    // 2. Avalia amostra dos valores reais na coluna para detectar símbolos monetários universais nas células
    const sample = (rows || []).slice(0, 100).map(r => String(r[colName] !== undefined ? r[colName] : ""));
    for (const val of sample) {
        if (/R\$|BRL/i.test(val)) { currencySymbol = "R$"; foundExplicit = true; break; }
        if (/\$|USD/i.test(val)) { currencySymbol = "$"; foundExplicit = true; break; }
        if (/€|EUR/i.test(val)) { currencySymbol = "€"; foundExplicit = true; break; }
        if (/£|GBP/i.test(val)) { currencySymbol = "£"; foundExplicit = true; break; }
        if (/¥|JPY|CNY/i.test(val)) { currencySymbol = "¥"; foundExplicit = true; break; }
        if (/CHF/i.test(val)) { currencySymbol = "CHF"; foundExplicit = true; break; }
    }

    // 3. Avalia raízes semânticas financeiras e contábeis universais no nome da coluna
    if (!foundExplicit) {
        const universalFinancialPattern = /(pre[çc]o|custo|cost|price|gasto|despesa|expense|sal[aá]rio|salary|fatur|receit|revenu|saldo|balance|or[çc]am|budget|empenh|liquid|pago|paid|payment|\bpay\b|_pay\b|invest|montant|fee|taxa|tax\b|renda|income|lucro|profit|reais|real|brl|usd|eur|gbp|autoriz|medid|medi[çc]|med\b|aprov|liber|cobr|venda|sales?)/i;
        if (universalFinancialPattern.test(nameLower)) {
            currencySymbol = "R$";
            foundExplicit = true;
        }
    }

    return {
        isCurrency: foundExplicit,
        symbol: currencySymbol || ""
    };
}

// 3.2. Formatador Inteligente Universal (Só coloca símbolo se REALMENTE for moeda)
function formatSmartValue(val, colType) {
    if (val === null || val === undefined || isNaN(val)) return "-";
    const num = Number(val);

    if (colType && colType.isCurrency) {
        if (colType.symbol === "$") {
            return "$" + num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        if (colType.symbol === "€") {
            return num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
        }
        if (colType.symbol === "£") {
            return "£" + num.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        if (colType.symbol === "¥") {
            return "¥" + num.toLocaleString("ja-JP", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
        }
        if (colType.symbol === "CHF") {
            return "CHF " + num.toLocaleString("de-CH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        return "R$ " + num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Se NÃO é moeda (Índices, Pontuações, Notas, Totais, Estatísticas gerais):
    if (Number.isInteger(num)) {
        return num.toLocaleString("pt-BR");
    }
    return num.toLocaleString("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 2 });
}

// 4. Parser Inteligente de Cabeçalho (Header Sniffer)
function extractRowsFromWorksheet(worksheet) {
    currentSheetColumnFormats = {};
    // Lê toda a planilha como matriz bruta de linhas
    const rawMatrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    if (!rawMatrix || rawMatrix.length === 0) return { rows: [], keys: [] };

    // Procura a linha com cara de cabeçalho nas primeiras 25 linhas
    let headerRowIndex = 0;
    let maxValidCols = 0;

    for (let r = 0; r < Math.min(rawMatrix.length, 25); r++) {
        const row = rawMatrix[r];
        if (!Array.isArray(row)) continue;
        
        // Conta quantas colunas nesta linha são texto não vazio e legível
        const textCols = row.filter(c => c !== null && c !== undefined && String(c).trim().length > 0);
        if (textCols.length > maxValidCols) {
            // Verifica se a linha seguinte tem dados
            const nextRow = rawMatrix[r + 1];
            if (nextRow && nextRow.some(c => c !== null && c !== undefined && String(c).trim() !== "")) {
                maxValidCols = textCols.length;
                headerRowIndex = r;
            }
        }
    }

    // Se encontrou cabeçalho
    const rawHeaders = rawMatrix[headerRowIndex] || [];
    const headers = rawHeaders.map((h, colIdx) => {
        const str = String(h || "").trim();
        return str.length > 0 ? str : `Coluna_${colIdx + 1}`;
    });

    // Inspeciona formatos nativos de células do Excel via SheetJS (cell.z e cell.w)
    if (worksheet && worksheet["!ref"]) {
        try {
            const range = XLSX.utils.decode_range(worksheet["!ref"]);
            for (let c = range.s.c; c <= range.e.c; c++) {
                const colHeader = headers[c - range.s.c] || `Coluna_${c + 1}`;
                for (let r = headerRowIndex + 1; r <= Math.min(range.e.r, headerRowIndex + 40); r++) {
                    const addr = XLSX.utils.encode_cell({ r, c });
                    const cell = worksheet[addr];
                    if (cell) {
                        const fmtStr = String(cell.z || "");
                        const textStr = String(cell.w || "");
                        if (/R\$|BRL|\[\$R/i.test(fmtStr) || /R\$/i.test(textStr)) {
                            currentSheetColumnFormats[colHeader] = "R$";
                            break;
                        }
                        if (/\$|USD/i.test(fmtStr) || /\$/i.test(textStr)) {
                            currentSheetColumnFormats[colHeader] = "$";
                            break;
                        }
                        if (/€|EUR/i.test(fmtStr) || /€/i.test(textStr)) {
                            currentSheetColumnFormats[colHeader] = "€";
                            break;
                        }
                        if (/£|GBP/i.test(fmtStr) || /£/i.test(textStr)) {
                            currentSheetColumnFormats[colHeader] = "£";
                            break;
                        }
                    }
                }
            }
        } catch (fmtErr) {
            console.warn("Nao foi possivel inspecionar formatos nativos do Excel:", fmtErr);
        }
    }

    const parsedRows = [];
    for (let r = headerRowIndex + 1; r < rawMatrix.length; r++) {
        const line = rawMatrix[r];
        if (!Array.isArray(line)) continue;

        // Ignora linhas totalmente vazias
        const hasContent = line.some(c => c !== null && c !== undefined && String(c).trim() !== "");
        if (!hasContent) continue;

        // Ignora linhas de rodapé ou totalizadores consolidados da planilha de forma universal
        const nonBlankCells = line.filter(c => c !== null && c !== undefined && String(c).trim() !== "");
        if (nonBlankCells.length > 0) {
            const firstCellText = String(nonBlankCells[0]).trim();
            if (/^(total(\s+geral|\s+gerais|\s+final|\s+acumulado)?|subtotal|grand\s*total|sum|soma|somme|gesamt|resumo\s*total)$/i.test(firstCellText)) {
                continue;
            }
        }

        const obj = {};
        headers.forEach((h, colIdx) => {
            obj[h] = line[colIdx] !== undefined ? line[colIdx] : "";
        });
        parsedRows.push(obj);
    }

    return { rows: parsedRows, keys: headers };
}

// 5. Parser Manual de CSV / Texto Tabular (Fallback Universal)
function parseDelimitedTextToWorkbook(text) {
    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length === 0) throw new Error("Arquivo vazio");

    // Detecta delimitador mais provável
    const firstLine = lines[0];
    const counts = {
        ";": (firstLine.match(/;/g) || []).length,
        ",": (firstLine.match(/,/g) || []).length,
        "\t": (firstLine.match(/\t/g) || []).length,
        "|": (firstLine.match(/\|/g) || []).length
    };
    let bestDelimiter = ";";
    let maxCount = -1;
    for (const [delim, count] of Object.entries(counts)) {
        if (count > maxCount) {
            maxCount = count;
            bestDelimiter = delim;
        }
    }

    const matrix = lines.map(line => {
        // Divisão básica por delimitador respeitando aspas
        const regex = new RegExp(`(?:${bestDelimiter}|\n|^)("(?:(?:"")*[^"]*)*"|[^"${bestDelimiter}\n]*)`, "gi");
        const matches = [];
        let match;
        while ((match = regex.exec(line)) !== null) {
            let val = match[1] || "";
            if (val.startsWith('"') && val.endsWith('"')) {
                val = val.slice(1, -1).replace(/""/g, '"');
            }
            matches.push(val.trim());
        }
        return matches.length > 0 ? matches : line.split(bestDelimiter);
    });

    const ws = XLSX.utils.aoa_to_sheet(matrix);
    return {
        SheetNames: ["Dados"],
        Sheets: { "Dados": ws }
    };
}

// 6. Processamento Universal com Multi-Pass Fallback
function processFile(file) {
    currentFileName = file.name;
    fileNameDisplay.textContent = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
        const buffer = e.target.result;
        currentWorkbook = null;

        // Estratégia 1: Leitura binária padrão do SheetJS
        try {
            const data = new Uint8Array(buffer);
            currentWorkbook = XLSX.read(data, { 
                type: "array", 
                cellDates: true, 
                cellNF: true, 
                cellText: true 
            });
        } catch (err1) {
            console.warn("Estratégia 1 (XLSX Array) falhou, tentando UTF-8...", err1);
        }

        // Estratégia 2: Se falhar ou for CSV/texto, decodifica como UTF-8
        if (!currentWorkbook) {
            try {
                const textUtf8 = new TextDecoder("utf-8").decode(buffer);
                currentWorkbook = XLSX.read(textUtf8, { type: "string" });
            } catch (err2) {
                console.warn("Estratégia 2 (UTF-8) falhou, tentando Windows-1252...", err2);
            }
        }

        // Estratégia 3: Decodifica com Windows-1252 / ANSI (padrão de sistemas e bancos no Brasil)
        if (!currentWorkbook) {
            try {
                const textAnsi = new TextDecoder("windows-1252").decode(buffer);
                currentWorkbook = XLSX.read(textAnsi, { type: "string" });
            } catch (err3) {
                console.warn("Estratégia 3 (Windows-1252) falhou, tentando parser manual...", err3);
            }
        }

        // Estratégia 4: Parser manual nativo com detecção de delimitador (; , \t)
        if (!currentWorkbook) {
            try {
                const textFallback = new TextDecoder("windows-1252").decode(buffer);
                currentWorkbook = parseDelimitedTextToWorkbook(textFallback);
            } catch (err4) {
                console.error("Todas as estratégias de leitura falharam.", err4);
            }
        }

        // Se conseguiu ler o workbook
        if (currentWorkbook && currentWorkbook.SheetNames && currentWorkbook.SheetNames.length > 0) {
            renderSheetButtons(currentWorkbook.SheetNames);

            // Procura a aba que realmente tem mais dados (evita abrir aba de capa vazia)
            let bestSheet = currentWorkbook.SheetNames[0];
            let maxRowsCount = 0;

            for (const name of currentWorkbook.SheetNames) {
                const ws = currentWorkbook.Sheets[name];
                if (ws && ws["!ref"]) {
                    const { rows } = extractRowsFromWorksheet(ws);
                    if (rows.length > maxRowsCount) {
                        maxRowsCount = rows.length;
                        bestSheet = name;
                    }
                }
            }

            dropZone.style.display = "none";
            resultsSection.style.display = "block";

            try {
                loadSheet(bestSheet);
            } catch (err) {
                console.error("Erro ao analisar planilha:", err);
                narrativeText.innerHTML = `<p>Não foi possível resumir esta aba automaticamente: ${err.message}</p>`;
            }
        } else {
            alert(`Não foi possível ler o arquivo "${file.name}".\n\nCertifique-se de que não é um arquivo corrompido ou protegido por senha no Excel.`);
        }
    };

    reader.onerror = () => {
        alert("Ocorreu um erro ao carregar o arquivo no navegador. Tente novamente.");
    };

    reader.readAsArrayBuffer(file);
}

function renderSheetButtons(sheetNames) {
    sheetSelectorArea.innerHTML = "";
    if (sheetNames.length <= 1) return;

    const label = document.createElement("span");
    label.textContent = " Aba: ";
    sheetSelectorArea.appendChild(label);

    sheetNames.forEach(sheetName => {
        const btn = document.createElement("button");
        btn.textContent = sheetName;
        btn.className = "btn-acc";
        btn.style.marginLeft = "4px";
        btn.onclick = () => loadSheet(sheetName);
        sheetSelectorArea.appendChild(btn);
    });
}

function loadSheet(sheetName) {
    const worksheet = currentWorkbook.Sheets[sheetName];
    if (!worksheet) {
        narrativeText.innerHTML = "<p>Esta aba não foi encontrada na planilha.</p>";
        return;
    }

    const { rows, keys } = extractRowsFromWorksheet(worksheet);
    if (!rows || rows.length === 0) {
        narrativeText.innerHTML = `<p>A aba <strong>"${sheetName}"</strong> parece estar vazia ou contém apenas títulos.</p>`;
        return;
    }
    analyzeData(rows, keys, sheetName);
}

// 7. Motor Universal de Análise dos Dados
// 6.5. Perfilamento Estatístico Universal de Colunas (Zero Hardcode)
function profileSpreadsheetColumns(rows, keys) {
    const sampleRows = (rows || []).slice(0, 200);
    const totalRows = sampleRows.length || 1;
    const profiles = {};

    keys.forEach(k => {
        let numericCount = 0;
        let stringCount = 0;
        let emptyCount = 0;
        let totalLength = 0;
        let hasDecimalsCount = 0;
        let valuesSum = 0;
        const distinctSet = new Set();
        const numericValues = [];

        sampleRows.forEach(row => {
            const raw = row[k];
            if (raw === null || raw === undefined || String(raw).trim() === "") {
                emptyCount++;
                return;
            }
            const str = String(raw).trim();
            distinctSet.add(str);

            const num = parseUniversalNumber(str);
            if (num !== null && !isNaN(num)) {
                numericCount++;
                numericValues.push(num);
                valuesSum += Math.abs(num);
                if (Math.floor(num) !== num) {
                    hasDecimalsCount++;
                }
            } else {
                stringCount++;
                totalLength += str.length;
            }
        });

        const numericRatio = numericCount / totalRows;
        const stringRatio = stringCount / totalRows;
        const distinctRatio = distinctSet.size / (numericCount + stringCount || 1);
        const avgLength = stringCount > 0 ? (totalLength / stringCount) : 0;

        // Detecta se é uma sequência estrita de IDs numéricos (ex: 1, 2, 3... ou 100, 101, 102...)
        let isSequentialId = false;
        if (numericValues.length >= 3 && hasDecimalsCount === 0) {
            let isSeq = true;
            for (let i = 1; i < numericValues.length; i++) {
                if (numericValues[i] - numericValues[i - 1] !== 1) {
                    isSeq = false;
                    break;
                }
            }
            isSequentialId = isSeq;
        }

        const isYearColumn = numericValues.length > 0 &&
                             hasDecimalsCount === 0 &&
                             numericValues.every(v => v >= 1900 && v <= 2100);

        profiles[k] = {
            key: k,
            numericRatio,
            stringRatio,
            distinctCount: distinctSet.size,
            distinctRatio,
            avgLength,
            hasDecimalsCount,
            isSequentialId,
            isYearColumn,
            valuesSum,
            numericCount
        };
    });

    return profiles;
}

function pickUniversalDefaultColumns(keys, profiles) {
    let bestValueCol = null;
    let maxValueScore = -Infinity;

    let bestDescCol = null;
    let maxDescScore = -Infinity;

    const allSums = keys.map(k => profiles[k]?.valuesSum || 0);
    const maxTotalSum = Math.max(...allSums) || 1;

    keys.forEach(k => {
        const p = profiles[k];
        if (!p) return;

        // --- PONTUAÇÃO ESTATÍSTICA PARA MÉTRICA DE VALOR ---
        let valueScore = 0;
        valueScore += p.numericRatio * 50;
        if (p.hasDecimalsCount > 0) valueScore += 25;
        if (p.isSequentialId) valueScore -= 100;
        if (p.isYearColumn) valueScore -= 50;
        if (p.valuesSum > 0) {
            valueScore += (p.valuesSum / maxTotalSum) * 35;
        }

        const lowerK = k.toLowerCase();
        // Raízes semânticas gerais de medidas em múltiplos idiomas
        if (/(val|amount|tot|sum|soma|price|prec|cost|cust|sald|balanc|pay|pag|fatur|rec|expens|desp|budg|orc|quant|qty|qtd|med|rate|tax|score|not|peso|weight)/i.test(lowerK)) {
            valueScore += 30;
        }

        if (valueScore > maxValueScore) {
            maxValueScore = valueScore;
            bestValueCol = k;
        }

        // --- PONTUAÇÃO ESTATÍSTICA PARA DESCRIÇÃO / CATEGORIA ---
        let descScore = 0;
        descScore += p.stringRatio * 50;
        if (p.distinctCount > 1) descScore += 20;
        if (p.avgLength >= 3 && p.avgLength <= 80) descScore += 20;
        if (p.numericRatio > 0.7) descScore -= 40;

        // Raízes semânticas gerais de rótulos e entidades
        if (/(desc|name|nom|titl|item|prod|serv|cat|typ|tip|label|rotul|client|forn|suppl|vend|loc|local|bairr|distr|city|cid|street|rua|lograd|end|addr|proj|dept|setor)/i.test(lowerK)) {
            descScore += 30;
        }

        if (p.isSequentialId && /item/i.test(lowerK)) {
            descScore -= 30;
        }

        if (descScore > maxDescScore) {
            maxDescScore = descScore;
            bestDescCol = k;
        }
    });

    if (bestDescCol === bestValueCol) {
        const otherDesc = keys.find(k => k !== bestValueCol && profiles[k].stringRatio > 0.3) ||
                          keys.find(k => k !== bestValueCol) ||
                          keys[0];
        bestDescCol = otherDesc;
    }

    return {
        valueColumn: bestValueCol || keys[keys.length - 1],
        descriptionColumn: bestDescCol || keys[0]
    };
}

// 7. Motor Universal de Análise dos Dados
let lastAnalyzedState = null;

function analyzeData(rows, keys, sheetName, forcedDescCol, forcedValCol) {
    currentActiveSheet = sheetName || "Principal";
    if (!keys || keys.length === 0) keys = Object.keys(rows[0] || {});

    // Perfilamento estatístico universal sem regras hardcoded
    const profiles = profileSpreadsheetColumns(rows, keys);
    const defaults = pickUniversalDefaultColumns(keys, profiles);

    let valueColumn = forcedValCol || defaults.valueColumn;
    let descriptionColumn = forcedDescCol || defaults.descriptionColumn;

    // Armazena o estado atual da análise para permitir re-renderização instantânea (ex: troca de moeda)
    lastAnalyzedState = { rows, keys, sheetName, descriptionColumn, valueColumn };

    // Calcula somatórios e destaques usando o parser numérico universal
    let total = 0;
    let count = 0;
    let items = [];

    rows.forEach(row => {
        const val = parseUniversalNumber(row[valueColumn]);
        if (val !== null && !isNaN(val)) {
            total += val;
            count++;
            items.push({ 
                desc: String(row[descriptionColumn] || "Item sem descrição").trim(), 
                val: val 
            });
        }
    });

    items.sort((a, b) => b.val - a.val);
    const topItem = items[0] || { desc: "Nenhum", val: 0 };
    const bottomItem = items[items.length - 1] || { desc: "Nenhum", val: 0 };
    const average = count > 0 ? (total / count) : 0;

    // Detecta se a coluna analisada é de moeda ou dados gerais (índices, pontuações, etc.)
    const colType = detectColumnType(valueColumn, rows);
    const fmt = (v) => formatSmartValue(v, colType);

    // Atualiza o botão de alternância rápida de formato de moeda
    const btnToggleCurr = document.getElementById("btnToggleCurrency");
    if (btnToggleCurr) {
        if (userForcedCurrencyMode === "currency_brl") {
            btnToggleCurr.innerHTML = "💰 R$ (Fixado)";
            btnToggleCurr.title = "Formato fixado em Reais (R$). Clique para mudar para Número Geral.";
        } else if (userForcedCurrencyMode === "number") {
            btnToggleCurr.innerHTML = "🔢 Nº Geral (Fixado)";
            btnToggleCurr.title = "Formato fixado em Número Geral. Clique para voltar para Detecção Automática.";
        } else {
            btnToggleCurr.innerHTML = colType.isCurrency ? "💵 R$ Ativado" : "🔢 Nº Geral";
            btnToggleCurr.title = `Detecção Automática (${colType.isCurrency ? "Moeda R$" : "Número Geral"}). Clique para alternar.`;
        }
    }

    let narrative = `<p>Olá! Analisamos a planilha com foco na coluna <strong>"${valueColumn}"</strong>. Encontramos um total de <strong>${count} registros</strong>.</p>`;
    
    if (colType.isCurrency) {
        narrative += `<p>O valor financeiro total consolidado é de <strong>${fmt(total)}</strong>.</p>`;
        if (topItem.val > 0) {
            narrative += `<p>O maior lançamento registrado foi <strong>"${topItem.desc}"</strong>, no valor de <strong>${fmt(topItem.val)}</strong>.</p>`;
        }
        narrative += `<p>O valor médio por lançamento ficou em <strong>${fmt(average)}</strong>.</p>`;
    } else {
        narrative += `<p>A pontuação / soma acumulada dos valores é de <strong>${fmt(total)}</strong>, com média geral de <strong>${fmt(average)}</strong>.</p>`;
        if (topItem.desc !== "Nenhum") {
            narrative += `<p>O maior registro identificado foi <strong>"${topItem.desc}"</strong>, com índice de <strong>${fmt(topItem.val)}</strong>.</p>`;
        }
        if (bottomItem.desc !== "Nenhum" && count > 1) {
            narrative += `<p>O menor registro foi <strong>"${bottomItem.desc}"</strong>, com índice de <strong>${fmt(bottomItem.val)}</strong>.</p>`;
        }
    }

    narrativeText.innerHTML = narrative;
    narrativeText.dataset.rawText = narrative.replace(/<[^>]*>?/gm, " ");

    const labelTotal = colType.isCurrency ? "Valor Total Geral" : "Soma dos Registros";
    const labelTop = colType.isCurrency ? "Maior Lançamento" : "Maior Registro";
    const labelAvg = colType.isCurrency ? "Valor Médio" : "Média dos Valores";
    const labelCount = colType.isCurrency ? "Total de Lançamentos" : "Total de Linhas";

    metricsGrid.innerHTML = `
        <div class="metric-card"><h3>${labelTotal}</h3><div class="metric-value">${fmt(total)}</div></div>
        <div class="metric-card"><h3>${labelTop}</h3><div class="metric-value highlight-red">${fmt(topItem.val)}</div></div>
        <div class="metric-card"><h3>${labelAvg}</h3><div class="metric-value">${fmt(average)}</div></div>
        <div class="metric-card"><h3>${labelCount}</h3><div class="metric-value">${count}</div></div>
    `;

    // C) Maiores Destaques (Top 5)
    const top5 = items.slice(0, 5);
    topItemsList.innerHTML = top5.map(item => `
        <div class="top-item-row">
            <span class="top-item-name">📌 ${item.desc}</span>
            <span class="top-item-value">${fmt(item.val)}</span>
        </div>
    `).join("");

    // D) Gráficos Automáticos (Barras e Categorias)
    renderAutomaticCharts(items, rows, keys, valueColumn);

    // E) Estúdio de Gráficos Selecionáveis (5 Templates)
    setupTemplateStudio(rows, keys, descriptionColumn, valueColumn);

    // F) Preenche e Configura o Comparador Interativo de Itens
    setupComparator(items, colType);

    // G) Tabela com Paginação
    setupTablePagination(rows, keys, descriptionColumn, valueColumn);
}

// 7. Controle da Tabela e Paginação Acessível
let currentRowsData = [];
let currentKeysData = [];
let tablePage = 1;
const PAGE_SIZE = 15;
let showingAllRows = false;

function setupTablePagination(rows, keys) {
    currentRowsData = rows;
    currentKeysData = keys;
    tablePage = 1;
    showingAllRows = false;
    renderTablePage();
}

function renderTablePage() {
    const head = document.getElementById("tableHead");
    const body = document.getElementById("tableBody");
    const btnPrev = document.getElementById("btnPrevPage");
    const btnNext = document.getElementById("btnNextPage");
    const pageIndicator = document.getElementById("pageIndicator");
    const btnToggleAll = document.getElementById("btnToggleAllRows");
    const countNote = document.getElementById("tableCountNote");

    // Mostra as primeiras colunas
    const visibleKeys = currentKeysData.slice(0, 6);
    head.innerHTML = `<tr>${visibleKeys.map(k => `<th>${k}</th>`).join("")}</tr>`;

    let displayRows = [];
    const total = currentRowsData.length;
    const totalPages = Math.ceil(total / PAGE_SIZE);

    if (showingAllRows) {
        displayRows = currentRowsData;
        btnPrev.style.display = "none";
        btnNext.style.display = "none";
        pageIndicator.style.display = "none";
        btnToggleAll.textContent = "📑 Voltar para 15 por página";
        countNote.textContent = `Exibindo todos os ${total} lançamentos.`;
    } else {
        const start = (tablePage - 1) * PAGE_SIZE;
        const end = Math.min(start + PAGE_SIZE, total);
        displayRows = currentRowsData.slice(start, end);

        btnPrev.style.display = totalPages > 1 ? "inline-block" : "none";
        btnNext.style.display = totalPages > 1 ? "inline-block" : "none";
        pageIndicator.style.display = totalPages > 1 ? "inline-block" : "none";
        
        btnPrev.disabled = (tablePage === 1);
        btnNext.disabled = (tablePage === totalPages);
        pageIndicator.textContent = ` Página ${tablePage} de ${totalPages} `;
        
        btnToggleAll.textContent = `📄 Ver Todos os ${total} Lançamentos`;
        countNote.textContent = `Exibindo do ${start + 1}º ao ${end}º de ${total} lançamentos.`;
    }

    body.innerHTML = displayRows.map(row => {
        const itemVal = row[currentValueCol];
        const itemDesc = row[currentDescCol];
        return `
        <tr>
            ${visibleKeys.map(k => `<td>${row[k] !== undefined ? row[k] : ""}</td>`).join("")}
            <td style="text-align: center;">
                <button class="btn-acc" style="font-size: 0.95rem; padding: 0.3rem 0.6rem;" 
                    onclick="selectItemForComparison('${encodeURIComponent(itemDesc || "")}')" 
                    title="Comparar este item">
                    ⚖️ Comparar
                </button>
            </td>
        </tr>`;
    }).join("");
}

let currentDescCol = "";
let currentValueCol = "";

function setupTablePagination(rows, keys, descCol, valCol) {
    currentRowsData = rows;
    currentKeysData = keys;
    currentDescCol = descCol;
    currentValueCol = valCol;
    tablePage = 1;
    showingAllRows = false;
    
    // Atualiza cabeçalho com coluna extra de ação
    const visibleKeys = currentKeysData.slice(0, 6);
    const head = document.getElementById("tableHead");
    head.innerHTML = `<tr>${visibleKeys.map(k => `<th>${k}</th>`).join("")}<th style="text-align:center;">Ação</th></tr>`;
    
    renderTablePage();
}

// 8. Gráficos Automáticos em Pizza (Chart.js)
let pieTopItemsInstance = null;
let categoryChartInstance = null;

function renderAutomaticCharts(items, rows, keys, valueColumn) {
    const canvasTop = document.getElementById("pieChartTopItems");
    if (!canvasTop) return;

    const fmt = (v) => v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
    const totalSoma = items.reduce((acc, i) => acc + i.val, 0);

    // A) Gráfico de Pizza dos Maiores Lançamentos
    const top5 = items.slice(0, 5);
    const top5Soma = top5.reduce((acc, i) => acc + i.val, 0);
    const outrosSoma = Math.max(0, totalSoma - top5Soma);

    const pieLabels = top5.map(i => i.desc.length > 22 ? i.desc.slice(0, 20) + "..." : i.desc);
    const pieData = top5.map(i => i.val);

    if (outrosSoma > 0) {
        pieLabels.push("Demais Lançamentos Combinados");
        pieData.push(outrosSoma);
    }

    const ctxPieTop = canvasTop.getContext("2d");
    if (pieTopItemsInstance) pieTopItemsInstance.destroy();

    pieTopItemsInstance = new Chart(ctxPieTop, {
        type: "pie",
        data: {
            labels: pieLabels,
            datasets: [{
                data: pieData,
                backgroundColor: [
                    "#1e40af", "#0284c7", "#10b981", "#f59e0b", "#8b5cf6", "#94a3b8"
                ],
                borderWidth: 3,
                borderColor: "#ffffff"
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 100,
            plugins: {
                legend: {
                    position: "bottom",
                    labels: {
                        font: { size: 14, weight: "bold" },
                        padding: 14
                    }
                },
                tooltip: {
                    titleFont: { size: 16 },
                    bodyFont: { size: 16 },
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.parsed;
                            const pct = totalSoma > 0 ? ((val / totalSoma) * 100).toFixed(1) : 0;
                            return ` ${ctx.label}: ${fmt(val)} (${pct}%)`;
                        }
                    }
                }
            }
        }
    });

    // B) Gráfico de Pizza por Categoria
    const categoryKey = keys.find(k => /categoria|tipo|setor|depto|departamento|grupo/i.test(k));
    const canvasCat = document.getElementById("pieChartCategories");
    if (!canvasCat) return;

    const ctxPieCat = canvasCat.getContext("2d");
    if (categoryChartInstance) categoryChartInstance.destroy();

    let catLabels = [];
    let catValues = [];

    if (categoryKey) {
        document.getElementById("categoryChartWrapper").style.display = "flex";
        const catMap = {};
        rows.forEach(r => {
            const cat = r[categoryKey] || "Geral";
            let val = r[valueColumn];
            if (typeof val === "string") val = parseFloat(val.replace("R$", "").replace(/\./g, "").replace(",", ".").trim());
            if (!isNaN(val) && val > 0) {
                catMap[cat] = (catMap[cat] || 0) + val;
            }
        });

        const sortedCats = Object.entries(catMap).sort((a, b) => b[1] - a[1]);
        catLabels = sortedCats.slice(0, 6).map(c => c[0]);
        catValues = sortedCats.slice(0, 6).map(c => c[1]);
    } else {
        document.getElementById("categoryChartWrapper").style.display = "flex";
        const top3Val = items.slice(0, 3).reduce((acc, i) => acc + i.val, 0);
        const othersVal = items.slice(3).reduce((acc, i) => acc + i.val, 0);
        catLabels = ["Top 3 Maiores Lançamentos", "Demais Lançamentos"];
        catValues = [top3Val, othersVal];
    }

    categoryChartInstance = new Chart(ctxPieCat, {
        type: "pie",
        data: {
            labels: catLabels,
            datasets: [{
                data: catValues,
                backgroundColor: [
                    "#2563eb", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#06b6d4"
                ],
                borderWidth: 3,
                borderColor: "#ffffff"
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 100,
            plugins: {
                legend: {
                    position: "bottom",
                    labels: {
                        font: { size: 14, weight: "bold" },
                        padding: 14
                    }
                },
                tooltip: {
                    titleFont: { size: 16 },
                    bodyFont: { size: 16 },
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.parsed;
                            const totalCat = catValues.reduce((a, b) => a + b, 0);
                            const pct = totalCat > 0 ? ((val / totalCat) * 100).toFixed(1) : 0;
                            return ` ${ctx.label}: ${fmt(val)} (${pct}%)`;
                        }
                    }
                }
            }
        }
    });
}

// 8.1. Estúdio de Gráficos Selecionáveis (5 Modelos) com Escalas Dinâmicas
let customTemplateChartInstance = null;
let currentTemplateSelected = "pie";
let studioRowsData = [];
let studioKeysData = [];

// Função para arredondar o teto para números "agradáveis" (nice numbers) com respiro visual
function calculateNiceCeiling(val) {
    if (val <= 0) return 0;
    if (val <= 1) return Math.ceil(val * 10) / 10 + 0.1;
    if (val <= 2) return 2.5;
    if (val <= 3) return 3.5;
    if (val <= 5) return 6;
    if (val <= 8) return 10;
    if (val <= 10) return 12; // Quando max <= 10, teto é 12 (deixa respiro visual de 20%)
    if (val <= 12) return 15;
    if (val <= 15) return 18;
    if (val <= 20) return 25;
    if (val <= 25) return 30;
    if (val <= 30) return 35;
    if (val <= 40) return 50;
    if (val <= 50) return 60;
    if (val <= 65) return 75;
    if (val <= 75) return 90;
    if (val <= 100) return 120;
    if (val <= 120) return 140; // Ex: Total = 111.9 -> teto 140 (respiro de 25%)
    if (val <= 140) return 160;
    if (val <= 160) return 180;
    if (val <= 180) return 200;
    if (val <= 200) return 250;
    if (val <= 250) return 300;
    if (val <= 300) return 350;
    if (val <= 400) return 450;
    if (val <= 500) return 600;
    if (val <= 750) return 900;
    if (val <= 1000) return 1200;

    const exponent = Math.floor(Math.log10(val));
    const magnitude = Math.pow(10, exponent);
    const fraction = val / magnitude;
    let niceFraction;
    if (fraction <= 1.2) niceFraction = 1.35;
    else if (fraction <= 1.4) niceFraction = 1.5;
    else if (fraction <= 1.6) niceFraction = 1.8;
    else if (fraction <= 2) niceFraction = 2.2;
    else if (fraction <= 2.5) niceFraction = 2.8;
    else if (fraction <= 3) niceFraction = 3.5;
    else if (fraction <= 5) niceFraction = 6;
    else if (fraction <= 7.5) niceFraction = 8.5;
    else niceFraction = 10;

    return Math.round(niceFraction * magnitude * 100) / 100;
}

// Calcula escala dinâmica do gráfico para que nenhum valor encoste ou ultrapasse o topo
function calculateDynamicScale(values, paddingRatio = 0.18) {
    const valid = (values || []).filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v));
    if (valid.length === 0) {
        return { min: 0, max: 12, rawMax: 10, rawMin: 0, beginAtZero: true, stepSize: 2 };
    }
    const min = Math.min(...valid);
    const max = Math.max(...valid);

    if (max === 0 && min === 0) {
        return { min: 0, max: 10, rawMax: 0, rawMin: 0, beginAtZero: true, stepSize: 2 };
    }

    const range = (max - min) || Math.abs(max) || 1;
    const padding = Math.max(range * paddingRatio, Math.abs(max) * paddingRatio, 0.5);

    const targetMax = max + padding;
    const targetMin = min >= 0 ? 0 : (min - padding);

    const finalMax = calculateNiceCeiling(targetMax);
    const finalMin = targetMin < 0 ? -calculateNiceCeiling(Math.abs(targetMin)) : 0;

    return {
        min: finalMin,
        max: finalMax,
        rawMax: max,
        rawMin: min,
        beginAtZero: min >= 0
    };
}

const TEMPLATES_INFO = {
    pie: {
        title: "🥧 Gráfico de Pizza (Proporções e Porcentagens)",
        desc: "Ideal para representar a divisão do todo, mostrando o percentual que cada fatia ocupa no orçamento."
    },
    bar: {
        title: "📊 Gráfico de Barras (Comparação com Escala Dinâmica)",
        desc: "Melhor para comparar dados entre categorias com limite superior automático que nunca corta as barras."
    },
    line: {
        title: "📈 Gráfico de Linha (Tendência Contínua com Folga Superior)",
        desc: "Ideal para exibir tendências e mudanças contínuas com respiro visual acima do ponto mais alto."
    },
    histogram: {
        title: "📶 Histograma (Distribuição e Frequência)",
        desc: "Visualiza a distribuição dos dados e a frequência de lançamentos dentro de faixas calculadas automaticamente."
    },
    scatter: {
        title: "🟣 Gráfico de Dispersão (Relações e Correlações)",
        desc: "Eficaz para demonstrar a dispersão entre registros com limites ajustados sem cortar pontos nas bordas."
    }
};

function setupTemplateStudio(rows, keys, descCol, valCol) {
    studioRowsData = rows;
    studioKeysData = keys;

    const selectCat = document.getElementById("chartSelectCategory");
    const selectVal = document.getElementById("chartSelectValue");
    const selectCount = document.getElementById("chartSelectCount");
    const selectOrder = document.getElementById("chartSelectOrder");

    selectCat.innerHTML = "";
    selectVal.innerHTML = "";

    keys.forEach(k => {
        selectCat.add(new Option(k, k));
        selectVal.add(new Option(k, k));
    });

    // Seleciona colunas sugeridas
    if (descCol && keys.includes(descCol)) selectCat.value = descCol;
    if (valCol && keys.includes(valCol)) selectVal.value = valCol;

    // Configura botões de templates
    document.querySelectorAll(".template-btn").forEach(btn => {
        btn.onclick = () => {
            document.querySelectorAll(".template-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            currentTemplateSelected = btn.dataset.template;
            
            // Atualiza explicação amigável
            const info = TEMPLATES_INFO[currentTemplateSelected];
            document.getElementById("templateTitle").textContent = info.title;
            document.getElementById("templateExplanation").textContent = info.desc;

            renderTemplateChart();
        };
    });

    selectCat.onchange = renderTemplateChart;
    selectVal.onchange = renderTemplateChart;
    if (selectCount) selectCount.onchange = renderTemplateChart;
    if (selectOrder) selectOrder.onchange = renderTemplateChart;

    renderTemplateChart();
}

function renderTemplateChart() {
    const canvas = document.getElementById("customTemplateChart");
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (customTemplateChartInstance) customTemplateChartInstance.destroy();

    const catKey = document.getElementById("chartSelectCategory").value;
    const valKey = document.getElementById("chartSelectValue").value;
    const countSelect = document.getElementById("chartSelectCount");
    const orderSelect = document.getElementById("chartSelectOrder");
    const countVal = countSelect ? countSelect.value : "10";
    const orderVal = orderSelect ? orderSelect.value : "desc";
    
    // Detecta dinamicamente se a coluna selecionada é moeda ou número puro (índice, nota, etc.)
    const tplColType = detectColumnType(valKey, studioRowsData);
    const fmt = (v) => formatSmartValue(v, tplColType);

    // Processa valores usando o conversor numérico universal
    const allData = [];
    studioRowsData.forEach((row, idx) => {
        const val = parseUniversalNumber(row[valKey]);
        if (val !== null && !isNaN(val)) {
            allData.push({
                index: idx + 1,
                label: String(row[catKey] || `Linha ${idx + 1}`).trim(),
                val: val
            });
        }
    });

    if (allData.length === 0) return;

    // Aplica ordenação escolhida pelo usuário
    let chartItems = [...allData];
    if (orderVal === "desc") {
        chartItems.sort((a, b) => b.val - a.val);
    } else if (orderVal === "asc") {
        chartItems.sort((a, b) => a.val - b.val);
    } // "original" mantém a ordem sequencial da planilha

    // Aplica limitação de quantidade escolhida pelo usuário
    if (countVal !== "all") {
        const limit = parseInt(countVal, 10) || 10;
        chartItems = chartItems.slice(0, limit);
    }

    // Calcula os limites dinâmicos da escala (Ceiling & Headroom) para que nenhum valor ultrapasse o topo
    const scaleBounds = calculateDynamicScale(chartItems.map(d => d.val), 0.20);

    // Atualiza o crachá de escala dinâmica visível para o usuário
    const scaleBadge = document.getElementById("chartScaleBadge");
    const scaleText = document.getElementById("chartScaleText");
    if (scaleBadge && scaleText) {
        if (currentTemplateSelected === "pie") {
            const sumValues = chartItems.reduce((acc, i) => acc + i.val, 0);
            scaleText.innerHTML = `🥧 <strong>Proporções em Pizza:</strong> Exibindo ${chartItems.length} registros que somam <strong>${fmt(sumValues)}</strong> (100% da amostra).`;
        } else if (currentTemplateSelected === "histogram") {
            scaleText.innerHTML = `📶 <strong>Histograma:</strong> Registros agrupados em 5 faixas calculadas automaticamente entre <strong>${fmt(scaleBounds.rawMin)}</strong> e <strong>${fmt(scaleBounds.rawMax)}</strong>.`;
        } else {
            scaleText.innerHTML = `📏 <strong>Escala Dinâmica do Eixo Y:</strong> Limite máximo ajustado para <strong>${fmt(scaleBounds.max)}</strong> (o maior registro é <strong>${fmt(scaleBounds.rawMax)}</strong>). Valores nunca tocam nem ultrapassam o teto do gráfico.`;
        }
        scaleBadge.style.display = "inline-flex";
    }

    let chartConfig = null;

    // 1. Template: Gráfico de Pizza (Pie Chart)
    if (currentTemplateSelected === "pie") {
        const groupMap = {};
        chartItems.forEach(d => {
            groupMap[d.label] = (groupMap[d.label] || 0) + d.val;
        });
        const sorted = Object.entries(groupMap).sort((a, b) => b[1] - a[1]);
        const top5 = sorted.slice(0, 5);
        const others = sorted.slice(5).reduce((acc, curr) => acc + curr[1], 0);

        const labels = top5.map(i => i[0]);
        const values = top5.map(i => i[1]);
        if (others > 0) {
            labels.push("Demais Registros");
            values.push(others);
        }

        const total = values.reduce((a, b) => a + b, 0);

        chartConfig = {
            type: "pie",
            data: {
                labels: labels,
                datasets: [{
                    data: values,
                    backgroundColor: ["#1e40af", "#0284c7", "#10b981", "#f59e0b", "#8b5cf6", "#94a3b8"],
                    borderWidth: 3,
                    borderColor: "#ffffff"
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 100,
                plugins: {
                    legend: { position: "bottom", labels: { font: { size: 14, weight: "bold" }, padding: 12 } },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => ` ${ctx.label}: ${fmt(ctx.parsed)} (${((ctx.parsed / (total || 1)) * 100).toFixed(1)}%)`
                        }
                    }
                }
            }
        };
    }

    // 2. Template: Gráfico de Barras (Bar Graph) com Teto Dinâmico
    else if (currentTemplateSelected === "bar") {
        chartConfig = {
            type: "bar",
            data: {
                labels: chartItems.map(d => d.label.length > 20 ? d.label.slice(0, 18) + "..." : d.label),
                datasets: [{
                    label: valKey,
                    data: chartItems.map(d => d.val),
                    backgroundColor: "#2563eb",
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 100,
                plugins: {
                    legend: { display: false },
                    tooltip: { 
                        titleFont: { size: 15 },
                        bodyFont: { size: 15 },
                        callbacks: { label: (ctx) => ` ${valKey}: ${fmt(ctx.parsed.y)}` } 
                    }
                },
                scales: {
                    y: { 
                        min: scaleBounds.min,
                        max: scaleBounds.max,
                        suggestedMax: scaleBounds.max,
                        beginAtZero: scaleBounds.beginAtZero,
                        title: { 
                            display: true, 
                            text: valKey, 
                            font: { size: 14, weight: "bold" },
                            color: "#0f172a" 
                        },
                        ticks: { 
                            font: { size: 13, weight: "bold" }, 
                            callback: (v) => fmt(v) 
                        },
                        grid: { color: "rgba(148, 163, 184, 0.25)" }
                    },
                    x: { 
                        ticks: { font: { size: 13, weight: "bold" }, maxRotation: 45, minRotation: 0 },
                        grid: { display: false }
                    }
                }
            }
        };
    }

    // 3. Template: Gráfico de Linha (Line Graph) com Folga Superior
    else if (currentTemplateSelected === "line") {
        chartConfig = {
            type: "line",
            data: {
                labels: chartItems.map(d => d.label),
                datasets: [{
                    label: valKey,
                    data: chartItems.map(d => d.val),
                    borderColor: "#2563eb",
                    backgroundColor: "rgba(37, 99, 235, 0.15)",
                    borderWidth: 3,
                    pointRadius: 6,
                    pointHoverRadius: 9,
                    pointBackgroundColor: "#1e40af",
                    fill: true,
                    tension: 0.3,
                    clip: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 100,
                plugins: {
                    legend: { labels: { font: { size: 14, weight: "bold" } } },
                    tooltip: { 
                        titleFont: { size: 15 },
                        bodyFont: { size: 15 },
                        callbacks: { label: (ctx) => ` ${valKey}: ${fmt(ctx.parsed.y)}` } 
                    }
                },
                scales: {
                    y: { 
                        min: scaleBounds.min,
                        max: scaleBounds.max,
                        suggestedMax: scaleBounds.max,
                        beginAtZero: scaleBounds.beginAtZero,
                        title: { 
                            display: true, 
                            text: valKey, 
                            font: { size: 14, weight: "bold" },
                            color: "#0f172a" 
                        },
                        ticks: { 
                            font: { size: 13, weight: "bold" }, 
                            callback: (v) => fmt(v) 
                        },
                        grid: { color: "rgba(148, 163, 184, 0.25)" }
                    },
                    x: { 
                        ticks: { font: { size: 12, weight: "bold" }, maxRotation: 45, minRotation: 0 },
                        grid: { display: false }
                    }
                }
            }
        };
    }

    // 4. Template: Histograma (Histogram - Distribuição e Frequência com Escala Dinâmica)
    else if (currentTemplateSelected === "histogram") {
        const values = chartItems.map(d => d.val);
        const minVal = Math.min(...values);
        const maxVal = Math.max(...values);
        const binCount = 5;
        const step = (maxVal - minVal) / binCount || 1;

        const bins = Array.from({ length: binCount }, (_, i) => {
            const start = minVal + i * step;
            const end = (i === binCount - 1) ? maxVal : start + step;
            return {
                start,
                end,
                label: `${fmt(start)} até ${fmt(end)}`,
                count: 0
            };
        });

        values.forEach(v => {
            for (let i = 0; i < bins.length; i++) {
                if (v >= bins[i].start && (i === bins.length - 1 ? v <= bins[i].end : v < bins[i].end)) {
                    bins[i].count++;
                    break;
                }
            }
        });

        const histCounts = bins.map(b => b.count);
        const histScale = calculateDynamicScale(histCounts, 0.25);

        chartConfig = {
            type: "bar",
            data: {
                labels: bins.map(b => b.label),
                datasets: [{
                    label: "Quantidade de Registros (Frequência)",
                    data: bins.map(b => b.count),
                    backgroundColor: "#10b981",
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 100,
                plugins: {
                    legend: { labels: { font: { size: 14, weight: "bold" } } },
                    tooltip: {
                        titleFont: { size: 15 },
                        bodyFont: { size: 15 },
                        callbacks: {
                            label: (ctx) => ` Frequência: ${ctx.parsed.y} registros nesta faixa`
                        }
                    }
                },
                scales: {
                    y: {
                        min: 0,
                        max: Math.max(histScale.max, 5),
                        suggestedMax: Math.max(histScale.max, 5),
                        beginAtZero: true,
                        ticks: { stepSize: 1, font: { size: 13, weight: "bold" }, precision: 0 },
                        grid: { color: "rgba(148, 163, 184, 0.25)" }
                    },
                    x: { 
                        ticks: { font: { size: 12, weight: "bold" }, maxRotation: 45, minRotation: 0 },
                        grid: { display: false }
                    }
                }
            }
        };
    }

    // 5. Template: Gráfico de Dispersão (Scatter Plot) com Limites Dinâmicos em X e Y
    else if (currentTemplateSelected === "scatter") {
        const indices = chartItems.map(d => d.index);
        const maxIndex = Math.max(...indices, 1);
        const maxX = Math.ceil(maxIndex * 1.08);

        chartConfig = {
            type: "scatter",
            data: {
                datasets: [{
                    label: "Distribuição dos Registros",
                    data: chartItems.map(d => ({ x: d.index, y: d.val, rawDesc: d.label })),
                    backgroundColor: "#8b5cf6",
                    borderColor: "#6d28d9",
                    pointRadius: 6,
                    pointHoverRadius: 10,
                    clip: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                resizeDelay: 100,
                plugins: {
                    legend: { labels: { font: { size: 14, weight: "bold" } } },
                    tooltip: {
                        titleFont: { size: 15 },
                        bodyFont: { size: 15 },
                        callbacks: {
                            label: (ctx) => {
                                const pt = ctx.raw;
                                return ` Linha #${pt.x} (${pt.rawDesc}): ${fmt(pt.y)}`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        min: 0,
                        max: maxX,
                        title: { display: true, text: "Ordem dos Registros na Planilha (Linha)", font: { size: 14, weight: "bold" }, color: "#0f172a" },
                        ticks: { font: { size: 12, weight: "bold" }, stepSize: Math.max(1, Math.floor(maxIndex / 10)) },
                        grid: { color: "rgba(148, 163, 184, 0.15)" }
                    },
                    y: {
                        min: scaleBounds.min,
                        max: scaleBounds.max,
                        suggestedMax: scaleBounds.max,
                        beginAtZero: scaleBounds.beginAtZero,
                        title: { display: true, text: valKey, font: { size: 14, weight: "bold" }, color: "#0f172a" },
                        ticks: { font: { size: 13, weight: "bold" }, callback: (v) => fmt(v) },
                        grid: { color: "rgba(148, 163, 184, 0.25)" }
                    }
                }
            }
        };
    }

    customTemplateChartInstance = new Chart(ctx, chartConfig);
}

// 9. Comparador Interativo de Itens
let comparatorList = [];
let comparisonChartInstance = null;
let lastComparisonSpeechText = "";
let comparatorColType = null;

function setupComparator(items, colType) {
    comparatorList = items;
    comparatorColType = colType || { isCurrency: false };
    const selectA = document.getElementById("selectItemA");
    const selectB = document.getElementById("selectItemB");
    const fmt = (v) => formatSmartValue(v, comparatorColType);

    selectA.innerHTML = "";
    selectB.innerHTML = "";

    items.forEach((item, index) => {
        const text = `${item.desc} (${fmt(item.val)})`;
        selectA.add(new Option(text, index));
        selectB.add(new Option(text, index));
    });

    // Pré-seleciona o 1º e 2º item
    if (items.length > 1) {
        selectA.selectedIndex = 0;
        selectB.selectedIndex = 1;
    }

    // Executa a comparação inicial
    executeComparison();
}

function selectItemForComparison(encodedDesc) {
    const desc = decodeURIComponent(encodedDesc);
    const index = comparatorList.findIndex(i => i.desc === desc);
    if (index !== -1) {
        const selectB = document.getElementById("selectItemB");
        selectB.selectedIndex = index;
        executeComparison();
        document.querySelector(".comparison-section").scrollIntoView({ behavior: "smooth" });
    }
}

function executeComparison() {
    if (comparatorList.length < 2) return;

    const selectA = document.getElementById("selectItemA");
    const selectB = document.getElementById("selectItemB");
    const itemA = comparatorList[selectA.value];
    const itemB = comparatorList[selectB.value];

    if (!itemA || !itemB) return;

    const fmt = (v) => formatSmartValue(v, comparatorColType);
    const diff = Math.abs(itemA.val - itemB.val);
    const isAGreater = itemA.val >= itemB.val;
    const ratio = itemB.val > 0 ? (itemA.val / itemB.val) : 1;
    const percentDiff = itemB.val > 0 ? Math.round(((itemA.val - itemB.val) / itemB.val) * 100) : 0;

    // A) Texto Narrativo Adaptativo (Moeda vs Número Geral)
    let narrative = "";
    if (itemA.val === itemB.val) {
        narrative = `Os dois itens selecionados (<strong>${itemA.desc}</strong> e <strong>${itemB.desc}</strong>) possuem exatamente o mesmo valor: <strong>${fmt(itemA.val)}</strong>.`;
    } else {
        const maior = isAGreater ? itemA : itemB;
        const menor = isAGreater ? itemB : itemA;
        const vezes = (maior.val / (menor.val || 1)).toFixed(1);

        if (comparatorColType && comparatorColType.isCurrency) {
            narrative = `O item <strong>"${maior.desc}"</strong> (${fmt(maior.val)}) é <strong>${fmt(diff)} mais caro</strong> que <strong>"${menor.desc}"</strong> (${fmt(menor.val)}). `;
        } else {
            narrative = `O registro <strong>"${maior.desc}"</strong> (${fmt(maior.val)}) tem um valor <strong>${fmt(diff)} maior</strong> que <strong>"${menor.desc}"</strong> (${fmt(menor.val)}). `;
        }
        
        narrative += `Isso significa que o primeiro é aproximadamente <strong>${vezes}x o valor</strong> do segundo`;
        if (menor.val > 0) {
            narrative += ` (uma diferença proporcional de <strong>${Math.abs(percentDiff)}%</strong>).`;
        } else {
            narrative += `.`;
        }
    }

    document.getElementById("comparisonNarrativeText").innerHTML = narrative;
    lastComparisonSpeechText = narrative.replace(/<[^>]*>?/gm, " ");

    const diffCardTitle = (comparatorColType && comparatorColType.isCurrency) ? "Diferença em Dinheiro" : "Diferença Numérica";

    // B) Cartões com Números Grandes
    document.getElementById("comparisonCardsGrid").innerHTML = `
        <div class="metric-card" style="border-top: 6px solid #2563eb;">
            <h3>1º: ${itemA.desc.slice(0, 22)}</h3>
            <div class="metric-value" style="color: #2563eb;">${fmt(itemA.val)}</div>
        </div>
        <div class="metric-card" style="border-top: 6px solid #f59e0b;">
            <h3>2º: ${itemB.desc.slice(0, 22)}</h3>
            <div class="metric-value" style="color: #d97706;">${fmt(itemB.val)}</div>
        </div>
        <div class="metric-card" style="border-top: 6px solid #dc2626;">
            <h3>${diffCardTitle}</h3>
            <div class="metric-value highlight-red">${fmt(diff)}</div>
        </div>
    `;

    // C) Gráfico de Pizza Comparativo (Fatia A vs Fatia B)
    const ctxComp = document.getElementById("comparisonPieChartCanvas").getContext("2d");
    if (comparisonChartInstance) comparisonChartInstance.destroy();

    const totalComp = (itemA.val + itemB.val) || 1;
    const pctA = ((itemA.val / totalComp) * 100).toFixed(1);
    const pctB = ((itemB.val / totalComp) * 100).toFixed(1);

    comparisonChartInstance = new Chart(ctxComp, {
        type: "pie",
        data: {
            labels: [
                `${itemA.desc} (${pctA}%)`,
                `${itemB.desc} (${pctB}%)`
            ],
            datasets: [{
                data: [itemA.val, itemB.val],
                backgroundColor: ["#2563eb", "#f59e0b"],
                borderWidth: 4,
                borderColor: "#ffffff"
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            resizeDelay: 100,
            plugins: {
                legend: {
                    position: "bottom",
                    labels: {
                        font: { size: 15, weight: "bold" },
                        padding: 16
                    }
                },
                tooltip: {
                    titleFont: { size: 16 },
                    bodyFont: { size: 16 },
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.parsed;
                            const pct = ((val / totalComp) * 100).toFixed(1);
                            return ` ${fmt(val)} (${pct}% da soma dos dois)`;
                        }
                    }
                }
            }
        }
    });

    document.getElementById("comparisonResultBox").style.display = "block";
    document.getElementById("btnSpeakComparison").style.display = "inline-block";
}

// Eventos do Comparador
document.getElementById("btnCompareItems").addEventListener("click", executeComparison);
document.getElementById("selectItemA").addEventListener("change", executeComparison);
document.getElementById("selectItemB").addEventListener("change", executeComparison);

document.getElementById("btnSpeakComparison").addEventListener("click", () => {
    if (!('speechSynthesis' in window)) return alert("Navegador sem suporte a voz.");
    if (window.speechSynthesis.speaking) {
        window.speechSynthesis.cancel();
        document.getElementById("btnSpeakComparison").textContent = "🔊 Ouvir Comparação";
        return;
    }
    const utterance = new SpeechSynthesisUtterance(lastComparisonSpeechText);
    utterance.lang = "pt-BR";
    utterance.rate = 0.9;
    utterance.onend = () => { document.getElementById("btnSpeakComparison").textContent = "🔊 Ouvir Comparação"; };
    window.speechSynthesis.speak(utterance);
    document.getElementById("btnSpeakComparison").textContent = "⏹️ Parar Voz";
});

// 10. Eventos de Navegação e Leitura
document.getElementById("btnPrevPage").addEventListener("click", () => {
    if (tablePage > 1) {
        tablePage--;
        renderTablePage();
    }
});

document.getElementById("btnNextPage").addEventListener("click", () => {
    const totalPages = Math.ceil(currentRowsData.length / PAGE_SIZE);
    if (tablePage < totalPages) {
        tablePage++;
        renderTablePage();
    }
});

document.getElementById("btnToggleAllRows").addEventListener("click", () => {
    showingAllRows = !showingAllRows;
    renderTablePage();
});

document.getElementById("btnSpeak").addEventListener("click", () => {
    if (!('speechSynthesis' in window)) return alert("Navegador sem suporte a voz.");
    if (window.speechSynthesis.speaking) {
        window.speechSynthesis.cancel();
        document.getElementById("btnSpeak").textContent = "🔊 Ouvir Resumo";
        return;
    }
    const textToSpeak = narrativeText.dataset.rawText || narrativeText.innerText;
    const utterance = new SpeechSynthesisUtterance(textToSpeak);
    utterance.lang = "pt-BR";
    utterance.rate = 0.9;
    utterance.onend = () => { document.getElementById("btnSpeak").textContent = "🔊 Ouvir Resumo"; };
    window.speechSynthesis.speak(utterance);
    document.getElementById("btnSpeak").textContent = "⏹️ Parar Voz";
});

document.getElementById("btnPrint").addEventListener("click", () => window.print());

// 11. Geração e Gerenciamento do Resumo em Texto (.txt)
function generateSummaryText() {
    if (!currentRowsData || currentRowsData.length === 0) {
        return "";
    }

    const now = new Date();
    const dataHora = now.toLocaleString("pt-BR", { dateStyle: "long", timeStyle: "short" });
    const fileName = currentFileName || document.getElementById("fileNameDisplay")?.textContent || "Planilha";
    const sheetName = currentActiveSheet || "Principal";

    const sep = "========================================================================\n";
    const subsep = "------------------------------------------------------------------------\n";

    let txt = sep;
    txt += "                RESUMO FÁCIL DE PLANILHAS - RELATÓRIO TXT\n";
    txt += "             Leitura Clara, Simplificada e Acessível de Dados\n";
    txt += sep + "\n";

    txt += `📄 Arquivo Analisado: ${fileName}\n`;
    txt += `📑 Aba / Pasta:       ${sheetName}\n`;
    txt += `📅 Data de Emissão:   ${dataHora}\n`;
    txt += `📊 Coluna em Foco:    ${currentValueCol || "Valor Principal"}\n`;
    txt += subsep + "\n";

    // 1. Resumo Narrativo Simplificado
    txt += "1. O QUE ESTA PLANILHA MOSTRA (RESUMO EM LINGUAGEM SIMPLES):\n\n";
    const rawNarrative = narrativeText?.dataset?.rawText || narrativeText?.innerText || "";
    if (rawNarrative) {
        const sentences = rawNarrative.split(/(?<=[.?!])\s+/);
        let line = "  ";
        sentences.forEach(sentence => {
            if (line.length + sentence.length > 70) {
                txt += line.trim() + "\n";
                line = "  " + sentence + " ";
            } else {
                line += sentence + " ";
            }
        });
        if (line.trim().length > 0) txt += line.trim() + "\n";
    }
    txt += "\n" + subsep + "\n";

    // 2. Indicadores e Números Principais
    txt += "2. INDICADORES PRINCIPAIS CONSOLIDADOS:\n\n";
    const cards = document.querySelectorAll("#metricsGrid .metric-card");
    cards.forEach(card => {
        const title = card.querySelector("h3")?.textContent || "";
        const val = card.querySelector(".metric-value")?.textContent || "";
        txt += `  • ${title.padEnd(28, " ")}: ${val}\n`;
    });
    txt += "\n" + subsep + "\n";

    // 3. Maiores Destaques Identificados
    txt += "3. MAIORES DESTAQUES IDENTIFICADOS (TOP 5):\n\n";
    const topRows = document.querySelectorAll("#topItemsList .top-item-row");
    topRows.forEach((row, idx) => {
        const name = row.querySelector(".top-item-name")?.textContent.replace(/^📌\s*/, "") || "";
        const val = row.querySelector(".top-item-value")?.textContent || "";
        txt += `  [${idx + 1}º Lugar] ${name}\n`;
        txt += `             Valor registrado: ${val}\n\n`;
    });
    txt += subsep + "\n";

    // 4. Comparador de Itens (se utilizado)
    const compBox = document.getElementById("comparisonResultBox");
    if (compBox && compBox.style.display !== "none" && lastComparisonSpeechText) {
        txt += "4. RESULTADO DA COMPARAÇÃO ENTRE ITENS:\n\n";
        txt += `  ${lastComparisonSpeechText.trim()}\n\n`;
        const compCards = document.querySelectorAll("#comparisonCardsGrid .metric-card");
        compCards.forEach(c => {
            const t = c.querySelector("h3")?.textContent || "";
            const v = c.querySelector(".metric-value")?.textContent || "";
            txt += `  • ${t.padEnd(28, " ")}: ${v}\n`;
        });
        txt += "\n" + subsep + "\n";
    }

    // 5. Configuração do Modelo Visual Ativo
    txt += "5. GRÁFICO SELECIONADO NA VISUALIZAÇÃO:\n\n";
    const tplTitle = document.getElementById("templateTitle")?.textContent || "";
    const scaleText = document.getElementById("chartScaleText")?.textContent || "";
    txt += `  • Modelo: ${tplTitle}\n`;
    if (scaleText) {
        txt += `  • Escala: ${scaleText}\n`;
    }
    txt += "\n" + subsep + "\n";

    // 6. Amostra dos Primeiros Lançamentos da Tabela
    txt += "6. AMOSTRA DOS PRIMEIROS LANÇAMENTOS (TABELA):\n\n";
    if (currentRowsData && currentRowsData.length > 0) {
        const sampleKeys = currentKeysData.slice(0, 5);
        txt += "  " + sampleKeys.map(k => k.slice(0, 16).padEnd(18, " ")).join("| ") + "\n";
        txt += "  " + "-".repeat(Math.min(90, sampleKeys.length * 20)) + "\n";
        currentRowsData.slice(0, 10).forEach(r => {
            const rowLine = sampleKeys.map(k => String(r[k] !== undefined ? r[k] : "").slice(0, 16).padEnd(18, " ")).join("| ");
            txt += `  ${rowLine}\n`;
        });
        if (currentRowsData.length > 10) {
            txt += `\n  ... e mais ${currentRowsData.length - 10} linhas registradas na planilha.\n`;
        }
    }
    txt += "\n" + sep;
    txt += "Documento gerado automaticamente pelo Resumo Fácil de Planilhas.\n";
    txt += "Compatível com qualquer leitor de tela, computador ou celular.\n";
    txt += sep;

    return txt;
}

function openTxtModal() {
    if (!currentRowsData || currentRowsData.length === 0) {
        alert("Carregue uma planilha primeiro para visualizar o resumo em texto.");
        return;
    }
    const txt = generateSummaryText();
    const modal = document.getElementById("txtSummaryModal");
    const textarea = document.getElementById("txtSummaryTextarea");
    const feedback = document.getElementById("copySuccessFeedback");
    
    if (modal && textarea) {
        textarea.value = txt;
        if (feedback) feedback.style.display = "none";
        modal.style.display = "flex";
        setTimeout(() => {
            textarea.focus();
            textarea.select();
        }, 100);
    }
}

function closeTxtModal() {
    const modal = document.getElementById("txtSummaryModal");
    if (modal) modal.style.display = "none";
}

function downloadTxtFile() {
    const txt = generateSummaryText();
    if (!txt) return;
    const fileName = currentFileName || document.getElementById("fileNameDisplay")?.textContent || "Planilha";
    const blob = new Blob([txt], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const safeName = (fileName.replace(/\.[^/.]+$/, "") || "resumo_planilha").replace(/[^a-zA-Z0-9_-]/g, "_");
    a.href = url;
    a.download = `Resumo_${safeName}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Função auxiliar para higienizar textos no PDF, removendo emojis e evitando caracteres corrompidos (mojibake)
function cleanTextForPdf(str) {
    if (!str) return "";
    let clean = String(str);
    
    // Remove todos os emojis unicode (surrogates e símbolos gráficos)
    clean = clean.replace(/[\u{10000}-\u{10FFFF}]/gu, "");
    clean = clean.replace(/[\u{1F300}-\u{1FAFF}]/gu, "");
    clean = clean.replace(/[\u{2600}-\u{27BF}]/gu, "");
    clean = clean.replace(/[\u{2300}-\u{23FF}]/gu, "");
    clean = clean.replace(/[\u{2B50}-\u{2B55}]/gu, "");
    clean = clean.replace(/[\u{FE00}-\u{FE0F}]/gu, "");
    clean = clean.replace(/[\u{E000}-\u{F8FF}]/gu, "");
    clean = clean.replace(/[🥧📊📈📶🟣📌📏💡⚖️📄📋📁🔍⏹️🔊🖨️🔄⬅️➡️⭐❌✅🔲]/g, "");

    // Normaliza acentuação e caracteres especiais para compatibilidade com fontes PDF padrão
    clean = clean
        .replace(/[áàãâä]/g, "a")
        .replace(/[ÁÀÃÂÄ]/g, "A")
        .replace(/[éèêë]/g, "e")
        .replace(/[ÉÈÊË]/g, "E")
        .replace(/[íìîï]/g, "i")
        .replace(/[ÍÌÎÏ]/g, "I")
        .replace(/[óòõôö]/g, "o")
        .replace(/[ÓÒÕÔÖ]/g, "O")
        .replace(/[úùûü]/g, "u")
        .replace(/[ÚÙÛÜ]/g, "U")
        .replace(/[ç]/g, "c")
        .replace(/[Ç]/g, "C")
        .replace(/[–—]/g, "-")
        .replace(/[º°]/g, "o.")
        .replace(/[ª]/g, "a.");

    return clean.replace(/\s+/g, " ").trim();
}

// 12. Exportação de Relatório em PDF com Gráficos Reais Embutidos (jsPDF)
function exportSummaryAsPdf() {
    if (!currentRowsData || currentRowsData.length === 0) {
        alert("Carregue uma planilha primeiro para exportar o resumo em PDF.");
        return;
    }

    if (!window.jspdf || !window.jspdf.jsPDF) {
        window.print();
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4"
    });

    const now = new Date();
    const dataHora = now.toLocaleString("pt-BR", { dateStyle: "long", timeStyle: "short" });
    const rawFileName = currentFileName || document.getElementById("fileNameDisplay")?.textContent || "Planilha";
    const fileName = cleanTextForPdf(rawFileName);
    const sheetName = cleanTextForPdf(currentActiveSheet || "Principal");
    const safeName = (rawFileName.replace(/\.[^/.]+$/, "") || "resumo_planilha").replace(/[^a-zA-Z0-9_-]/g, "_");

    let cursorY = 15;
    const pageWidth = 210;
    const margin = 15;
    const contentWidth = pageWidth - (margin * 2); // 180mm

    function checkPageBreak(spaceNeeded) {
        if (cursorY + spaceNeeded > 275) {
            doc.addPage();
            cursorY = 20;
            return true;
        }
        return false;
    }

    // --- CABEÇALHO AZUL ---
    doc.setFillColor(30, 64, 175);
    doc.rect(0, 0, pageWidth, 24, "F");

    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("RESUMO FACIL DE PLANILHAS", margin, 11);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text("Relatorio Sintetico e Acessivel de Dados", margin, 18);

    cursorY = 32;

    // --- QUADRO DE METADADOS ---
    doc.setFillColor(241, 245, 249);
    doc.setDrawColor(203, 213, 225);
    doc.roundedRect(margin, cursorY, contentWidth, 22, 2, 2, "FD");

    doc.setTextColor(30, 41, 59);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("Arquivo:", margin + 4, cursorY + 7);
    doc.setFont("helvetica", "normal");
    doc.text(`${fileName.slice(0, 45)} (Aba: ${sheetName})`, margin + 22, cursorY + 7);

    doc.setFont("helvetica", "bold");
    doc.text("Data/Hora:", margin + 4, cursorY + 15);
    doc.setFont("helvetica", "normal");
    const metricaClean = cleanTextForPdf(currentValueCol || "Valor Principal");
    doc.text(`${dataHora} | Metrica em Foco: ${metricaClean}`, margin + 25, cursorY + 15);

    cursorY += 30;

    // --- 1. RESUMO EXPLICADO ---
    doc.setTextColor(30, 64, 175);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("1. O QUE ESTA PLANILHA MOSTRA (RESUMO SIMPLIFICADO)", margin, cursorY);
    cursorY += 6;

    const rawNarrative = narrativeText?.dataset?.rawText || narrativeText?.innerText || "";
    if (rawNarrative) {
        doc.setFillColor(236, 253, 245);
        doc.setDrawColor(16, 185, 129);
        
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10.5);
        doc.setTextColor(6, 78, 59);
        const narrativeClean = cleanTextForPdf(rawNarrative);
        const narrativeLines = doc.splitTextToSize(narrativeClean, contentWidth - 8);
        const boxHeight = (narrativeLines.length * 5.2) + 8;
        
        checkPageBreak(boxHeight + 5);
        doc.roundedRect(margin, cursorY, contentWidth, boxHeight, 2, 2, "FD");
        doc.text(narrativeLines, margin + 4, cursorY + 6);
        cursorY += boxHeight + 8;
    }

    // --- 2. INDICADORES PRINCIPAIS ---
    checkPageBreak(35);
    doc.setTextColor(30, 64, 175);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("2. INDICADORES E NUMEROS PRINCIPAIS", margin, cursorY);
    cursorY += 6;

    const metricCards = document.querySelectorAll("#metricsGrid .metric-card");
    if (metricCards.length > 0) {
        const cardWidth = (contentWidth - 6) / 2;
        let cardY = cursorY;
        let rowMaxY = cursorY;

        metricCards.forEach((c, idx) => {
            const title = cleanTextForPdf(c.querySelector("h3")?.textContent || "");
            const val = cleanTextForPdf(c.querySelector(".metric-value")?.textContent || "");
            
            const isRight = (idx % 2 === 1);
            const x = isRight ? (margin + cardWidth + 6) : margin;
            const y = cardY;

            doc.setFillColor(248, 250, 252);
            doc.setDrawColor(203, 213, 225);
            doc.roundedRect(x, y, cardWidth, 16, 2, 2, "FD");

            doc.setFont("helvetica", "normal");
            doc.setFontSize(9);
            doc.setTextColor(71, 85, 105);
            doc.text(title.slice(0, 28), x + 4, y + 6);

            doc.setFont("helvetica", "bold");
            doc.setFontSize(12);
            doc.setTextColor(30, 64, 175);
            doc.text(val, x + 4, y + 13);

            if (isRight) {
                cardY += 20;
            }
            rowMaxY = Math.max(rowMaxY, y + 18);
        });
        cursorY = rowMaxY + 8;
    }

    // --- 3. GRÁFICO VISUAL GERADO (CAPTURA REAL DO CANVAS COM PROPORÇÃO PRESERVADA) ---
    const chartCanvas = document.getElementById("customTemplateChart");
    let chartImgData = null;
    let renderW = 0;
    let renderH = 0;
    let renderX = margin;

    if (chartCanvas && chartCanvas.width > 0 && chartCanvas.height > 0) {
        try {
            // Desenha em canvas temporário com fundo branco sólido para o PDF
            const tempCanvas = document.createElement("canvas");
            tempCanvas.width = chartCanvas.width;
            tempCanvas.height = chartCanvas.height;
            const tempCtx = tempCanvas.getContext("2d");
            tempCtx.fillStyle = "#ffffff";
            tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
            tempCtx.drawImage(chartCanvas, 0, 0);
            chartImgData = tempCanvas.toDataURL("image/png", 1.0);

            // Preservação estrita e matemática da proporção original (aspect ratio):
            // O gráfico nunca é esticado ou achatado, mantendo círculos e textos 100% fiéis
            const canvasAspect = chartCanvas.width / chartCanvas.height; // largura / altura original
            const maxW = 160; // mm (largura máxima confortável na folha A4)
            const maxH = 85;  // mm (altura máxima)

            renderW = maxW;
            renderH = renderW / canvasAspect;

            if (renderH > maxH) {
                renderH = maxH;
                renderW = renderH * canvasAspect;
            }

            // Centraliza horizontalmente na área de conteúdo
            renderX = margin + (contentWidth - renderW) / 2;
        } catch (err) {
            console.warn("Nao foi possivel converter o grafico para imagem:", err);
        }
    }

    const tplTitle = cleanTextForPdf(document.getElementById("templateTitle")?.textContent || "Grafico");
    const scaleText = cleanTextForPdf(document.getElementById("chartScaleText")?.textContent || "");

    checkPageBreak(chartImgData ? (renderH + 35) : 25);
    doc.setTextColor(30, 64, 175);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("3. GRAFICO VISUAL GERADO", margin, cursorY);
    cursorY += 6;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.setTextColor(30, 41, 59);
    doc.text(`Modelo Ativo: ${tplTitle}`, margin, cursorY);
    cursorY += 5;

    if (scaleText) {
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.setTextColor(71, 85, 105);
        const sLines = doc.splitTextToSize(`Escala / Configuracao: ${scaleText}`, contentWidth);
        doc.text(sLines, margin, cursorY);
        cursorY += (sLines.length * 4.5) + 3;
    }

    if (chartImgData && renderW > 0 && renderH > 0) {
        // Moldura em volta do gráfico dimensionada proporcionalmente
        doc.setDrawColor(203, 213, 225);
        doc.setFillColor(255, 255, 255);
        doc.roundedRect(renderX - 2, cursorY, renderW + 4, renderH + 4, 2, 2, "FD");

        doc.addImage(chartImgData, "PNG", renderX, cursorY + 2, renderW, renderH);
        cursorY += renderH + 10;
    }

    // --- 4. MAIORES DESTAQUES ---
    checkPageBreak(40);
    doc.setTextColor(30, 64, 175);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("4. MAIORES DESTAQUES IDENTIFICADOS (TOP 5)", margin, cursorY);
    cursorY += 6;

    const topRows = document.querySelectorAll("#topItemsList .top-item-row");
    topRows.forEach((row, i) => {
        checkPageBreak(12);
        const rawName = row.querySelector(".top-item-name")?.textContent || "";
        const name = cleanTextForPdf(rawName);
        const val = cleanTextForPdf(row.querySelector(".top-item-value")?.textContent || "");

        doc.setFillColor(i % 2 === 0 ? 255 : 248, i % 2 === 0 ? 255 : 250, i % 2 === 0 ? 255 : 252);
        doc.setDrawColor(226, 232, 240);
        doc.rect(margin, cursorY, contentWidth, 8, "FD");

        doc.setFont("helvetica", "bold");
        doc.setFontSize(9.5);
        doc.setTextColor(30, 41, 59);
        doc.text(`${i + 1}o. ${name.slice(0, 50)}`, margin + 3, cursorY + 5.5);

        doc.setFont("helvetica", "bold");
        doc.setTextColor(185, 28, 28);
        doc.text(val, margin + contentWidth - 4, cursorY + 5.5, { align: "right" });

        cursorY += 8;
    });
    cursorY += 8;

    // --- 5. COMPARACAO (SE ATIVA) ---
    const compBox = document.getElementById("comparisonResultBox");
    if (compBox && compBox.style.display !== "none" && lastComparisonSpeechText) {
        checkPageBreak(50);
        doc.setTextColor(30, 64, 175);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(12);
        doc.text("5. COMPARACAO ENTRE ITENS SELECIONADOS", margin, cursorY);
        cursorY += 6;

        doc.setFillColor(239, 246, 255);
        doc.setDrawColor(59, 130, 246);
        const compClean = cleanTextForPdf(lastComparisonSpeechText);
        const compLines = doc.splitTextToSize(compClean, contentWidth - 8);
        const cBoxH = (compLines.length * 5) + 8;
        doc.roundedRect(margin, cursorY, contentWidth, cBoxH, 2, 2, "FD");

        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);
        doc.setTextColor(30, 58, 138);
        doc.text(compLines, margin + 4, cursorY + 6);
        cursorY += cBoxH + 6;

        // Imagem do Gráfico de Comparação com proporção preservada
        const compCanvas = document.getElementById("comparisonPieChartCanvas");
        if (compCanvas && compCanvas.width > 0 && compCanvas.height > 0) {
            try {
                const tempC = document.createElement("canvas");
                tempC.width = compCanvas.width;
                tempC.height = compCanvas.height;
                const tCtx = tempC.getContext("2d");
                tCtx.fillStyle = "#ffffff";
                tCtx.fillRect(0, 0, tempC.width, tempC.height);
                tCtx.drawImage(compCanvas, 0, 0);
                const compImg = tempC.toDataURL("image/png", 1.0);

                const compAspect = compCanvas.width / compCanvas.height;
                const compMaxW = 85;
                const compMaxH = 75;
                let cImgW = compMaxW;
                let cImgH = cImgW / compAspect;
                if (cImgH > compMaxH) {
                    cImgH = compMaxH;
                    cImgW = cImgH * compAspect;
                }
                const compX = margin + (contentWidth - cImgW) / 2;

                checkPageBreak(cImgH + 10);
                doc.setDrawColor(203, 213, 225);
                doc.setFillColor(255, 255, 255);
                doc.roundedRect(compX - 2, cursorY, cImgW + 4, cImgH + 4, 2, 2, "FD");
                doc.addImage(compImg, "PNG", compX, cursorY + 2, cImgW, cImgH);
                cursorY += cImgH + 10;
            } catch (e) {
                console.warn(e);
            }
        }
    }

    // --- RODAPÉ EM TODAS AS PÁGINAS ---
    const totalPages = doc.internal.getNumberOfPages();
    for (let p = 1; p <= totalPages; p++) {
        doc.setPage(p);
        doc.setDrawColor(203, 213, 225);
        doc.line(margin, 285, margin + contentWidth, 285);

        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        doc.setTextColor(148, 163, 184);
        doc.text("Resumo Facil de Planilhas - Relatorio PDF Acessivel", margin, 290);
        doc.text(`Pagina ${p} de ${totalPages}`, margin + contentWidth, 290, { align: "right" });
    }

    doc.save(`Resumo_${safeName}.pdf`);
}

// Vinculação dos Eventos de Interface
const btnToggleCurrencyEl = document.getElementById("btnToggleCurrency");
if (btnToggleCurrencyEl) {
    btnToggleCurrencyEl.addEventListener("click", () => {
        if (userForcedCurrencyMode === null) {
            userForcedCurrencyMode = "currency_brl";
        } else if (userForcedCurrencyMode === "currency_brl") {
            userForcedCurrencyMode = "number";
        } else {
            userForcedCurrencyMode = null;
        }

        if (lastAnalyzedState) {
            analyzeData(
                lastAnalyzedState.rows,
                lastAnalyzedState.keys,
                lastAnalyzedState.sheetName,
                lastAnalyzedState.descriptionColumn,
                lastAnalyzedState.valueColumn
            );
        }
    });
}

const btnExportTxtEl = document.getElementById("btnExportTxt");
if (btnExportTxtEl) {
    btnExportTxtEl.addEventListener("click", openTxtModal);
}

const btnExportPdfEl = document.getElementById("btnExportPdf");
if (btnExportPdfEl) {
    btnExportPdfEl.addEventListener("click", exportSummaryAsPdf);
}

const btnCloseTxtModalEl = document.getElementById("btnCloseTxtModal");
if (btnCloseTxtModalEl) {
    btnCloseTxtModalEl.addEventListener("click", closeTxtModal);
}

const btnCloseTxtModalBottomEl = document.getElementById("btnCloseTxtModalBottom");
if (btnCloseTxtModalBottomEl) {
    btnCloseTxtModalBottomEl.addEventListener("click", closeTxtModal);
}

const modalOverlayEl = document.getElementById("txtSummaryModal");
if (modalOverlayEl) {
    modalOverlayEl.addEventListener("click", (e) => {
        if (e.target === modalOverlayEl) closeTxtModal();
    });
}

window.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modalOverlayEl && modalOverlayEl.style.display !== "none") {
        closeTxtModal();
    }
});

const btnCopyTxtContentEl = document.getElementById("btnCopyTxtContent");
if (btnCopyTxtContentEl) {
    btnCopyTxtContentEl.addEventListener("click", () => {
        const textarea = document.getElementById("txtSummaryTextarea");
        if (!textarea) return;
        
        textarea.select();
        textarea.setSelectionRange(0, 99999);

        const textToCopy = textarea.value;
        const feedback = document.getElementById("copySuccessFeedback");

        const showSuccess = () => {
            if (feedback) feedback.style.display = "block";
            const originalText = btnCopyTxtContentEl.innerHTML;
            btnCopyTxtContentEl.innerHTML = "✅ Copiado!";
            setTimeout(() => {
                btnCopyTxtContentEl.innerHTML = originalText;
                if (feedback) feedback.style.display = "none";
            }, 3500);
        };

        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(textToCopy).then(showSuccess).catch(() => {
                document.execCommand("copy");
                showSuccess();
            });
        } else {
            document.execCommand("copy");
            showSuccess();
        }
    });
}

const btnSelectAllTxtEl = document.getElementById("btnSelectAllTxt");
if (btnSelectAllTxtEl) {
    btnSelectAllTxtEl.addEventListener("click", () => {
        const textarea = document.getElementById("txtSummaryTextarea");
        if (textarea) {
            textarea.focus();
            textarea.select();
        }
    });
}

const btnDownloadTxtDirectEl = document.getElementById("btnDownloadTxtDirect");
if (btnDownloadTxtDirectEl) {
    btnDownloadTxtDirectEl.addEventListener("click", downloadTxtFile);
}

document.getElementById("btnNewFile").addEventListener("click", () => {
    resultsSection.style.display = "none";
    dropZone.style.display = "block";
    fileInput.value = "";
    closeTxtModal();
    if (customTemplateChartInstance) customTemplateChartInstance.destroy();
    if (window.speechSynthesis.speaking) window.speechSynthesis.cancel();
});

