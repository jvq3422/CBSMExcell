// Servidor HTTP leve e sem dependências para visualização do Resumo Fácil de Planilhas
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');

const PORT = process.env.PORT || 8080;
const BASE_DIR = __dirname;

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.csv': 'text/csv; charset=utf-8',
    '.txt': 'text/plain; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xls': 'application/vnd.ms-excel'
};

function getNetworkAddresses() {
    const interfaces = os.networkInterfaces();
    const addresses = [];
    for (const name of Object.keys(interfaces)) {
        for (const iface of interfaces[name]) {
            if (iface.family === 'IPv4' && !iface.internal) {
                addresses.push(iface.address);
            }
        }
    }
    return addresses;
}

const server = http.createServer((req, res) => {
    // Normaliza URL
    let safeUrl = req.url.split('?')[0].split('#')[0];
    if (safeUrl === '/' || safeUrl === '') {
        safeUrl = '/index.html';
    }

    const decodedPath = decodeURIComponent(safeUrl);
    const filePath = path.normalize(path.join(BASE_DIR, decodedPath));

    // Proteção de segurança contra Directory Traversal
    if (!filePath.startsWith(BASE_DIR)) {
        res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Acesso proibido');
        return;
    }

    fs.stat(filePath, (err, stats) => {
        if (err || !stats.isFile()) {
            res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            res.end('Arquivo não encontrado: ' + safeUrl);
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME_TYPES[ext] || 'application/octet-stream';

        res.writeHead(200, {
            'Content-Type': contentType,
            'Content-Length': stats.size,
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Access-Control-Allow-Origin': '*'
        });

        const readStream = fs.createReadStream(filePath);
        readStream.pipe(res);
    });
});

server.listen(PORT, '0.0.0.0', () => {
    const ips = getNetworkAddresses();
    console.log(`========================================================`);
    console.log(` 🚀 RESUMO FÁCIL DE PLANILHAS - SERVIDOR ATIVO`);
    console.log(`========================================================`);
    console.log(` 👉 Link Local (Navegador):    http://localhost:${PORT}`);
    ips.forEach(ip => {
        console.log(` 📱 Link Mobile (Wi-Fi/Rede):  http://${ip}:${PORT}`);
    });
    console.log(`========================================================\n`);
});
