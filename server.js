const http = require('http');
const fs   = require('fs');
const path = require('path');

const PORT = 3000;
const MIME = {
  '.html':'text/html', '.css':'text/css', '.js':'application/javascript',
  '.json':'application/json', '.png':'image/png', '.xml':'application/xml'
};

http.createServer((req, res) => {
  const url  = req.url === '/' ? '/src/taskpane/taskpane.html' : req.url;
  const file = path.join(__dirname, url);
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('No encontrado: ' + url); return; }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(file)] || 'text/plain',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(data);
  });
}).listen(PORT, () => {
  console.log('Servidor corriendo en http://localhost:' + PORT);
  console.log('Mantén esta ventana abierta mientras uses el complemento en Excel.');
});
