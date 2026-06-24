const http = require('http');
const port = process.env.PORT || 3000;
http.createServer(function (req, res) {
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end('OK\n');
}).listen(port, '0.0.0.0', function () {
  console.log('TEST: server listening on ' + port);
});
