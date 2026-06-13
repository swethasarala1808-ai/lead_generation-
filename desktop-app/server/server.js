const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = 3000;

// ── OSM proxy (bypasses browser CORS) ─────────────────────────────
function osmQuery(query) {
  return new Promise((resolve, reject) => {
    const endpoints = [
      { host: 'overpass-api.de',    path: '/api/interpreter' },
      { host: 'overpass.kumi.systems', path: '/api/interpreter' },
    ];

    function tryNext(i) {
      if (i >= endpoints.length) { reject(new Error('All OSM servers failed')); return; }
      const ep = endpoints[i];
      const body = 'data=' + encodeURIComponent(query);
      const options = {
        hostname: ep.host, path: ep.path, method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(body),
          'User-Agent': 'bizaxl-LeadGen/2.0 (sales@bizaxl.com)',
        },
        timeout: 40000,
      };
      const req = https.request(options, (res) => {
        let data = '';
        res.on('data', c => data += c);
        res.on('end', () => {
          if (res.statusCode === 200) resolve(data);
          else tryNext(i + 1);
        });
      });
      req.on('error', () => tryNext(i + 1));
      req.on('timeout', () => { req.destroy(); tryNext(i + 1); });
      req.write(body); req.end();
    }
    tryNext(0);
  });
}

// ── Nominatim geocode ──────────────────────────────────────────────
function geocode(address) {
  return new Promise((resolve) => {
    const q = encodeURIComponent(address + ', India');
    const options = {
      hostname: 'nominatim.openstreetmap.org',
      path: `/search?q=${q}&format=json&limit=1`,
      headers: { 'User-Agent': 'bizaxl-LeadGen/2.0 (sales@bizaxl.com)' },
      timeout: 12000,
    };
    const req = https.get(options, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const d = JSON.parse(data);
          if (d && d[0]) resolve({ lat: parseFloat(d[0].lat), lon: parseFloat(d[0].lon) });
          else resolve(null);
        } catch { resolve(null); }
      });
    });
    req.on('error', () => resolve(null));
    req.on('timeout', () => { req.destroy(); resolve(null); });
  });
}

// ── HTTP Server ───────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);

  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  // API: geocode
  if (parsed.pathname === '/api/geocode' && req.method === 'GET') {
    const result = await geocode(parsed.query.q || '');
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(result || { error: 'Not found' }));
    return;
  }

  // API: OSM search
  if (parsed.pathname === '/api/osm' && req.method === 'POST') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const { query } = JSON.parse(body);
        const result = await osmQuery(query);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(result);
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  // Serve index.html
  if (parsed.pathname === '/' || parsed.pathname === '/index.html') {
    const filePath = path.join(__dirname, 'index.html');
    fs.readFile(filePath, (err, data) => {
      if (err) { res.writeHead(404); res.end('Not found'); return; }
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(data);
    });
    return;
  }

  res.writeHead(404); res.end('Not found');
});

server.listen(PORT, '127.0.0.1', () => {
  console.log('\n================================================');
  console.log('  bizaxl LeadGen is RUNNING!');
  console.log('================================================');
  console.log(`\n  Open this in Chrome on Windows:`);
  console.log(`  → http://localhost:${PORT}\n`);
  console.log('  Press Ctrl+C to stop');
  console.log('================================================\n');

  // Auto-open browser on Windows from WSL
  try {
    const { exec } = require('child_process');
    exec(`cmd.exe /c start http://localhost:${PORT}`);
  } catch(e) {}
});
