const express = require('express');
const http = require('http');
const app = express();

function proxyTo(service, path) {
    return (req, res) => {
        // Attention: ici tes services user et product écoutent sur le port 3000
        http.get(`http://${service}:3000${path}`, (proxyRes) => {
            let data = '';
            proxyRes.on('data', c => data += c);
            proxyRes.on('end', () => res.json(JSON.parse(data)));
        }).on('error', () => res.status(502).json({ error: `${service} unavailable` }));
    };
}

app.get('/users', proxyTo('user-service', '/users'));
app.get('/products', proxyTo('product-service', '/products'));
app.get('/health', (req, res) => res.json({ service: 'gateway', status: 'up' }));

// 👇 AJOUTE '0.0.0.0' pour écouter sur toutes les interfaces du conteneur
app.listen(5001, '0.0.0.0', () => console.log('Gateway on :5001'));