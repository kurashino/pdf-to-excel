const { TEMPLATE_B64 } = require('../../lib/template');

export default function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.json({ b64: TEMPLATE_B64 });
}
