const https = require('https');

function extractResponseText(payload) {
  if (typeof payload.output_text === 'string') return payload.output_text;
  const chunks = [];
  for (const item of payload.output || []) {
    for (const content of item.content || []) {
      if (typeof content.text === 'string') chunks.push(content.text);
    }
  }
  return chunks.join('\n').trim();
}

function extractGeminiText(payload) {
  const parts = [];
  for (const candidate of payload.candidates || []) {
    for (const part of candidate.content?.parts || []) {
      if (typeof part.text === 'string') parts.push(part.text);
    }
  }
  return parts.join('\n').trim();
}

function postJson(url, headers, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const target = new URL(url);
    const req = https.request({
      method: 'POST',
      hostname: target.hostname,
      path: target.pathname + target.search,
      headers: {
        ...headers,
        'Content-Length': Buffer.byteLength(body)
      },
      timeout: 45000
    }, res => {
      let text = '';
      res.setEncoding('utf8');
      res.on('data', chunk => {
        text += chunk;
      });
      res.on('end', () => {
        let json = {};
        try {
          json = text ? JSON.parse(text) : {};
        } catch (error) {
          reject(new Error(`OpenAI trả về dữ liệu không phải JSON: ${error.message}`));
          return;
        }
        resolve({ status: res.statusCode || 0, ok: res.statusCode >= 200 && res.statusCode < 300, json });
      });
    });
    req.on('timeout', () => {
      req.destroy(new Error('Kết nối OpenAI quá thời gian chờ.'));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function main() {
  const inputText = await new Promise((resolve, reject) => {
    let text = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', chunk => {
      text += chunk;
    });
    process.stdin.on('end', () => resolve(text));
    process.stdin.on('error', reject);
  });
  const request = JSON.parse(inputText || '{}');
  let response;
  if (request.provider === 'gemini') {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) throw new Error('Chưa cấu hình GEMINI_API_KEY.');
    const prompt = (request.input || [])
      .flatMap(item => item.content || [])
      .map(item => item.text || '')
      .filter(Boolean)
      .join('\n\n');
    response = await postJson(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(request.model)}:generateContent`, {
      'x-goog-api-key': apiKey,
      'Content-Type': 'application/json'
    }, {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: request.max_output_tokens || 900 }
    });
    const payload = response.json || {};
    if (!response.ok) {
      const error = new Error(payload.error?.message || 'Không gọi được Gemini API.');
      error.status = response.status;
      throw error;
    }
    process.stdout.write(JSON.stringify({
      answer: extractGeminiText(payload) || 'AI không trả về nội dung.',
      model: request.model
    }));
    return;
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('Chưa cấu hình OPENAI_API_KEY.');
  response = await postJson('https://api.openai.com/v1/responses', {
    'Authorization': `Bearer ${apiKey}`,
    'Content-Type': 'application/json'
  }, request);
  const payload = response.json || {};
  if (!response.ok) {
    const error = new Error(payload.error?.message || 'Không gọi được OpenAI API.');
    error.status = response.status;
    throw error;
  }
  process.stdout.write(JSON.stringify({
    answer: extractResponseText(payload) || 'AI không trả về nội dung.',
    model: payload.model || request.model
  }));
}

main().catch(error => {
  console.error(error.message || String(error));
  process.exit(1);
});
