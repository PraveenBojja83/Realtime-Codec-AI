import pptxgen from 'pptxgenjs';

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI Realtime example';
pptx.subject = 'How realtime.mts works';
pptx.title = 'Realtime Audio with AI SDK Gateway';
pptx.company = 'Vercel';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US',
};

const C = {
  ink: '17202A',
  muted: '52606D',
  blue: '146C94',
  cyan: '19A7CE',
  pale: 'EAF6FA',
  orange: 'E76F51',
  paleOrange: 'FFF0EA',
  green: '2A9D8F',
  paleGreen: 'E9F7F4',
  white: 'FFFFFF',
  line: 'D9E2EC',
};

function addBase(slide, title, kicker = '') {
  slide.background = { color: C.white };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.16, fill: { color: C.cyan }, line: { color: C.cyan } });
  if (kicker) slide.addText(kicker.toUpperCase(), { x: 0.65, y: 0.45, w: 5, h: 0.25, fontFace: 'Aptos', fontSize: 10, bold: true, color: C.cyan, charSpacing: 1.2, margin: 0 });
  slide.addText(title, { x: 0.65, y: kicker ? 0.76 : 0.55, w: 11.8, h: 0.55, fontFace: 'Aptos Display', fontSize: 26, bold: true, color: C.ink, margin: 0 });
  slide.addText('realtime.mts  |  AI SDK Gateway + WebSocket', { x: 0.65, y: 7.08, w: 6, h: 0.18, fontSize: 8, color: C.muted, margin: 0 });
}

function addBulletList(slide, items, x, y, w, h, color = C.ink, size = 18) {
  slide.addText(items.map((text) => ({ text, options: { bullet: { indent: size }, hanging: size / 2 } })), {
    x, y, w, h, fontSize: size, color, breakLine: true, paraSpaceAfterPt: 12, margin: 0.08,
  });
}

function addCode(slide, code, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.04, fill: { color: '14202B' }, line: { color: '14202B' } });
  slide.addText(code, { x: x + 0.2, y: y + 0.18, w: w - 0.4, h: h - 0.35, fontFace: 'Cascadia Mono', fontSize: 12, color: 'D8F3FA', breakLine: false, margin: 0, fit: 'shrink' });
}

function addBox(slide, text, x, y, w, h, fill = C.pale, line = C.cyan, color = C.ink) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.04, fill: { color: fill }, line: { color: line, width: 1.2 } });
  slide.addText(text, { x: x + 0.12, y: y + 0.12, w: w - 0.24, h: h - 0.24, align: 'center', valign: 'mid', fontSize: 15, bold: true, color, margin: 0.04, fit: 'shrink' });
}

{
  const slide = pptx.addSlide();
  slide.background = { color: C.ink };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.18, fill: { color: C.cyan }, line: { color: C.cyan } });
  slide.addText('REALTIME AUDIO', { x: 0.8, y: 1.2, w: 4, h: 0.3, fontSize: 12, bold: true, color: C.cyan, charSpacing: 2, margin: 0 });
  slide.addText('How realtime.mts works', { x: 0.8, y: 1.65, w: 8.8, h: 0.8, fontFace: 'Aptos Display', fontSize: 34, bold: true, color: C.white, margin: 0 });
  slide.addText('A small TypeScript example that sends text to a realtime model, receives streamed audio, and saves a playable WAV file.', { x: 0.83, y: 2.75, w: 7.6, h: 0.8, fontSize: 19, color: 'D8E7EF', breakLine: false, margin: 0 });
  slide.addShape(pptx.ShapeType.arc, { x: 9.2, y: 1.1, w: 2.8, h: 2.8, adjustPoint: 0.3, line: { color: C.cyan, width: 3, transparency: 15 }, fill: { color: C.ink, transparency: 100 } });
  slide.addShape(pptx.ShapeType.arc, { x: 9.8, y: 1.7, w: 1.6, h: 1.6, adjustPoint: 0.3, line: { color: C.orange, width: 3 }, fill: { color: C.ink, transparency: 100 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: 10.35, y: 2.25, w: 0.5, h: 0.5, fill: { color: C.cyan }, line: { color: C.cyan } });
  slide.addText('Purpose, flow, and runtime behavior', { x: 0.83, y: 6.35, w: 6, h: 0.3, fontSize: 12, color: 'A9C6D3', margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'What is the purpose of this code?', '01  |  purpose');
  addBulletList(slide, [
    'Connect to a realtime AI voice model through the Vercel AI Gateway.',
    'Send a text prompt over a persistent WebSocket connection.',
    'Receive the answer as two streams: transcript text and audio bytes.',
    'Convert the raw PCM16 audio into a standard reply.wav file.',
  ], 0.85, 1.7, 7.1, 4.5, C.ink, 18);
  addBox(slide, 'Input\ntext prompt', 8.5, 1.85, 1.65, 1.05, C.pale, C.cyan);
  slide.addText('→', { x: 10.2, y: 2.1, w: 0.5, h: 0.4, fontSize: 24, color: C.cyan, bold: true, margin: 0 });
  addBox(slide, 'Realtime\nmodel', 10.75, 1.85, 1.65, 1.05, C.paleOrange, C.orange);
  slide.addText('↓', { x: 11.35, y: 3.05, w: 0.5, h: 0.4, fontSize: 24, color: C.orange, bold: true, margin: 0 });
  addBox(slide, 'reply.wav', 10.75, 3.75, 1.65, 1.05, C.paleGreen, C.green);
  slide.addText('The program is a command-line demo, not a web UI.', { x: 8.55, y: 5.45, w: 3.8, h: 0.7, fontSize: 15, color: C.muted, italic: true, margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'End-to-end architecture', '02  |  system flow');
  addBox(slide, 'realtime.mts', 0.8, 2.2, 2.05, 1.2, C.pale, C.cyan);
  addBox(slide, 'AI SDK\nGateway', 3.55, 2.2, 2.05, 1.2, C.paleOrange, C.orange);
  addBox(slide, 'Realtime\nprovider', 6.3, 2.2, 2.05, 1.2, C.paleGreen, C.green);
  addBox(slide, 'reply.wav', 9.05, 2.2, 2.05, 1.2, 'F2F0FA', '8874C7');
  slide.addText('API key\n(token request)', { x: 2.85, y: 1.45, w: 1.1, h: 0.55, fontSize: 11, color: C.muted, align: 'center', margin: 0 });
  slide.addText('WebSocket\n(JSON events)', { x: 5.55, y: 1.45, w: 1.25, h: 0.55, fontSize: 11, color: C.muted, align: 'center', margin: 0 });
  slide.addText('PCM16\nbase64 audio', { x: 8.25, y: 1.45, w: 1.25, h: 0.55, fontSize: 11, color: C.muted, align: 'center', margin: 0 });
  for (const x of [2.9, 5.65, 8.4]) slide.addText('→', { x, y: 2.58, w: 0.45, h: 0.35, fontSize: 24, bold: true, color: C.cyan, margin: 0 });
  slide.addText('The Gateway handles provider authentication and translates normalized AI SDK events into the provider wire format.', { x: 1.2, y: 4.45, w: 10.9, h: 0.65, fontSize: 19, color: C.ink, align: 'center', margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Startup and configuration', '03  |  dependencies');
  addCode(slide, "import 'dotenv/config';\nimport { gateway } from '@ai-sdk/gateway';\nimport WebSocket from 'ws';\nimport { writeFileSync } from 'node:fs';", 0.8, 1.55, 6.2, 2.1);
  addBulletList(slide, [
    'dotenv loads local environment variables from .env.',
    '@ai-sdk/gateway provides the realtime model adapter.',
    'ws opens the Node.js WebSocket connection.',
    'node:fs writes the final WAV file.',
  ], 7.45, 1.6, 4.9, 3.2, C.ink, 16);
  slide.addText('The API key stays server-side. It is never placed in the source code or sent to a browser.', { x: 0.85, y: 5.1, w: 11.5, h: 0.55, fontSize: 18, bold: true, color: C.blue, margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Step 1: obtain a temporary realtime token', '04  |  authentication');
  addCode(slide, "const { token, url } =\n  await gateway.experimental_realtime.getToken({\n    model: modelId,\n  });", 0.85, 1.55, 5.5, 2.1);
  addBulletList(slide, [
    'The server uses AI_GATEWAY_API_KEY through the environment.',
    'The Gateway validates the account and selected model.',
    'The response contains a short-lived token and WebSocket URL.',
    'The long-lived API key is not used as the browser/session token.',
  ], 7.0, 1.55, 5.2, 3.7, C.ink, 16);
  slide.addShape(pptx.ShapeType.line, { x: 3.6, y: 4.15, w: 0, h: 1, line: { color: C.cyan, width: 2, beginArrowType: 'none', endArrowType: 'triangle' } });
  slide.addText('Security boundary', { x: 2.5, y: 5.25, w: 2.3, h: 0.3, fontSize: 14, bold: true, color: C.orange, align: 'center', margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Step 2: open the WebSocket and send a prompt', '05  |  client events');
  addCode(slide, "const ws = new WebSocket(config.url, config.protocols);\n\nawait send({\n  type: 'conversation-item-create',\n  item: { type: 'text-message', role: 'user',\n    text: 'Say hello in one sentence.' },\n});\nawait send({ type: 'response-create' });", 0.8, 1.45, 7.15, 3.85);
  addBox(slide, 'open', 8.65, 1.7, 1.25, 0.8, C.paleGreen, C.green);
  slide.addText('1', { x: 9.15, y: 2.75, w: 0.25, h: 0.3, fontSize: 17, bold: true, color: C.cyan, margin: 0 });
  addBox(slide, 'conversation-item-create', 9.55, 2.55, 2.35, 0.8, C.pale, C.cyan, C.ink);
  slide.addText('2', { x: 9.15, y: 3.85, w: 0.25, h: 0.3, fontSize: 17, bold: true, color: C.cyan, margin: 0 });
  addBox(slide, 'response-create', 9.55, 3.65, 2.35, 0.8, C.paleOrange, C.orange, C.ink);
  slide.addText('Events are serialized as JSON strings over the socket.', { x: 8.55, y: 5.3, w: 3.7, h: 0.65, fontSize: 15, color: C.muted, margin: 0, align: 'center' });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Step 3: process streamed server events', '06  |  response handling');
  addBox(slide, 'audio-transcript-delta', 0.85, 1.7, 3.15, 0.85, C.pale, C.cyan);
  slide.addText('Print text immediately', { x: 4.35, y: 1.93, w: 2.45, h: 0.3, fontSize: 16, color: C.ink, margin: 0 });
  addBox(slide, 'audio-delta', 0.85, 2.95, 3.15, 0.85, C.paleOrange, C.orange);
  slide.addText('Decode base64 and buffer PCM16', { x: 4.35, y: 3.18, w: 3.1, h: 0.3, fontSize: 16, color: C.ink, margin: 0 });
  addBox(slide, 'response-done', 0.85, 4.2, 3.15, 0.85, C.paleGreen, C.green);
  slide.addText('Build WAV, save, close socket', { x: 4.35, y: 4.43, w: 3.1, h: 0.3, fontSize: 16, color: C.ink, margin: 0 });
  addBox(slide, 'error', 8.7, 2.75, 2.15, 0.85, C.paleOrange, C.orange);
  slide.addText('Log message and close', { x: 8.15, y: 4.05, w: 3.3, h: 0.4, fontSize: 16, color: C.ink, align: 'center', margin: 0 });
  slide.addText('The event loop turns a realtime stream into visible text and a complete audio file.', { x: 0.9, y: 5.85, w: 10.6, h: 0.4, fontSize: 18, bold: true, color: C.blue, margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Step 4: wrap raw PCM audio as WAV', '07  |  file output');
  addCode(slide, "function toWav(pcm: Buffer, sampleRate: number): Buffer {\n  const header = Buffer.alloc(44);\n  header.write('RIFF', 0);\n  ...\n  header.writeUInt32LE(pcm.length, 40);\n  return Buffer.concat([header, pcm]);\n}", 0.8, 1.5, 6.5, 2.8);
  addBulletList(slide, [
    'Audio format: PCM16, mono.',
    'Sample rate: 24,000 Hz.',
    'The 44-byte header describes the audio data.',
    'The result is written to reply.wav.',
  ], 7.8, 1.65, 4.3, 2.8, C.ink, 17);
  slide.addText('WAV = header metadata + raw PCM samples', { x: 2.1, y: 5.45, w: 8.9, h: 0.45, fontSize: 22, bold: true, color: C.green, align: 'center', margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Current runtime result and limitation', '08  |  troubleshooting');
  addBox(slide, 'Code starts', 0.9, 1.75, 2.2, 0.95, C.paleGreen, C.green);
  slide.addText('→', { x: 3.25, y: 2.02, w: 0.45, h: 0.3, fontSize: 22, color: C.cyan, margin: 0 });
  addBox(slide, 'Token request', 3.8, 1.75, 2.2, 0.95, C.pale, C.cyan);
  slide.addText('→', { x: 6.15, y: 2.02, w: 0.45, h: 0.3, fontSize: 22, color: C.cyan, margin: 0 });
  addBox(slide, '403 Gateway', 6.7, 1.75, 2.2, 0.95, C.paleOrange, C.orange);
  slide.addText('Credit card required', { x: 9.35, y: 2.03, w: 2.6, h: 0.3, fontSize: 17, bold: true, color: C.orange, margin: 0 });
  addBulletList(slide, [
    'The TypeScript setup is valid and the .env loader runs.',
    'The request reaches Vercel AI Gateway successfully.',
    'Gateway blocks token creation until the Vercel team has a valid card on file.',
    'This is an account/billing restriction, not a WebSocket or WAV bug.',
  ], 1.0, 3.65, 10.8, 2.1, C.ink, 16);
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'How to run it', '09  |  local usage');
  addCode(slide, "# .env\nAI_GATEWAY_API_KEY=your-key\n\n# install dependencies\nnpm install\n\n# run\nnpx tsx realtime.mts", 0.85, 1.45, 5.25, 3.35);
  addBulletList(slide, [
    'Keep the real key in .env, never in .env.example or source code.',
    'The .env file is excluded by .gitignore.',
    'Use a Vercel environment variable for deployed applications.',
    'After a successful response, listen to reply.wav.',
  ], 6.8, 1.6, 5.2, 3.4, C.ink, 16);
  slide.addText('Expected success output: streamed transcript + “Saved reply.wav”', { x: 1.1, y: 5.65, w: 10.8, h: 0.4, fontSize: 19, bold: true, color: C.green, align: 'center', margin: 0 });
}

{
  const slide = pptx.addSlide();
  addBase(slide, 'Key takeaways', '10  |  summary');
  addBulletList(slide, [
    'realtime.mts is a server-side realtime audio demo.',
    'The Gateway mints a temporary token, then the WebSocket carries JSON events.',
    'Transcript deltas are printed while audio deltas are collected.',
    'A minimal WAV header makes the raw PCM audio playable.',
    'The current blocker is Vercel Gateway account verification for billing.',
  ], 1.0, 1.65, 11.0, 3.8, C.ink, 20);
  slide.addText('One small file demonstrates authentication, streaming, event handling, and audio file encoding.', { x: 1.0, y: 5.75, w: 11.1, h: 0.5, fontSize: 19, bold: true, color: C.blue, align: 'center', margin: 0 });
}

pptx.writeFile({ fileName: 'realtime-code-explained.pptx' });