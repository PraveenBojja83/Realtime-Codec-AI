import 'dotenv/config';
import { gateway } from '@ai-sdk/gateway';
import WebSocket from 'ws';
import { writeFileSync } from 'node:fs';

const modelId = 'openai/gpt-realtime-2';

async function main() {
  // getToken runs on the server, where your API key lives. It returns a token
  // and the WebSocket URL to connect with.
  const { token, url } = await gateway.experimental_realtime.getToken({
    model: modelId,
  });

  // The realtime model is a codec: it builds the WebSocket config and
  // translates between normalized AI SDK events and the provider wire format.
  const model = gateway.experimental_realtime(modelId);
  const getWebSocketConfig = model.getWebSocketConfig;
  if (!getWebSocketConfig) {
    throw new Error('The selected realtime model does not support WebSockets.');
  }
  const config = getWebSocketConfig({ token, url });

  const ws = new WebSocket(config.url, config.protocols);
  const audioChunks: Buffer[] = [];

  const send = async (
    event: Parameters<typeof model.serializeClientEvent>[0],
  ) => ws.send(JSON.stringify(await model.serializeClientEvent(event)));

  ws.on('open', async () => {
    await send({
      type: 'conversation-item-create',
      item: { type: 'text-message', role: 'user', text: 'Say hello in one sentence.' },
    });
    await send({ type: 'response-create' });
  });

  ws.on('message', (data) => {
    const parsed = model.parseServerEvent(JSON.parse(data.toString()));

    for (const event of Array.isArray(parsed) ? parsed : [parsed]) {
      switch (event.type) {
        case 'audio-transcript-delta':
          process.stdout.write(event.delta);
          break;
        case 'audio-delta':
          audioChunks.push(Buffer.from(event.delta, 'base64'));
          break;
        case 'response-done':
          writeFileSync('reply.wav', toWav(Buffer.concat(audioChunks), 24000));
          console.log('\nSaved reply.wav');
          ws.close();
          break;
        case 'error':
          console.error(event.message);
          ws.close();
          break;
      }
    }
  });
}

main().catch(console.error);

// Wrap raw PCM16 mono audio in a minimal WAV header so the file is playable
function toWav(pcm: Buffer, sampleRate: number): Buffer {
  const header = Buffer.alloc(44);
  header.write('RIFF', 0);
  header.writeUInt32LE(36 + pcm.length, 4);
  header.write('WAVE', 8);
  header.write('fmt ', 12);
  header.writeUInt32LE(16, 16);
  header.writeUInt16LE(1, 20);
  header.writeUInt16LE(1, 22);
  header.writeUInt32LE(sampleRate, 24);
  header.writeUInt32LE(sampleRate * 2, 28);
  header.writeUInt16LE(2, 32);
  header.writeUInt16LE(16, 34);
  header.write('data', 36);
  header.writeUInt32LE(pcm.length, 40);
  return Buffer.concat([header, pcm]);
}