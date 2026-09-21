# Realtime Audio Example

This TypeScript example connects to an OpenAI realtime model through the Vercel AI Gateway. It sends a text prompt over WebSocket, prints the streamed transcript, collects streamed PCM16 audio, and saves the response as `reply.wav`.

## Setup

Install dependencies:

```powershell
npm install
```

Create a local environment file from the template:

```powershell
Copy-Item .env.example .env
```

Set `AI_GATEWAY_API_KEY` in `.env`. The `.env` file is ignored by Git and must not be committed.

## Run

```powershell
npx tsx realtime.mts
```

The example requires a Vercel AI Gateway account with a valid payment method. A successful run prints the transcript and creates `reply.wav`.

## Typecheck

```powershell
npx tsc --noEmit
```

## Presentation

The code explanation is available in `realtime-code-explained.pptx`. Regenerate it with:

```powershell
node create-presentation.mjs
```

Yes, provided that:

The valid payment method is added to the correct Vercel team linked to your API key.
The new key is stored in .env:

.env is in vercel and not .env.example.
You run:

The code should then authenticate, connect to the realtime model, print the transcript, and create reply.wav.
