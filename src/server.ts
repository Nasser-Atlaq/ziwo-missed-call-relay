import express from 'express';

const app = express();
app.use(express.json({ limit: '1mb' }));

const { TEAMS_WEBHOOK_URL, ZIWO_SHARED_SECRET, PORT = '3000' } = process.env;
if (!TEAMS_WEBHOOK_URL || !ZIWO_SHARED_SECRET) {
  throw new Error('Missing TEAMS_WEBHOOK_URL or ZIWO_SHARED_SECRET');
}

type ZiwoCallResult = 'Answered' | 'Busy' | 'Cancel' | 'Lose-race' | 'Failed' | 'Timeout';

interface ZiwoCallEnded {
  callID: string;
  callerIdNumber: string;
  callerIdName?: string;
  calleeIdNumber?: string;
  calleeIdName?: string;
  startedAt: string;
  callResult: ZiwoCallResult;
  data?: Record<string, unknown>;
  flags?: Record<string, unknown>;
}

app.post('/ziwo/call-ended', async (req, res) => {
  if (req.query.token !== ZIWO_SHARED_SECRET) {
    return res.status(401).send('unauthorized');
  }

  const event = req.body as ZiwoCallEnded;

  console.log(JSON.stringify({ at: new Date().toISOString(), callID: event.callID, callResult: event.callResult }));

  if (event.callResult !== 'Cancel') {
    return res.status(200).send('ignored');
  }

  const localTime = new Date(event.startedAt).toLocaleString('en-GB', {
    timeZone: 'Asia/Dubai',
    dateStyle: 'medium',
    timeStyle: 'short',
  });

  const card = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.4',
        body: [
          { type: 'TextBlock', text: 'Missed Call', weight: 'Bolder', size: 'Medium', color: 'Attention' },
          {
            type: 'FactSet',
            facts: [
              { title: 'Number', value: event.callerIdNumber },
              { title: 'Name', value: event.callerIdName || 'Unknown' },
              { title: 'Time', value: localTime },
              { title: 'Dialed', value: event.calleeIdNumber || '—' },
              { title: 'Call ID', value: event.callID },
            ],
          },
        ],
        actions: [
          { type: 'Action.OpenUrl', title: 'Call Back', url: `tel:${event.callerIdNumber}` },
        ],
      },
    }],
  };

  const resp = await fetch(TEAMS_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(card),
  });

  if (!resp.ok) {
    console.error('Teams post failed', resp.status, await resp.text());
    return res.status(502).send('teams_failed');
  }

  return res.status(200).send('posted');
});

app.get('/health', (_req, res) => res.status(200).send('ok'));

app.listen(Number(PORT), () => console.log(`listening on ${PORT}`));
