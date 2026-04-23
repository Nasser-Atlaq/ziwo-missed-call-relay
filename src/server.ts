import express from 'express';

const app = express();
app.use(express.json({ limit: '1mb' }));

const {
  ENTRA_TENANT_ID,
  ENTRA_CLIENT_ID,
  ENTRA_CLIENT_SECRET,
  ONEDRIVE_USER,
  EXCEL_FILE_PATH,
  EXCEL_TABLE_NAME,
  ZIWO_SHARED_SECRET,
  PORT = '3000',
} = process.env;

const required = {
  ENTRA_TENANT_ID,
  ENTRA_CLIENT_ID,
  ENTRA_CLIENT_SECRET,
  ONEDRIVE_USER,
  EXCEL_FILE_PATH,
  EXCEL_TABLE_NAME,
  ZIWO_SHARED_SECRET,
};
for (const [k, v] of Object.entries(required)) {
  if (!v) throw new Error(`Missing env var: ${k}`);
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

let cachedToken: { token: string; expiresAt: number } | null = null;

async function getGraphToken(): Promise<string> {
  if (cachedToken && Date.now() < cachedToken.expiresAt - 60_000) {
    return cachedToken.token;
  }

  const resp = await fetch(
    `https://login.microsoftonline.com/${ENTRA_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: ENTRA_CLIENT_ID!,
        client_secret: ENTRA_CLIENT_SECRET!,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    },
  );

  if (!resp.ok) {
    throw new Error(`Entra token fetch failed: ${resp.status} ${await resp.text()}`);
  }

  const data = (await resp.json()) as { access_token: string; expires_in: number };
  cachedToken = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
  return cachedToken.token;
}

function encodeExcelPath(path: string): string {
  const trimmed = path.startsWith('/') ? path.slice(1) : path;
  return trimmed.split('/').map(encodeURIComponent).join('/');
}

async function appendMissedCallToExcel(event: ZiwoCallEnded): Promise<void> {
  const token = await getGraphToken();

  const localTime = new Date(event.startedAt).toLocaleString('en-GB', {
    timeZone: 'Asia/Dubai',
    dateStyle: 'short',
    timeStyle: 'medium',
  });

  const row = [
    localTime,
    event.callerIdNumber,
    event.callerIdName || 'Unknown',
    event.calleeIdNumber || '',
    event.callResult,
    event.callID,
    '',
    '',
  ];

  const encodedPath = encodeExcelPath(EXCEL_FILE_PATH!);
  const url =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(ONEDRIVE_USER!)}` +
    `/drive/root:/${encodedPath}:/workbook/tables/${encodeURIComponent(EXCEL_TABLE_NAME!)}/rows/add`;

  const resp = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ values: [row] }),
  });

  if (!resp.ok) {
    throw new Error(`Excel append failed: ${resp.status} ${await resp.text()}`);
  }
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

  try {
    await appendMissedCallToExcel(event);
    return res.status(200).send('appended');
  } catch (e) {
    console.error('Excel append error', e);
    return res.status(502).send('excel_error');
  }
});

app.get('/health', (_req, res) => res.status(200).send('ok'));

app.listen(Number(PORT), () => console.log(`listening on ${PORT}`));
