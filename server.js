'use strict';

/* eslint-env node */

const path = require('node:path');
const express = require('express');
const appConfig = require('./config/app-config.js');

const DEFAULT_APPS_SCRIPT_URL = appConfig.API_URL;
const DEFAULT_TIMEOUT_MS = 10000;
const DEFAULT_ACTION = 'metaWebhook';

function buildForwardUrl(baseUrl, originalUrl, defaultAction = DEFAULT_ACTION){
  const source = typeof originalUrl === 'string' ? originalUrl : '';
  const upstream = new URL(baseUrl);
  const queryStart = source.indexOf('?');
  if(queryStart !== -1){
    const queryString = source.slice(queryStart + 1);
    const params = new URLSearchParams(queryString);
    for(const [key, value] of params.entries()){
      upstream.searchParams.append(key, value);
    }
  }
  if(defaultAction && !upstream.searchParams.has('action')){
    upstream.searchParams.set('action', defaultAction);
  }
  return upstream;
}

function pickForwardHeaders(headers = {}){
  const normalized = {};
  const pick = name => {
    const target = String(name || '').toLowerCase();
    const keys = Object.keys(headers);
    for(let i = 0; i < keys.length; i += 1){
      const key = keys[i];
      if(!key) continue;
      if(String(key).toLowerCase() === target){
        const value = headers[key];
        return Array.isArray(value) ? value[0] : value;
      }
    }
    return undefined;
  };
  const contentType = pick('content-type');
  if(contentType){
    normalized['Content-Type'] = contentType;
  }
  const sig256 = pick('x-hub-signature-256');
  if(sig256){
    normalized['X-Hub-Signature-256'] = sig256;
  }
  const sigLegacy = pick('x-hub-signature');
  if(sigLegacy){
    normalized['X-Hub-Signature'] = sigLegacy;
  }
  return normalized;
}

function selectForwardBody(req){
  if(req.rawBody){
    return req.rawBody;
  }
  if(req.body === undefined || req.body === null) return undefined;
  if(typeof req.body === 'string') return req.body;
  try{
    return Buffer.from(JSON.stringify(req.body));
  }catch(err){
    return undefined;
  }
}

function createServer(options = {}){
  const config = {
    appsScriptUrl: options.appsScriptUrl || process.env.APPS_SCRIPT_URL || DEFAULT_APPS_SCRIPT_URL,
    forwardTimeoutMs: Number.isFinite(options.forwardTimeoutMs)
      ? options.forwardTimeoutMs
      : Number.isFinite(Number(process.env.APPS_SCRIPT_TIMEOUT_MS))
        ? Number(process.env.APPS_SCRIPT_TIMEOUT_MS)
        : DEFAULT_TIMEOUT_MS,
    upstreamAction: options.upstreamAction || DEFAULT_ACTION,
    logRequests: options.logRequests !== undefined
      ? options.logRequests
      : process.env.LOG_WEBHOOK_REQUESTS === 'true',
    staticDir: options.staticDir || __dirname
  };

  const app = express();
  app.disable('x-powered-by');
  app.set('trust proxy', true);

  const rawBodySaver = (req, _res, buf) => {
    if(buf && buf.length){
      req.rawBody = Buffer.from(buf);
    }
  };

  app.use(express.json({ limit: '10mb', verify: rawBodySaver }));
  app.use(express.urlencoded({ limit: '10mb', extended: true, verify: rawBodySaver }));

  async function forwardToAppsScript(req, res){
    try{
      if(config.logRequests){
        console.log(`[webhook] ${req.method} ${req.originalUrl}`);
      }
      const upstreamUrl = buildForwardUrl(config.appsScriptUrl, req.originalUrl, config.upstreamAction);
      const headers = pickForwardHeaders(req.headers);
      const method = req.method.toUpperCase();
      let body;
      if(method !== 'GET' && method !== 'HEAD' && method !== 'OPTIONS'){
        body = selectForwardBody(req);
        if(body && !headers['Content-Type']){
          headers['Content-Type'] = 'application/json';
        }
      }
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), config.forwardTimeoutMs);
      const response = await fetch(upstreamUrl, {
        method,
        headers,
        body,
        signal: controller.signal
      });
      clearTimeout(timeout);
      const text = await response.text();
      const contentType = response.headers.get('content-type');
      if(contentType) res.set('Content-Type', contentType);
      const cacheControl = response.headers.get('cache-control');
      if(cacheControl) res.set('Cache-Control', cacheControl);
      res.status(response.status).send(text);
    }catch(err){
      if(config.logRequests){
        console.error('[webhook] Error reenviando a Apps Script', err);
      }
      const timeoutError = err && err.name === 'AbortError';
      const status = timeoutError ? 504 : 502;
      res
        .status(status)
        .json({
          ok: false,
          error: 'No se pudo reenviar el webhook a Apps Script.',
          details: timeoutError ? 'Tiempo de espera agotado.' : (err && err.message) || 'Error desconocido.'
        });
    }
  }

  app.get('/webhook', forwardToAppsScript);
  app.post('/webhook', forwardToAppsScript);
  app.all('/webhook', forwardToAppsScript);

  const staticOptions = { fallthrough: true, index: false, extensions: ['html'] };
  app.use(express.static(config.staticDir, staticOptions));

  app.get('*', (req, res, next) => {
    if(req.path.startsWith('/webhook')){
      next();
      return;
    }
    res.sendFile(path.join(config.staticDir, 'index.html'));
  });

  return app;
}

if(require.main === module){
  const port = Number(process.env.PORT) || 3000;
  const server = createServer();
  server.listen(port, () => {
    console.log(`ReLead server escuchando en puerto ${port}`);
  });
}

module.exports = {
  createServer,
  buildForwardUrl,
  pickForwardHeaders,
  selectForwardBody,
  DEFAULT_APPS_SCRIPT_URL,
  DEFAULT_TIMEOUT_MS,
  DEFAULT_ACTION
};
