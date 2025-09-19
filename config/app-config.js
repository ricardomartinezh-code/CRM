(function (global) {
  const source =
    global.__RELEAD_CONFIG__ && typeof global.__RELEAD_CONFIG__ === 'object'
      ? global.__RELEAD_CONFIG__
      : global.ReLeadConfig && typeof global.ReLeadConfig === 'object'
        ? global.ReLeadConfig
        : {};
  const config = {
    API_URL:
      source.API_URL ||
      'https://script.google.com/macros/s/AKfycbz_DsHnZdKlQ0a19SELkTqYTUUgXEnvKpn14nCjsGCt1fxTK0vcFxWzL9iraxSFNK3R/exec',
    REQUEST_TIMEOUT_MS: Number.isFinite(source.REQUEST_TIMEOUT_MS)
      ? source.REQUEST_TIMEOUT_MS
      : 45000,
    VERSION:
      typeof source.VERSION === 'string' && source.VERSION.trim() ? source.VERSION.trim() : '2.6.2',
  };
  if (typeof module === 'object' && module.exports) {
    module.exports = config;
  } else {
    global.ReLeadConfig = config;
  }
})(typeof self !== 'undefined' ? self : globalThis);
