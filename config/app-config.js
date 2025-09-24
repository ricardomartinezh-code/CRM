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
      'https://script.google.com/macros/s/AKfycbzg0XHsW_fYiVhEcmp0WB3zmqQ9hV-usT0EDsnhwBlD8CSat3Gc_-lDOhMyGIUMjbcq/exec',
    REQUEST_TIMEOUT_MS: Number.isFinite(source.REQUEST_TIMEOUT_MS)
      ? source.REQUEST_TIMEOUT_MS
      : 45000,
    VERSION:
      typeof source.VERSION === 'string' && source.VERSION.trim() ? source.VERSION.trim() : '2.8.0',
  };
  if (typeof module === 'object' && module.exports) {
    module.exports = config;
  } else {
    global.ReLeadConfig = config;
  }
})(typeof self !== 'undefined' ? self : globalThis);
