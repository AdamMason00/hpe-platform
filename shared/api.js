/* =====================================================================
 * HPE Service Platform — Backend API client
 * ---------------------------------------------------------------------
 * Talks to the Google Apps Script web app (see HPE_CONFIG.BACKEND.API_URL).
 *
 * GET  -> JSONP (script tag). Avoids CORS preflight entirely, which the
 *         Apps Script ContentService cannot satisfy for cross-origin
 *         fetch requests from GitHub Pages.
 * POST -> fetch() with Content-Type "text/plain" so it stays a "simple"
 *         request (no preflight). Apps Script reads e.postData.contents.
 *
 * Every call resolves to the parsed JSON the backend returns, or rejects
 * with an Error. All endpoints listed in the spec are wrapped below.
 * ===================================================================== */

(function (global) {
  'use strict';

  var CFG = global.HPE_CONFIG || {};
  var API_URL = (CFG.BACKEND && CFG.BACKEND.API_URL) || '';

  var _jsonpSeq = 0;

  /* ---- low level: JSONP GET --------------------------------------- */
  function jsonpGet(params) {
    return new Promise(function (resolve, reject) {
      if (!API_URL) { reject(new Error('API_URL not configured')); return; }
      _jsonpSeq += 1;
      var cb = '__hpe_jsonp_' + _jsonpSeq + '_' + (new Date().getTime());
      var timeoutId;
      var script = document.createElement('script');

      function cleanup() {
        if (timeoutId) clearTimeout(timeoutId);
        try { delete global[cb]; } catch (e) { global[cb] = undefined; }
        if (script && script.parentNode) script.parentNode.removeChild(script);
      }

      global[cb] = function (data) {
        cleanup();
        resolve(data);
      };

      var q = ['callback=' + encodeURIComponent(cb)];
      Object.keys(params || {}).forEach(function (k) {
        q.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
      });

      script.src = API_URL + '?' + q.join('&');
      script.onerror = function () {
        cleanup();
        reject(new Error('Network error calling backend (' + (params.action || '?') + ')'));
      };
      timeoutId = setTimeout(function () {
        cleanup();
        reject(new Error('Backend timeout (' + (params.action || '?') + ')'));
      }, 30000);

      document.body.appendChild(script);
    });
  }

  /* ---- low level: POST (text/plain, no preflight) ----------------- */
  function postAction(action, payload) {
    if (!API_URL) return Promise.reject(new Error('API_URL not configured'));
    var body = JSON.stringify({ action: action, payload: payload || {} });
    return fetch(API_URL, {
      method: 'POST',
      // "text/plain" keeps this a CORS-simple request (no OPTIONS preflight).
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: body
    }).then(function (res) {
      return res.text();
    }).then(function (txt) {
      var data;
      try { data = JSON.parse(txt); }
      catch (e) { throw new Error('Bad backend response: ' + txt.slice(0, 200)); }
      if (data && data.ok === false) {
        throw new Error(data.error || 'Backend returned an error');
      }
      return data;
    });
  }

  function get(action, extra) {
    var params = { action: action };
    if (extra) Object.keys(extra).forEach(function (k) { params[k] = extra[k]; });
    return jsonpGet(params).then(function (data) {
      if (data && data.ok === false) {
        throw new Error(data.error || 'Backend returned an error');
      }
      return data;
    });
  }

  /* ---- public API ------------------------------------------------- */
  var API = {
    raw: { get: get, post: postAction },

    // --- GET actions ---
    ping:           function ()       { return get('ping'); },
    getUser:        function ()       { return get('getUser'); },
    loadKPI:        function ()       { return get('loadKPI'); },
    loadWarranty:   function ()       { return get('loadWarranty'); },
    loadPOS:        function ()       { return get('loadPOS'); },
    loadEfficiency: function ()       { return get('loadEfficiency'); },
    loadExclusions: function ()       { return get('loadExclusions'); },
    loadStaff:      function ()       { return get('loadStaff'); },

    // --- POST actions ---
    saveKPI:        function (p)      { return postAction('saveKPI', p); },
    saveWarranty:   function (p)      { return postAction('saveWarranty', p); },
    savePOS:        function (p)      { return postAction('savePOS', p); },
    saveEfficiency: function (p)      { return postAction('saveEfficiency', p); },
    saveExclusions: function (p)      { return postAction('saveExclusions', p); },
    saveStaff:      function (p)      { return postAction('saveStaff', p); },
    savePINs:       function (p)      { return postAction('savePINs', p); },
    uploadFile:     function (p)      { return postAction('uploadFile', p); },
    backup:         function (p)      { return postAction('backup', p); },
    importCSV:      function (p)      { return postAction('importCSV', p); }
  };

  global.HPE_API = API;

})(typeof window !== 'undefined' ? window : this);
