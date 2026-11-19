/* TQ Excel Connector - setup page: overview, NetSuite connection, licence, logo, seats, updates and the AI add-on,
 * one section at a time. Secrets are sent to the local server only, which writes them to its own config file;
 * they are never read back. */
(function () {
  'use strict';

  try {   // the panel's theme (it knows Excel's); outside Excel the page follows Windows
    var theme = localStorage.getItem('tq.theme');
    if (theme) document.documentElement.setAttribute('data-theme', theme);
  } catch (e) { /* private storage */ }

  var $ = function (id) { return document.getElementById(id); };
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
    'October', 'November', 'December'];

  function box(id, text, kind) {
    var el = $(id);
    el.hidden = false;
    el.textContent = text;
    el.className = 'box' + (kind ? ' ' + kind : '');
  }

  function showLicence(licence) {
    if (!licence) box('licenceState', 'No licence checked yet: it is checked once the NetSuite account is saved.');
    else box('licenceState', licence.message, licence.ok ? 'ok' : 'bad');
  }

  /* One section at a time; Save shows where there is something to save. */
  var SAVED = ['connection', 'licence', 'updates', 'ai'];
  function section(name) {
    Array.prototype.forEach.call(document.querySelectorAll('main section[data-section]'), function (s) {
      s.hidden = s.getAttribute('data-section') !== name;
    });
    Array.prototype.forEach.call(document.querySelectorAll('#snav button'), function (b) {
      b.classList.toggle('on', b.getAttribute('data-section') === name);
    });
    var on = document.querySelector('#snav button.on');
    if (on) on.parentNode.scrollLeft = Math.max(0, on.offsetLeft - 40);   // the pills row, when the page is narrow
    $('savebar').hidden = SAVED.indexOf(name) === -1;
    if (name === 'seats') $('seatRefresh').click();
    window.scrollTo(0, 0);
  }
  Array.prototype.forEach.call(document.querySelectorAll('#snav button'), function (b) {
    b.addEventListener('click', function () { section(b.getAttribute('data-section')); });
  });

  function stats(data) {
    var on = data.mode === 'netsuite';
    var s = data.signIn || {};
    $('statDot').className = 'dot ' + (on ? 'ok' : '');
    $('statNetsuite').textContent = on ? 'Connected' : 'Not connected';
    $('statAccount').textContent = on ? data.accountId : 'sample data';
    $('statUser').textContent = s.signedIn ? (s.user || 'Yes') : s.mode === 'oauth2' ? 'Not signed in' : (on ? 'Token' : '—');
    $('statRole').textContent = s.signedIn ? (s.role || '') : '';
    var l = data.licence;
    $('statLicence').textContent = l ? (l.ok ? (l.users ? l.users + ' users' : 'Valid') : 'Not valid') : 'None';
    $('statLicenceNote').textContent = l && l.ok && l.expires ? 'until ' + l.expires : l && !l.ok ? 'see Licence' : '';
    $('statVersion').textContent = data.version || '';
    $('statUpdate').textContent = data.updateUrl ? 'updates on' : 'updates by hand';
    $('navConnection').textContent = on ? '✓' : '';
    $('navLicence').textContent = l ? (l.ok ? '✓' : '!') : '';
    $('navLicence').className = l && !l.ok ? 'bad' : '';
  }

  function load() {
    fetch('/api/setup').then(function (r) { return r.json(); }).then(function (data) {
      stats(data);
      if (data.certExpires) {
        var soon = data.certDaysLeft !== null && data.certDaysLeft <= 60;
        box('certState', 'This PC’s certificate for https://localhost is valid until ' + data.certExpires +
          (soon ? '. It is renewed the next time the connector starts (Windows asks you to confirm).' : '. It renews itself 60 days before that.'), soon ? 'bad' : '');
      }
      $('advanced').open = !data.restletUrl || !data.clientId;
      ['accountId', 'restletUrl', 'consumerKey', 'tokenId'].forEach(function (k) { $(k).value = data[k] || ''; });
      $('fiscal').value = String(data.fiscalYearStartMonth || 1);
      $('clientId').value = data.clientId || '';
      $('logoState').textContent = data.logo === 'custom' ? 'Using the uploaded logo.' : 'Using the logo from NetSuite (if one is set there).';
      $(data.authMode === 'tba' && data.consumerKey ? 'modeTba' : 'modeOauth').checked = true;
      showMode();
      if (data.redirectUri) $('redirectUri').textContent = data.redirectUri;
      var s = data.signIn;
      $('signInState').hidden = !s;
      $('signOut').hidden = !(s && s.signedIn);
      if (s) box('signInState', s.signedIn ? 'Signed in to NetSuite as ' + (s.user || 'you') + (s.role ? ' (' + s.role + ')' : '') +
        ' since ' + s.since + '.' : 'Not signed in yet: press Sign in with NetSuite.', s.signedIn ? 'ok' : '');
      $('updateUrl').value = data.updateUrl || '';
      $('secretsNote').hidden = !data.hasSecrets;
      $('folder').textContent = data.dataFolder || 'the data folder';
      box('state', data.mode === 'netsuite'
        ? 'Connected to NetSuite account ' + data.accountId + '. Version ' + data.version + '.'
        : 'Not connected to NetSuite yet: the panel shows sample data. Version ' + data.version + '.',
        data.mode === 'netsuite' ? 'ok' : '');
      $('aiOff').checked = false;
      $('aiKey').placeholder = data.aiOn ? 'A key is saved (the add-in is on)' : 'Not set (the add-in is off)';
      showLicence(data.licence);
    }).catch(function (error) { box('state', 'The connector is not answering: ' + error.message, 'bad'); });
  }

  function save() {
    var body = {
      accountId: $('accountId').value, restletUrl: $('restletUrl').value, consumerKey: $('consumerKey').value,
      consumerSecret: $('consumerSecret').value, tokenId: $('tokenId').value, tokenSecret: $('tokenSecret').value,
      fiscalYearStartMonth: Number($('fiscal').value), licence: $('licence').value, updateUrl: $('updateUrl').value,
      authMode: $('modeOauth').checked ? 'oauth2' : 'tba', clientId: $('clientId').value,
      aiKey: $('aiKey').value, aiOff: $('aiOff').checked
    };
    box('result', 'Saving and testing…');
    fetch('/api/setup', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) { box('result', data.error, 'bad'); return; }
        ['consumerSecret', 'tokenSecret', 'aiKey', 'licence'].forEach(function (k) { $(k).value = ''; });
        if (data.signIn) box('result', data.error || 'Saved. Now sign in with NetSuite.', '');
        else if (!data.connected) box('result', 'Saved, but NetSuite did not answer: ' + data.error, 'bad');
        else if (data.licence && !data.licence.ok) box('result', 'Connected as ' + data.user + ', but: ' + data.licence.message, 'bad');
        else box('result', 'Saved. Connected to NetSuite as ' + data.user + '. Go back to the panel and use Refresh all.', 'ok');
        load();
      })
      .catch(function (error) { box('result', error.message, 'bad'); });
  }

  MONTHS.forEach(function (name, n) {
    var o = document.createElement('option');
    o.value = String(n + 1);
    o.textContent = name;
    $('fiscal').appendChild(o);
  });
  $('save').addEventListener('click', save);

  /* Logo: an uploaded picture replaces the NetSuite logo; "Use the NetSuite logo" removes it. */
  function sendLogo(body) {
    fetch('/api/logo', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        $('logoState').textContent = data.ok ? (data.logo === 'custom' ? 'Saved: the uploaded logo is used.' : 'The NetSuite logo is used.') : data.error;
        $('logoPreview').src = '/logo?t=' + Date.now();
      });
  }
  $('logoFile').addEventListener('change', function () {
    var file = $('logoFile').files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function () { sendLogo({ op: 'save', data: String(reader.result).split(',')[1] }); };
    reader.readAsDataURL(file);
  });
  $('logoRemove').addEventListener('click', function () { sendLogo({ op: 'remove' }); });
  $('logoPreview').addEventListener('error', function () { $('logoPreview').hidden = true; $('logoNone').hidden = false; });
  $('logoPreview').addEventListener('load', function () { $('logoPreview').hidden = false; $('logoNone').hidden = true; });

  /* Seats: who uses the licence, with Release (your own seat, or anyone's for an administrator). */
  function seatCall(body) {
    return fetch('/api', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json(); });
  }
  function showSeats(data) {
    var box = $('seatList');
    box.hidden = false;
    box.className = 'box';
    box.textContent = '';
    if (!data.ok) { box.textContent = data.error || 'Could not read the seats.'; box.className = 'box bad'; return; }
    $('navSeats').textContent = data.seats.length ? String(data.seats.length) : '';
    if (!data.seats.length) { box.textContent = 'No seat is in use.'; return; }
    var table = document.createElement('table');
    box.className = '';
    table.innerHTML = '<tr><th>Who</th><th>PC</th><th>Last used</th><th></th></tr>';
    data.seats.forEach(function (seat) {
      var tr = document.createElement('tr');
      [seat.user + (seat.mine ? ' (you)' : ''), seat.label, seat.last].forEach(function (text) {
        var td = document.createElement('td');
        td.textContent = text || '';
        tr.appendChild(td);
      });
      var td = document.createElement('td');
      if (seat.mine || data.canRelease) {
        var b = document.createElement('button');
        b.className = 'btn';
        b.textContent = 'Release';
        b.addEventListener('click', function () {
          box.textContent = 'Releasing…';
          seatCall({ action: 'seatRelease', key: seat.key }).then(showSeats);
        });
        td.appendChild(b);
      }
      tr.appendChild(td);
      table.appendChild(tr);
    });
    box.appendChild(table);
  }
  $('seatRefresh').addEventListener('click', function () {
    $('seatList').textContent = 'Reading the seats…';
    seatCall({ action: 'seatList' }).then(showSeats).catch(function (e) { $('seatList').textContent = e.message; });
  });
  function showMode() {
    $('oauthBox').hidden = !$('modeOauth').checked;
    $('tbaBox').hidden = $('modeOauth').checked;
  }
  $('modeOauth').addEventListener('change', showMode);
  $('modeTba').addEventListener('change', showMode);
  $('signOut').addEventListener('click', function () {
    fetch('/api/signout', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
      .then(function () { load(); });
  });
  window.addEventListener('focus', load);   // back from the NetSuite sign-in tab: show who is signed in
  /* Updates: the address is saved with Save; Check reads the signed description, Install runs the new setup. */
  $('checkUpdate').addEventListener('click', function () {
    box('updateState', 'Checking…');
    fetch('/api/update').then(function (r) { return r.json(); }).then(function (data) {
      if (!data.ok) { box('updateState', data.error, 'bad'); return; }
      $('installUpdate').hidden = !data.available;
      box('updateState', data.message + (data.notes ? ' ' + data.notes : '') + ' (this PC has ' + data.current + ')',
        data.available ? '' : 'ok');
    }).catch(function (e) { box('updateState', e.message, 'bad'); });
  });
  $('installUpdate').addEventListener('click', function () {
    box('updateState', 'Downloading and checking the update…');
    fetch('/api/update', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{"op":"install"}' })
      .then(function (r) { return r.json(); })
      .then(function (data) { box('updateState', data.ok ? data.message : data.error, data.ok ? 'ok' : 'bad'); })
      .catch(function (e) { box('updateState', e.message, 'bad'); });
  });
  $('diagSave').addEventListener('click', function () {
    box('diagState', 'Saving…');
    fetch('/api/diagnostics', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
      .then(function (r) { return r.json(); })
      .then(function (d) { box('diagState', d.ok ? 'Saved: ' + d.path + '. Email this file to Techno Q support.' : d.error, d.ok ? 'ok' : 'bad'); })
      .catch(function (e) { box('diagState', e.message, 'bad'); });
  });
  load();
  var start = (location.hash || '').replace('#', '');
  if (start && document.querySelector('main section[data-section="' + start + '"]')) section(start);
})();
