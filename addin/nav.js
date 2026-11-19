/* TQ Excel Connector - the panel's screens.
 * One screen shows at a time (Home, This report, Add a filter, Layout, Dashboards, Send, Upload, Formulas, Data,
 * Drill down, Chat, New report, Welcome). Buttons with data-go="name" open a screen; data-back goes back.
 * A screen, or a block inside it, marked "lazy" gets open = true and a "toggle" event when it is shown, the same
 * signal the older fold-out sections gave, so the scripts that load their lists on first open work unchanged. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var current = 'home';
  var trail = [];   // the screens visited before the current one, for Back

  function screens() { return Array.prototype.slice.call(document.querySelectorAll('.screen[data-screen]')); }

  function lazyIn(screen) {
    var list = Array.prototype.slice.call(screen.querySelectorAll('.lazy'));
    if (screen.classList.contains('lazy')) list.unshift(screen);
    return list;
  }

  function show(name) {
    var target = document.querySelector('.screen[data-screen="' + name + '"]');
    if (!target) return;
    screens().forEach(function (s) {
      if (s === target) return;
      s.hidden = true;
      lazyIn(s).forEach(function (el) { el.open = false; });
    });
    target.hidden = false;
    current = name;
    var box = document.querySelector('.screens');
    if (box) box.scrollTop = 0;
    lazyIn(target).forEach(function (el) {
      el.open = true;
      el.dispatchEvent(new Event('toggle'));
    });
    if (name === 'filter' && window.TQFilters) window.TQFilters.load();
    var chat = $('openChat');
    if (chat) chat.classList.toggle('on', name === 'chat');
    document.dispatchEvent(new CustomEvent('tq:screen', { detail: name }));
  }

  /* Keyboard and screen-reader users land on the new screen's heading (only when they opened it themselves:
   * a screen that changes because another sheet was clicked leaves the focus in Excel). */
  function focusHeading(name) {
    var target = document.querySelector('.screen[data-screen="' + name + '"]');
    var h = target && (target.querySelector('h2') || target.querySelector('.hero b'));
    if (!h) return;
    h.tabIndex = -1;
    try { h.focus({ preventScroll: true }); } catch (e) { h.focus(); }
  }

  /* Opens a screen; Back returns to where you were. */
  function go(name) {
    if (name !== current) {
      trail.push(current);
      if (trail.length > 20) trail.shift();
    }
    show(name);
    focusHeading(name);
  }

  /* Back to the screen before (or to a named one, or Home). */
  function back(to) {
    var name = to || trail.pop() || 'home';
    if (to) trail = [];
    if (name === current) name = 'home';
    show(name);
    focusHeading(name);
  }

  /* Replaces the current screen without adding a Back step (used when the open sheet changes). */
  function replace(name) {
    trail = [];
    show(name);
  }

  // ------------------------------------------------------------- the settings menu --

  function closeMenu() {
    $('menu').hidden = true;
    $('menuButton').setAttribute('aria-expanded', 'false');
  }

  // --------------------------------------------- segment pills over the filter field --

  /* The "Filter on" dropdown (filled by filters.js) shown as pills: the first five fields, then More. */
  function pills() {
    var select = $('fField');
    var box = $('fPills');
    if (!select || !box) return;
    var draw = function () {
      box.textContent = '';
      var options = Array.prototype.slice.call(select.options);
      if (!options.length) { select.hidden = false; return; }
      var shown = options.slice(0, 5);
      var chosen = options[select.selectedIndex];
      if (chosen && shown.indexOf(chosen) === -1) shown.push(chosen);
      shown.forEach(function (o) {
        var b = document.createElement('button');
        b.type = 'button';
        b.textContent = o.textContent.replace(/ \((custom segment|customer, vendor, employee)\)$/, '');
        b.title = o.textContent;
        if (o === chosen) b.className = 'on';
        b.addEventListener('click', function () {
          select.value = o.value;
          select.dispatchEvent(new Event('change'));
          draw();
        });
        box.appendChild(b);
      });
      if (options.length > shown.length) {
        var more = document.createElement('button');
        more.type = 'button';
        more.textContent = 'More ▾';
        more.addEventListener('click', function () { select.hidden = !select.hidden; if (!select.hidden) select.focus(); });
        box.appendChild(more);
      }
      select.hidden = true;
    };
    select.addEventListener('change', function () { setTimeout(draw, 0); });
    new MutationObserver(draw).observe(select, { childList: true });
    draw();
  }

  // ------------------------------------------------------------------- the theme --

  /* Dark when Excel's theme is dark (Office.context.officeTheme), otherwise as Windows says. */
  function theme() {
    try {
      var t = Office.context.officeTheme;
      if (!t || !t.bodyBackgroundColor) return;
      var hex = t.bodyBackgroundColor.replace('#', '');
      if (hex.length === 3) hex = hex.replace(/./g, '$&$&');
      var n = parseInt(hex, 16);
      var light = (0.299 * (n >> 16 & 255) + 0.587 * (n >> 8 & 255) + 0.114 * (n & 255)) > 140;
      document.documentElement.setAttribute('data-theme', light ? 'light' : 'dark');
      try { localStorage.setItem('tq.theme', light ? 'light' : 'dark'); } catch (e) { /* the setup page then follows Windows */ }
    } catch (e) { /* outside Excel: the CSS follows Windows */ }
  }

  function start() {
    document.addEventListener('click', function (event) {
      var b = event.target.closest('[data-go]');
      if (b) { closeMenu(); go(b.getAttribute('data-go')); return; }
      var k = event.target.closest('[data-back]');
      if (k) { back(k.getAttribute('data-back') || ''); return; }
      if (!event.target.closest('#menu') && !event.target.closest('#menuButton')) closeMenu();
    });
    $('menuButton').addEventListener('click', function () {
      var open = $('menu').hidden;
      $('menu').hidden = !open;
      $('menuButton').setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    document.addEventListener('keydown', function (event) { if (event.key === 'Escape') closeMenu(); });
    $('openChat').addEventListener('click', function () { if (current === 'chat') back(); else go('chat'); });
    pills();
    if (window.Office && Office.onReady) Office.onReady(theme);
  }

  window.TQNav = { go: go, back: back, replace: replace, current: function () { return current; } };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
