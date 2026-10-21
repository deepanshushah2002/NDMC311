/* ==========================================================================
   NDMC 311 — shared pricing data + "plans" popup
   Used by index.html (main tool) and pricing.html.
   The popup opens on EVERY page load / refresh (nothing is remembered).

   To take payments: paste a payment link into `paymentUrl` for each plan
   (Razorpay / Cashfree payment link, UPI link, Google Form, WhatsApp, etc.)
   ========================================================================== */
(function () {
  'use strict';

  var BASE_MONTHLY = 199;

  var PLANS = [
    { id: 'monthly', name: '1 Month', months: 1, price: 199, tag: '', paymentUrl: '' },
    { id: 'half', name: '6 Months', months: 6, price: 999, tag: 'Most popular', featured: true, paymentUrl: '' },
    { id: 'yearly', name: '1 Year', months: 12, price: 1999, tag: 'Best savings', paymentUrl: '' }
  ];

  var FEATURES = [
    'Management dashboard (PDF)',
    'Officer-wise aging report (PDF)',
    'Excel exports incl. pivot lists',
    'Unlimited complaint files',
    'Light & dark themes'
  ];

  var inr = new Intl.NumberFormat('en-IN');
  function rupees(n) { return '₹' + inr.format(n); }
  function perMonth(p) { return Math.round(p.price / p.months); }
  function saving(p) { return p.months * BASE_MONTHLY - p.price; }
  function savingPct(p) { return Math.round((saving(p) / (p.months * BASE_MONTHLY)) * 100); }
  function periodLabel(p) { return p.months === 1 ? '/ month' : p.months === 12 ? '/ year' : '/ ' + p.months + ' months'; }

  var REDUCED_MOTION = !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);
  var ON_PRICING_PAGE = /\/pricing(\.html)?\/?$/i.test(location.pathname);

  /* ---------- Styles (injected so any page only needs styles.css + this file) ---------- */
  var CSS = [
    '.btn-outline{background:var(--bg-card);color:var(--ink-heading);border:1px solid var(--border-subtle);box-shadow:var(--shadow-sm)}',
    '.btn-outline:hover{border-color:var(--primary-indigo);color:var(--primary-indigo);transform:translateY(-2px);box-shadow:var(--shadow-md)}',
    '.modal-backdrop{position:fixed;inset:0;z-index:1000;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(9,13,22,.6);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);opacity:0;transition:opacity .25s ease}',
    '.modal-backdrop[hidden]{display:none}',
    '.modal-backdrop.open{opacity:1}',
    '.modal{position:relative;width:100%;max-width:520px;max-height:calc(100vh - 32px);overflow-y:auto;padding:30px 28px 24px;border-radius:22px;background:var(--bg-main);border:1px solid var(--border-card);box-shadow:var(--shadow-3d),0 30px 60px -20px rgba(0,0,0,.45);transform:translateY(16px) scale(.97);transition:transform .3s cubic-bezier(.16,1,.3,1)}',
    '.modal::before{content:"";position:absolute;inset:0 0 auto 0;height:5px;background:var(--grad-primary)}',
    '.modal-backdrop.open .modal{transform:none}',
    '.modal-close{position:absolute;top:14px;right:14px;width:34px;height:34px;border-radius:50%;border:1px solid var(--border-subtle);background:var(--bg-card);color:var(--ink-heading);font-size:18px;line-height:1;cursor:pointer;transition:all .2s ease}',
    '.modal-close:hover{border-color:var(--primary-rose);color:var(--primary-rose)}',
    '.modal-badge{display:inline-block;padding:3px 12px;border-radius:9999px;font-size:12px;font-weight:800;letter-spacing:.4px;text-transform:uppercase;color:var(--primary-indigo);background:rgba(99,102,241,.12);border:1px solid rgba(99,102,241,.3)}',
    '.modal h2{margin:12px 0 6px;padding-right:36px;font-size:22px;font-weight:800;letter-spacing:-.02em;color:var(--ink-heading)}',
    '.modal p.lead{font-size:14px;color:var(--ink-muted);margin-bottom:18px}',
    '.promo-list{display:flex;flex-direction:column;gap:10px}',
    '.promo-row{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;border-radius:14px;background:var(--bg-card);border:1px solid var(--border-subtle)}',
    '.promo-row.featured{border-color:rgba(99,102,241,.6);box-shadow:0 0 0 1px rgba(99,102,241,.4)}',
    '.promo-info{display:flex;flex-direction:column;min-width:0}',
    '.promo-info strong{font-size:15px;color:var(--ink-heading)}',
    '.promo-info span{font-size:12.5px;color:var(--ink-muted)}',
    '.promo-price{font-size:20px;font-weight:800;color:var(--ink-heading);white-space:nowrap}',
    '.promo-row .btn{padding:9px 16px;font-size:13px}',
    '.promo-right{display:flex;align-items:center;gap:12px}',
    '.modal-actions{display:flex;align-items:center;justify-content:center;flex-wrap:wrap;gap:6px 20px;margin-top:16px}',
    '.modal-later{background:none;border:none;font:inherit;font-size:13.5px;font-weight:600;color:var(--ink-muted);text-decoration:underline;text-underline-offset:3px;cursor:pointer}',
    '.modal-later:hover{color:var(--primary-indigo)}',
    '.toast{position:fixed;left:50%;bottom:28px;z-index:1100;max-width:calc(100vw - 32px);padding:12px 20px;border-radius:12px;font-size:14px;font-weight:600;color:var(--ink-heading);background:var(--bg-main);border:1px solid var(--border-subtle);box-shadow:var(--shadow-lg);transform:translate(-50%,20px);opacity:0;pointer-events:none;transition:all .3s ease}',
    '.toast.show{opacity:1;transform:translate(-50%,0)}',
    '@media (max-width:480px){.promo-row{flex-wrap:wrap}.promo-right{width:100%;justify-content:space-between}}'
  ].join('\n');

  var modal, toastEl, lastFocus = null, toastTimer = 0;

  function injectStyles() {
    var s = document.createElement('style');
    s.id = 'promoStyles';
    s.textContent = CSS;
    document.head.appendChild(s);
  }

  function buildDom() {
    modal = document.createElement('div');
    modal.className = 'modal-backdrop';
    modal.id = 'promoModal';
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    modal.setAttribute('aria-labelledby', 'promoTitle');
    modal.hidden = true;

    var rows = PLANS.map(function (p) {
      var sub = p.months === 1 ? 'Pay monthly' : rupees(perMonth(p)) + '/month · save ' + rupees(saving(p));
      return '<div class="promo-row' + (p.featured ? ' featured' : '') + '">' +
        '<div class="promo-info"><strong>' + p.name + '</strong><span>' + sub + '</span></div>' +
        '<div class="promo-right"><span class="promo-price">' + rupees(p.price) + '</span>' +
        '<button class="btn ' + (p.featured ? 'btn-primary' : 'btn-outline') + '" type="button" data-plan="' + p.id + '">Choose</button></div>' +
        '</div>';
    }).join('');

    modal.innerHTML =
      '<div class="modal">' +
      '<button class="modal-close" id="promoClose" type="button" aria-label="Close">&times;</button>' +
      '<span class="modal-badge">Get full access</span>' +
      '<h2 id="promoTitle">Unlock NDMC 311 Analytics Studio</h2>' +
      '<p class="lead">Dashboards, officer-wise reports and Excel exports in a few clicks. Choose a plan:</p>' +
      '<div class="promo-list">' + rows + '</div>' +
      '<div class="modal-actions">' +
      (ON_PRICING_PAGE ? '' : '<a class="modal-later" href="pricing.html">See full plan details</a>') +
      '<button class="modal-later" id="promoLater" type="button">Maybe later</button>' +
      '</div></div>';
    document.body.appendChild(modal);

    toastEl = document.createElement('div');
    toastEl.className = 'toast';
    toastEl.setAttribute('role', 'status');
    toastEl.setAttribute('aria-live', 'polite');
    document.body.appendChild(toastEl);

    modal.querySelector('#promoClose').addEventListener('click', closeModal);
    modal.querySelector('#promoLater').addEventListener('click', closeModal);
    modal.addEventListener('click', function (e) { if (e.target === modal) closeModal(); });
  }

  function toast(msg) {
    if (!toastEl) return;
    toastEl.textContent = msg;
    toastEl.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove('show'); }, 3500);
  }

  function openModal() {
    if (!modal || !modal.hidden) return;
    lastFocus = document.activeElement;
    modal.hidden = false;
    document.body.style.overflow = 'hidden';
    requestAnimationFrame(function () {
      modal.classList.add('open');
      var c = modal.querySelector('#promoClose');
      if (c) c.focus();
    });
  }

  function closeModal() {
    if (!modal || modal.hidden) return;
    modal.classList.remove('open');
    document.body.style.overflow = '';
    setTimeout(function () { modal.hidden = true; }, REDUCED_MOTION ? 0 : 250);
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }

  function choosePlan(id) {
    var p = PLANS.filter(function (x) { return x.id === id; })[0];
    if (!p) return;
    closeModal();
    if (p.paymentUrl) {
      window.open(p.paymentUrl, '_blank', 'noopener');
    } else {
      toast(p.name + ' (' + rupees(p.price) + ') selected. Payment link is not set up yet.');
    }
  }

  // Any element with data-plan="<id>" (popup buttons, pricing-page cards) starts checkout
  document.addEventListener('click', function (e) {
    var btn = e.target.closest && e.target.closest('[data-plan]');
    if (btn) choosePlan(btn.getAttribute('data-plan'));
  });

  document.addEventListener('keydown', function (e) {
    if (!modal || modal.hidden) return;
    if (e.key === 'Escape') { closeModal(); return; }
    if (e.key === 'Tab') { // keep focus inside the popup
      var f = modal.querySelectorAll('button, a[href]');
      if (!f.length) return;
      var first = f[0], last = f[f.length - 1];
      if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
  });

  // Public API (used by pricing.html to draw its plan cards from the same data)
  window.NDMCPricing = {
    PLANS: PLANS, FEATURES: FEATURES,
    rupees: rupees, perMonth: perMonth, saving: saving, savingPct: savingPct, periodLabel: periodLabel,
    choosePlan: choosePlan, toast: toast, open: openModal, close: closeModal
  };

  function boot() {
    try {
      injectStyles();
      buildDom();
      // Shown on EVERY load / refresh of the main page — never on the pricing page itself
      if (!ON_PRICING_PAGE) setTimeout(openModal, 400);
    } catch (err) {
      console.warn('Pricing popup failed to start:', err);
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
