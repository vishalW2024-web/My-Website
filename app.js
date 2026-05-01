/* ============================================================
   app.js — 3D QR Menu Portfolio
   Excel (SheetJS) + Multi-step Form + Image Library Overlay
   ============================================================ */

'use strict';

// ────────────────────────────────────────────────────────────
// CONFIGURATION
// ────────────────────────────────────────────────────────────
const CONFIG = {
  excelFile: 'Excel for my website.xlsx',
  portfolioSheet: 'Working QR',
  imageLibrarySheet: 'Image Library',
  whatsappNumber: '918459987710',
  businessName: '3D QR Menu',
  googleForm: {
    actionUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSdXQWDQXMM5goCsnBhL9H3cb9s2UR7VawB8y7TcmI8e3wtyFQ/formResponse',
    fields: {
      name:          'entry.548929712',
      email:         'entry.641265205',
      phone:         'entry.319542046',
      business_name: 'entry.1577217556',
      outlet_type:   'entry.1183475351',
      locations:     'entry.1988247197',
      city:          'entry.398739428',
      country:       'entry.1401824165',
      goals:         'entry.879905860',
      source:        'entry.97475968',
      message:       'entry.1383922222',
    }
  },
};

// ────────────────────────────────────────────────────────────
// UTILITY
// ────────────────────────────────────────────────────────────
function escHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function convertDriveLink(url) {
  if (!url) return '';
  let fileId = null;
  const m1 = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) fileId = m1[1];
  const m2 = url.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (m2) fileId = m2[1];
  const m3 = url.match(/drive\.google\.com\/uc\?.*id=([a-zA-Z0-9_-]+)/);
  if (m3) fileId = m3[1];
  if (fileId) return `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;
  return url;
}

// ────────────────────────────────────────────────────────────
// EXCEL FILE READER (SheetJS)
// ────────────────────────────────────────────────────────────
let excelWorkbook = null;

async function loadExcel() {
  if (excelWorkbook) return excelWorkbook;
  try {
    const resp = await fetch(CONFIG.excelFile, { cache: 'no-store' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.arrayBuffer();
    excelWorkbook = XLSX.read(data, { type: 'array' });
    return excelWorkbook;
  } catch (err) {
    console.warn('Failed to load Excel file:', err.message);
    return null;
  }
}

function getSheetData(wb, sheetName) {
  if (!wb) return [];
  // Try exact match first, then trimmed match (Excel sheet names may have trailing spaces)
  let sheet = wb.Sheets[sheetName];
  if (!sheet) {
    const found = wb.SheetNames.find(n => n.trim() === sheetName.trim());
    if (found) sheet = wb.Sheets[found];
  }
  if (!sheet) return [];
  const rawData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  return rawData.map(row => {
    const cleaned = {};
    for (const [key, val] of Object.entries(row)) {
      cleaned[key.trim()] = typeof val === 'string' ? val.trim() : val;
    }
    return cleaned;
  });
}

// ────────────────────────────────────────────────────────────
// PORTFOLIO (from Excel "Working QR" sheet)
// ────────────────────────────────────────────────────────────
let portfolioData = [];

async function loadPortfolio() {
  const wb = await loadExcel();
  if (!wb) return [];
  const rows = getSheetData(wb, CONFIG.portfolioSheet);
  portfolioData = rows.map(row => ({
    name: row['Name of the Outlet'] || '',
    type: row['outlet Type'] || row['Outlet Type'] || '',
    qrImageUrl: convertDriveLink(row['QR Image Link'] || ''),
    website: row['Website'] || '',
  })).filter(r => r.name);
  return portfolioData;
}

function renderPortfolio(items) {
  const grid = document.getElementById('portfolio-grid');
  const viewAllBtn = document.getElementById('portfolio-view-all-btn');
  if (!grid) return;

  if (items.length === 0) {
    grid.innerHTML = '<p style="color:var(--text-3);text-align:center;grid-column:1/-1;padding:40px">Portfolio data loading...</p>';
    return;
  }

  // Show first 3 cards initially
  const initialCount = 3;
  const buildCard = (item, i) => `
    <div class="glass-card portfolio-card tilt-3d reveal" style="overflow:hidden;width:calc(33.333% - 24px);min-width:280px;max-width:380px;flex-shrink:0">
      <div style="aspect-ratio:1/1;background:var(--surface-2);display:flex;align-items:center;justify-content:center;overflow:hidden;border-radius:12px 12px 0 0">
        <img src="${escHtml(item.qrImageUrl)}" alt="${escHtml(item.name)} QR Menu" loading="lazy" style="width:100%;height:100%;object-fit:cover" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'">
        <div style="display:none;flex-direction:column;align-items:center;gap:8px;color:var(--text-3)">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1" style="width:32px;height:32px;opacity:0.4"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/></svg>
          <span style="font-size:0.75rem">Image loading...</span>
        </div>
      </div>
      <div style="padding:20px 24px">
        <h3 style="font-family:var(--font-head);font-size:1.15rem;font-weight:700;color:var(--text-1);margin-bottom:4px">${escHtml(item.name)}</h3>
        ${item.type ? `<p style="font-size:0.85rem;color:var(--text-3);margin-bottom:${(item.type.toLowerCase().includes('dessert') || item.name.toLowerCase().includes('dessert')) ? '6px' : '16px'}">(${escHtml(item.type)})</p>` : ''}
        ${(item.type.toLowerCase().includes('dessert') || item.name.toLowerCase().includes('dessert')) ? `<p style="font-size:0.72rem;color:var(--gold);font-style:italic;margin-bottom:12px;line-height:1.4">(Click on the image in the website for best experience)</p>` : ''}
        ${item.website ? `<a href="${escHtml(item.website)}" target="_blank" rel="noopener noreferrer" class="btn btn-outline" style="padding:8px 20px;font-size:0.85rem">Experience Our Work →</a>` : ''}
      </div>
    </div>`;

  // Render initial cards
  grid.innerHTML = items.slice(0, initialCount).map((item, i) => buildCard(item, i)).join('');

  // View All button
  if (items.length > initialCount && viewAllBtn) {
    viewAllBtn.style.display = 'inline-flex';
    viewAllBtn.addEventListener('click', () => {
      // Add remaining cards
      const remaining = items.slice(initialCount).map((item, i) => buildCard(item, i + initialCount)).join('');
      grid.innerHTML += remaining;
      viewAllBtn.style.display = 'none';
      // Re-init reveal and tilt animations for new cards
      observeReveal();
      init3DTilt();
    });
  }
}

// ────────────────────────────────────────────────────────────
// IMAGE LIBRARY (from Excel "Image Library" sheet)
// ────────────────────────────────────────────────────────────
let imageLibraryData = [];

async function loadImageLibrary() {
  const wb = await loadExcel();
  if (!wb) return [];
  const rows = getSheetData(wb, CONFIG.imageLibrarySheet);
  imageLibraryData = rows.map(row => ({
    rawUrl: convertDriveLink(row['Raw image Url'] || row['Raw Image Link'] || ''),
    enhUrl: convertDriveLink(row['Enhanced Image Url'] || row['Enhanced Image Link'] || ''),
    itemName: row['Item Name'] || row['Item name'] || row['item name'] || '',
  })).filter(r => r.rawUrl || r.enhUrl);
  return imageLibraryData;
}

// Render the inline preview (only first pair + bold View More button)
function renderBeforeAfter(images) {
  const container = document.getElementById('before-after-container');
  if (!container) return;

  if (images.length === 0) {
    container.innerHTML = '<p style="color:var(--text-3);text-align:center;padding:40px">Loading image pairs...</p>';
    return;
  }

  // Always show the first pair (A2, B2, C2) in the same card layout as the overlay
  const img = images[0];
  let html = `
    <div class="il-card ba-inline-card">
      <div class="il-card-headers">
        <div class="il-header-raw">Raw Image</div>
        <div class="il-header-enh">Enhanced Image</div>
      </div>
      <div class="il-card-images">
        <div class="il-img-box il-raw-box">
          <img src="${escHtml(img.rawUrl)}" alt="Raw" loading="lazy" onerror="this.style.display='none'">
        </div>
        <div class="il-img-box il-enh-box">
          <img src="${escHtml(img.enhUrl)}" alt="Enhanced" loading="lazy" onerror="this.style.display='none'">
        </div>
      </div>
      ${img.itemName ? `<div class="il-card-item-name">${escHtml(img.itemName)}</div>` : ''}
    </div>
    <div style="text-align:center;margin-top:16px">
      <button class="ba-view-more" id="ba-open-library-btn" style="font-weight:700">View More +</button>
    </div>`;

  container.innerHTML = html;

  // Click "View More +" → open Image Library overlay
  document.getElementById('ba-open-library-btn')?.addEventListener('click', () => openImageLibrary());
}

// Open full Image Library overlay
function openImageLibrary() {
  const overlay = document.getElementById('image-library-overlay');
  if (!overlay || imageLibraryData.length === 0) return;

  document.body.style.overflow = 'hidden';

  let html = `
    <div class="il-container">
      <div class="il-header-bar">
        <h2 class="il-title">Image Library</h2>
        <button class="il-close-btn" id="il-close-btn" aria-label="Close">✕</button>
      </div>
      <p class="il-subtitle">Professional enhancement for every image. See the difference.</p>

      <div class="il-grid">
  `;

  // Start from row 2 (index 1) — skip the first pair which is shown inline
  const overlayData = imageLibraryData.slice(1);
  const firstBatch = overlayData.slice(0, 2);
  firstBatch.forEach((img, i) => {
    html += buildILCard(img, i);
  });

  html += `</div>`;

  // VIEW ALL button (if more than 2 pairs in overlay)
  if (overlayData.length > 2) {
    html += `
      <div style="text-align:center;margin:32px 0" id="il-view-all-wrap">
        <button class="btn btn-primary" id="il-view-all-btn" style="padding:12px 40px;font-size:1rem">VIEW ALL</button>
      </div>
      <div class="il-grid il-extra-grid" id="il-extra-grid" style="display:none">
    `;
    overlayData.slice(2).forEach((img, i) => {
      html += buildILCard(img, i + 2);
    });
    html += `</div>`;
  }

  // Selling / convincing content below
  html += `
      <div class="il-selling-section">
        <div class="il-sell-card">
          <div class="il-sell-icon">📸</div>
          <h4>Send Any Photo</h4>
          <p>Even low-quality phone camera shots work. We accept everything.</p>
        </div>
        <div class="il-sell-card">
          <div class="il-sell-icon">✨</div>
          <h4>Professional Enhancement</h4>
          <p>Color-grading, upscaling, and theme-matching — included free.</p>
        </div>
        <div class="il-sell-card">
          <div class="il-sell-icon">🎨</div>
          <h4>Brand-Matched Results</h4>
          <p>Every image is tailored to your outlet's unique theme and vibe.</p>
        </div>
      </div>
      <div style="text-align:center;padding:24px 0">
        <a href="#contact" class="btn btn-primary" onclick="closeImageLibrary()" style="padding:14px 36px;font-size:1rem">Get Your Menu Enhanced →</a>
      </div>
    </div>
  `;

  overlay.innerHTML = html;
  overlay.style.display = 'block';
  overlay.style.opacity = '0';
  requestAnimationFrame(() => { overlay.style.opacity = '1'; });

  // Close button
  document.getElementById('il-close-btn')?.addEventListener('click', closeImageLibrary);
  overlay.addEventListener('click', e => { if (e.target === overlay) closeImageLibrary(); });

  // VIEW ALL button
  document.getElementById('il-view-all-btn')?.addEventListener('click', () => {
    const extra = document.getElementById('il-extra-grid');
    if (extra) { extra.style.display = ''; }
    document.getElementById('il-view-all-wrap')?.remove();
    initImageLightbox(); // Bind lightbox to newly revealed images
  });

  // Bind lightbox to overlay images
  initImageLightbox();
}

function buildILCard(img, index) {
  return `
    <div class="il-card">
      <div class="il-card-headers">
        <div class="il-header-raw">Raw Image</div>
        <div class="il-header-enh">Enhanced Image</div>
      </div>
      <div class="il-card-images">
        <div class="il-img-box il-raw-box">
          <img src="${escHtml(img.rawUrl)}" alt="Raw" loading="lazy" onerror="this.style.display='none'">
        </div>
        <div class="il-img-box il-enh-box">
          <img src="${escHtml(img.enhUrl)}" alt="Enhanced" loading="lazy" onerror="this.style.display='none'">
        </div>
      </div>
      ${img.itemName ? `<div class="il-card-item-name">${escHtml(img.itemName)}</div>` : ''}
    </div>`;
}

function closeImageLibrary() {
  const overlay = document.getElementById('image-library-overlay');
  if (overlay) {
    overlay.style.opacity = '0';
    setTimeout(() => { overlay.style.display = 'none'; }, 300);
  }
  document.body.style.overflow = '';
}

window.openImageLibrary = openImageLibrary;
window.closeImageLibrary = closeImageLibrary;

// ────────────────────────────────────────────────────────────
// PORTFOLIO MODAL
// ────────────────────────────────────────────────────────────
function openPortfolioModal(index) {
  const item = portfolioData[index];
  if (!item) return;
  const backdrop = document.getElementById('portfolio-modal-backdrop');
  document.getElementById('modal-restaurant-name').textContent = item.name;
  document.getElementById('modal-restaurant-category').textContent = item.type;
  document.getElementById('modal-qr-img').src = item.qrImageUrl;
  document.getElementById('modal-qr-img').alt = `QR Code for ${item.name}`;
  const waLink = `https://wa.me/${CONFIG.whatsappNumber}?text=${encodeURIComponent(`Hi! I saw ${item.name}'s QR menu. I'd like one too!`)}`;
  document.getElementById('modal-whatsapp-btn').href = waLink;
  backdrop.classList.add('open');
  backdrop.setAttribute('aria-hidden', 'false');
  document.body.style.overflow = 'hidden';
}

function closePortfolioModal() {
  const backdrop = document.getElementById('portfolio-modal-backdrop');
  if (!backdrop) return;
  backdrop.classList.remove('open');
  backdrop.setAttribute('aria-hidden', 'true');
  document.body.style.overflow = '';
}
window.openPortfolioModal = openPortfolioModal;
window.closePortfolioModal = closePortfolioModal;

// ────────────────────────────────────────────────────────────
// MULTI-STEP FORM
// ────────────────────────────────────────────────────────────
function initMultiStepForm() {
  const form = document.getElementById('contact-form');
  if (!form) return;

  let currentStep = 1;
  const totalSteps = 3;

  function showStep(step) {
    for (let i = 1; i <= totalSteps; i++) {
      const s = document.getElementById(`form-step-${i}`);
      if (s) s.classList.toggle('active', i === step);
    }
    document.querySelectorAll('.step-dot').forEach(dot => {
      const n = parseInt(dot.dataset.step);
      dot.classList.toggle('active', n === step);
      dot.classList.toggle('done', n < step);
    });
    currentStep = step;
  }

  // Next buttons
  form.querySelectorAll('.form-next-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const nextStep = parseInt(btn.dataset.next);
      // Validate current step required fields
      const currentStepEl = document.getElementById(`form-step-${currentStep}`);
      const required = currentStepEl.querySelectorAll('[required]');
      let valid = true;
      required.forEach(field => {
        if (!field.value.trim()) {
          valid = false;
          field.style.borderColor = 'var(--error)';
          const err = field.closest('.form-group')?.querySelector('.form-field-error');
          if (err) err.style.display = 'block';
        } else {
          field.style.borderColor = '';
          const err = field.closest('.form-group')?.querySelector('.form-field-error');
          if (err) err.style.display = 'none';
        }
      });
      if (valid) showStep(nextStep);
    });
  });

  // Back buttons
  form.querySelectorAll('.form-prev-btn').forEach(btn => {
    btn.addEventListener('click', () => showStep(parseInt(btn.dataset.prev)));
  });

  // Submit form
  form.addEventListener('submit', async e => {
    e.preventDefault();
    const submitBtn = document.getElementById('form-submit-btn');
    const status = document.getElementById('form-status');
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="animation:spin 1s linear infinite"><path d="M21 12a9 9 0 11-6.219-8.56"/></svg> Sending…';

    try {
      const gf = CONFIG.googleForm;

      // Collect goals checkboxes as separate values
      const goalsChecked = Array.from(form.querySelectorAll('[name="goals"]:checked')).map(cb => cb.value);

      // Build URL-encoded body
      const params = new URLSearchParams();
      params.append(gf.fields.name, form.querySelector('[name="name"]').value);
      params.append(gf.fields.email, form.querySelector('[name="email"]').value);
      params.append(gf.fields.phone, form.querySelector('[name="phone"]').value);
      params.append(gf.fields.business_name, form.querySelector('[name="business_name"]').value);
      params.append(gf.fields.outlet_type, form.querySelector('[name="outlet_type"]').value);
      params.append(gf.fields.locations, form.querySelector('[name="locations"]').value);
      params.append(gf.fields.city, form.querySelector('[name="city"]')?.value || '');
      params.append(gf.fields.country, form.querySelector('[name="country"]')?.value || '');
      // Send each checkbox value as a separate entry (Google Forms expects this for checkboxes)
      goalsChecked.forEach(goal => {
        params.append(gf.fields.goals, goal);
      });
      params.append(gf.fields.source, form.querySelector('[name="source"]')?.value || '');
      params.append(gf.fields.message, form.querySelector('[name="message"]')?.value || '');

      // Multi-page Google Form requires pageHistory to submit all fields
      params.append('pageHistory', '0,1,2,3');
      params.append('fvv', '1');
      params.append('partialResponse', '[null,null,null,null]');
      params.append('submit', 'Submit');

      // Submit via fetch with no-cors (Google Forms doesn't allow CORS)
      await fetch(gf.actionUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      });

      status.className = 'form-status success';
      status.textContent = '✓ Message received! We\'ll reach out within 24 hours.';
      form.reset();
      showStep(1);
    } catch (err) {
      status.className = 'form-status error';
      status.textContent = '✗ Something went wrong. Please try WhatsApp instead.';
    } finally {
      submitBtn.disabled = false;
      submitBtn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 2L11 13"/><path d="M22 2L15 22 11 13 2 9l20-7z"/></svg> Submit';
    }
  });
}

// ────────────────────────────────────────────────────────────
// NAVBAR
// ────────────────────────────────────────────────────────────
function initNavbar() {
  const navbar = document.getElementById('navbar');
  const hamburger = document.getElementById('nav-hamburger');
  const mobileNav = document.getElementById('nav-mobile');
  const mobileClose = document.getElementById('nav-mobile-close');

  let ticking = false;
  window.addEventListener('scroll', () => {
    if (!ticking) {
      requestAnimationFrame(() => {
        navbar.classList.toggle('scrolled', window.scrollY > 40);
        ticking = false;
      });
      ticking = true;
    }
  }, { passive: true });

  hamburger.addEventListener('click', () => {
    if (mobileNav.classList.contains('open')) {
      closeMobileNav();
    } else {
      mobileNav.classList.add('open');
      mobileNav.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
    }
  });

  function closeMobileNav() {
    mobileNav.classList.remove('open');
    mobileNav.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  }

  mobileClose.addEventListener('click', closeMobileNav);
  mobileNav.querySelectorAll('a').forEach(a => a.addEventListener('click', closeMobileNav));
  mobileNav.addEventListener('click', e => { if (e.target === mobileNav) closeMobileNav(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && mobileNav.classList.contains('open')) closeMobileNav(); });
}

// ────────────────────────────────────────────────────────────
// SCROLL REVEAL
// ────────────────────────────────────────────────────────────
function observeReveal() {
  const observer = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('revealed');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
  document.querySelectorAll('.reveal:not(.revealed)').forEach(el => observer.observe(el));
}

// ────────────────────────────────────────────────────────────
// FAQ ACCORDION (only one open at a time)
// ────────────────────────────────────────────────────────────
function initFaq() {
  document.querySelectorAll('.faq-question').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = btn.closest('.faq-item');
      const isOpen = item.classList.contains('open');
      document.querySelectorAll('.faq-item.open').forEach(i => {
        i.classList.remove('open');
        i.querySelector('.faq-question').setAttribute('aria-expanded', 'false');
      });
      if (!isOpen) {
        item.classList.add('open');
        btn.setAttribute('aria-expanded', 'true');
      }
    });
  });
}

// ────────────────────────────────────────────────────────────
// SMOOTH SCROLL
// ────────────────────────────────────────────────────────────
function initSmoothScroll() {
  document.querySelectorAll('a[href^="#"]').forEach(a => {
    a.addEventListener('click', e => {
      const target = document.querySelector(a.getAttribute('href'));
      if (target) {
        e.preventDefault();
        const top = target.getBoundingClientRect().top + window.scrollY - 80;
        window.scrollTo({ top, behavior: 'smooth' });
      }
    });
  });
}

// ────────────────────────────────────────────────────────────
// COUNTER ANIMATION
// ────────────────────────────────────────────────────────────
function animateCounters() {
  const counters = document.querySelectorAll('[data-counter]');
  if (!counters.length) return;
  const observer = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      const el = entry.target;
      const target = parseFloat(el.dataset.counter);
      const suffix = el.dataset.suffix || '';
      const prefix = el.dataset.prefix || '';
      const start = performance.now();
      function update(now) {
        const progress = Math.min((now - start) / 1800, 1);
        const eased = 1 - Math.pow(1 - progress, 3);
        el.textContent = prefix + Math.round(eased * target).toLocaleString() + suffix;
        if (progress < 1) requestAnimationFrame(update);
      }
      requestAnimationFrame(update);
      observer.unobserve(el);
    });
  }, { threshold: 0.5 });
  counters.forEach(el => observer.observe(el));
}

const spinStyle = document.createElement('style');
spinStyle.textContent = '@keyframes spin{to{transform:rotate(360deg)}} @keyframes fadeIn{from{opacity:0}to{opacity:1}}';
document.head.appendChild(spinStyle);

// ────────────────────────────────────────────────────────────
// ESC KEY / MODAL
// ────────────────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    closePortfolioModal();
    closeImageLibrary();
  }
});

document.addEventListener('DOMContentLoaded', () => {
  const backdrop = document.getElementById('portfolio-modal-backdrop');
  if (backdrop) backdrop.addEventListener('click', e => { if (e.target === backdrop) closePortfolioModal(); });
});

// ────────────────────────────────────────────────────────────
// INTRO SPLASH
// ────────────────────────────────────────────────────────────
function initIntroSplash() {
  const splash = document.getElementById('intro-splash');
  if (!splash) return;
  document.body.style.overflow = 'hidden';

  const scene = splash.querySelector('.intro-scene');
  if (scene) {
    for (let i = 0; i < 8; i++) {
      const qr = document.createElement('div');
      qr.className = 'floating-qr';
      qr.style.cssText = `position:absolute;width:${30+Math.random()*40}px;height:${30+Math.random()*40}px;top:${Math.random()*100}%;left:${Math.random()*100}%;opacity:${0.08+Math.random()*0.15};animation:float-qr ${6+Math.random()*8}s ${Math.random()*4}s ease-in-out infinite,spin-qr ${10+Math.random()*15}s linear infinite;`;
      qr.innerHTML = `<svg viewBox="0 0 21 21" fill="rgba(212,175,55,0.6)" style="width:100%;height:100%"><rect x="0" y="0" width="7" height="7"/><rect x="2" y="2" width="3" height="3" fill="#020408"/><rect x="14" y="0" width="7" height="7"/><rect x="16" y="2" width="3" height="3" fill="#020408"/><rect x="0" y="14" width="7" height="7"/><rect x="2" y="16" width="3" height="3" fill="#020408"/><rect x="8" y="8" width="2" height="2"/><rect x="11" y="8" width="2" height="2"/><rect x="8" y="11" width="2" height="2"/><rect x="14" y="14" width="2" height="2"/><rect x="17" y="14" width="2" height="2"/><rect x="14" y="17" width="2" height="2"/><rect x="17" y="17" width="4" height="4"/></svg>`;
      scene.appendChild(qr);
    }
  }

  setTimeout(() => {
    splash.classList.add('dismiss');
    document.body.style.overflow = '';

    // Show hero video, hide placeholder
    const heroVideo = document.getElementById('hero-demo-video');
    const heroPlaceholder = document.getElementById('hero-video-placeholder');
    if (heroVideo) {
      heroVideo.style.display = 'block';
      heroVideo.play().catch(() => {}); // autoplay may be blocked
    }
    if (heroPlaceholder) {
      heroPlaceholder.style.display = 'none';
    }

    // Init mute/unmute toggle
    initHeroMuteToggle();

    setTimeout(() => splash.remove(), 1000);
  }, 4200);
}

// ────────────────────────────────────────────────────────────
// HERO VIDEO MUTE TOGGLE
// ────────────────────────────────────────────────────────────
function initHeroMuteToggle() {
  const btn = document.getElementById('hero-mute-toggle');
  const video = document.getElementById('hero-demo-video');
  const iconOff = document.getElementById('mute-icon-off');
  const iconOn = document.getElementById('mute-icon-on');
  if (!btn || !video) return;

  btn.addEventListener('click', () => {
    video.muted = !video.muted;
    if (video.muted) {
      iconOff.style.display = '';
      iconOn.style.display = 'none';
      btn.setAttribute('aria-label', 'Unmute video');
      btn.setAttribute('title', 'Unmute video');
    } else {
      iconOff.style.display = 'none';
      iconOn.style.display = '';
      btn.setAttribute('aria-label', 'Mute video');
      btn.setAttribute('title', 'Mute video');
    }
  });
}

// ────────────────────────────────────────────────────────────
// 3D TILT
// ────────────────────────────────────────────────────────────
function init3DTilt() {
  document.querySelectorAll('.tilt-3d').forEach(card => {
    card.addEventListener('mousemove', e => {
      const rect = card.getBoundingClientRect();
      const rX = ((e.clientY - rect.top - rect.height/2) / (rect.height/2)) * -8;
      const rY = ((e.clientX - rect.left - rect.width/2) / (rect.width/2)) * 8;
      card.style.transform = `perspective(800px) rotateX(${rX}deg) rotateY(${rY}deg) scale3d(1.03,1.03,1.03)`;
    });
    card.addEventListener('mouseleave', () => {
      card.style.transform = 'perspective(800px) rotateX(0) rotateY(0) scale3d(1,1,1)';
    });
  });
}

// ────────────────────────────────────────────────────────────
// CURSOR FOLLOWER
// ────────────────────────────────────────────────────────────
function initCursorFollower() {
  if (window.matchMedia('(hover: none)').matches) return;
  const follower = document.createElement('div');
  follower.id = 'cursor-follower';
  follower.style.cssText = 'position:fixed;pointer-events:none;z-index:9998;width:20px;height:20px;border-radius:50%;border:1.5px solid var(--gold);transition:width 0.3s,height 0.3s,border-radius 0.3s,opacity 0.3s,background 0.3s;transform:translate(-50%,-50%);opacity:0;mix-blend-mode:screen;';
  document.body.appendChild(follower);
  const glow = document.createElement('div');
  glow.id = 'cursor-glow';
  glow.style.cssText = 'position:fixed;pointer-events:none;z-index:9997;width:200px;height:200px;border-radius:50%;background:radial-gradient(circle,rgba(212,175,55,0.06) 0%,transparent 70%);transform:translate(-50%,-50%);transition:opacity 0.5s;opacity:0;';
  document.body.appendChild(glow);
  let mouseX=0,mouseY=0,fx=0,fy=0;
  document.addEventListener('mousemove', e => {
    mouseX=e.clientX; mouseY=e.clientY;
    follower.style.opacity='1'; glow.style.opacity='1';
    const el = document.elementFromPoint(e.clientX, e.clientY);
    if (el?.closest('a, button, .tilt-3d, .portfolio-card, .faq-item')) {
      follower.style.width='50px'; follower.style.height='50px'; follower.style.background='rgba(212,175,55,0.08)';
    } else {
      follower.style.width='20px'; follower.style.height='20px'; follower.style.background='transparent';
    }
  });
  document.addEventListener('mouseleave', () => { follower.style.opacity='0'; glow.style.opacity='0'; });
  (function animate() {
    fx+=(mouseX-fx)*0.15; fy+=(mouseY-fy)*0.15;
    follower.style.left=fx+'px'; follower.style.top=fy+'px';
    glow.style.left=mouseX+'px'; glow.style.top=mouseY+'px';
    requestAnimationFrame(animate);
  })();
}

// ────────────────────────────────────────────────────────────
// ────────────────────────────────────────────────────────────
// REPRINT COST CALCULATOR
// ────────────────────────────────────────────────────────────

// Default placeholder values for initial calculation
const CALC_DEFAULTS = {
  qty: 8,
  cost: 30,
  reprints: 6,
  seasonalCount: 4,
  seasonalQty: 8,
  seasonalCost: 30
};

function calcUpdate() {
  const qtyEl = document.getElementById('calc-qty');
  const costEl = document.getElementById('calc-cost');
  const reprintsEl = document.getElementById('calc-reprints');
  const sCountEl = document.getElementById('calc-seasonal-count');
  const sQtyEl = document.getElementById('calc-seasonal-qty');
  const sCostEl = document.getElementById('calc-seasonal-cost');

  // Check if ANY field has user-entered data
  const anyFieldHasData = [qtyEl, costEl, reprintsEl, sCountEl, sQtyEl, sCostEl]
    .some(el => el && el.value !== '');

  let qty, cost, reprints, sCount, sQty, sCost;

  if (anyFieldHasData) {
    // If user started entering data, empty fields = 0
    qty = qtyEl?.value ? parseFloat(qtyEl.value) : 0;
    cost = costEl?.value ? parseFloat(costEl.value) : 0;
    reprints = reprintsEl?.value ? parseFloat(reprintsEl.value) : 0;
    sCount = sCountEl?.value ? parseFloat(sCountEl.value) : 0;
    sQty = sQtyEl?.value ? parseFloat(sQtyEl.value) : 0;
    sCost = sCostEl?.value ? parseFloat(sCostEl.value) : 0;
  } else {
    // All fields empty — use placeholder defaults
    qty = CALC_DEFAULTS.qty;
    cost = CALC_DEFAULTS.cost;
    reprints = CALC_DEFAULTS.reprints;
    sCount = CALC_DEFAULTS.seasonalCount;
    sQty = CALC_DEFAULTS.seasonalQty;
    sCost = CALC_DEFAULTS.seasonalCost;
  }

  const yearlyStandard = qty * cost * reprints;
  const yearlySeasonal = sCount * sQty * sCost;
  const yearly = yearlyStandard + yearlySeasonal;
  const fiveYear = yearly * 5;

  const resultsEl = document.getElementById('calc-results');
  const closingEl = document.getElementById('calc-closing');

  // Always show results
  resultsEl.style.display = 'flex';
  closingEl.style.display = 'block';
  document.getElementById('calc-yearly').textContent = '$' + yearly.toLocaleString();
  document.getElementById('calc-5year').textContent = '$' + fiveYear.toLocaleString();
}

function initCalc() {
  const inputs = document.querySelectorAll('#reprint-calc input');
  inputs.forEach(inp => inp.addEventListener('input', calcUpdate));
  // Run initial calculation with placeholder defaults
  calcUpdate();
}

// ────────────────────────────────────────────────────────────
// IMAGE LIGHTBOX — fullscreen zoom overlay
// ────────────────────────────────────────────────────────────
let lightboxOverlay = null;

function createLightboxOverlay() {
  if (lightboxOverlay) return;
  lightboxOverlay = document.createElement('div');
  lightboxOverlay.className = 'img-lightbox-overlay';
  lightboxOverlay.innerHTML = `
    <button class="img-lightbox-close" aria-label="Close">✕</button>
    <img class="img-lightbox-img" src="" alt="Zoomed image">
  `;
  document.body.appendChild(lightboxOverlay);

  const closeBtn = lightboxOverlay.querySelector('.img-lightbox-close');
  const img = lightboxOverlay.querySelector('.img-lightbox-img');

  // Close on X button
  closeBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    closeLightbox();
  });

  // Close on backdrop click (outside image)
  lightboxOverlay.addEventListener('click', (e) => {
    if (e.target === lightboxOverlay) closeLightbox();
  });

  // Prevent closing when clicking on the image itself
  img.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  // Close on Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && lightboxOverlay.classList.contains('active')) {
      closeLightbox();
    }
  });

  // Close on Android back button (popstate)
  window.addEventListener('popstate', (e) => {
    if (lightboxOverlay.classList.contains('active')) {
      closeLightboxInternal(true);
    }
  });
}

function openLightbox(src) {
  createLightboxOverlay();
  const img = lightboxOverlay.querySelector('.img-lightbox-img');
  img.src = src;
  img.classList.remove('zoomed');

  // Push history state so Android back button can close it
  history.pushState({ lightbox: true }, '');

  lightboxOverlay.style.display = 'flex';
  requestAnimationFrame(() => {
    lightboxOverlay.classList.add('active');
  });
  document.body.style.overflow = 'hidden';
}

function closeLightboxInternal(fromBackBtn) {
  if (!lightboxOverlay || !lightboxOverlay.classList.contains('active')) return;
  lightboxOverlay.classList.remove('active');

  // Only call history.back() if NOT triggered by back button (to avoid loop)
  if (!fromBackBtn && history.state && history.state.lightbox) {
    history.back();
  }

  setTimeout(() => {
    lightboxOverlay.style.display = 'none';
    lightboxOverlay.querySelector('.img-lightbox-img').src = '';
  }, 300);
  document.body.style.overflow = '';
}

function closeLightbox() {
  closeLightboxInternal(false);
}

function initImageLightbox() {
  // Use event delegation on body for maximum reliability across mobile + desktop
  // This catches clicks on dynamically rendered images too
  if (document._lightboxDelegated) return;
  document._lightboxDelegated = true;

  document.body.addEventListener('click', (e) => {
    const img = e.target.closest('.il-img-box img');
    if (img && img.src && img.complete && img.naturalWidth > 0) {
      e.preventDefault();
      e.stopPropagation();
      openLightbox(img.src);
    }
  });
}

// ────────────────────────────────────────────────────────────
// MAIN INIT
// ────────────────────────────────────────────────────────────
async function init() {
  initIntroSplash();
  initNavbar();
  initFaq();
  initMultiStepForm();
  initSmoothScroll();
  observeReveal();
  animateCounters();
  init3DTilt();
  initCursorFollower();
  initCalc();

  try {
    const portfolioItems = await loadPortfolio();
    renderPortfolio(portfolioItems);
    const baImages = await loadImageLibrary();
    renderBeforeAfter(baImages);
    init3DTilt();
    observeReveal();
    initImageLightbox();
  } catch (err) {
    console.warn('Data loading error:', err.message);
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
