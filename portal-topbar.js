/**
 * TV Europe Sales Management Portal - Unified Smart TopBar Component
 */
(function () {
  const isIframe = (function () {
    try {
      return window.self !== window.top;
    } catch (e) {
      return true;
    }
  })();

  if (isIframe) {
    document.addEventListener('DOMContentLoaded', function () {
      const legacyBtns = document.querySelectorAll('.legacy-portal-return, #btn-return-portal-local');
      legacyBtns.forEach(el => el.style.display = 'none');
    });
    return;
  }

  function initUnifiedTopBar() {
    if (!document.getElementById('portal-unified-fonts')) {
      const fontLink = document.createElement('link');
      fontLink.id = 'portal-unified-fonts';
      fontLink.rel = 'stylesheet';
      fontLink.href = 'https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap';
      document.head.appendChild(fontLink);
    }

    if (!document.getElementById('portal-unified-icons')) {
      const iconLink = document.createElement('link');
      iconLink.id = 'portal-unified-icons';
      iconLink.rel = 'stylesheet';
      iconLink.href = 'https://cdn.jsdelivr.net/npm/remixicon@3.5.0/fonts/remixicon.css';
      document.head.appendChild(iconLink);
    }

    let pageTitle = document.title || 'LGE Sales Dashboard';
    pageTitle = pageTitle.replace(/Executive Portal\s*\|\s*/i, '').replace(/Dashboard\s*\|\s*/i, '');

    const topBar = document.createElement('div');
    topBar.id = 'portal-unified-topbar';
    topBar.style.cssText = `
      position: sticky;
      top: 0;
      left: 0;
      right: 0;
      width: 100%;
      height: 56px;
      background-color: #0F172A;
      border-bottom: 1px solid #1E293B;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 24px;
      box-sizing: border-box;
      z-index: 99999;
      font-family: 'Inter', system-ui, -apple-system, sans-serif;
      box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
    `;

    let portalUrl = '../00. TV Europe Sales Management Portal/index.html';
    if (window.location.hostname.includes('web.app') || window.location.hostname.includes('firebase')) {
      portalUrl = 'https://tv-sales-portal-2026.web.app';
    }

    topBar.innerHTML = `
      <style>
        #portal-unified-topbar .portal-back-btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 7px 14px;
          background: rgba(255, 255, 255, 0.08);
          border: 1px solid rgba(255, 255, 255, 0.18);
          border-radius: 6px;
          color: #FFFFFF;
          font-family: 'Inter', sans-serif;
          font-size: 13.5px;
          font-weight: 500;
          text-decoration: none;
          cursor: pointer;
          transition: all 0.2s ease-in-out;
          line-height: 1;
        }
        #portal-unified-topbar .portal-back-btn:hover {
          background-color: #0D9488;
          border-color: #0D9488;
          color: #FFFFFF;
          box-shadow: 0 2px 8px rgba(13, 148, 136, 0.3);
          transform: translateY(-1px);
        }
        #portal-unified-topbar .portal-dash-title {
          font-family: 'IBM Plex Sans', 'Inter', sans-serif;
          font-size: 15px;
          font-weight: 600;
          color: #F8FAFC;
          display: flex;
          align-items: center;
          gap: 10px;
        }
        #portal-unified-topbar .portal-badge {
          background-color: rgba(13, 148, 136, 0.2);
          border: 1px solid rgba(13, 148, 136, 0.4);
          color: #2DD4BF;
          font-size: 11px;
          font-weight: 600;
          padding: 2px 8px;
          border-radius: 4px;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }
      </style>
      
      <div style="display: flex; align-items: center; gap: 16px;">
        <a href="${portalUrl}" class="portal-back-btn" id="btn-portal-back-action">
          <i class="ri-arrow-left-line" style="font-size: 16px;"></i>
          <span>포탈로 돌아가기</span>
        </a>
      </div>

      <div class="portal-dash-title">
        <span class="portal-badge">STANDALONE VIEW</span>
        <span>${pageTitle}</span>
      </div>
    `;

    const backBtn = topBar.querySelector('#btn-portal-back-action');
    backBtn.addEventListener('click', function (e) {
      if (window.history.length > 1 && document.referrer.includes('Portal')) {
        e.preventDefault();
        window.history.back();
      }
    });

    if (document.body) {
      document.body.insertBefore(topBar, document.body.firstChild);
    } else {
      document.addEventListener('DOMContentLoaded', function () {
        document.body.insertBefore(topBar, document.body.firstChild);
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initUnifiedTopBar);
  } else {
    initUnifiedTopBar();
  }
})();
