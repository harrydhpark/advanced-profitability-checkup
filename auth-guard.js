/**
 * LGE TV Europe/CIS Sub-Dashboard Unified Auth Guard & SSO Handler
 * - SHA-256 Credentials Verification (LGE135 / LGE246)
 * - 30-Minute Session Timeout with Activity Auto-Extension
 * - Portal SSO URL Token Parser & Address Bar Cleanup
 * - Guaranteed Body-First Injection & Initial Render Blocking
 */
(function() {
  const AUTH_CONFIG = {
    ID_HASH: '4ed89d4c95cd896421176fe47e4c9ee9a0baad02dae39f57ce32eef58ec8e942', // 'LGE135'
    PW_HASH: '618db43b60c434d96fa46606dfd55e64a4ee7321f05b6dd846f5b5bea2e7cade', // 'LGE246'
    SESSION_KEY: 'lge_portal_authenticated_user',
    TIMESTAMP_KEY: 'lge_portal_auth_timestamp',
    TIMEOUT_MS: 30 * 60 * 1000, // 30 minutes
    SECRET: 'LGE_TV_EU_PORTAL_SECRET_2026'
  };

  // Block page rendering immediately until auth status is determined
  if (document.documentElement) {
    document.documentElement.classList.add('ag-authenticating');
  }

  async function sha256(message) {
    const msgBuffer = new TextEncoder().encode(message);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }

  // Inject critical CSS
  const style = document.createElement('style');
  style.id = 'authGuardStyles';
  style.textContent = `
    html.ag-authenticating body {
      visibility: hidden !important;
      opacity: 0 !important;
    }
    .ag-login-overlay {
      position: fixed !important;
      top: 0 !important;
      left: 0 !important;
      width: 100vw !important;
      height: 100vh !important;
      background-color: #0f172a !important;
      background-image: 
        radial-gradient(at 0% 0%, rgba(13, 148, 136, 0.3) 0px, transparent 50%),
        radial-gradient(at 100% 100%, rgba(59, 130, 246, 0.2) 0px, transparent 50%) !important;
      z-index: 2147483647 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      padding: 20px !important;
      box-sizing: border-box !important;
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
      opacity: 1 !important;
      visibility: visible !important;
      transition: opacity 0.2s ease, visibility 0.2s ease !important;
    }
    .ag-login-overlay.hidden {
      opacity: 0 !important;
      visibility: hidden !important;
      pointer-events: none !important;
      display: none !important;
    }
    .ag-login-card {
      background: #ffffff !important;
      border-radius: 12px !important;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7) !important;
      padding: 36px 32px !important;
      width: 100% !important;
      max-width: 420px !important;
      box-sizing: border-box !important;
    }
    .ag-login-header {
      text-align: center !important;
      margin-bottom: 24px !important;
    }
    .ag-login-icon {
      width: 52px !important;
      height: 52px !important;
      border-radius: 14px !important;
      background: #0f172a !important;
      color: #0d9488 !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      font-size: 26px !important;
      margin: 0 auto 14px !important;
    }
    .ag-login-title {
      font-size: 20px !important;
      font-weight: 700 !important;
      color: #0f172a !important;
      margin: 0 0 6px !important;
    }
    .ag-login-subtitle {
      font-size: 13px !important;
      color: #64748b !important;
      margin: 0 !important;
    }
    .ag-input-group {
      margin-bottom: 16px !important;
      text-align: left !important;
    }
    .ag-input-group label {
      display: block !important;
      font-size: 12px !important;
      font-weight: 600 !important;
      color: #0f172a !important;
      margin-bottom: 6px !important;
    }
    .ag-input-wrapper {
      position: relative !important;
    }
    .ag-input-wrapper input {
      width: 100% !important;
      padding: 11px 14px !important;
      font-size: 14px !important;
      color: #0f172a !important;
      background: #f8fafc !important;
      border: 1px solid #cbd5e1 !important;
      border-radius: 8px !important;
      outline: none !important;
      box-sizing: border-box !important;
      transition: border-color 0.2s !important;
    }
    .ag-input-wrapper input:focus {
      border-color: #0d9488 !important;
      background: #ffffff !important;
    }
    .ag-error-msg {
      display: none;
      align-items: center;
      gap: 6px;
      padding: 8px 12px;
      background: #fef2f2;
      border: 1px solid #fecaca;
      border-radius: 6px;
      color: #dc2626;
      font-size: 12px;
      font-weight: 600;
      margin-bottom: 16px;
    }
    .ag-submit-btn {
      width: 100% !important;
      padding: 12px !important;
      background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%) !important;
      color: #ffffff !important;
      border: none !important;
      border-radius: 8px !important;
      font-size: 14px !important;
      font-weight: 600 !important;
      cursor: pointer !important;
      transition: transform 0.15s, background 0.2s !important;
    }
    .ag-submit-btn:hover {
      transform: translateY(-1px) !important;
    }
    .ag-footer-info {
      font-size: 11px !important;
      color: #94a3b8 !important;
      border-top: 1px solid #e2e8f0 !important;
      margin-top: 20px !important;
      padding-top: 14px !important;
      text-align: center !important;
    }
  `;

  if (document.head) {
    document.head.appendChild(style);
  } else {
    document.addEventListener('DOMContentLoaded', () => document.head.appendChild(style));
  }

  function injectOverlayHtml() {
    if (document.getElementById('authGuardLoginOverlay')) return;
    const overlay = document.createElement('div');
    overlay.className = 'ag-login-overlay';
    overlay.id = 'authGuardLoginOverlay';
    overlay.innerHTML = `
      <div class="ag-login-card">
        <div class="ag-login-header">
          <div class="ag-login-icon">🔒</div>
          <h2 class="ag-login-title">TV EU/CIS Dashboard System</h2>
          <p class="ag-login-subtitle">보안 대시보드 접속을 위해 인증 정보를 입력하세요.</p>
        </div>
        <form id="agLoginForm" autocomplete="off" onsubmit="return false;">
          <div class="ag-input-group">
            <label for="agUsernameInput">아이디 (ID)</label>
            <div class="ag-input-wrapper">
              <input type="text" id="agUsernameInput" placeholder="아이디 입력 (대소문자 구분)" required autocomplete="username">
            </div>
          </div>
          <div class="ag-input-group">
            <label for="agPasswordInput">비밀번호 (Password)</label>
            <div class="ag-input-wrapper">
              <input type="password" id="agPasswordInput" placeholder="비밀번호 입력 (대소문자 구분)" required autocomplete="current-password">
            </div>
          </div>
          <div id="agErrorMsg" class="ag-error-msg">
            <span>아이디 또는 비밀번호가 일치하지 않습니다.</span>
          </div>
          <button type="submit" id="agSubmitBtn" class="ag-submit-btn">대시보드 접속하기</button>
        </form>
        <div class="ag-footer-info">LGE TV EU/CIS Sales & Marketing Security Portal</div>
      </div>
    `;

    if (document.body) {
      document.body.prepend(overlay);
    } else {
      document.addEventListener('DOMContentLoaded', () => {
        if (document.body) document.body.prepend(overlay);
      });
    }

    setTimeout(() => {
      const form = document.getElementById('agLoginForm');
      if (form) form.addEventListener('submit', handleLogin);
    }, 0);
  }

  function updateAuthActivity() {
    const authUser = sessionStorage.getItem(AUTH_CONFIG.SESSION_KEY);
    if (authUser) {
      sessionStorage.setItem(AUTH_CONFIG.TIMESTAMP_KEY, Date.now().toString());
    }
  }

  function clearAuthSession() {
    sessionStorage.removeItem(AUTH_CONFIG.SESSION_KEY);
    sessionStorage.removeItem(AUTH_CONFIG.TIMESTAMP_KEY);
  }

  async function checkSsoUrlToken() {
    const params = new URLSearchParams(window.location.search);
    const token = params.get('auth_token');
    const tsStr = params.get('auth_ts');
    const user = params.get('auth_user');

    if (token && tsStr && user) {
      const ts = parseInt(tsStr, 10);
      const now = Date.now();
      if (!isNaN(ts) && Math.abs(now - ts) < 5 * 60 * 1000) {
        const expectedSig = await sha256(user + ":" + ts + ":" + AUTH_CONFIG.SECRET);
        if (expectedSig === token) {
          sessionStorage.setItem(AUTH_CONFIG.SESSION_KEY, user);
          sessionStorage.setItem(AUTH_CONFIG.TIMESTAMP_KEY, now.toString());

          const cleanUrl = window.location.pathname + window.location.hash;
          window.history.replaceState({}, document.title, cleanUrl);
          return true;
        }
      }
    }
    return false;
  }

  async function checkAuthStatus() {
    injectOverlayHtml();
    const ssoSuccess = await checkSsoUrlToken();
    const authUser = sessionStorage.getItem(AUTH_CONFIG.SESSION_KEY);
    const lastActivity = sessionStorage.getItem(AUTH_CONFIG.TIMESTAMP_KEY);
    const now = Date.now();

    const isValid = ssoSuccess || (authUser && lastActivity && (now - parseInt(lastActivity, 10) < AUTH_CONFIG.TIMEOUT_MS));

    const overlay = document.getElementById('authGuardLoginOverlay');
    if (isValid) {
      updateAuthActivity();
      if (overlay) overlay.classList.add('hidden');
    } else {
      clearAuthSession();
      if (overlay) overlay.classList.remove('hidden');
      const input = document.getElementById('agUsernameInput');
      if (input) setTimeout(() => input.focus(), 100);
    }

    // Unblock page rendering after auth evaluation
    if (document.documentElement) {
      document.documentElement.classList.remove('ag-authenticating');
    }
  }

  async function handleLogin(e) {
    if (e) e.preventDefault();
    const userInput = document.getElementById('agUsernameInput');
    const pwInput = document.getElementById('agPasswordInput');
    const errorMsg = document.getElementById('agErrorMsg');
    const submitBtn = document.getElementById('agSubmitBtn');

    if (!userInput || !pwInput) return;
    const enteredId = userInput.value.trim();
    const enteredPw = pwInput.value.trim();

    if (!enteredId || !enteredPw) return;
    if (submitBtn) submitBtn.disabled = true;

    try {
      const idHash = await sha256(enteredId);
      const pwHash = await sha256(enteredPw);

      if (idHash === AUTH_CONFIG.ID_HASH && pwHash === AUTH_CONFIG.PW_HASH) {
        sessionStorage.setItem(AUTH_CONFIG.SESSION_KEY, enteredId);
        sessionStorage.setItem(AUTH_CONFIG.TIMESTAMP_KEY, Date.now().toString());
        if (errorMsg) errorMsg.style.display = 'none';
        userInput.value = '';
        pwInput.value = '';
        checkAuthStatus();
      } else {
        if (errorMsg) errorMsg.style.display = 'flex';
        pwInput.value = '';
        pwInput.focus();
      }
    } catch (err) {
      console.error("Auth Error", err);
    } finally {
      if (submitBtn) submitBtn.disabled = false;
    }
  }

  let activityThrottleTimer = null;
  function setupActivityListeners() {
    const events = ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'];
    events.forEach(evt => {
      window.addEventListener(evt, () => {
        if (!activityThrottleTimer) {
          updateAuthActivity();
          activityThrottleTimer = setTimeout(() => {
            activityThrottleTimer = null;
          }, 10000);
        }
      }, { passive: true });
    });

    setInterval(checkAuthStatus, 60000);
  }

  function init() {
    setupActivityListeners();
    checkAuthStatus();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
