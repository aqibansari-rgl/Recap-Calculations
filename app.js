/**
 * Renaissance Global Limited — Automations Portal Logic
 * Single-page login handling, validation, password toggle, and interactive workspace state.
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM Elements
  const loginForm = document.getElementById('login-form');
  const userEmailInput = document.getElementById('user-email');
  const userPasswordInput = document.getElementById('user-password');
  const togglePasswordBtn = document.getElementById('toggle-password-btn');
  const eyeIcon = document.getElementById('eye-icon');
  const emailError = document.getElementById('email-error');
  const passwordError = document.getElementById('password-error');
  const loginSubmitBtn = document.getElementById('login-submit-btn');
  const btnSpinner = document.getElementById('btn-spinner');
  const btnText = document.getElementById('btn-text');
  const btnArrow = document.getElementById('btn-arrow');
  const demoFillBtn = document.getElementById('demo-fill-btn');
  const forgotPasswordBtn = document.getElementById('forgot-password-btn');

  // Views
  const loginFormView = document.getElementById('login-form-view');
  const authenticatedView = document.getElementById('authenticated-view');
  const loggedUserEmail = document.getElementById('logged-user-email');
  const launchHubBtn = document.getElementById('launch-hub-btn');
  const logoutBtn = document.getElementById('logout-btn');
  const toastContainer = document.getElementById('toast-container');

  // --- Toggle Password Visibility ---
  if (togglePasswordBtn) {
    togglePasswordBtn.addEventListener('click', () => {
      const isPassword = userPasswordInput.type === 'password';
      userPasswordInput.type = isPassword ? 'text' : 'password';

      if (isPassword) {
        // Eye-off icon
        eyeIcon.innerHTML = `
          <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path>
          <line x1="1" y1="1" x2="23" y2="23"></line>
        `;
      } else {
        // Normal eye icon
        eyeIcon.innerHTML = `
          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
          <circle cx="12" cy="12" r="3"></circle>
        `;
      }
    });
  }

  // --- Demo Credentials Autofill ---
  if (demoFillBtn) {
    demoFillBtn.addEventListener('click', () => {
      userEmailInput.value = 'operations@renaissanceglobal.com';
      userPasswordInput.value = 'RG-Automations2026';
      clearErrors();
      showToast('Demo credentials populated.', 'tan');
      userEmailInput.focus();
    });
  }

  // --- Input Validation Helpers ---
  function validateEmail(email) {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(String(email).toLowerCase());
  }

  function clearErrors() {
    userEmailInput.classList.remove('invalid');
    userPasswordInput.classList.remove('invalid');
    emailError.style.display = 'none';
    passwordError.style.display = 'none';
  }

  userEmailInput.addEventListener('input', () => {
    if (userEmailInput.classList.contains('invalid')) {
      if (validateEmail(userEmailInput.value)) {
        userEmailInput.classList.remove('invalid');
        emailError.style.display = 'none';
      }
    }
  });

  userPasswordInput.addEventListener('input', () => {
    if (userPasswordInput.classList.contains('invalid')) {
      if (userPasswordInput.value.length >= 6) {
        userPasswordInput.classList.remove('invalid');
        passwordError.style.display = 'none';
      }
    }
  });

  // --- Form Submission / Login ---
  loginForm.addEventListener('submit', (e) => {
    e.preventDefault();
    clearErrors();

    const emailVal = userEmailInput.value.trim();
    const passVal = userPasswordInput.value.trim();
    let hasError = false;

    if (!emailVal || !validateEmail(emailVal)) {
      userEmailInput.classList.add('invalid');
      emailError.style.display = 'block';
      hasError = true;
    }

    if (!passVal || passVal.length < 6) {
      userPasswordInput.classList.add('invalid');
      passwordError.style.display = 'block';
      hasError = true;
    }

    if (hasError) return;

    // Simulate Authentication Request
    loginSubmitBtn.disabled = true;
    btnSpinner.style.display = 'inline-block';
    btnArrow.style.display = 'none';
    btnText.textContent = 'Authenticating...';

    setTimeout(() => {
      loginSubmitBtn.disabled = false;
      btnSpinner.style.display = 'none';
      btnArrow.style.display = 'inline-flex';
      btnText.textContent = 'Sign In to Automations';

      showToast(`Authenticated! Redirecting to Summary Sheet Creation...`, 'green');

      setTimeout(() => {
        window.location.href = 'summary-sheet-creation.html';
      }, 700);
    }, 900);
  });

  // --- Logout / Sign Out ---
  if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
      authenticatedView.style.display = 'none';
      loginFormView.style.display = 'block';
      userPasswordInput.value = '';
      clearErrors();
      showToast('Signed out of automation session.', 'tan');
    });
  }

  // --- Launch Hub Action ---
  if (launchHubBtn) {
    launchHubBtn.addEventListener('click', () => {
      showToast('Launching active automation workspace: 4 pipelines connected.', 'green');
    });
  }

  // --- Forgot Password ---
  if (forgotPasswordBtn) {
    forgotPasswordBtn.addEventListener('click', () => {
      const email = userEmailInput.value.trim() || 'your registered work email';
      showToast(`Password recovery link sent to ${email}`, 'tan');
    });
  }

  // --- Toast Notification Helper ---
  function showToast(message, type = 'green') {
    const toast = document.createElement('div');
    toast.className = `toast ${type === 'tan' ? 'toast-tan' : ''}`;
    toast.innerHTML = `
      <svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
        <polyline points="22 4 12 14.01 9 11.01"></polyline>
      </svg>
      <span>${message}</span>
    `;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(30px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 3800);
  }
});
