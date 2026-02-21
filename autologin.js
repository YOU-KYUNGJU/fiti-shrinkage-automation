// RDMS auto login helper (prompt version)
// Usage (browser console): run rdmsPromptLogin();

(function (global) {
  const LOGIN_URL_KEYWORD = 'rdms.fiti.re.kr';

  function ensurePage() {
    if (!location.href.includes(LOGIN_URL_KEYWORD)) {
      console.warn('현재 페이지가 RDMS가 아닙니다:', location.href);
    }
  }

  function queryLoginElements() {
    const idInput = document.querySelector('#userid_t');
    const pwInput = document.querySelector('#userpw_t');
    const form =
      (idInput && idInput.form) ||
      (pwInput && pwInput.form) ||
      document.querySelector('form#loginForm') ||
      document.querySelector('form[name="loginForm"]') ||
      document.querySelector('form');

    const loginButton =
      document.querySelector('a img[src*="button_login"]') ||
      document.querySelector('img[src*="button_login"]') ||
      document.querySelector('a[onclick*="login" i]') ||
      document.querySelector('button[type="submit"], input[type="submit"]');

    return { idInput, pwInput, form, loginButton };
  }

  function fillInput(input, value) {
    input.focus();
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function submitLogin(form, loginButton) {
    if (loginButton) {
      const anchor = loginButton.closest ? loginButton.closest('a') : null;
      if (anchor) {
        anchor.click();
        return true;
      }
      loginButton.click();
      return true;
    }

    if (form) {
      const evt = new Event('submit', { bubbles: true, cancelable: true });
      form.dispatchEvent(evt);
      if (!evt.defaultPrevented && typeof form.submit === 'function') {
        form.submit();
      }
      return true;
    }

    return false;
  }

  function rdmsPromptLogin() {
    ensurePage();

    const id = prompt('RDMS ID를 입력하세요');
    if (id === null || id.trim() === '') {
      throw new Error('ID 입력이 취소되었거나 비어 있습니다.');
    }

    const pw = prompt('RDMS PW를 입력하세요');
    if (pw === null || pw.trim() === '') {
      throw new Error('PW 입력이 취소되었거나 비어 있습니다.');
    }

    const { idInput, pwInput, form, loginButton } = queryLoginElements();

    if (!idInput || !pwInput) {
      throw new Error('로그인 입력 필드(#userid_t, #userpw_t)를 찾지 못했습니다.');
    }

    fillInput(idInput, id.trim());
    fillInput(pwInput, pw);

    const submitted = submitLogin(form, loginButton);
    if (!submitted) {
      throw new Error('로그인 버튼 또는 폼을 찾지 못했습니다.');
    }

    return true;
  }

  global.rdmsPromptLogin = rdmsPromptLogin;
})(window);
