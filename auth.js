/* 管理画面・従業員画面の Google ログイン(Google Identity Services)
   - ログインすると Google の ID トークンを受け取り、サーバーへの依頼に付けて送る(サーバーで本物か・会社のアカウントかを確かめる)
   - トークンは約1時間で切れる。切れたら自動で取り直し、だめならログイン画面に戻す
   - トークンはこのタブの中(sessionStorage)だけに置き、タブを閉じると消える */
(function () {
  const CONF = window.WASH_CONFIG || {}, KEY = 'wash_id_token';
  const A = { enabled: !!CONF.googleClientId, token: '', onSignIn: null };
  try { A.token = sessionStorage.getItem(KEY) || ''; } catch (e) {}
  let ready = null, waiters = [];
  const save = t => { A.token = t || ''; try { if (t) sessionStorage.setItem(KEY, t); else sessionStorage.removeItem(KEY); } catch (e) {} };
  A.load = () => ready || (ready = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client'; s.async = true;
    s.onload = () => {
      google.accounts.id.initialize({
        client_id: CONF.googleClientId, auto_select: true, cancel_on_tap_outside: false, use_fedcm_for_prompt: true,
        hd: CONF.domain || undefined,
        callback: r => {
          save(r.credential);
          const w = waiters; waiters = [];
          if (w.length) w.forEach(f => f()); else if (A.onSignIn) A.onSignIn();
        }
      });
      res();
    };
    s.onerror = () => { ready = null; rej(new Error('Google のログイン画面を読み込めませんでした。通信状態を確かめてください')); };
    document.head.appendChild(s);
  }));
  /** ログインボタンを表示する(前にログインした端末なら、自動でログインを試みる) */
  A.button = el => A.load().then(() => {
    el.innerHTML = '';
    google.accounts.id.renderButton(el, { theme: 'filled_black', size: 'large', shape: 'pill', text: 'signin_with', locale: 'ja', width: 260 });
    google.accounts.id.prompt();
  }).catch(e => { el.textContent = e.message; });
  /** トークンを取り直す(切れたとき)。取れなければ失敗 */
  A.refresh = () => A.load().then(() => new Promise((res, rej) => {
    const t = setTimeout(() => { waiters = waiters.filter(f => f !== done); rej(new Error('LOGIN')); }, 12000);
    const done = () => { clearTimeout(t); res(); };
    waiters.push(done);
    google.accounts.id.prompt();
  }));
  A.clear = () => save('');
  A.signOut = () => { save(''); try { google.accounts.id.disableAutoSelect(); } catch (e) {} };
  window.WashAuth = A;
})();
