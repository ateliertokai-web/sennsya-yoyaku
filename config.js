/* 洗車予約の設定(お客様画面・管理画面・従業員画面で共通)
   api            … サーバー(Google Apps Script)のウェブアプリURL
   googleClientId … 管理画面・従業員画面の Google ログイン用 OAuth クライアントID。
                    空欄のあいだは、これまでどおりパスワードでログインします。
                    入れるときは、サーバーのスクリプトプロパティ GOOGLE_CLIENT_ID にも同じ値を入れてください。
   domain         … 会社のドメイン(例: example.co.jp)。ログイン画面で会社のアカウントを優先して出します */
window.WASH_CONFIG = {
  api: 'https://script.google.com/macros/s/AKfycbzHj9e15-2k4kkNa63ctM0P2KBXVbKfzGJGvhBNKXl0Er2u5_75qR7Kmev8Z0TzTdcg/exec',
  googleClientId: '',
  domain: ''
};
