/* Sécurité CRM HECG — chargé en premier par toutes les pages connectées.
   Depuis nov. 2026, le serveur (doPost) n'accepte plus que des appels portant un jeton de session
   valide (sauf login / publicRegister). Ce script ajoute automatiquement le jeton de l'utilisateur
   connecté (localStorage « user ».sessionToken) à chaque appel fetch vers l'Apps Script,
   sans avoir à modifier chaque appel des pages.
   Session expirée → retour à la page de connexion (sauf si window.AUTH_FETCH_NO_REDIRECT, ex. salon.html
   qui garde ses saisies hors ligne et affiche son propre bandeau). */
(function () {
  if (window.__hecgAuthFetch) return;
  window.__hecgAuthFetch = true;
  const nativeFetch = window.fetch.bind(window);
  let redirige = false;
  const estApi = url => typeof url === 'string' && url.indexOf('script.google.com/macros/') !== -1;

  window.fetch = function (input, init) {
    const url = typeof input === 'string' ? input : (input && input.url) || '';
    if (!estApi(url)) return nativeFetch(input, init);
    try {
      if (init && typeof init.body === 'string' && init.body.charAt(0) === '{') {
        const u = JSON.parse(localStorage.getItem('user') || '{}');
        if (u.sessionToken) {
          const b = JSON.parse(init.body);
          if (b && !b.sessionToken) { b.sessionToken = u.sessionToken; init = Object.assign({}, init, { body: JSON.stringify(b) }); }
        }
      }
    } catch (e) { /* corps non JSON : envoyé tel quel */ }
    return nativeFetch(input, init).then(function (res) {
      if (!window.AUTH_FETCH_NO_REDIRECT && !redirige) {
        res.clone().text().then(function (t) {
          if (t.length < 400 && t.indexOf('"sessionExpired":true') !== -1 && !redirige && !/index\.html$|\/$/.test(location.pathname)) {
            redirige = true;
            try { localStorage.removeItem('user'); } catch (e) {}
            alert('Votre session a expiré ou n\'est plus valide. Merci de vous reconnecter.');
            location.href = 'index.html';
          }
        }).catch(function () {});
      }
      return res;
    });
  };
})();
