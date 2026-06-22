<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Connexion</title>
</head>
<body>
  <h1>Connexion</h1>

  <input type="email" id="email" placeholder="Email">
  <input type="password" id="password" placeholder="Mot de passe">
  <button onclick="login()">Entrer</button>

  <p id="msg"></p>

  <script>
    async function login() {
      const res = await fetch('/api/login', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
          email: document.getElementById('email').value,
          password: document.getElementById('password').value
        })
      });

      if (res.ok) {
        window.location.href = '/';
      } else {
        document.getElementById('msg').innerText = 'Email ou mot de passe incorrect';
      }
    }
  </script>
</body>
</html>
