#!/bin/bash
# keep_alive.sh — Impulse Academy
# Lance Flask + tunnel Cloudflare avec auto-restart permanent

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
LOG_DIR="$APP_DIR/logs"
mkdir -p "$LOG_DIR"

FLASK_LOG="$LOG_DIR/flask.log"
TUNNEL_LOG="$LOG_DIR/tunnel.log"
URL_FILE="$LOG_DIR/tunnel_url.txt"

echo "=== Impulse Academy — Démarrage permanent ===" | tee "$FLASK_LOG"
echo "Répertoire : $APP_DIR"

# Tuer les processus précédents proprement
pkill -f "python3 app.py" 2>/dev/null
pkill -f "cloudflared tunnel" 2>/dev/null
sleep 1

# Boucle Flask avec auto-restart
restart_flask() {
  while true; do
    echo "[$(date '+%H:%M:%S')] Flask démarré" >> "$FLASK_LOG"
    cd "$APP_DIR" && python3 app.py >> "$FLASK_LOG" 2>&1
    echo "[$(date '+%H:%M:%S')] Flask arrêté — redémarrage dans 3s…" >> "$FLASK_LOG"
    sleep 3
  done
}

# Boucle tunnel avec auto-restart
restart_tunnel() {
  sleep 4  # Attendre que Flask soit up
  while true; do
    echo "[$(date '+%H:%M:%S')] Tunnel démarré" >> "$TUNNEL_LOG"
    cloudflared tunnel --url http://localhost:5000 >> "$TUNNEL_LOG" 2>&1 &
    TUNNEL_PID=$!

    # Attendre l'URL (max 30s)
    for i in $(seq 1 30); do
      sleep 1
      URL=$(grep -o 'https://[a-z0-9-]*\.trycloudflare\.com' "$TUNNEL_LOG" 2>/dev/null | tail -1)
      if [ -n "$URL" ]; then
        echo "$URL" > "$URL_FILE"
        echo "[$(date '+%H:%M:%S')] URL publique : $URL" | tee -a "$TUNNEL_LOG"
        break
      fi
    done

    wait $TUNNEL_PID
    echo "[$(date '+%H:%M:%S')] Tunnel arrêté — redémarrage dans 5s…" >> "$TUNNEL_LOG"
    sleep 5
  done
}

# Démarrer les deux boucles en arrière-plan
restart_flask &
FLASK_LOOP_PID=$!

restart_tunnel &
TUNNEL_LOOP_PID=$!

echo "Flask loop PID  : $FLASK_LOOP_PID"
echo "Tunnel loop PID : $TUNNEL_LOOP_PID"
echo "$FLASK_LOOP_PID $TUNNEL_LOOP_PID" > "$LOG_DIR/pids.txt"

echo ""
echo "Serveur démarré. URL publique dans : $URL_FILE"
echo "Logs Flask  : $FLASK_LOG"
echo "Logs Tunnel : $TUNNEL_LOG"
echo ""
echo "Pour arrêter : kill \$(cat $LOG_DIR/pids.txt)"
echo "Pour voir l'URL : cat $URL_FILE"

# Attendre l'URL et l'afficher
sleep 12
if [ -f "$URL_FILE" ]; then
  URL=$(cat "$URL_FILE")
  echo ""
  echo "╔══════════════════════════════════════════════════╗"
  echo "║  IMPULSE ACADEMY — Accès public                  ║"
  echo "║  $URL  ║"
  echo "╚══════════════════════════════════════════════════╝"
fi

# Rester actif pour superviser
wait
