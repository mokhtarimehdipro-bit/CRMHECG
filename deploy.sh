#!/usr/bin/env bash
# =============================================================================
#  deploy.sh — mise en production complète du CRM HECG, sans action manuelle
#
#    bash deploy.sh "Message de version"
#
#  1. Sauvegarde du dossier (hors .git) dans ../_backups/<horodatage>
#  2. Commit de toutes les modifications locales
#  3. Récupère GitHub (uploads faits depuis le web) et fusionne ;
#     en cas de conflit : arrêt propre, rien n'est perdu (tout est commité)
#  4. git push  → GitHub Pages (frontend .html)
#  5. clasp push + clasp deploy sur le déploiement de PRODUCTION
#     → même URL /exec, nouvelle version (plus besoin de passer par l'éditeur)
# =============================================================================
set -euo pipefail
cd "$(dirname "$0")"

PROD_DEPLOYMENT_ID="AKfycbwfwxfVa2968AgKLdwwV8VIazVACVPVypKXrEUTc-zTuQ08gda8m5_qPfC2BgMyJMwb"
MSG="${1:-Mise à jour CRM $(date '+%Y-%m-%d %H:%M')}"
ONLY="${2:-all}"   # all | front | back

log() { printf '\n\033[1;34m▶ %s\033[0m\n' "$*"; }

# 1. Sauvegarde
BK="../_backups/$(date +%Y%m%d-%H%M%S)"
mkdir -p "$BK"
tar --exclude=.git -cf - . | (cd "$BK" && tar xf -)
log "Sauvegarde : $BK"
# on ne garde que les 15 dernières sauvegardes
ls -1d ../_backups/*/ 2>/dev/null | sort | head -n -15 | xargs -r rm -rf

if [ "$ONLY" != "back" ]; then
  # 2. Commit local
  git add -A
  if ! git diff --cached --quiet; then
    git commit -q -m "$MSG"
    log "Commit : $MSG"
  else
    log "Aucune modification à commiter"
  fi

  # 3. Synchronisation avec GitHub
  git fetch -q origin main
  if ! git merge-base --is-ancestor origin/main HEAD; then
    log "GitHub contient des changements faits depuis le web → fusion"
    if ! git merge -q --no-edit origin/main; then
      git merge --abort || true
      echo "❌ Conflit de fusion avec GitHub. Rien n'a été publié ni perdu (commit local intact)."
      exit 2
    fi
  fi

  # 4. Publication GitHub Pages
  git push -q origin HEAD:main
  log "GitHub Pages : poussé ($(git rev-parse --short HEAD))"
fi

if [ "$ONLY" != "front" ]; then
  # 5. Apps Script
  clasp push -f
  VERSION=$(clasp create-version "$MSG" 2>&1 | grep -o '[0-9]\+' | tail -1)
  log "Apps Script : version $VERSION créée"
  if ! OUT=$(clasp update-deployment "$PROD_DEPLOYMENT_ID" -V "$VERSION" -d "$MSG" 2>&1) || echo "$OUT" | grep -qi "only users\|error"; then
    echo "$OUT"
    echo "❌ Mise en production refusée par Google : clasp n'est pas connecté avec le compte propriétaire du script."
    echo "   → une seule fois : clasp login (choisir le compte propriétaire), puis relancer : bash deploy.sh \"$MSG\" back"
    exit 3
  fi
  log "Apps Script : déploiement de production → version $VERSION"
fi

log "✅ Terminé"
