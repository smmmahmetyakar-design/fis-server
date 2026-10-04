#!/bin/bash
# fis-server güncelleme betiği
# Kullanım: bash /srv/apps/fis-server/guncelle.sh

set -e
cd /srv/apps/fis-server

echo "== GitHub'dan son sürüm çekiliyor =="
sudo git pull origin main

echo "== Docker container yeniden başlatılıyor =="
sudo docker stop fis-otomasyon 2>/dev/null || true
sudo docker rm fis-otomasyon 2>/dev/null || true
sudo docker build -t fis-otomasyon .
sudo docker run -d --name fis-otomasyon --restart unless-stopped \
  -p 8091:8091 -v /srv/apps/fis-server/data:/data \
  -v /srv/veri/firmalar:/firmalar:ro \
  fis-otomasyon

echo "== Bitti =="
sudo docker ps | grep fis-otomasyon
