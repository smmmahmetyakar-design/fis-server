#!/bin/bash
# Fiş Aktarım Aracı — sunucuda güncelleme scripti.
# Kullanım: cd ~/fis-server && bash guncelle.sh
set -e
cd "$(dirname "$0")"

echo "== git pull =="
git pull --ff-only

echo "== docker build =="
docker build -t fis-server-fis:latest .

echo "== container yeniden başlatılıyor =="
docker stop fis-otomasyon 2>/dev/null || true
docker rm fis-otomasyon 2>/dev/null || true
docker run -d --name fis-otomasyon -p 8091:8091 \
  -v /home/ahmet/fis-server/data:/srv/data \
  -v e6f12da3f5acc64d7ea20aac21055891ca1f82efc4121ccc1fc5bfd5cd5db59a:/data \
  -e FIS_DATA=/data \
  fis-server-fis:latest

echo "== bitti =="
docker ps --filter name=fis-otomasyon
