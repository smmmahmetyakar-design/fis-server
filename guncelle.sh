#!/bin/bash
# fis-server güncelleme betiği
# Kullanım: bash /srv/apps/fis-server/guncelle.sh
#
# Sunucuya özel ayarlar (git'e girmez) — /srv/apps/fis-server/ayarlar.env :
#   OLLAMA_URL=http://host.docker.internal:11434   # Ollama adresi; boş bırakılırsa yapay zekâ kapalı
#   OLLAMA_MODEL=qwen2.5:14b

# Tüm betik bir fonksiyonda: bash önce tamamını okur, sonra çalıştırır.
# (Aşağıdaki git pull bu dosyanın kendisini değiştirdiğinde yarıda kalan
#  eski sürüm, yeni dosyanın ortasından okumaya devam etmesin diye.)
guncelle() {
set -e
cd /srv/apps/fis-server

# varsayılanlar; ayarlar.env varsa onlar geçerli
OLLAMA_URL="http://host.docker.internal:11434"
OLLAMA_MODEL="qwen2.5:14b"
[ -f ayarlar.env ] && . ./ayarlar.env

echo "== GitHub'dan son sürüm çekiliyor =="
sudo git pull origin main

echo "== Docker container yeniden başlatılıyor =="
sudo docker stop fis-otomasyon 2>/dev/null || true
sudo docker rm fis-otomasyon 2>/dev/null || true
sudo docker build -t fis-otomasyon .
sudo docker run -d --name fis-otomasyon --restart unless-stopped \
  -p 8091:8091 -v /srv/apps/fis-server/data:/data \
  -v /srv/veri/firmalar:/firmalar:ro \
  --add-host=host.docker.internal:host-gateway \
  -e OLLAMA_URL="$OLLAMA_URL" -e OLLAMA_MODEL="$OLLAMA_MODEL" \
  fis-otomasyon

echo "== Bitti =="
sudo docker ps | grep fis-otomasyon
sleep 3
echo "== Yapay zekâ bağlantısı =="
curl -s localhost:8091/api/yapay-zeka || echo "(araç henüz açılmadı, birkaç saniye sonra tekrar deneyin)"
echo
}
guncelle "$@"
exit $?
