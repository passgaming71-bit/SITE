ODAX Estoque - PWA aplicado

Incluído:
- public/manifest.json
- public/icon-192.png
- public/icon-512.png
- public/apple-touch-icon.png
- public/sw.js
- registro do Service Worker em src/main.jsx
- tags PWA no index.html

Como testar:
npm install
npm run dev

Como gerar para publicar:
npm run build

Depois publique normalmente na Vercel.

No celular:
- Android/Chrome: abrir o site e tocar em "Instalar app" ou "Adicionar à tela inicial".
- iPhone/Safari: compartilhar > "Adicionar à Tela de Início".

Observação:
O PWA precisa estar publicado em HTTPS para instalação funcionar corretamente. Vercel já fornece HTTPS.
