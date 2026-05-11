# ODAX Estoque Web

Versão web separada do ODAX Estoque, adaptada para navegador.

## Como rodar localmente

```powershell
npm install
npm run dev
```

Abra o endereço que aparecer no terminal.

## Como gerar para publicar

```powershell
npm run build
```

A pasta `dist/` será gerada. Você pode postar essa pasta no Vercel, Netlify, Hostinger ou outro servidor.

## Observações

- Importação de Excel funciona pelo navegador.
- Exportação gera download `.xlsx`.
- Salvamento automático usa `localStorage` do navegador.


## Novo recurso

Na tela de contagem, clique em `Item X / Total` para escolher manualmente o número do item.
