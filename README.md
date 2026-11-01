# Wrapnation Courier Tracker

Shopify orders → local courier ka hisaab → Google Sheet (database).
Google Apps Script par chalta hai, free hai, alag server nahi chahiye.

## Kya karta hai
- Shopify se orders fetch (pichhle 30 din), duplicate nahi banta
- Status: New / Asked / Delivered / Returned, courier ka naam, phone, note
- Delivered par amount mila + courier charge (default 250) = Net
- Sab kuch Google Sheet ke `Orders` tab mein save hota hai
- Roz raat 11 baje khud sync + `Daily` tab mein roz ka hisaab

## Setup
1. Naya Google Sheet banayen. **Extensions → Apps Script**.
2. `src/Code.gs`, `src/Index.html`, `src/appsscript.json` ka content wahan daalen
   (ya `clasp` use karein, neeche dekhein).
3. Shopify admin: **Settings → Apps and sales channels → Develop apps**, nayi app,
   scope `read_orders`, install, Admin API access token copy karein.
4. Apps Script **Project Settings → Script Properties**: `SHOP` = `yourstore.myshopify.com`,
   `TOKEN` = `shpat_...` (token kabhi git mein na daalein).
5. Function `createTriggers` ek baar run karein (roz raat 11 baje auto sync).
6. **Deploy → New deployment → Web app**, Execute as: Me, Access: Only myself.
   Link bookmark kar lein.

## clasp se (git ke saath)
```
npm i -g @google/clasp
clasp login
clasp create --type sheets --title "Wrapnation Courier" --rootDir src
clasp push
```

## Git par push
```
git init
git add .
git commit -m "Wrapnation courier tracker"
git branch -M main
git remote add origin https://github.com/<username>/wrapnation-courier.git
git push -u origin main
```
