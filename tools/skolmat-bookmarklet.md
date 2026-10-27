# Hämta rätter från skolmat.info som JSON

Planeraren kan importera aktuella rätter från [skolmat.info](https://www.skolmat.info) via en
`.json`-fil. Webbläsaren kan inte hämta API:t direkt från planeraren (CORS), så vi använder ett
**bokmärke (bookmarklet)** som körs *på* skolmat.info-sidan. Det hämtar valda veckor och
sparar en `.json`-fil som du sedan importerar i planeraren.

## 1. Skapa bokmärket

1. Högerklicka på din bokmärkesfält (eller "Bokmärken") → **Lägg till bokmärke** / **Ny bokmärke**.
2. Namn: `Skolmat → JSON`
3. I **URL**-fältet: klistra in raden nedan (hela raden, som en enda rad):

```
javascript:(function(){var m=location.search.match(/facilityId=([A-Za-z0-9-]+)/);var id=m?m[1]:null;if(!id){alert('Öppna en matsedelsida med facilityId i adressen.');return;}function isoW(d){var t=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));t.setUTCDate(t.getUTCDate()+4-(t.getUTCDay()||7));return{y:t.getUTCFullYear(),w:Math.ceil(((t-new Date(Date.UTC(t.getUTCFullYear(),0,1)))/864e5+1)/7)};}var now=isoW(new Date());var sw=prompt('Startvecka (ISO, t.ex. 38):',now.w);if(sw===null)return;var sy=prompt('År:',now.y);if(sy===null)return;var n=prompt('Antal veckor (standard 8):','8');if(n===null)return;sw=parseInt(sw,10);sy=parseInt(sy,10);n=parseInt(n,10);if(!sw||!sy||!n){alert('Ogiltigt värde.');return;}var base=document.createElement('div');base.style.cssText='position:fixed;top:0;left:0;right:0;z-index:99999;background:#fff;color:#000;font:13px sans-serif;padding:8px;box-shadow:0 2px 6px rgba(0,0,0,.3)';base.textContent='Hämtar '+n+' veckor...';document.body.appendChild(base);var out=[];(async function(){for(var i=0;i<n;i++){var url='/api/public/matsedlar/'+encodeURIComponent(id)+'?year='+sy+'&week='+(sw+i);try{var r=await fetch(url,{headers:{'Accept':'application/json'}});var j=await r.json();if(j&&j.ok)out.push(j);}catch(e){}}base.remove();var blob=new Blob([JSON.stringify(out,null,2)],{type:'application/json'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='skolmat-'+sy+'-v'+sw+'-v'+(sw+n-1)+'.json';document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(a.href);},1000);alert('Klart! Hamtade '+out.length+' av '+n+' veckor. Spara filen och importera den i planeraren.');})();})();
```

## 2. Hämta rätter

1. Öppna matsedeln på [skolmat.info](https://www.skolmat.info) – adressen måste innehålla `facilityId=...`
   (t.ex. `https://www.skolmat.info/matsedlar?countyCode=10&municipalityId=1081&facilityId=1C59B6EA-...`).
2. Klicka på bokmärket **Skolmat → JSON**.
3. Svara på frågorna:
   - **Startvecka (ISO)** – veckan som planen ska börja i (standard = nuvarande vecka).
   - **År** – t.ex. `2026`.
   - **Antal veckor** – standard `8`, ändra fritt.
4. En fil `skolmat-<år>-v<start>-v<slut>.json` laddas ner automatiskt.

## 3. Importera i planeraren

I **steg 1 (Data)** klicka på **📥 Importera matsedel (JSON)** och välj filen.
Planeraren bygger då måltidspoolen ur de aktuella rätterna (vanlig rätt + vegetariskt alternativ
per dag) och du kan generera 8-veckorsplanen som vanligt.

## Detaljer

- Hämtar `/api/public/matsedlar/{facilityId}?year=YYYY&week=NN` för varje vecka (publikt API, ingen inloggning).
- Veckor räknas som **ISO-veckor** (samma som vecka.nu) – samma numrering som i planeraren.
- Filen innehåller ett objekt per vecka: `{ ok, facility, week: { year, weekNumber, days: { MONDAY: { isOpen, rows: [{ text, allergens }] } } } }`.
- Om en vecka saknas (t.ex. framtida vecka som inte är publicerad) hoppas den över – bokmärket
  rapporterar hur många av veckorna som hämtades.
