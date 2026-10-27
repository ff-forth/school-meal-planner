// Skolmat → JSON  (bookmarklet)
//
// Spara som ett bokmärke:
//   1. Högerklicka på bokmärkesfältet → Lägg till bokmärke / Ny bokmärke.
//   2. Namn:  Skolmat → JSON
//   3. URL:   klistra in HELA raden nedan (en enda rad, från "javascript:" till "})();").
//
// Användning:
//   Öppna en matsedelsida på skolmat.info (adressen ska innehålla facilityId=...),
//   klicka på bokmärket, svara på vecka/år/antal, och filen laddas ner som
//   skolmat-<år>-v<start>-v<slut>-<skolanamn>.json. Importera den i planeraren.

javascript:(function(){var m=location.search.match(/facilityId=([A-Za-z0-9-]+)/);var id=m?m[1]:null;if(!id){alert('Öppna en matsedelsida med facilityId i adressen.');return;}function isoW(d){var t=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));t.setUTCDate(t.getUTCDate()+4-(t.getUTCDay()||7));return{y:t.getUTCFullYear(),w:Math.ceil(((t-new Date(Date.UTC(t.getUTCFullYear(),0,1)))/864e5+1)/7)};}var now=isoW(new Date());var sw=prompt('Startvecka (ISO, t.ex. 38):',now.w);if(sw===null)return;var sy=prompt('År:',now.y);if(sy===null)return;var n=prompt('Antal veckor (standard 8):','8');if(n===null)return;sw=parseInt(sw,10);sy=parseInt(sy,10);n=parseInt(n,10);if(!sw||!sy||!n){alert('Ogiltigt värde.');return;}var base=document.createElement('div');base.style.cssText='position:fixed;top:0;left:0;right:0;z-index:99999;background:#fff;color:#000;font:13px sans-serif;padding:8px;box-shadow:0 2px 6px rgba(0,0,0,.3)';base.textContent='Hämtar '+n+' veckor...';document.body.appendChild(base);var out=[];(async function(){for(var i=0;i<n;i++){var url='/api/public/matsedlar/'+encodeURIComponent(id)+'?year='+sy+'&week='+(sw+i);try{var r=await fetch(url,{headers:{'Accept':'application/json'}});var j=await r.json();if(j&&j.ok)out.push(j);}catch(e){}}base.remove();var nm=(out.length&&out[0].facility&&out[0].facility.name)||'';var slug=nm.replace(/[^a-z0-9]+/gi,'-').replace(/^-|-$/g,'').toLowerCase();var blob=new Blob([JSON.stringify(out,null,2)],{type:'application/json'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='skolmat-'+sy+'-v'+sw+'-v'+(sw+n-1)+(slug?'-'+slug:'')+'.json';document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(a.href);},1000);alert('Klart! Hamtade '+out.length+' av '+n+' veckor. Spara filen och importera den i planeraren.');})();})();
