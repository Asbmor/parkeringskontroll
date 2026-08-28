const $=s=>document.querySelector(s);let db=[],stream=null,scanning=false,ocrBusy=false,scanTimer=null,worker=null,workerReady=false,workerPromise=null;
const norm=v=>String(v??'').toUpperCase().replace(/[^A-ZÆØÅ0-9]/g,'');
function save(){localStorage.setItem('parkingDb',JSON.stringify(db));renderStatus()}
function renderStatus(){$('#count').textContent=`${db.length} registreringsnummer`;$('#status').textContent=db.length?`Liste klar • ${db.length} biler`:'Ingen liste importert'}
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function formatReg(r){r=norm(r);return r.length>2?r.slice(0,2)+' '+r.slice(2):r}
function validNorwegian(r){return /^[A-ZÆØÅ]{2}[0-9]{5}$/.test(norm(r))}
function lookup(){const q=norm($('#q').value),box=$('#result');if(!q){box.className='result idle';box.innerHTML='<strong>Klar for oppslag</strong><span>Skann eller skriv inn et bilnummer.</span>';return}if(!validNorwegian(q)){box.className='result wait';const need=Math.max(0,7-q.length);box.innerHTML=`<strong>Venter på komplett skilt</strong><span>${esc(formatReg(q))}${need?' • mangler '+need+' tegn':''}</span>`;return}const hits=db.filter(x=>x.reg===q);if(hits.length){box.className='result ok';box.innerHTML=`<strong>✓ GODKJENT</strong><span>${esc(formatReg(q))} • ${hits.map(x=>esc(x.name)).join(', ')}</span>`}else{box.className='result no';box.innerHTML=`<strong>✕ IKKE FUNNET</strong><span>${esc(formatReg(q))}</span>`}}
$('#q').addEventListener('input',lookup);$('#clear').onclick=()=>{$('#q').value='';lookup();$('#q').focus()};$('#forget').onclick=()=>{if(confirm('Slette den lokalt lagrede parkeringslisten?')){db=[];localStorage.removeItem('parkingDb');renderStatus();lookup()}};
$('#file').addEventListener('change',async e=>{try{if(!e.target.files[0])return;const data=await e.target.files[0].arrayBuffer(),wb=XLSX.read(data),rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});if(!rows.length)throw Error('Arket er tomt');const keys=Object.keys(rows[0]),nk=keys.find(k=>norm(k)==='NAVN'),rk=keys.find(k=>['REGISTRERINGSNUMMER','BILNUMMER','REGNR','SKILT'].includes(norm(k)));if(!nk||!rk)throw Error('Fant ikke kolonnene Navn og Registreringsnummer');db=rows.map(r=>({name:String(r[nk]).trim(),reg:norm(r[rk])})).filter(x=>x.reg);save();$('#q').value='';lookup();alert(`Importerte ${db.length} registreringsnummer.`)}catch(err){alert('Kunne ikke importere: '+err.message)}finally{e.target.value=''}});
function setScanStatus(t){const s=$('#scanStatus');s.hidden=false;s.textContent=t}
async function initOCR(){if(workerReady)return worker;if(workerPromise)return workerPromise;workerPromise=(async()=>{try{setScanStatus('Klargjør skiltleser…');worker=await Tesseract.createWorker('eng',1,{logger:m=>{if(m.status==='loading tesseract core')setScanStatus('Klargjør skiltleser…');if(m.status==='recognizing text'&&scanning)setScanStatus(`Leser skilt… ${Math.round((m.progress||0)*100)} %`)}});await worker.setParameters({tessedit_char_whitelist:'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',preserve_interword_spaces:'0'});workerReady=true;setScanStatus('Skiltleser klar');return worker}catch(e){workerPromise=null;setScanStatus('Kunne ikke klargjøre OCR: '+e.message);throw e}})();return workerPromise}
function cropFromVideo(){const v=$('#video'),base=$('#crop');if(!v.videoWidth)return null;const vw=v.videoWidth,vh=v.videoHeight,cropW=Math.round(vw*.84),cropH=Math.round(cropW*.24),sx=Math.round((vw-cropW)/2),sy=Math.round((vh-cropH)/2);base.width=Math.max(1,cropW*2);base.height=Math.max(1,cropH*2);const x=base.getContext('2d',{willReadFrequently:true});x.imageSmoothingEnabled=true;x.drawImage(v,sx,sy,cropW,cropH,0,0,base.width,base.height);return base}
function variantCanvas(base,type){const c=document.createElement('canvas');c.width=base.width;c.height=base.height;const x=c.getContext('2d',{willReadFrequently:true});x.drawImage(base,0,0);if(type==='original')return c;const im=x.getImageData(0,0,c.width,c.height),d=im.data;let sum=0;for(let i=0;i<d.length;i+=4){const g=.299*d[i]+.587*d[i+1]+.114*d[i+2];sum+=g;d[i]=d[i+1]=d[i+2]=g}if(type==='gray'){x.putImageData(im,0,0);return c}const avg=sum/(d.length/4),threshold=Math.max(105,Math.min(190,avg*.86));for(let i=0;i<d.length;i+=4){const g=d[i],val=g>threshold?255:0;d[i]=d[i+1]=d[i+2]=val}x.putImageData(im,0,0);return c}
function candidatesFromOCR(text){
  const raw=String(text).toUpperCase(),compact=raw.replace(/[^A-Z0-9]/g,'');
  const full=[],partial=[];
  // Look at every possible start position instead of trusting the first two letters.
  for(let i=0;i<compact.length;i++){
    if(i+1>=compact.length||!/^[A-Z]{2}$/.test(compact.slice(i,i+2)))continue;
    let digits='';
    for(let j=i+2;j<compact.length&&digits.length<5;j++){
      if(/[0-9]/.test(compact[j]))digits+=compact[j];
      else if(digits.length)break;
      else break;
    }
    if(digits.length===5)full.push({reg:compact.slice(i,i+2)+digits,start:i});
    else if(digits.length)partial.push({reg:compact.slice(i,i+2)+digits,start:i});
  }
  return{raw:raw.trim(),compact,full,partial};
}
function chooseCandidate(results){
  const votes=new Map();
  for(const r of results)for(const c of r.full){
    const x=votes.get(c.reg)||{reg:c.reg,count:0,types:[],bestStart:999};
    x.count++;x.types.push(r.type);x.bestStart=Math.min(x.bestStart,c.start);votes.set(c.reg,x);
  }
  const ranked=[...votes.values()].sort((a,b)=>b.count-a.count||a.bestStart-b.bestStart);
  return ranked[0]?.reg||null;
}
function bestPartial(results){
  return results.flatMap(r=>r.partial.map(x=>x.reg)).sort((a,b)=>b.length-a.length)[0]||null;
}
async function recognize(canvas){const w=await initOCR();const r=await w.recognize(canvas,{}, {text:true});return r.data.text}
function showDebug(base,results){const img=$('#debugImage');img.src=base.toDataURL('image/jpeg',.88);img.hidden=false;$('#debugText').textContent=results.map(r=>`${r.type}: ${r.text.replace(/\s+/g,' ').trim()||'(ingen tekst)'}${r.full.length?'\n  kandidater: '+r.full.map(x=>formatReg(x.reg)).join(', '):''}`).join('\n')}
async function readPlate(){
  if(!stream||ocrBusy)return;
  ocrBusy=true;$('#readBtn').disabled=true;
  try{
    const base=cropFromVideo();if(!base)throw Error('Kamerabildet er ikke klart ennå.');
    setScanStatus('Leser skilt…');
    const results=[];
    for(const type of ['original','gray','threshold']){
      const canvas=variantCanvas(base,type),text=await recognize(canvas),parsed=candidatesFromOCR(text);
      results.push({type,text,...parsed});
    }
    showDebug(base,results);
    const full=chooseCandidate(results),partial=bestPartial(results);
    if(full){$('#q').value=formatReg(full);lookup();setScanStatus(`Leste: ${formatReg(full)} • kameraet er fortsatt aktivt`);$('#result').scrollIntoView({behavior:'smooth',block:'center'});}
    else if(partial){$('#q').value=formatReg(partial);lookup();setScanStatus(`Leste bare: ${formatReg(partial)} • prøv igjen eller kompletter manuelt.`);}
    else setScanStatus('Fant ikke 2 bokstaver + 5 sifre. Juster skiltet i rammen og trykk LES SKILT igjen.');
  }catch(e){setScanStatus('OCR-feil: '+e.message)}finally{ocrBusy=false;$('#readBtn').disabled=false}
}
async function startCamera(){try{await initOCR();if(!navigator.mediaDevices?.getUserMedia)throw Error('Nettleseren gir ikke tilgang til live-kamera.');stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'},width:{ideal:1920},height:{ideal:1080}},audio:false});const v=$('#video');v.srcObject=stream;await v.play();$('#liveWrap').hidden=false;$('#liveControls').hidden=false;$('#liveBtn').hidden=true;setScanStatus('Kamera klart • plasser skiltet i rammen og trykk LES SKILT')}catch(e){setScanStatus('Kunne ikke starte kamera: '+e.message)}}
function stopCamera(){if(stream){stream.getTracks().forEach(t=>t.stop());stream=null}$('#video').srcObject=null;$('#liveWrap').hidden=true;$('#liveControls').hidden=true;$('#liveBtn').hidden=false;setScanStatus('Kamera stoppet')}
$('#liveBtn').onclick=startCamera;$('#readBtn').onclick=readPlate;$('#stopBtn').onclick=stopCamera;document.addEventListener('visibilitychange',()=>{if(document.hidden&&stream)stopCamera()});
$('#camera').addEventListener('change',async e=>{const file=e.target.files[0];if(!file)return;const preview=$('#preview');preview.src=URL.createObjectURL(file);preview.hidden=false;try{await initOCR();setScanStatus('Leser bilde…');const text=await recognize(file),p=parseOCR(text),got=p.full||p.partial;if(!got){setScanStatus('Fant ikke et registreringsnummer. Prøv et nærmere bilde.');return}$('#q').value=formatReg(got);lookup();setScanStatus(p.full?`Leste: ${formatReg(p.full)}`:`Leste bare: ${formatReg(p.partial)} • kompletter manuelt.`)}catch(err){setScanStatus('Kunne ikke lese bildet: '+err.message)}finally{e.target.value=''}});
try{db=JSON.parse(localStorage.getItem('parkingDb')||'[]')}catch{}renderStatus();lookup();if('serviceWorker'in navigator&&location.protocol!=='file:')navigator.serviceWorker.register('service-worker.js');setTimeout(()=>initOCR().catch(()=>{}),700);
