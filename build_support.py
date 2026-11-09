#!/usr/bin/env python3
"""Génère l'outil « Support XML » (Excel -> XML + Compléter un XML) à partir de index.html (outil SGC).
Usage : python3 build_support.py index.html xml.html
Les blocs XMLC, GENX, GENUI et XMLUI sont repris tels quels : toute amélioration du moteur XML faite dans index.html est donc reprise ici en relançant ce script."""
import re,sys
src=open(sys.argv[1],encoding='utf-8').read()
lines=src.split('\n')
def block(name):
    a=src.index('/*%s-START*/'%name); b=src.index('/*%s-END*/'%name)+len('/*%s-END*/'%name)
    return src[a:b]
# --- en-tête (CSS inclus) ---
end_style=next(i for i,l in enumerate(lines) if l.strip()=='</style>')
head='\n'.join(lines[:end_style+1])
head=re.sub(r'<title>.*?</title>','<title>Outils XML ePack</title>',head,count=1)
brand=next(l for l in lines if l.startswith('<div class="brandbar">'))
brand=brand.replace('Outil de suivi des mises à jour Grands Comptes','Outils XML · Support')
code=''
i18n=block('I18N')
# dictionnaire : textes propres à la page Support
extra='''Outils XML ⇒ XML tools
Génère ou complète les XML des modules à partir des fichiers du client. Rien n'est enregistré : les fichiers restent dans ton navigateur. ⇒ Generate or complete the module XML files from the client's files. Nothing is saved: files stay in your browser.
Excel → XML ⇒ Excel → XML
Dépose l'Excel du client, choisis le type d'XML (Warming, Stickers, Contrôles divers…), récupère le XML dans un ZIP. ⇒ Drop the client's Excel file, choose the XML type (Warming, Stickers, Miscellaneous controls…), get the XML in a ZIP.
Un onglet = un XML ⇒ One sheet = one XML
XML → Excel ⇒ XML → Excel
Dépose le ZIP de la configuration du client (ou un seul XML) : j'en fais un Excel, un onglet par module. ⇒ Drop the client's configuration ZIP (or a single XML): I turn it into an Excel file, one sheet per module.
Un onglet = un module ⇒ One sheet = one module
Compléter un XML ⇒ Complete an XML
Dépose un export XML du configurateur : j'ajoute les balises manquantes et je normalise le fichier. ⇒ Drop an XML export from the configurator: I add the missing tags and normalise the file.
Produits, checklists, plans de nettoyage ⇒ Products, checklists, cleaning plans
Choisis le type d'XML, dépose l'Excel du client puis génère. ⇒ Choose the XML type, drop the client's Excel file, then generate.
'''
k=i18n.index('`;',i18n.index('I18N_RAW'))
i18n=i18n[:k]+'\n'+extra+i18n[k:]
fl=block('FLAGS')
src_lines=src
a=src.index('const LSTH='); b=src.index('const meKey=')
prof=src[a:b]
code+=i18n+'\n'+fl+'\n'
for n in ['XMLC','GENX','GENUI','XMLUI']: code+=block(n)+'\n'
code+=prof+'\n'
code+=open(__import__('os').path.join(__import__('os').path.dirname(__import__('os').path.abspath(__file__)),'x2e.js'),encoding='utf-8').read()+'\n'
# adaptations propres à l'outil Support (aucun contexte de MAJ)
a='md.append(hd,$("p","gsubp",M.name+" · "+regLabel));'
assert code.count(a)==1
code=code.replace(a,'const subP=$("p","gsubp gfile");md.append(hd,subP);')
a='const fl=$("p",null);fl.append("Fichier : ",$("b",null,st.file)," ");'
assert code.count(a)==1
code=code.replace(a,'const fl=subP;fl.append("Fichier : ",$("b",null,st.file)," ");')
a='fl.append(ch,inp);md.appendChild(fl);'
assert code.count(a)==1
code=code.replace(a,'fl.append(ch,inp);')
a='if(wd.length===1)st.sel.add(wd[0][1])'
assert code.count(a)==1,code.count(a)
code=code.replace(a,'wd.forEach(([,i])=>st.sel.add(i))')
a='al.appendChild($("small",null,"Génération bloquée : demande les DLC manquantes au siège (mail ci-dessous), puis relance. Si un produit n\'a vraiment pas de DLC, saisis -1 dans l\'Excel."));'
assert code.count(a)==1,code.count(a)
code=code.replace(a,'al.appendChild($("small",null,"Génération bloquée : il est nécessaire de demander au client l\'information des DLC manquantes (mail ci-dessous) puis de relancer la génération du fichier XML."));al.appendChild($("small",null,"Si un produit n\'a vraiment pas de DLC, saisis -1 dans l\'Excel."));')
a='$("strong",null,"Mail au siège : DLC manquantes")'
assert code.count(a)==1
code=code.replace(a,'$("strong",null,"Mail au client : DLC manquantes")')
a='ta.setAttribute("aria-label","Mail au siège")'
assert code.count(a)==1
code=code.replace(a,'ta.setAttribute("aria-label","Mail au client")')
a='const w=$("div");w.append(d,inp);md.appendChild(w);return;'
assert code.count(a)==1
code=code.replace(a,'const w=$("div");w.append(d,inp);md.appendChild(w);const g0=md.querySelector(".gtl");if(g0)g0.before(w);const er=md.querySelector("p.msg.err");if(er)w.after(er);return;')
prelude=r'''
const SUPPORT=true;
let majs={},cur=null;
const $=(t,c,x)=>{const e=document.createElement(t);if(c)e.className=c;if(x!=null)e.textContent=x;return e};
const slug=s=>s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"");
const fold=s=>String(s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"");
const persist=()=>{},groupOf=()=>"",zipGroup=()=>"",zipName=(id,b)=>b+".zip",meKey=()=>null;
function toast(t){const d=$("div","spToast",t);d.setAttribute("role","status");document.body.appendChild(d);setTimeout(()=>d.classList.add("out"),4200);setTimeout(()=>d.remove(),4600)}
'''
main=r'''
const ctx={title:"",group:""},XM={name:"",gen:true,regions:[]},XR={label:"—"};
function buildBar(){
  const br=document.getElementById("brandright");if(!br)return;br.textContent="";
  const rg=$("div","actions");const av=$("button","avchip");av.type="button";
  {const nm=myName();const ini=nm?nm.split(" ").filter(Boolean).slice(0,2).map(w=>w[0].toUpperCase()).join(""):"?";
   av.append($("span","avc",ini),$("span","avn",nm||"Ton nom"));av.title=nm?"Profil : "+nm+" (modifier dans Paramètres)":"Indique ton nom";av.setAttribute("aria-label",av.title);
   av.setAttribute("aria-haspopup","dialog");av.onclick=()=>profileMenu(av)}
  const gb=$("button","gearbtn");gb.type="button";gb.title="Réglages : thème et langue";gb.setAttribute("aria-label","Réglages : thème et langue");gb.setAttribute("aria-haspopup","dialog");
  gb.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>';
  gb.onclick=()=>settingsMenu(gb);
  rg.append(av,gb);br.appendChild(rg);
}
function settingsMenu(anchor){
  const old=document.getElementById("pmenu");if(old){pmClose(old);return}
  const pm=$("div","pmenu");pm.id="pmenu";pm.setAttribute("role","dialog");pm.setAttribute("aria-label","Réglages");
  const trw=$("div","pmrow");trw.style.borderTop="0";trw.style.paddingTop="0";trw.append($("span",null,"Thème"),themeBtn());
  const lw=$("div","pmrow");lw.style.borderTop="0";lw.style.paddingTop="0";lw.append($("span",null,"Langue"),langBtn());
  pm.append(trw,lw);
  const r=anchor.getBoundingClientRect();pm.style.top=(r.bottom+window.scrollY+8)+"px";pm.style.right=Math.max(8,document.documentElement.clientWidth-r.right)+"px";
  document.body.appendChild(pm);
  const away=e=>{if(!pm.contains(e.target)&&!anchor.contains(e.target)){pmClose(pm);document.removeEventListener("mousedown",away);document.removeEventListener("keydown",esc)}};
  const esc=e=>{if(e.key==="Escape"){pmClose(pm);document.removeEventListener("mousedown",away);document.removeEventListener("keydown",esc);anchor.focus()}};
  document.addEventListener("mousedown",away);document.addEventListener("keydown",esc);
}
function render(){
  buildBar();
  const app=document.getElementById("app");app.textContent="";
  const ver=document.getElementById("appver");if(ver){ver.textContent=SUPPORT_BUILD;ver.title="Version en ligne : "+SUPPORT_BUILD+" (moteur "+APP_VERSION+")"}
  app.appendChild($("h1",null,"Outils XML"));
  app.appendChild($("p","sub","Génère ou complète les XML des modules à partir des fichiers du client. Rien n'est enregistré : les fichiers restent dans ton navigateur."));
  const tl=$("div","tools");
  const one=(t,d,small,fn,ic,wip)=>{const b=$("button","tool");b.type="button";const i=$("span","ticon");i.setAttribute("aria-hidden","true");i.innerHTML=ic;i.dataset.nt="1";b.append(i,$("strong",null,t),$("span",null,d),$("small",null,small));if(wip){const w=$("span","wip","BÊTA");w.title="Version bêta : utilisable, mais ton retour est attendu";b.appendChild(w)}b.onclick=fn;tl.appendChild(b)};
  const IC1='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="10" height="14" rx="2"/><path d="M3 9h10M8 9v9"/><g class="tarr"><path d="M15 11h6M18.5 8.5 21 11l-2.5 2.5"/></g></svg>';
  const IC2='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="m9 14 2 2 4-4"/><path class="tspk" d="M19 2v3M17.5 3.5h3"/></svg>';
  const IC3='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="M9 12h6M9 16h6M12 11v7"/></svg>';
  one("Excel → XML","Dépose l'Excel du client, choisis le type d'XML (Warming, Stickers, Contrôles divers…), récupère le XML dans un ZIP.","Un onglet = un XML",()=>openGen(ctx,null,null,XM,XR),IC1);
  one("Compléter un XML","Dépose un export XML du configurateur : j'ajoute les balises manquantes et je normalise le fichier.","Produits, checklists, plans de nettoyage",()=>openXml(ctx,null,null),IC2);
  one("XML → Excel","Dépose le ZIP de la configuration du client (ou un seul XML) : j'en fais un Excel, un onglet par module.","Un onglet = un module",()=>openX2E(),IC3);
  app.appendChild(tl);
  trSchedule();
}
render();
'''
import datetime,zoneinfo
SB='1.0.3'
build_label='v'+SB
ver=re.search(r'const APP_VERSION="(v[\d.]+)"',src).group(1)
extra_css='''<style>
/* animations propres à la page Support */
.tool{position:relative;overflow:hidden}
.tool .ticon{width:44px;height:44px;border-radius:12px;display:grid;place-items:center;background:rgba(155,249,68,.16);color:var(--accent);margin-bottom:4px;transition:transform .25s,background .25s}
.tool .ticon svg{width:26px;height:26px;overflow:visible}
.tool:hover .ticon,.tool:focus-visible .ticon{background:rgba(155,249,68,.3);transform:rotate(-6deg) scale(1.08)}
.tool .tarr{transition:transform .3s}.tool:hover .tarr,.tool:focus-visible .tarr{animation:tnudge .9s ease-in-out infinite}
.tool .tspk{transform-origin:center;transform-box:fill-box}.tool:hover .tspk,.tool:focus-visible .tspk{animation:tspin 1.2s linear infinite}
.tool::after{content:"";position:absolute;inset:0;background:linear-gradient(105deg,transparent 35%,rgba(155,249,68,.22) 50%,transparent 65%);transform:translateX(-120%);pointer-events:none}
.tool:hover::after,.tool:focus-visible::after{transform:translateX(120%);transition:transform .8s ease}
.tool:active{transform:translateY(0) scale(.98)}
.tools .tool{animation:tin .5s cubic-bezier(.2,.8,.2,1) both}.tools .tool:nth-child(2){animation-delay:.12s}
.btn,.gtile,.chip{transition:transform .15s,box-shadow .2s,background .2s,border-color .2s}
.btn:not(:disabled):hover,.gtile:hover{transform:translateY(-1px)}
.btn:not(:disabled):active,.gtile:active,.chip:active{transform:scale(.97)}
.gtile.on .gtk{animation:tpop .35s cubic-bezier(.3,1.6,.5,1)}
@keyframes tin{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
@keyframes tnudge{0%,100%{transform:translateX(0)}50%{transform:translateX(5px)}}
@keyframes tspin{to{transform:rotate(360deg)}}
@keyframes tpop{from{transform:scale(.4)}to{transform:scale(1)}}
@media (prefers-reduced-motion:reduce){.tools .tool,.tool .tarr,.tool .tspk,.gtile.on .gtk{animation:none!important}.tool::after{display:none}.tool,.btn,.gtile,.chip,.tool .ticon{transition:none!important}}
/* XML → Excel */
.gearbtn{font:inherit;width:40px;height:40px;display:inline-grid;place-items:center;border-radius:999px;border:1px solid var(--line,#8886);background:transparent;color:inherit;cursor:pointer;margin-left:8px;transition:transform .35s,background .2s}.gearbtn svg{width:20px;height:20px}.gearbtn:hover{background:rgba(155,249,68,.18);transform:rotate(60deg)}.gearbtn:focus-visible{outline:2px solid var(--accent);outline-offset:2px}\n.spToast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#12201e;color:#fff;border:1px solid #9fe03a;border-radius:10px;padding:10px 16px;z-index:200;max-width:min(92vw,520px);box-shadow:0 8px 28px rgba(0,0,0,.35);font-size:14px;transition:opacity .35s}.spToast.out{opacity:0}.gsubp.gfile:empty{display:none}.gsubp.gfile{display:flex;align-items:center;flex-wrap:wrap;gap:6px;overflow-wrap:anywhere}\n.x2beta{font-size:11px;font-weight:700;letter-spacing:.04em;padding:2px 8px;border-radius:999px;background:rgba(255,176,32,.18);color:#ffb020;border:1px solid rgba(255,176,32,.55);vertical-align:middle}
.x2bn{font-size:12.5px;color:var(--muted,#9aa39e);border-left:3px solid #ffb020;padding:4px 10px;margin:8px 0}
.wip{position:absolute;top:12px;right:12px;font-size:11px;font-weight:700;letter-spacing:.02em;padding:3px 9px;border-radius:999px;background:rgba(255,176,32,.18);color:#ffb020;border:1px solid rgba(255,176,32,.55)}
.x2list{flex:none;min-height:min(150px,30vh);display:flex;flex-direction:column;gap:2px;max-height:34vh;overflow:auto;border:1px solid var(--border,#ccc);border-radius:10px;padding:6px}
.x2row{display:flex;align-items:center;gap:10px;padding:6px 8px;border-radius:8px;cursor:pointer}
.x2row:hover{background:rgba(155,249,68,.12)}.x2row.empty{opacity:.55}
.modal:has(.x2list){max-height:92vh;overflow:auto}
.x2row small{margin-left:auto;opacity:.7}.x2pv{padding:2px 10px;font-size:.78rem}.x2pv.on{border-color:#9BF944;background:rgba(155,249,68,.18)}.x2pv:disabled{opacity:.35;cursor:default}
.x2prev{margin:10px 0;border:1px solid var(--border,#ccc);border-radius:10px;padding:8px 10px;animation:ddin .18s ease-out}.x2pt{margin:0 0 6px;font-weight:700}
.x2sc{overflow:auto;max-width:100%}.x2prev table{border-collapse:collapse;font-size:.8rem;min-width:100%}.x2prev th,.x2prev td{border:1px solid var(--border,#ccc);padding:3px 8px;text-align:left;white-space:nowrap}.x2prev th{background:rgba(155,249,68,.16)}
.x2sum{display:inline-flex;align-items:center;gap:6px;font-size:.88rem;margin-left:10px}
.x2n{flex:1}
</style>
'''
out=head+'\n'+extra_css+brand+'\n<div class="wrap" id="app"></div>\n<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>\n<script>\nconst APP_VERSION="%s";\nconst SUPPORT_BUILD="%s";\n'%(ver,build_label)+prelude+code+main+'\n</script>\n'
open(sys.argv[2],'w',encoding='utf-8').write(out)
print('OK',ver,len(out))
