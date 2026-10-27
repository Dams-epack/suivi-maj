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
code=code.replace(a,'md.append(hd,$("p",null,"Choisis le type d\'XML, dépose l\'Excel du client puis génère."));')
a='if(wd.length===1)st.sel.add(wd[0][1])'
assert code.count(a)==1,code.count(a)
code=code.replace(a,'wd.forEach(([,i])=>st.sel.add(i))')
prelude=r'''
const SUPPORT=true;
let majs={},cur=null;
const $=(t,c,x)=>{const e=document.createElement(t);if(c)e.className=c;if(x!=null)e.textContent=x;return e};
const slug=s=>s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"");
const fold=s=>String(s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"");
const persist=()=>{},groupOf=()=>"",zipGroup=()=>"",zipName=(id,b)=>b+".zip",meKey=()=>null;
'''
main=r'''
const ctx={title:"",group:""},XM={name:"",gen:true,regions:[]},XR={label:"—"};
function buildBar(){
  const br=document.getElementById("brandright");if(!br)return;br.textContent="";
  const rg=$("div","actions");const av=$("button","avchip");av.type="button";
  {const nm=myName();const ini=nm?nm.split(" ").filter(Boolean).slice(0,2).map(w=>w[0].toUpperCase()).join(""):"?";
   av.append($("span","avc",ini),$("span","avn",nm||"Ton nom"));av.title=nm?"Profil : "+nm+" (modifier dans Paramètres)":"Indique ton nom";av.setAttribute("aria-label",av.title);
   av.setAttribute("aria-haspopup","dialog");av.onclick=()=>profileMenu(av)}
  rg.append(av);br.appendChild(rg);
}
function render(){
  buildBar();
  const app=document.getElementById("app");app.textContent="";
  const ver=document.getElementById("appver");if(ver)ver.textContent="Outils XML ePack · "+APP_VERSION;
  app.appendChild($("h1",null,"Outils XML"));
  app.appendChild($("p","sub","Génère ou complète les XML des modules à partir des fichiers du client. Rien n'est enregistré : les fichiers restent dans ton navigateur."));
  const tl=$("div","tools");
  const one=(t,d,small,fn,ic,wip)=>{const b=$("button","tool");b.type="button";const i=$("span","ticon");i.setAttribute("aria-hidden","true");i.innerHTML=ic;i.dataset.nt="1";b.append(i,$("strong",null,t),$("span",null,d),$("small",null,small));if(wip){const w=$("span","wip","Work in progress");w.title="Fonctionnalité en cours de développement";b.appendChild(w)}b.onclick=fn;tl.appendChild(b)};
  const IC1='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="10" height="14" rx="2"/><path d="M3 9h10M8 9v9"/><g class="tarr"><path d="M15 11h6M18.5 8.5 21 11l-2.5 2.5"/></g></svg>';
  const IC2='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="m9 14 2 2 4-4"/><path class="tspk" d="M19 2v3M17.5 3.5h3"/></svg>';
  const IC3='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="M9 12h6M9 16h6M12 11v7"/></svg>';
  one("Excel → XML","Dépose l'Excel du client, choisis le type d'XML (Warming, Stickers, Contrôles divers…), récupère le XML dans un ZIP.","Un onglet = un XML",()=>openGen(ctx,null,null,XM,XR),IC1);
  one("Compléter un XML","Dépose un export XML du configurateur : j'ajoute les balises manquantes et je normalise le fichier.","Produits, checklists, plans de nettoyage",()=>openXml(ctx,null,null),IC2);
  one("XML → Excel","Dépose le ZIP de la configuration du client (ou un seul XML) : j'en fais un Excel, un onglet par module.","Un onglet = un module",()=>openX2E(),IC3,true);
  app.appendChild(tl);
  trSchedule();
}
render();
'''
ver=re.search(r'const APP_VERSION="(v\d+)"',src).group(1)
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
.wip{position:absolute;top:12px;right:12px;font-size:11px;font-weight:700;letter-spacing:.02em;padding:3px 9px;border-radius:999px;background:rgba(255,176,32,.18);color:#ffb020;border:1px solid rgba(255,176,32,.55)}
.x2list{display:flex;flex-direction:column;gap:2px;max-height:46vh;overflow:auto;border:1px solid var(--border,#ccc);border-radius:10px;padding:6px}
.x2row{display:flex;align-items:center;gap:10px;padding:6px 8px;border-radius:8px;cursor:pointer}
.x2row:hover{background:rgba(155,249,68,.12)}.x2row.empty{opacity:.55}
.x2n{flex:1}
</style>
'''
out=head+'\n'+extra_css+brand+'\n<div class="wrap" id="app"></div>\n<div class="appver" id="appver"></div>\n<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>\n<script>\nconst APP_VERSION="%s";\n'%ver+prelude+code+main+'\n</script>\n'
open(sys.argv[2],'w',encoding='utf-8').write(out)
print('OK',ver,len(out))
