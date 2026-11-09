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
  if(fbDb){const fb=$("button","fbbtn");fb.type="button";fb.title="Un retour sur l'outil ?";fb.setAttribute("aria-label","Retours sur l'outil");fb.setAttribute("aria-haspopup","dialog");
    fb.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15a2 2 0 0 1-2 2H8l-5 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>';
    if(fbCanEdit&&fbNew()){const bd=$("span","fbbadge",String(fbNew()));bd.setAttribute("aria-hidden","true");fb.appendChild(bd);fb.setAttribute("aria-label","Retours sur l'outil ("+fbNew()+" nouveaux)")}
    fb.onclick=fbOpenDlg;rg.appendChild(fb)}
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
  one("Excel → XML","Dépose l'Excel du client, choisis le type d'XML (Warming, Stickers, Contrôles divers…), récupère le XML dans un ZIP.","Un onglet = un XML",()=>{fbTool="Excel → XML";openGen(ctx,null,null,XM,XR)},IC1);
  one("Compléter un XML","Dépose un export XML du configurateur : j'ajoute les balises manquantes et je normalise le fichier.","Produits, checklists, plans de nettoyage",()=>{fbTool="Compléter un XML";openXml(ctx,null,null)},IC2);
  one("XML → Excel","Dépose le ZIP de la configuration du client (ou un seul XML) : j'en fais un Excel, un onglet par module.","Un onglet = un module",()=>{fbTool="XML → Excel";openX2E()},IC3);
  app.appendChild(tl);
  trSchedule();
}

/* ===== Retours de l'équipe (base partagée de la page) ===== */
const FB_TOOLS=["Excel → XML","XML → Excel","Compléter un XML","Autre"];
const FB_KINDS=[["bug","Bug"],["idea","Idée"],["label","Libellé / nom"],["other","Autre"]];
const FB_ST=[["new","Nouveau"],["seen","Vu"],["done","Traité"]];
let fbDb=null,fbUser=null,fbMe=null,fbList=[],fbDlg=null,fbTab="send",fbTool="Excel → XML",fbKind="bug",fbDraft="",fbCanWrite=null,fbCanEdit=false,fbFilter="open",fbInfo="";
const fbDate=t=>{try{return new Date(t).toLocaleString("fr-FR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}catch(e){return ""}};
const fbNew=()=>fbList.filter(x=>(x.status||"new")==="new").length;
async function fbName(){const n=(typeof myName==="function"?myName():"")||"";if(n)return n;try{const m=fbUser&&await fbUser.me();return (m&&m.name)||""}catch(e){return ""}}
(async()=>{try{
  const c=window.claude;if(!c||!c.use)return;
  const u=await c.use("user");fbUser=u||null;
  if(u){try{fbMe=await u.id()}catch(e){}try{fbCanEdit=!!(await u.canEdit())}catch(e){}try{fbCanWrite=u.can?await u.can("data.write"):null}catch(e){}}
  fbDb=await c.use("db");if(!fbDb)return;
  buildBar();
  fbDb.collection("feedback").orderBy("ts","desc").limit(300).onSnapshot(s=>{fbList=s.docs.map(d=>Object.assign({id:d.id},d.data()));buildBar();if(fbDlg)fbPaint()},()=>{});
}catch(e){}})();
function fbClose(){if(!fbDlg)return;const bgn=fbDlg.parentNode;(bgn&&bgn.classList&&bgn.classList.contains("modal-bg")?bgn:fbDlg).remove();fbDlg=null;document.removeEventListener("keydown",fbKey,true)}
function fbKey(e){if(e.key==="Escape"){e.stopPropagation();fbClose()}}
function fbOpenDlg(){
  if(fbDlg)return;
  const bg=$("div","modal-bg");bg.setAttribute("role","dialog");bg.setAttribute("aria-modal","true");bg.setAttribute("aria-label","Retours sur l'outil");
  bg.addEventListener("mousedown",e=>{if(e.target===bg)fbClose()});
  const md=$("div","modal xml fbm");bg.appendChild(md);document.body.appendChild(bg);fbDlg=md;
  document.addEventListener("keydown",fbKey,true);fbTab=fbCanEdit?"list":"send";fbPaint();
}
function fbPaint(){
  const md=fbDlg;if(!md)return;const keep=md.scrollTop;md.textContent="";
  const hd=$("div","xhead sticky");const h=$("h2",null,"Retours sur l'outil");const x=$("button","xclose","×");x.type="button";x.setAttribute("aria-label","Fermer");x.onclick=fbClose;hd.append(h,x);md.appendChild(hd);
  const tabs=$("div","fbtabs");tabs.setAttribute("role","tablist");
  [["send","Envoyer un retour"],["list","Retours de l'équipe ("+fbList.length+")"]].forEach(([k,t])=>{const b=$("button","fbtab"+(fbTab===k?" on":""),t);b.type="button";b.setAttribute("role","tab");b.setAttribute("aria-selected",fbTab===k?"true":"false");b.onclick=()=>{fbTab=k;fbPaint()};tabs.appendChild(b)});
  md.appendChild(tabs);
  if(fbTab==="send")fbSendPane(md);else fbListPane(md);
  md.scrollTop=keep;
}
function fbSendPane(md){
  md.appendChild($("p","gsubp","Un bug, une idée, un nom à corriger ? Dis-le ici : Dam's le voit directement."));
  const f1=$("label","fld");f1.appendChild($("span",null,"Où ?"));const sl=$("select");FB_TOOLS.forEach(t=>{const o=$("option",null,t);o.value=t;if(t===fbTool)o.selected=true;sl.appendChild(o)});sl.onchange=()=>{fbTool=sl.value};f1.appendChild(sl);md.appendChild(f1);
  const f2=$("div","fld");f2.appendChild($("span",null,"Type"));const ch=$("div","fbkinds");ch.setAttribute("role","radiogroup");
  FB_KINDS.forEach(([k,t])=>{const b=$("button","fbkind"+(fbKind===k?" on":""),t);b.type="button";b.setAttribute("role","radio");b.setAttribute("aria-checked",fbKind===k?"true":"false");b.onclick=()=>{fbKind=k;ch.querySelectorAll(".fbkind").forEach(z=>{const on=z===b;z.classList.toggle("on",on);z.setAttribute("aria-checked",on?"true":"false")})};ch.appendChild(b)});
  f2.appendChild(ch);md.appendChild(f2);
  const f3=$("label","fld");f3.appendChild($("span",null,"Ton message"));const ta=$("textarea");ta.rows=5;ta.maxLength=1500;ta.value=fbDraft;ta.placeholder="Ce qui ne va pas, ce que tu voudrais, le nom à corriger… Ne mets pas de données sensibles du client.";
  const cnt=$("small","fbcnt",ta.value.length+" / 1500");ta.oninput=()=>{fbDraft=ta.value;cnt.textContent=ta.value.length+" / 1500";sb.disabled=!ta.value.trim()||fbCanWrite===false};
  f3.appendChild(ta);f3.appendChild(cnt);md.appendChild(f3);
  md.appendChild($("small","fbnote","La version de l'outil ("+SUPPORT_BUILD+"), ton nom et l'heure sont joints automatiquement."));
  const msg=$("p","msg"+(fbInfo?"":" ")," ");msg.textContent=fbInfo;
  if(fbCanWrite===false){msg.className="msg err";msg.textContent="Tu peux lire les retours mais pas en envoyer : demande à Dam's de te passer en Contributeur."}
  md.appendChild(msg);
  const ac=$("div","adqact");const sb=$("button","btn","Envoyer");sb.type="button";sb.disabled=!fbDraft.trim()||fbCanWrite===false;
  sb.onclick=async()=>{const t=fbDraft.trim();if(!t||!fbDb)return;sb.disabled=true;msg.className="msg";msg.textContent="Envoi…";
    try{await fbDb.collection("feedback").add({ts:Date.now(),authorId:fbMe||"",authorName:await fbName(),tool:fbTool,kind:fbKind,text:t,version:SUPPORT_BUILD,status:"new"});
      fbDraft="";fbInfo="Merci, ton retour est envoyé.";fbPaint();toast("Merci, retour envoyé")}
    catch(e){sb.disabled=false;msg.className="msg err";msg.textContent=(e&&e.code==="invalid_argument")?"Envoi refusé : tu n'as pas les droits d'écriture. Demande à Dam's de te passer en Contributeur.":"Envoi impossible pour le moment, réessaie dans un instant."}};
  ac.appendChild(sb);md.appendChild(ac);setTimeout(()=>ta.focus(),30);
}
function fbListPane(md){
  const bar=$("div","fbfilters");
  [["open","À traiter"],["all","Tous"],["done","Traités"]].forEach(([k,t])=>{const b=$("button","fbf"+(fbFilter===k?" on":""),t);b.type="button";b.onclick=()=>{fbFilter=k;fbPaint()};bar.appendChild(b)});
  md.appendChild(bar);
  const items=fbList.filter(x=>fbFilter==="all"||(fbFilter==="done"?x.status==="done":x.status!=="done"));
  if(!items.length){md.appendChild($("p","gsubp",fbList.length?"Aucun retour dans cette vue.":"Aucun retour pour l'instant."));return}
  const kl=Object.fromEntries(FB_KINDS),sl=Object.fromEntries(FB_ST);
  items.forEach(it=>{
    const c=$("article","fbitem st-"+(it.status||"new"));
    const top=$("div","fbtop");top.append($("b",null,it.authorName||"Anonyme"),$("span","fbmeta",fbDate(it.ts)+" · "+(it.version||"")));c.appendChild(top);
    const tg=$("div","fbtags");tg.append($("span","fbtag",it.tool||"—"),$("span","fbtag k-"+(it.kind||"other"),kl[it.kind]||"Autre"),$("span","fbtag s-"+(it.status||"new"),sl[it.status||"new"]||"Nouveau"));c.appendChild(tg);
    c.appendChild($("p","fbtext",it.text||""));
    if(it.reply){const r=$("p","fbreply");r.append($("b",null,"Réponse : "),it.reply);c.appendChild(r)}
    if(fbCanEdit){
      const ac=$("div","fbact");
      const s=$("select");FB_ST.forEach(([v,t])=>{const o=$("option",null,t);o.value=v;if(v===(it.status||"new"))o.selected=true;s.appendChild(o)});s.setAttribute("aria-label","Statut");
      s.onchange=()=>fbDb.doc("feedback/"+it.id).update({status:s.value}).catch(()=>toast("Modification impossible"));
      const rp=$("input");rp.type="text";rp.placeholder="Répondre (visible par l'équipe)";rp.value=it.reply||"";rp.maxLength=400;rp.setAttribute("aria-label","Réponse");
      rp.onchange=()=>{fbDb.doc("feedback/"+it.id).update({reply:rp.value.trim(),replyTs:Date.now()}).catch(()=>toast("Modification impossible"))};
      const dl=$("button","btn ghost","Supprimer");dl.type="button";dl.onclick=()=>{if(dl.dataset.s!=="1"){dl.dataset.s="1";dl.textContent="Confirmer ?";setTimeout(()=>{dl.dataset.s="";dl.textContent="Supprimer"},3000);return}fbDb.doc("feedback/"+it.id).delete().catch(()=>toast("Suppression impossible"))};
      ac.append(s,rp,dl);c.appendChild(ac);
    }
    md.appendChild(c);
  });
}
render();
'''
import datetime,zoneinfo
SB='1.2.1'
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
.fbbtn{font:inherit;position:relative;width:40px;height:40px;display:inline-grid;place-items:center;border-radius:999px;border:1px solid var(--line,#8886);background:transparent;color:inherit;cursor:pointer}
.fbbtn svg{width:20px;height:20px}
.fbbtn:hover{background:rgba(155,249,68,.16)}
.fbbadge{position:absolute;top:-4px;right:-4px;min-width:18px;height:18px;padding:0 5px;border-radius:999px;background:var(--accent,#9bf944);color:#10201f;font-size:.7rem;font-weight:700;display:grid;place-items:center}
.modal.fbm{max-width:640px}
.fbtabs{display:flex;gap:6px;margin:6px 0 14px;border-bottom:1px solid var(--line,#8884)}
.fbtab{font:inherit;background:transparent;border:0;border-bottom:2px solid transparent;color:inherit;padding:8px 12px;cursor:pointer;opacity:.7}
.fbtab.on{border-bottom-color:var(--accent,#9bf944);opacity:1;font-weight:600}
.fbm .fld{display:block;margin:0 0 12px}
.fbm .fld>span{display:block;font-size:.8rem;opacity:.75;margin-bottom:4px;font-weight:600}
.fbm textarea{width:100%;box-sizing:border-box;font:inherit;padding:10px 12px;border-radius:10px;border:1px solid var(--line,#8886);background:transparent;color:inherit;resize:vertical}
.fbkinds{display:flex;flex-wrap:wrap;gap:8px}
.fbkind,.fbf{font:inherit;padding:6px 12px;border-radius:999px;border:1px solid var(--line,#8886);background:transparent;color:inherit;cursor:pointer}
.fbkind.on,.fbf.on{background:rgba(155,249,68,.22);border-color:var(--accent,#9bf944);font-weight:600}
.fbcnt,.fbnote{display:block;opacity:.65;font-size:.78rem;margin-top:4px}
.fbfilters{display:flex;gap:8px;margin-bottom:12px}
.fbitem{border:1px solid var(--line,#8886);border-radius:12px;padding:12px 14px;margin-bottom:10px;display:flex;flex-direction:column;gap:6px}
.fbitem.st-new{border-left:4px solid var(--accent,#9bf944)}
.fbitem.st-done{opacity:.7}
.fbtop{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.fbmeta{opacity:.65;font-size:.8rem}
.fbtags{display:flex;gap:6px;flex-wrap:wrap}
.fbtag{font-size:.75rem;padding:2px 9px;border-radius:999px;border:1px solid var(--line,#8886)}
.fbtag.k-bug{border-color:#e5484d;color:#e5484d}.fbtag.s-new{background:rgba(155,249,68,.22)}.fbtag.s-done{opacity:.7}
.fbtext{margin:0;white-space:pre-wrap;overflow-wrap:anywhere}
.fbreply{margin:0;padding:8px 10px;border-radius:8px;background:rgba(155,249,68,.12);font-size:.9rem}
.fbact{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.fbact input{flex:1 1 200px;min-width:0;font:inherit;padding:7px 10px;border-radius:8px;border:1px solid var(--line,#8886);background:transparent;color:inherit}
.fbact select{font:inherit;padding:6px 8px;border-radius:8px}
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
