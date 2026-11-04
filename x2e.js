/*X2E-START*/
/* XML → Excel : moteur (lecture ZIP sans dépendance, conversion des XML de configuration en classeur) + dialogue */
const X2E=(()=>{
  const SHEETS={Fridges:"Température frigo",Cleanings:"Nettoyage des postes",ControlsList:"Liste de contrôles",ControlsList2:"Liste de contrôles 2",ControlsList3:"Liste de contrôles 3",ControlsManager:"Contrôles manager",CheckLists:"Check lists",ProductsStickers:"Étiquettes produits",ProductsProduction:"Traçabilité produits",ProductsStockIn:"Réception marchandises",ProductsService:"Température de service",ProductsFreezing:"Mise en congélation",ProductsFrozen:"Décongélation",ProductsWarming:"Mise en température",ProductsStickersDelivery:"Étiquettes livraisons",ProductsStickersDefrosting:"Étiquettes décongélation",ProductsCookingSousVide:"Cuisson sous vide",Cookings:"Cuissons",ProductionsTemp:"Température de production",ProductionsTempType:"Types de productions",PlatsTemoins:"Plats témoins",ControlesDivers:"Contrôles divers",Actions:"Actions correctives",Methods:"Méthodes",Notes:"Notes",Steps:"Étapes",PresetText:"Textes prédéfinis",PresetTemp:"Températures prédéfinies",Moments:"Moments",Timers:"Minuteurs",AddressBook:"Carnet d'adresses",PrintersMapping:"Imprimantes",OptionsUI:"Liste des modules",Usernames:"Utilisateurs",Company:"Société",Options:"Options"};
  const LBL={Name:"Désignation",Firstname:"Prénom",Category:"Catégorie",TempMin:"Temp. min",TempMax:"Temp. max",TempLimit:"Temp. limite",TempMaxExcu:"Temp. max exception",DureeMaxExcu:"Durée max exception (min)",Sensor_Name:"Sonde",Sensor_Id:"Id sonde",Sensor_CanChangeValue:"Valeur modifiable (sonde)",Recurrence:"Récurrence",CheckList_Name:"Check list",CheckList_AllMandatoryToValid:"Tous les éléments obligatoires",DateLimit:"DLC",DLType:"Unité DLC",Fridge:"Frigo",Printer:"Imprimante",MustIndicateOrigin:"Origine à indiquer",IsDDM:"DDM",Dlc:"DLC",Ddm:"DDM",Du:"Unité",TypeFunction:"Module",State:"État",ButtonText:"Texte du bouton",ButtonImage:"Image",ButtonAction:"Action",ModulePosition:"Position",ProductReference:"Référence produit",DLCTheorique:"DLC théorique",NumLot:"N° de lot",Quantity:"Quantité",Job:"Poste",Phone:"Téléphone",Admin:"Administrateur",eMail:"E-mail",Type:"Type",ProviderCode:"Code fournisseur",PrinterName:"Nom imprimante",Copies:"Copies",MesureMin:"Mesure min",MesureMax:"Mesure max",MesureUnit:"Unité",Conservation:"Conservation (jours)",TimeFrom:"Heure de début",Libelle:"Libellé",Message:"Message",Jours:"Jours",Son:"Son",Couleur:"Couleur",TempStart_Min:"Temp. départ min",TempStart_Max:"Temp. départ max",TempFinish_Min:"Temp. fin min",TempFinish_Max:"Temp. fin max",Alert_MaxTime:"Durée max alerte",ConditionTempInf:"Condition temp. inférieure",ConditionTempSup:"Condition temp. supérieure",ConditionCheckList:"Condition check list",ConditionCheckListElements:"Éléments check list conditionnels",IsTempMeth:"Méthode de température",Items:"Éléments",Language_Default:"Langue",ProductCode:"Code produit",CookingTime:"Durée de cuisson"};
  const HIDE=new Set(["MyDate","MyNote_Text","MyNote_Photo","MyNote_Username","UserName","Id","TS","Delete","Disable","idFileAttente","Photo","_IsSensorError","IsSensorError","Password","RoleId","Icon","Icone","ButtonImage","TimeStartTicks","TimeEndTicks","IsWebEPH","CanGoBackHome","Paper_Size","Paper_Source","IsCollate","Duplex","IsColor","MAC","IsLandscape"]);
  /* TypeFunction (colonne « Module ») : code technique -> nom du module en français */
  const FUNC={Fridge:"Température frigo",Frozen:"Décongélation",Defrosting:"Étiquettes décongélation",StockIn:"Réception marchandises",DeliveryTemp:"Température à la livraison",Tracking:"Traçabilité",Oil:"Suivi huiles",Warming:"Mise en température",ServiceTemp:"Température de service",StickerProduct:"Étiquettes produits",Cleaning:"Nettoyage des postes",Scheduler:"Planning",Custom:"Personnalisé",ALL:"Tous les modules"};
  const TYPE_AB={Supplier:"Fournisseur",Customer:"Client"};
  const HIDE_BY={Actions:["State"]};
  const DLT={"-1":"Selon l'interface","0":"Jours","1":"Services","2":"Heures","3":"Minutes"};
  const DAYS=["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"];
  const DOW={Monday:"Lun",Tuesday:"Mar",Wednesday:"Mer",Thursday:"Jeu",Friday:"Ven",Saturday:"Sam",Sunday:"Dim"};
  const OFF_BY_DEFAULT=new Set(["Options","Company"]);
  /* dates ISO -> lisibles ; année 0001 = « non renseigné » ; heures seules pour les débuts de moment */
  const ISO=/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/;
  function fmtDate(v,tag){
    const m=ISO.exec(v);if(!m)return v;
    if(+m[1]<1900)return "";
    if(/^Time(From|To|Start|End)?$|^Heure/.test(tag))return m[4]+":"+m[5];
    const d=m[3]+"/"+m[2]+"/"+m[1];
    return (m[4]==="00"&&m[5]==="00"&&(!m[6]||m[6]==="00"))?d:d+" "+m[4]+":"+m[5];
  }
  const ln=n=>n.replace(/^\{.*\}/,"").replace(/^.*:/,"");
  const kids=e=>Array.from(e.children);
  const txt=e=>(e.textContent||"").trim();

  /* ---- ZIP (lecture) ---- */
  async function inflate(raw){
    const ds=new DecompressionStream("deflate-raw");
    const w=ds.writable.getWriter();w.write(raw);w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
  }
  async function unzip(b){
    const dv=new DataView(b.buffer,b.byteOffset,b.byteLength);let e=-1;
    for(let i=b.length-22;i>=Math.max(0,b.length-70000);i--){if(dv.getUint32(i,true)===0x06054b50){e=i;break}}
    if(e<0)throw new Error("Ce fichier n'est pas une archive ZIP valide.");
    const n=dv.getUint16(e+10,true);let p=dv.getUint32(e+16,true);const out=[];
    for(let k=0;k<n;k++){
      if(dv.getUint32(p,true)!==0x02014b50)break;
      const flag=dv.getUint16(p+8,true),meth=dv.getUint16(p+10,true),csz=dv.getUint32(p+20,true),usz=dv.getUint32(p+24,true);
      const nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),lo=dv.getUint32(p+42,true);
      const name=new TextDecoder(flag&2048?"utf-8":"utf-8").decode(b.subarray(p+46,p+46+nl));
      p+=46+nl+el+cl;
      if(name.endsWith("/"))continue;
      const lnl=dv.getUint16(lo+26,true),lel=dv.getUint16(lo+28,true),st=lo+30+lnl+lel;
      const raw=b.subarray(st,st+csz);
      let data;
      if(meth===0)data=raw;else if(meth===8)data=await inflate(raw);else continue;
      out.push({name,bytes:data});
    }
    return out;
  }

  /* ---- XML ---- */
  function decode(bytes){
    let s=new TextDecoder("utf-8").decode(bytes).replace(/^﻿/,"");
    const m=/^<\?xml[^>]*encoding=["']([^"']+)/i.exec(s);
    if(m&&!/utf-?8/i.test(m[1])){try{s=new TextDecoder(m[1]).decode(bytes).replace(/^﻿/,"")}catch(e){}}
    return s;
  }
  function parse(bytes){
    const d=new DOMParser().parseFromString(decode(bytes),"application/xml");
    if(d.getElementsByTagName("parsererror").length)return null;
    return d.documentElement;
  }
  const hum=t=>LBL[t]||t.replace(/_/g," ").replace(/([a-zà-ÿ])([A-Z])/g,"$1 $2");
  const base=n=>n.split("/").pop().replace(/\.xml$/i,"");

  /* ---- Valeurs ---- */
  function recur(r,mom){
    const ty=txt(r.querySelector(":scope > Type")||{textContent:""});
    const nb=parseInt(txt(r.querySelector(":scope > Nb")||{textContent:"0"}),10)||0;
    const dis=txt(r.querySelector(":scope > IsDisable")||{textContent:""})==="true";
    const ms=[];const UNK="(moment supprimé)";
    kids(r).filter(c=>ln(c.tagName)==="Moments").forEach(M=>kids(M).forEach(m=>{
      const id=txt(m.querySelector(":scope > MomentId")||{textContent:""});
      const days=[1,2,3,4,5,6,7].map(i=>txt(m.querySelector(":scope > Day"+i)||{textContent:""})==="true");
      ms.push({n:mom[id]||UNK,days});
    }));
    let s;
    if(ty==="None"||!ty)return dis?"Aucune (désactivée)":"Aucune";
    if(ty==="xDays")s="Tous les "+nb+" jour"+(nb>1?"s":"");
    else if(ty==="Daily")s="Quotidienne";
    else s=ty+(nb?" ("+nb+")":"");
    const parts=ms.map(m=>{
      const on=m.days.map((d,i)=>d?DAYS[i]:null).filter(Boolean);
      if(!on.length)return null;
      return m.n+" ("+(on.length===7?"tous les jours":on.join(", "))+")";
    }).filter(Boolean);
    if(ty==="xDays"&&!parts.length&&ms.length)return s+" – "+ms.map(m=>m.n).join(", ")+(dis?" [désactivée]":"");
    if(parts.length){const cnt=new Map();parts.forEach(p=>cnt.set(p,(cnt.get(p)||0)+1));s+=" – "+[...cnt].map(([p,n])=>n>1?p+" ×"+n:p).join(" ; ")}
    else if(ty==="Daily")s+=" – aucun jour coché";
    return s+(dis?" [désactivée]":"");
  }
  function generic(e,mom,depth){
    const ks=kids(e);
    if(!ks.length)return txt(e);
    return ks.map(c=>{const v=generic(c,mom,(depth||0)+1);return ks.length>1&&kids(c).length===0?hum(ln(c.tagName))+" : "+v:v}).filter(v=>v!=="").join(" ; ");
  }
  function cell(tag,e,mom,type){
    const k=kids(e);
    if(tag==="Recurrence")return recur(e,mom);
    if(tag==="Jours")return k.map(c=>DOW[txt(c)]||txt(c)).join(", ");
    if(tag==="DLType")return DLT[txt(e)]??txt(e);
    if(tag==="MomentId"||tag==="Moment"){const q=txt(e);return q?(mom[q]||"(moment supprimé)"):""}
    if(k.length)return generic(e,mom,0);
    let v=txt(e).replace(/[\uE000-\uF8FF]/g,"").trim();
    if(tag==="Language_Default")v=v.replace(/^Language\./,"");
    if(tag==="TypeFunction"&&FUNC[v])return FUNC[v];
    if(tag==="Type"&&TYPE_AB[v])return TYPE_AB[v];
    if(ISO.test(v))return fmtDate(v,tag);
    if(v==="true")return "Oui";if(v==="false")return "Non";
    if(/^-?\d+(\.\d+)?$/.test(v)&&v.length<16&&!/^0\d/.test(v))return Number(v);
    return v;
  }

  /* ---- Conversion d'un module ---- */
  function convert(name,root,mom,opt){
    const rn=ln(root.tagName);
    const sheetName=SHEETS[name]||hum(name);
    // fichiers de type « une fiche » (DefApp, DefCompany) : Paramètre / Valeur
    if(!/^ArrayOf/.test(rn)){
      const rows=kids(root).filter(c=>!kids(c).length&&!/password|motdepasse/i.test(ln(c.tagName))).map(c=>[ln(c.tagName),cell(ln(c.tagName),c,mom)]);
      return {name,sheet:sheetName,head:["Paramètre","Valeur"],rows,count:rows.length};
    }
    let recs=kids(root).filter(r=>txt(r.querySelector(":scope > Delete")||{textContent:"false"})!=="true");
    if(name==="Moments")recs=recs;
    const cols=[],seen=new Set();
    const rows=[];
    const expand=name==="CheckLists";
    const addCol=t=>{if(!seen.has(t)){seen.add(t);cols.push(t)}};
    const recRows=[];
    recs.forEach(r=>{
      const o={};
      kids(r).forEach(c=>{
        let t=ln(c.tagName);
        if(HIDE.has(t)||(HIDE_BY[name]&&HIDE_BY[name].includes(t)))return;
        if(name==="Usernames"&&/^CanAccess_/.test(t)){
          if(txt(c)==="true"){o.__acc=(o.__acc?o.__acc+", ":"")+t.replace("CanAccess_","");}
          return;
        }
        if(expand&&t==="Items")return;
        o[t]=cell(t,c,mom);addCol(t);
      });
      if(o.__acc){o["Accès aux modules"]=o.__acc;addCol("Accès aux modules")}
      delete o.__acc;
      if(expand){
        const it=r.querySelector(":scope > Items");
        const its=it?kids(it).filter(i=>txt(i.querySelector(":scope > Delete")||{textContent:"false"})!=="true"):[];
        addCol("Élément");addCol("Obligatoire");
        if(!its.length)recRows.push(Object.assign({},o,{"Élément":"","Obligatoire":""}));
        its.forEach(i=>recRows.push(Object.assign({},o,{"Élément":txt(i.querySelector(":scope > Name")||{textContent:""}),"Obligatoire":txt(i.querySelector(":scope > IsMandatory")||{textContent:""})==="true"?"Oui":"Non"})));
      }else recRows.push(o);
    });
    const used=cols.filter(t=>recRows.some(o=>o[t]!==undefined&&o[t]!==""&&o[t]!=="Non"||(o[t]==="Non"&&false)));
    // une colonne de « Non » uniquement est du bruit (indicateurs inactifs) : on la retire
    const keep=used.filter(t=>recRows.some(o=>o[t]!==undefined&&o[t]!==""&&o[t]!=="Non")||/^Obligatoire$/.test(t));
    // Désignation en premier
    keep.sort((a,b)=>(a==="Name"?-1:0)-(b==="Name"?-1:0)||(a==="Category"?-1:0)-(b==="Category"?-1:0)||0);
    const stable=cols.filter(t=>keep.includes(t));
    const ordered=["Name","Category"].filter(t=>stable.includes(t)).concat(stable.filter(t=>t!=="Name"&&t!=="Category"));
    const head=ordered.map(t=>t==="Name"?(name==="Moments"?"Moment":name==="Usernames"?"Nom":hum(t)):hum(t));
    const body=recRows.map(o=>ordered.map(t=>o[t]===undefined?"":o[t]));
    if(name==="OptionsUI"){const pi=ordered.indexOf("ModulePosition");if(pi>=0)body.sort((a,b)=>(Number(a[pi])||0)-(Number(b[pi])||0)||String(a[0]).localeCompare(String(b[0]),"fr"))}
    return {name,sheet:sheetName,head,rows:body,count:recs.length};
  }

  function sheetTitle(s,used){
    let t=s.replace(/[\[\]:*?\/\\]/g," ").trim().slice(0,31)||"Feuille";
    let k=2,b=t;while(used.has(t.toLowerCase())){const suf=" ("+k++ +")";t=b.slice(0,31-suf.length)+suf}
    used.add(t.toLowerCase());return t;
  }

  async function load(files){
    // files : [{name, bytes}] — zips ou xml
    const xml=[];
    for(const f of files){
      if(/\.zip$/i.test(f.name)){(await unzip(f.bytes)).forEach(z=>{if(/\.xml$/i.test(z.name)&&!/__MACOSX/.test(z.name))xml.push(z)})}
      else if(/\.xml$/i.test(f.name))xml.push(f);
    }
    const mom={};
    const mf=xml.find(x=>base(x.name).toLowerCase()==="moments");
    if(mf){const r=parse(mf.bytes);if(r)kids(r).forEach(m=>{const id=txt(m.querySelector(":scope > Id")||{textContent:""}),n=txt(m.querySelector(":scope > Name")||{textContent:""});if(id)mom[id]=n})}
    const mods=[],bad=[];
    xml.forEach(x=>{
      const nm=base(x.name);const r=parse(x.bytes);
      if(!r){bad.push(nm);return}
      const m=convert(nm,r,mom);
      mods.push(m);
    });
    mods.sort((a,b)=>(b.name==="OptionsUI")-(a.name==="OptionsUI")||a.sheet.localeCompare(b.sheet,"fr"));
    return {mods,bad,moments:mf?Object.keys(mom).length:0};
  }

  function build(mods,opt){
    const wb=XLSX.utils.book_new(),used=new Set();
    const titles=[];
    mods.forEach(m=>{
      const ws=XLSX.utils.aoa_to_sheet([m.head].concat(m.rows));
      ws["!cols"]=m.head.map((h,i)=>{let w=String(h).length;m.rows.slice(0,300).forEach(r=>{const l=String(r[i]??"").length;if(l>w)w=l});return {wch:Math.min(60,Math.max(10,w+2))}});
      ws["!freeze"]={xSplit:0,ySplit:1};
      if(m.head.length)ws["!autofilter"]={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(0,m.rows.length),c:m.head.length-1}})};
      const tt=sheetTitle(m.sheet,used);titles.push(tt);
      XLSX.utils.book_append_sheet(wb,ws,tt);
    });
    if(opt&&opt.summary&&mods.length>1){
      const rows=[["Module","Onglet","Lignes"]].concat(mods.map((m,i)=>[m.name,titles[i],m.rows.length]));
      const sm=XLSX.utils.aoa_to_sheet(rows);
      mods.forEach((m,i)=>{sm["B"+(i+2)].l={Target:"#'"+titles[i].replace(/'/g,"''")+"'!A1",Tooltip:"Aller à l'onglet"}});
      sm["!cols"]=[{wch:28},{wch:34},{wch:9}];
      const sn=sheetTitle("Sommaire",used);XLSX.utils.book_append_sheet(wb,sm,sn);
      wb.SheetNames.unshift(wb.SheetNames.pop());
    }
    return new Uint8Array(XLSX.write(wb,{bookType:"xlsx",type:"array"}));
  }
  return {unzip,load,build,convert,SHEETS,OFF_BY_DEFAULT};
})();

function openX2E(){
  const st={showEmpty:false,prev:"",summary:false,mods:null,sel:new Set(),name:"",err:"",msg:"",bad:[],moments:0,busy:false,lastBytes:null,lastName:""};
  const bg=$("div","modal-bg");bg.setAttribute("role","dialog");bg.setAttribute("aria-modal","true");bg.setAttribute("aria-labelledby","x2t");
  const md=$("div","modal xml");bg.appendChild(md);document.body.appendChild(bg);
  const close=()=>{bg.remove();document.removeEventListener("keydown",kd)};
  const kd=e=>{if(e.key==="Escape")close()};document.addEventListener("keydown",kd);
  const readFile=f=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(new Uint8Array(r.result));r.onerror=()=>rej(r.error);r.readAsArrayBuffer(f)});
  async function onFiles(fl){
    st.err="";st.msg="";st.busy=true;paint();
    try{
      const files=[];for(const f of fl)files.push({name:f.name,bytes:await readFile(f)});
      const r=await X2E.load(files);
      if(!r.mods.length)throw new Error("Aucun fichier XML trouvé dans ce que tu as déposé.");
      st.mods=r.mods;st.bad=r.bad;st.moments=r.moments;
      st.name=(fl.length===1?fl[0].name.replace(/\.(zip|xml)$/i,""):"Configuration");
      st.zip=fl.length===1&&/\.zip$/i.test(fl[0].name);
      st.sel=new Set(r.mods.filter(m=>m.count>0&&!(r.mods.length>1&&X2E.OFF_BY_DEFAULT.has(m.name))).map(m=>m.name));
    }catch(e){st.mods=null;st.err=(e&&e.message)||"Lecture impossible."}
    st.busy=false;paint();
  }
  const pick=()=>{
    const inp=$("input");inp.type="file";inp.accept=".zip,.xml,application/zip,text/xml";inp.multiple=true;inp.hidden=true;
    const d=$("button","drop");d.type="button";d.append($("strong",null,"Archive ZIP de la configuration (ou fichier XML)"),$("span",null,"Dépose le fichier ici ou clique pour le choisir"));
    d.onclick=()=>inp.click();
    d.ondragover=e=>{e.preventDefault();d.classList.add("over")};d.ondragleave=()=>d.classList.remove("over");
    d.ondrop=e=>{e.preventDefault();d.classList.remove("over");if(e.dataTransfer.files.length)onFiles(e.dataTransfer.files)};
    inp.onchange=()=>{if(inp.files.length)onFiles(inp.files)};
    const w=$("div");w.append(d,inp);return w;
  };
  function chosen(){return st.mods.filter(m=>st.sel.has(m.name))}
  /* « 20210406242 - BOULANGERIE X - VILLE.zip » -> « BOULANGERIE X - VILLE - Configuration de la solution.xlsx » */
  function clientOf(n){
    let c=n.replace(/\.zip$/i,"").replace(/\s*\(\d+\)\s*$/,"");
    c=c.replace(/^[\s\d]+[\s_-]+/,"");
    if(!/\s/.test(c))c=c.replace(/_+-?_*/g,m=>m.includes("-")?" - ":" ");
    else c=c.replace(/_/g," ");
    return c.replace(/[\\\/:*?"<>|]/g," ").replace(/\s+/g," ").replace(/^[\s-]+|[\s-]+$/g,"");
  }
  function fname(){
    const c=chosen();
    if(st.zip){const cl=clientOf(st.name);if(cl)return cl+" - Configuration de la solution.xlsx"}
    const b=c.length===1?c[0].sheet:st.name.replace(/[\\\/:*?"<>|]/g," ").trim()||"Configuration";
    return b+".xlsx";
  }
  async function dl(){
    st.msg="";const c=chosen();
    if(!c.length){st.msg="Coche au moins un module.";paint();return}
    let data;
    try{data=X2E.build(c,{summary:st.summary})}catch(e){st.msg="Création du fichier Excel impossible : "+((e&&e.message)||"erreur");paint();return}
    const filename=fname();
    try{
      const d=window.claude&&typeof window.claude.use==="function"?await window.claude.use("downloads"):null;
      if(!d)throw {code:"nodl"};
      await d.save({filename,data});
      st.msg="Fichier téléchargé : "+filename+" ("+c.length+" onglet"+(c.length>1?"s":"")+").";
    }catch(e){
      if(e&&e.code==="declined")st.msg="Téléchargement annulé.";
      else{
        try{const u=URL.createObjectURL(new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));const a=document.createElement("a");a.href=u;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);st.msg="Téléchargement lancé : "+filename+"."}
        catch(x){st.msg="Téléchargement impossible."}
      }
    }
    paint();
  }
  function paint(){
    md.textContent="";
    const h=$("h2",null,"XML → Excel");h.id="x2t";
    const hd=$("div","xhead");const x=$("button","xclose","×");x.type="button";x.setAttribute("aria-label","Fermer");x.onclick=close;hd.append(h,x);
    md.append(hd,$("p",null,"Dépose le ZIP de la configuration du client (ou un seul XML) : je produis un Excel avec un onglet par module. Rien n'est envoyé, tout reste dans ton navigateur."));
    if(st.busy){md.appendChild($("p","msg","Lecture en cours…"));return}
    if(!st.mods){
      if(st.err){const e=$("p","msg err",st.err);e.setAttribute("role","alert");md.appendChild(e)}
      md.appendChild(pick());
      const ac=$("div","actions");const no=$("button","btn ghost","Fermer");no.onclick=close;ac.appendChild(no);md.appendChild(ac);
      return;
    }
    const nonEmpty=st.mods.filter(m=>m.count>0).length;
    md.appendChild($("p",null,st.mods.length+" module(s) trouvé(s), dont "+nonEmpty+" avec des données."+(st.moments?" Les moments sont affichés par leur nom.":"")));
    const lst=$("div","x2list");
    st.mods.filter(m=>m.count>0||st.showEmpty).forEach(m=>{
      const l=$("label","x2row"+(m.count?"":" empty"));
      const cb=$("input");cb.type="checkbox";cb.checked=st.sel.has(m.name);cb.onchange=()=>{cb.checked?st.sel.add(m.name):st.sel.delete(m.name);paint()};
      const pv=$("button","chip x2pv"+(st.prev===m.name?" on":""),"Aperçu");pv.type="button";pv.disabled=!m.count;pv.onclick=e=>{e.preventDefault();e.stopPropagation();st.prev=st.prev===m.name?"":m.name;paint()};
      l.append(cb,$("span","x2n",m.sheet),$("small",null,m.count?m.count+" ligne"+(m.count>1?"s":""):"vide"),pv);
      lst.appendChild(l);
    });
    md.appendChild(lst);
    if(nonEmpty<st.mods.length){const te=$("button","chip",st.showEmpty?"Masquer les modules vides":"Afficher les "+(st.mods.length-nonEmpty)+" modules vides");te.type="button";te.style.marginTop="8px";te.style.alignSelf="flex-start";te.onclick=()=>{st.showEmpty=!st.showEmpty;paint()};md.appendChild(te)}
    {const pm=st.mods.find(m=>m.name===st.prev);if(pm){
      const bx=$("div","x2prev");bx.appendChild($("p","x2pt","Aperçu : "+pm.sheet+" ("+pm.rows.length+" ligne"+(pm.rows.length>1?"s":"")+")"));
      const sc=$("div","x2sc"),tb=$("table");const th=$("tr");pm.head.forEach(h=>th.appendChild($("th",null,String(h))));tb.appendChild(th);
      pm.rows.slice(0,5).forEach(r=>{const tr=$("tr");r.forEach(v=>tr.appendChild($("td",null,String(v==null?"":v).slice(0,70))));tb.appendChild(tr)});
      sc.appendChild(tb);bx.appendChild(sc);md.appendChild(bx)}}
    const q=$("div","actions");
    const a1=$("button","chip","Tout cocher");a1.type="button";a1.onclick=()=>{st.sel=new Set(st.mods.filter(m=>m.count>0).map(m=>m.name));paint()};
    const a2=$("button","chip","Tout décocher");a2.type="button";a2.onclick=()=>{st.sel=new Set();paint()};
    q.append(a1,a2);
    md.appendChild(q);
    md.appendChild($("p","ghint","Les éléments supprimés ne sont pas repris. Les mots de passe des utilisateurs ne sont jamais exportés."));
    if(st.bad.length)md.appendChild($("p","warnmsg","Fichier(s) illisible(s), ignoré(s) : "+st.bad.join(", ")));
    if(st.msg)md.appendChild($("p","msg",st.msg));
    const ac=$("div","actions");
    const c=chosen();
    const d=$("button","btn ok","Télécharger l'Excel ("+c.length+" onglet"+(c.length>1?"s":"")+")");d.disabled=!c.length;d.onclick=dl;
    const ch=$("button","btn ghost","Changer de fichier");ch.onclick=()=>{st.mods=null;st.msg="";paint()};
    const no=$("button","btn ghost","Fermer");no.onclick=close;
    ac.append(d,ch,no);md.appendChild(ac);
  }
  paint();
}
/*X2E-END*/
