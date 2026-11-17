const BILAN={annee:null,cache:{}},MOIS_COURTS=["janv.","févr.","mars","avr.","mai","juin","juil.","août","sept.","oct.","nov.","déc."],moisCourt=s=>MOIS_COURTS[+s.slice(5)-1],moisLong=s=>MOIS_MOTS[+s.slice(5)-1],prevCle=s=>{let[p,c]=s.split("-").map(Number);return c--,c<1&&(c=12,p--),p+"-"+String(c).padStart(2,"0")},signe=s=>(s>.004?"+":s<-.004?"−":"")+$$(Math.abs(s)),pct1=s=>s==null?"—":(Math.round(s*1e3)/10).toString().replace(".",",")+" %",ecartTxt=(s,p,c)=>{if(!p)return"";const r=s-p,u=r/Math.abs(p);return`<span class="${r>0==!!c?"bl-bon":"bl-mauvais"}">${r>0?"▲":"▼"} ${Math.abs(Math.round(u*100))} %</span>`};async function vueBilanPro(s){const p=BILAN.annee||+today().slice(0,4);s.innerHTML='<button class="p-retour" id="pRet">‹ Plus</button><div id="blZ"><div class="card"><div class="empty">Calcul du bilan…</div></div></div>',$("#pRet",s).onclick=()=>vaVers("plus");let c=BILAN.cache[p];try{c||(c=BILAN.cache[p]=await api("/api/bilan?annee="+p))}catch(l){$("#blZ",s).innerHTML=`<div class="banner err">Bilan impossible : ${esc(l.message)}</div>`;return}const r=$("#blZ",s);if(!r||!r.isConnected)return;r.innerHTML=bilanHTML(c),r.querySelectorAll("[data-an]").forEach(l=>l.onclick=()=>{BILAN.annee=+l.dataset.an,vueBilanPro(s)});const u=$("#blMois",r);u&&graphBilan(u,c.mois)}function bilanHTML(s){const p=s.annees.indexOf(s.annee),c=s.annees[p-1],r=s.annees[p+1],u=`<div class="card bl-tete"><button type="button" class="bl-nav" ${c?`data-an="${c}"`:"disabled"} aria-label="Année précédente">‹</button>
    <div><div class="lbl">Bilan de l'année</div><div class="bl-an">${s.annee}</div>
    ${s.vide?"":`<div class="small muted">${s.n===12?"Année complète":`${moisLong(s.de)} à ${moisLong(s.a)}`}${s.en_cours?` · à ce jour (${dateLongue(s.auj)})`:""}</div>`}</div>
    <button type="button" class="bl-nav" ${r?`data-an="${r}"`:"disabled"} aria-label="Année suivante">›</button></div>`;if(s.vide)return u+`<div class="spacer"></div><div class="card"><div class="empty">Aucune donnée pour ${s.annee}.</div></div>`;const l=s.totaux,n=s.comparaison,b=(e,f,h,x="")=>`<div class="bl-k"><div class="lbl">${e}</div><div class="v num ${x}">${f}</div>${h?`<div class="s">${h}</div>`:""}</div>`,g=(e,f,h)=>n?`${ecartTxt(e,f,h)} vs ${n.n===s.n?n.annee:`${moisCourt(n.de)}–${moisCourt(n.a)} ${n.annee}`}`:"";let m=u+`<div class="spacer"></div>
   <div class="bl-kpis">
     ${b("Revenus",$$(l.revenus),`${$$(l.moyenne_revenus)} / mois ${n&&n.n===s.n?g(l.revenus,n.revenus,!0):""}`)}
     ${b("Dépenses (ta part)",$$(l.depenses),`${$$(l.moyenne_depenses)} / mois ${n&&n.n===s.n?g(l.depenses,n.depenses,!1):""}`)}
     ${b("Revenu net",signe(l.net),`Taux d'épargne ${pct1(l.taux)}`,l.net<0?"neg":"pos")}
     ${s.placements?b("Placements",$$(s.placements.fin),`Rendement ${signe(s.placements.rendement)}`):b("Placé",$$(l.epargne_placee),"Versements de l'année")}
   </div>
   ${n&&n.n!==s.n?`<div class="card bl-comp"><div class="lbl">${s.annee} comparé à ${n.annee}, mêmes mois (${n.n===1?esc(moisLong(n.de)):`${esc(moisCourt(n.de))} à ${esc(moisCourt(n.a))}`})</div>
     <div class="bl-l"><span>Revenus</span><span><b class="num">${$$(n.revenus_ici)}</b> <small class="muted">(${$$(n.revenus)})</small> ${ecartTxt(n.revenus_ici,n.revenus,!0)}</span></div>
     <div class="bl-l"><span>Dépenses</span><span><b class="num">${$$(n.depenses_ici)}</b> <small class="muted">(${$$(n.depenses)})</small> ${ecartTxt(n.depenses_ici,n.depenses,!1)}</span></div></div>`:""}
   <div class="spacer"></div>`;m+=`<div class="card cadre"><h3>Mois par mois</h3><div id="blMois"></div>
    <div class="legend"><span><i style="background:var(--pos)"></i>Revenus</span><span><i style="background:var(--neg)"></i>Dépenses</span></div>
    <div class="small muted" style="margin-top:6px">${s.n_complets?`${s.positifs} mois sur ${s.n_complets} avec un revenu net positif${s.en_cours?" (mois en cours exclu)":""}.`:""}
    ${s.extremes?` Le plus cher : <b>${esc(moisLong(s.extremes.plus.cle))}</b> (${$$(s.extremes.plus.depenses)}) ; le moins cher : <b>${esc(moisLong(s.extremes.moins.cle))}</b> (${$$(s.extremes.moins.depenses)}).`:""}</div></div>
   <div class="spacer"></div>`;const t=s.revenus,v=t.pourboires;m+=`<div class="card"><h3>Revenus <span class="sp num">${$$(t.total)}</span></h3>
    ${t.principal?`<div class="bl-l"><span>${esc(t.nom_principal)}${t.detail_de?` <small class="muted">depuis ${esc(moisLong(t.detail_de))}</small>`:""}</span><b class="num">${$$(t.principal)}</b></div>`:""}
    ${v?`<div class="bl-sous">
       ${v.de>s.de?`<div class="bl-l"><span><small class="muted">Pourboires depuis ${esc(moisLong(v.de))}</small></span></div>`:""}
       <div class="bl-l"><span>Pourboires en espèces</span><b class="num">${$$(v.especes)}</b></div>
       <div class="bl-l"><span>Pourboires par carte</span><b class="num">${$$(v.carte)}</b></div>
       <div class="bl-l"><span>${v.quarts} quart${v.quarts>1?"s":""}${v.heures?` · ${String(v.heures).replace(".",",")} h`:""}</span><b class="num">${v.par_heure?$$(v.par_heure)+" / h":""}</b></div></div>`:""}
    ${t.groupes.filter(e=>!e.principal).map(e=>`<div class="bl-l"><span>${esc(e.nom)}${e.n>1?` <small class="muted">× ${e.n}</small>`:""}</span><b class="num">${$$(e.total)}</b></div>`).join("")}
    ${t.groupes.some(e=>e.nom==="Revenus repris de l'Excel")?`<p class="hint" style="margin:8px 0 0">Les mois repris de ton ancien Excel n'ont qu'un total de revenus (sans détail).</p>`:""}</div>
   <div class="spacer"></div>`;const P=Math.max(...s.categories.map(e=>e.total),1);m+=`<div class="card"><h3>Dépenses par catégorie <span class="sp num">${$$(l.depenses)}</span></h3>
    ${s.categories.map(e=>`<div class="bl-cat"><div class="bl-l"><span>${iconeCat(e.nom)} ${esc(e.nom)}</span><b class="num">${$$(e.total)}</b></div>
      <div class="bl-barre"><i style="width:${Math.max(1,e.total/P*100).toFixed(1)}%"></i></div>
      <div class="bl-meta"><span>${pct1(e.pct)} · ${$$(e.moyenne)} / mois</span>${e.prec!=null&&e.prec>.004?`<span>${ecartTxt(e.total,e.prec,!1)} vs ${n.annee}</span>`:""}</div></div>`).join("")}</div>
   <div class="spacer"></div>`,s.marchands.length&&(m+=`<div class="card"><h3>Où va ton argent</h3>
    ${s.marchands.map((e,f)=>`<div class="bl-l"><span><b class="bl-rang">${f+1}</b>${esc(e.nom)} <small class="muted">× ${e.n}</small></span><b class="num">${$$(e.total)}</b></div>`).join("")}
    ${s.grosses.length?`<div class="bl-t">Plus grosses dépenses</div>${s.grosses.map(e=>`<div class="bl-l"><span>${esc(e.description)} <small class="muted">${esc(dateLongue(e.date))}</small></span><b class="num">${$$(e.montant)}</b></div>`).join("")}`:""}</div>
   <div class="spacer"></div>`);const i=s.partage;i&&MOD("partage")&&(m+=`<div class="card"><h3>Partage avec ${esc(i.partenaire)}</h3>
    <div class="bl-l"><span>Dépenses partagées que tu as payées <small class="muted">× ${i.n}</small></span><b class="num">${$$(i.total)}</b></div>
    <div class="bl-sous"><div class="bl-l"><span>Ta part</span><b class="num">${$$(i.ta_part)}</b></div>
      <div class="bl-l"><span>Part de ${esc(i.partenaire)}</span><b class="num">${$$(i.sa_part)}</b></div></div>
    ${i.n_lui?`<div class="bl-l"><span>Payé par ${esc(i.partenaire)} pour vous deux <small class="muted">× ${i.n_lui}</small></span><b class="num">${$$(i.payees_par_lui)}</b></div>`:""}
    <div class="bl-l"><span>Soldes réglés</span><b class="num">${i.reglements}</b></div>
    ${i.du_aujourdhui!=null&&Math.abs(i.du_aujourdhui)>=.005?`<div class="bl-l"><span>${i.du_aujourdhui>0?esc(i.partenaire)+" te doit aujourd'hui":"Tu dois à "+esc(i.partenaire)+" aujourd'hui"}</span><b class="num">${$$(Math.abs(i.du_aujourdhui))}</b></div>`:""}</div>
   <div class="spacer"></div>`);const d=s.placements;d&&MOD("placements")&&(m+=`<div class="card"><h3>Placements</h3>
    <div class="bl-pl"><div><div class="lbl">1<sup>er</sup> ${esc(moisLong(d.de))}</div><b class="num">${$$(d.debut)}</b></div><div class="bl-fl">→</div>
      <div><div class="lbl">${d.a===s.a&&s.en_cours?"Maintenant":`1<sup>er</sup> ${esc(moisLong(d.a))}`}</div><b class="num">${$$(d.fin)}</b></div></div>
    <div class="bl-l"><span>Versé pendant l'année</span><b class="num">${signe(d.verse)}</b></div>
    <div class="bl-l"><span>Rendement${d.rendement_pct!=null?` <small class="muted">≈ ${pct1(d.rendement_pct)}</small>`:""}</span><b class="num ${d.rendement<0?"neg":"pos"}">${signe(d.rendement)}</b></div>
    ${d.comptes.length>1?`<div class="bl-t">Par compte</div>${d.comptes.map(e=>`<div class="bl-cpt"><div class="bl-l"><span><b>${esc(e.nom)}</b></span><b class="num">${$$(e.fin)}</b></div>
      <div class="bl-meta"><span>${$$(e.debut)} au départ · versé ${$$(e.verse)}</span>${e.fixe?"":`<span class="${e.rendement<0?"neg":"pos"}">${signe(e.rendement)}</span>`}</div></div>`).join("")}`:""}
    <p class="hint" style="margin:8px 0 0">Valeur au 1<sup>er</sup> de chaque mois (+ versements ajoutés). Les versements ne comptent pas comme rendement.</p></div>
   <div class="spacer"></div>`);const y=s.patrimoine;m+=`<div class="card"><h3>Compte et patrimoine</h3>
    <div class="bl-l"><span>Compte au 1<sup>er</sup> ${esc(moisLong(s.de))}</span><b class="num">${$$(s.compte.debut)}</b></div>
    <div class="bl-l"><span>Compte ${s.en_cours?"aujourd'hui":"à la fin de "+esc(moisLong(s.a))}</span><b class="num">${$$(s.compte.fin)}</b></div>
    ${s.financements.map(e=>`<div class="bl-l"><span>${esc(e.nom)} : solde dû${e.recu_annee?` <small class="muted">(reçu cette année : ${$$(e.recu_annee)})</small>`:""}</span><b class="num neg">−${$$(e.solde_du)}</b></div>`).join("")}
    <div class="bl-l bl-tot"><span>Valeur nette <small class="muted">compte + placements − dettes</small></span><b class="num">${$$(y.net)}</b></div></div>
   <div class="spacer"></div>`;const o=s.impots;return m+=`<div class="card"><h3>Pour tes impôts ${s.annee}</h3>
    <p class="small muted" style="margin:0 0 8px">Les chiffres de l'app, à comparer avec tes feuillets (T4, RL-1, relevés de placements). Ce n'est pas un avis fiscal.</p>
    ${t.detail_de?`<p class="small" style="margin:0 0 6px"><b>Attention</b> : de ${esc(moisLong(s.de))} à ${esc(moisLong(prevCle(t.detail_de)))}, l'app n'a que les totaux repris de ton Excel. Les montants ci-dessous commencent en ${esc(moisLong(t.detail_de))} : ils ne remplacent pas tes feuillets.</p>`:""}
    ${t.principal?`<div class="bl-l"><span>${esc(t.nom_principal)}${t.detail_de?` <small class="muted">depuis ${esc(moisLong(t.detail_de))}</small>`:""}</span><b class="num">${$$(t.principal)}</b></div>`:""}
    ${o.pourboires?`<div class="bl-sous"><div class="bl-l"><span>dont pourboires en espèces</span><b class="num">${$$(o.pourboires.especes)}</b></div>
      <div class="bl-l"><span>dont pourboires par carte</span><b class="num">${$$(o.pourboires.carte)}</b></div></div>`:""}
    ${o.celiapp?`<div class="bl-l"><span>Versé au CÉLIAPP</span><b class="num">${$$(o.celiapp)}</b></div>`:""}
    ${o.celi?`<div class="bl-l"><span>Versé au CÉLI</span><b class="num">${$$(o.celi)}</b></div>`:""}
    ${o.reer?`<div class="bl-l"><span>Versé au REER</span><b class="num">${$$(o.reer)}</b></div>`:""}
    ${o.autres.length?`<div class="bl-t">Autres revenus</div>${o.autres.map(e=>`<div class="bl-l"><span>${esc(e.nom)}</span><b class="num">${$$(e.total)}</b></div>`).join("")}`:""}
    ${s.en_cours&&+s.auj.slice(5,7)>=10&&MOD("placements")?`<p class="hint" style="margin:8px 0 0">Un versement au CÉLIAPP compte pour l'année où il est fait : jusqu'au 31 décembre pour ${s.annee}.</p>`:""}</div>`,m}function graphBilan(s,p){const c=p.length,r=Math.max(300,Math.round(s.clientWidth||600)),u=220,l=44,n=6,b=10,g=26,m=Math.max(...p.flatMap(a=>[a.revenus,a.depenses]),1),t=Math.pow(10,Math.floor(Math.log10(m/4))),v=[1,2,2.5,5,10].map(a=>a*t).find(a=>a*4>=m*1.04),P=v*4,i=a=>u-g-(u-g-b)*Math.max(0,a)/P,d=(r-l-n)/c,y=Math.max(3,Math.min(18,d*.32)),o=2,e=a=>(Math.round(a/100)/10).toString().replace(".",",")+" k$",f=(a,_,M,L)=>{const j=i(_),C=u-g-j,R=Math.min(3,C/2);return C<=0?"":`<path d="M${a},${u-g} V${j+R} Q${a},${j} ${a+R},${j} H${a+y-R} Q${a+y},${j} ${a+y},${j+R} V${u-g} Z" fill="${M}" ${L?'fill-opacity=".45"':""}/>`};let h=`<svg viewBox="0 0 ${r} ${u}" width="100%" role="img" aria-label="Revenus et dépenses de chaque mois">`;for(let a=0;a<=4;a++){const _=P/4*a,M=i(_);h+=`<line x1="${l}" x2="${r-n}" y1="${M}" y2="${M}" stroke="var(--line)"/><text x="${l-6}" y="${M+4}" text-anchor="end">${e(_)}</text>`}const x=r<420&&c>8?2:1;p.forEach((a,_)=>{const M=l+d*_,L=M+d/2;h+=`<g>${f(L-y-o/2,a.revenus,"var(--pos)",a.partiel)}${f(L+o/2,a.depenses,"var(--neg)",a.partiel)}<rect x="${M}" y="${b}" width="${d}" height="${u-b}" fill="transparent"/><title>${esc(a.nom)}${a.partiel?" (en cours)":""}
Revenus : ${$$(a.revenus)}
Dépenses : ${$$(a.depenses)}
Revenu net : ${$$(a.net)}</title></g>`,(_%x===0||_===c-1)&&(h+=`<text x="${L}" y="${u-8}" text-anchor="middle">${esc(moisCourt(a.cle))}</text>`)}),s.innerHTML=h+"</svg>"}