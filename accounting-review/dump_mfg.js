const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();const p=await b.newPage();
await p.goto('file://'+__dirname+'/mccoys-fleet-forecast-2027.html');await p.waitForTimeout(1500);
const o=await p.evaluate(()=>{const A=DATA.adaptive;const out=[];
for(const r of SCEN.rows)for(const pt of (r.parts||[{loc:r.loc,op:r.op}])){if((A.fmt[pt.loc]||'op')!=='op')continue;
 const src=r.a.status==='Budget adjustment'?'Budget adjustment (accounting revision)':r.a.status==='Book only'?'Book-only depreciation (not on roster)':(RMAP[r.a.id]&&S.rpOn?'Unit (replaced in 2027)':(r.a.status==='Planned'?'Planned new unit':'Unit'));
 for(const c in pt.op){const v=pt.op[c].reduce((a,b)=>a+b,0);if(v)out.push([pt.loc,A.levels[pt.loc]||A.fleetLevel,c,src,v,r.a.id])}}
const ar=adaptiveRows();return {out,rows:ar.rows,proceeds:ar.proceeds,acc:A.accounts}});
require('fs').writeFileSync(process.argv[2],JSON.stringify(o));await b.close();})();
