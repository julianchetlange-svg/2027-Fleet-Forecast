const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();const p=await b.newPage();const e=[];p.on('pageerror',x=>e.push(x.message));
await p.goto('file://'+__dirname+'/mccoys-fleet-forecast-2027.html');await p.waitForTimeout(1500);
const o=await p.evaluate(()=>{const o={};for(const r of BASE.rows){if(r.a.status==='Budget adjustment')continue;for(const pt of (r.parts||[{loc:r.loc,op:r.op}])){for(const c in pt.op){const v=pt.op[c].reduce((a,b)=>a+b,0);if(!v)continue;(o[pt.loc]=o[pt.loc]||{});o[pt.loc][c]=(o[pt.loc][c]||0)+v}}}return o});
require('fs').writeFileSync(process.argv[2],JSON.stringify(o));console.log('ERR',e);await b.close();})();
