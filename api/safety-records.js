// v21.0: separate, protected safety records. QR/public access intentionally not enabled.
import { timingSafeEqual } from 'node:crypto';
const equal=(a,b)=>{const x=Buffer.from(String(a||'')),y=Buffer.from(String(b||''));return x.length===y.length&&timingSafeEqual(x,y);};
function role(req){const secret=String(req.headers.authorization||'').replace(/^Bearer\s+/i,'');if(process.env.SAFETY_MANAGER_KEY&&equal(secret,process.env.SAFETY_MANAGER_KEY))return 'manager';if(process.env.SAFETY_STAFF_KEY&&equal(secret,process.env.SAFETY_STAFF_KEY))return 'staff';return null;}
async function db(method,body){const url=process.env.SUPABASE_URL,key=process.env.SUPABASE_SERVICE_KEY;if(!url||!key)throw new Error('Supabase is not configured');const res=await fetch(`${url}/rest/v1/safety_data?id=eq.1${method==='GET'?'&select=id,payload,revision':''}`,{method,headers:{apikey:key,Authorization:`Bearer ${key}`,'Content-Type':'application/json',Prefer:'return=representation'},body:body?JSON.stringify(body):undefined});const txt=await res.text();if(!res.ok)throw new Error(`Safety storage request failed (${res.status}): ${txt.slice(0,300)}`);return txt?JSON.parse(txt):[];}
export default async function handler(req,res){res.setHeader('Cache-Control','no-store');const who=role(req);if(!who)return res.status(401).json({error:'Safety staff sign-in required'});if(!['GET','POST'].includes(req.method))return res.status(405).json({error:'Unsupported method'});
try{const rows=await db('GET');if(rows.length!==1)return res.status(409).json({error:'Missing safety_data row. Run the included SQL migration.'});const record=rows[0];if(req.method==='GET')return res.status(200).json({...record,role:who});
const {action,payload,revision,workerId,decision,reviewNote}=req.body||{};if(Number(revision)!==Number(record.revision))return res.status(409).json({error:'Someone updated these records. Refresh before saving.'});let next;
if(action==='save'){
if(!payload||!Array.isArray(payload.workers)||!Array.isArray(payload.orientations))return res.status(400).json({error:'Invalid record format'});
// Only a safety manager can change verification fields. Staff may maintain worker and orientation records.
const old=new Map((record.payload?.orientations||[]).map(o=>[String(o.id),o]));
for(const entry of payload.orientations){const previous=old.get(String(entry.id));if(previous&&(entry.approval_status!==previous.approval_status||entry.approved_by!==previous.approved_by||entry.approved_at!==previous.approved_at))return res.status(403).json({error:'Use the manager approval action to change orientation verification'});if(!previous&&entry.approval_status==='approved')return res.status(403).json({error:'New orientations cannot start approved'});}
// Protect approved records from being removed by ordinary staff.
for(const p of old.values())if(p.approval_status==='approved'&&!payload.orientations.some(o=>String(o.id)===String(p.id)))return res.status(403).json({error:'Approved orientations cannot be deleted'});
next={workers:payload.workers,orientations:payload.orientations};
}else if(action==='review'){
if(who!=='manager')return res.status(403).json({error:'Safety manager credentials required'});
if(!['approve','return'].includes(decision))return res.status(400).json({error:'Invalid review decision'});
const orientations=(record.payload?.orientations||[]).map(o=>{if(String(o.id)!==String(workerId))return o;if(decision==='approve'&&(o.status!=='completed'||!o.signature||!o.completed_at))throw new Error('Worker must complete and sign the orientation first');return {...o,approval_status:decision==='approve'?'approved':'returned',approved_at:decision==='approve'?new Date().toISOString():null,approved_by:decision==='approve'?'Safety manager':null,review_note:String(reviewNote||'').slice(0,1000),reviewed_at:new Date().toISOString()};});
if(!orientations.some(o=>String(o.id)===String(workerId)))return res.status(404).json({error:'Orientation not found'});
next={...(record.payload||{}),orientations};
}else return res.status(400).json({error:'Invalid action'});
// An optimistic version check; prevents most stale writes. For production-scale concurrency, move to a database RPC compare-and-swap.
const updated=await db('PATCH',{payload:next,revision:Number(record.revision)+1});if(updated.length!==1)return res.status(409).json({error:'Safety data not updated'});return res.status(200).json({...updated[0],role:who});
}catch(e){console.error('safety-records:',e);return res.status(500).json({error:e.message||'Safety records request failed'});}}
