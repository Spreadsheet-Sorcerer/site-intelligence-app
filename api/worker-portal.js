// v21.1 PREVIEW-ONLY worker self-service intake. No staff data is returned.
// Production requires identity verification, abuse throttling, separate private file storage and security review.
import {createHmac, randomUUID, timingSafeEqual} from 'node:crypto';
const allowedOrigins = null;
// Standalone portal: use the configured session key when available, otherwise
// reuse the existing server-only WORKER_DB_SECRET for session signing.
// This fallback does not grant additional database permissions.
const secret = () => process.env.SAFETY_PORTAL_SIGNING_KEY || process.env.WORKER_DB_SECRET;
const sign = s => createHmac('sha256',secret()).update(s).digest('base64url');
const safeEq=(a,b)=>{const x=Buffer.from(String(a||'')),y=Buffer.from(String(b||''));return x.length===y.length&&timingSafeEqual(x,y)};
function createToken(workerId){const payload=Buffer.from(JSON.stringify({workerId,exp:Date.now()+1000*60*60*12})).toString('base64url');return `${payload}.${sign(payload)}`;}
function decodeToken(token){try{const [p,mac]=String(token||'').split('.');if(!p||!mac||!safeEq(mac,sign(p)))return null;const o=JSON.parse(Buffer.from(p,'base64url'));return o.exp>Date.now()&&typeof o.workerId==='string'?o.workerId:null;}catch{return null;}}
const txt=(v,max=140)=>String(v??'').trim().slice(0,max);
const isoDate=(v)=> !v||/^\d{4}-\d{2}-\d{2}$/.test(v);
const normalizeIdentity=v=>txt(v).replace(/\s+/g,' ').trim();
// Restricted PostgREST RPC: uses publishable/anon key, NOT service-role key.
// The RPC validates WORKER_DB_SECRET and never returns the full safety_data payload.
async function portalDb(action,workerId,input={}){
 const root=process.env.SUPABASE_URL;
 const anon=process.env.SUPABASE_ANON_KEY;
 const dbSecret=process.env.WORKER_DB_SECRET;
 if(!root||!anon||!dbSecret)throw Object.assign(new Error('Restricted worker storage is not configured'),{status:503});
 // Build from the origin so a trailing slash or copied Data API suffix cannot
 // produce a malformed /rest/v1/rest/v1 or //rest/v1 path.
 let endpoint;
 try {
  const parsed=new URL(root.trim());
  if(parsed.protocol!=='https:'||!parsed.hostname.endsWith('.supabase.co'))throw Error('invalid project URL');
  endpoint=`${parsed.origin}/rest/v1/rpc/fortuna_worker_action`;
 }catch{
  throw Object.assign(new Error('Supabase project URL is invalid. Check SUPABASE_URL in Vercel.'),{status:503});
 }
 const response=await fetch(endpoint,{
  method:'POST',cache:'no-store',headers:{apikey:anon,...(anon.startsWith('sb_publishable_')?{}:{Authorization:`Bearer ${anon}`}), 'Content-Type':'application/json'},
  body:JSON.stringify({p_secret:dbSecret,p_action:action,p_worker_id:workerId,p_input:input})
 });
 if(!response.ok){
  const detail=await response.json().catch(()=>({}));
  const code=String(detail.code||'');
  if(code==='23505')throw Object.assign(new Error('A matching worker already exists. Ask Southwest staff for a returning-worker link.'),{status:409});
  if(code==='40001')throw Object.assign(new Error('Please retry: the record changed while saving.'),{status:409});
  console.error('restricted worker storage error',response.status,code,'path=',new URL(endpoint).pathname,'message=',String(detail.message||'').slice(0,160));
  throw Object.assign(new Error('The safety record could not be accessed. Contact Southwest safety staff.'),{status:response.status===401||response.status===403?503:500});
 }
 return response.json();
}

async function callClaude(content,max_tokens=1400){
 const apiKey=process.env.ANTHROPIC_API_KEY;
 if(process.env.ENABLE_WORKER_AI!=='true'||!apiKey)throw Object.assign(new Error('Document and translation service is not configured'),{status:503});
 const r=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01'},body:JSON.stringify({model:'claude-sonnet-4-6',max_tokens,messages:[{role:'user',content}]})});
 const data=await r.json();if(!r.ok)throw Object.assign(new Error('AI processing is temporarily unavailable'),{status:503});return (data.content||[]).filter(c=>c.type==='text').map(c=>c.text).join('');
}
function parseJson(text){let cleaned=String(text||'').replace(/^```(?:json)?/i,'').replace(/```$/,'').trim();const start=cleaned.indexOf('{'),end=cleaned.lastIndexOf('}');if(start<0||end<start)throw Error('Cannot parse document response');return JSON.parse(cleaned.slice(start,end+1));}
function validAttachment(a){return !!(a?.data&&/^data:(image\/(jpeg|png|webp)|application\/pdf);base64,/.test(a.data)&&a.data.length<=2900000);}
async function extractTicket(attachment){
 const match=attachment.data.match(/^data:([^;]+);base64,(.*)$/s);const mime=match[1],data=match[2];
 const fileBlock=mime==='application/pdf'?{type:'document',source:{type:'base64',media_type:mime,data}}:{type:'image',source:{type:'base64',media_type:mime,data}};
 const prompt=`Read this construction worker safety training certificate. Return a single JSON object containing: {"type":"certificate name or null","issued_date":"YYYY-MM-DD or null","expiry_date":"YYYY-MM-DD or null","certificate_number":"number or null","issuing_body":"provider or null","holder_name":"name or null"}. Only copy dates clearly printed; never invent an expiry date or claim a certificate is valid. If unsure return null. Return JSON only.`;
 try {const extracted=parseJson(await callClaude([fileBlock,{type:'text',text:prompt}],650));const date=v=>typeof v==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(v)?v:'';return {type:txt(extracted.type,100)||'Unidentified certificate — review required',issued_date:date(extracted.issued_date),expiry_date:date(extracted.expiry_date),cert_number:txt(extracted.certificate_number,100),issuing_body:txt(extracted.issuing_body,120),holder_name:txt(extracted.holder_name,140),extraction_status:'extracted_unverified'};}
 catch(e){return {type:'Unidentified certificate — review required',issued_date:'',expiry_date:'',extraction_status:'needs_manual_extraction'};}
}

function bad(res,message,status=400){return res.status(status).json({error:message});}
export default async function handler(req,res){res.setHeader('Cache-Control','no-store');if(process.env.ENABLE_PUBLIC_WORKER_PORTAL!=='true')return res.status(503).json({error:'Worker intake is not enabled yet. Contact Southwest safety staff.'});if(req.method!=='POST')return bad(res,'POST required',405);if(!secret()||secret().length<32)return bad(res,'Worker portal is not configured',503);
 try{const body=req.body||{},action=body.action;
 if(action==='register'){
 const name=normalizeIdentity(body.name),company=normalizeIdentity(body.company),trade=txt(body.trade),occupation=txt(body.occupation),consent=body.consent===true;
 if(!name||!company||!trade||!consent)return bad(res,'Name, company, trade and acknowledgement are required');
 const worker={id:randomUUID(),name,company,trade,occupation,preferred_language:['en','fr','es','ar','uk','de'].includes(body.language)?body.language:'en',supervisor:'',tickets:[],created_at:new Date().toISOString(),source:'worker_portal',intake_status:'needs_staff_review'};
 const summary=await portalDb('register',worker.id,worker);
 return res.status(200).json({token:createToken(worker.id),worker:summary});}
 const workerId=decodeToken(body.token);if(!workerId)return bad(res,'This worker session has expired. Ask Southwest staff for a new link.',401);
 if(action==='profile'){const summary=await portalDb('profile',workerId);return res.status(200).json({worker:summary});}
 if(action==='translate'){
  const lang=txt(body.language,4);if(!['fr','es','ar','uk','de'].includes(lang))return bad(res,'Unsupported language');
  const index=Number(body.section_index);if(!Number.isInteger(index)||index<0||index>5)return bad(res,'Invalid section');
  const qs=body.questions;if(!Array.isArray(qs)||qs.length!==2||qs.some(q=>typeof q.prompt!=='string'||q.prompt.length>300||!Array.isArray(q.options)||q.options.length>5||q.options.some(v=>typeof v!=='string'||v.length>220)))return bad(res,'Invalid questions');
  // The English safety source is passed through by the worker UI. This translation is a DRAFT pending human review.
  const source={title:txt(body.title,150),points:Array.isArray(body.points)?body.points.slice(0,8).map(x=>txt(x,600)):[],questions:qs};
  if(!source.title||source.points.length<2)return bad(res,'Section source missing');
  const sourceText=JSON.stringify(source);
  if(sourceText.length>7000)return bad(res,'Section too long');
  const raw=await callClaude([{type:'text',text:`Translate this site safety orientation JSON to language code ${lang}. Keep all JSON keys exactly unchanged; translate all string values, including the incorrect joke answers with natural culturally appropriate humour but never alter the correct safety meaning. Return only JSON. This is a draft for safety-manager review. Source: ${sourceText}`}],3100);
  const translated=parseJson(raw);
  if(typeof translated.title!=='string'||!Array.isArray(translated.points)||translated.points.length!==source.points.length||!Array.isArray(translated.questions)||translated.questions.length!==2||translated.questions.some((q,i)=>typeof q.prompt!=='string'||!Array.isArray(q.options)||q.options.length!==qs[i].options.length))return bad(res,'Translation incomplete',502);
  return res.status(200).json({translation:translated,translation_status:'machine_draft_requires_safety_review'});
 }
 if(action==='ticket_auto'){
  const attachment=body.attachment;if(!validAttachment(attachment))return bad(res,'Upload a JPEG, PNG, WebP or PDF smaller than 2 MB');
  const extracted=await extractTicket(attachment);
  const t={id:randomUUID(),...extracted,notes:'Automatically extracted from worker upload. Verify against original document.',status:'needs_review',attachment:{name:txt(attachment.name,100),type:txt(attachment.type,80),data:attachment.data},added_at:new Date().toISOString(),source:'worker_portal_auto'};
  const summary=await portalDb('ticket',workerId,{ticket:t});
  return res.status(200).json({worker:summary});
 }
 if(action==='ticket'){
 const type=txt(body.type,100),issue=txt(body.issued_date,10),expiry=txt(body.expiry_date,10),notes=txt(body.notes,300);
 if(!type||!isoDate(issue)||!isoDate(expiry))return bad(res,'Check certificate type and dates');
 const attachment=body.attachment;
 if(!attachment?.data||!/^data:(image\/(jpeg|png|webp)|application\/pdf);base64,/.test(attachment.data)||attachment.data.length>1450000)return bad(res,'Upload a JPEG, PNG, WebP or PDF smaller than 1 MB');
 const t={id:randomUUID(),type,issued_date:issue,expiry_date:expiry,notes,status:'needs_review',attachment:{name:txt(attachment.name,100),type:txt(attachment.type,80),data:attachment.data},added_at:new Date().toISOString(),source:'worker_portal'};
 const summary=await portalDb('ticket',workerId,{ticket:t});
 return res.status(200).json({worker:summary});}
 if(action==='orientation'){
 const sections=body.sections,signature=txt(body.signature),ack=body.acknowledged===true;
 if(!Array.isArray(sections)||sections.length!==6||!ack||!signature||sections.some(s=>!s.title||!(s.elapsed_seconds>=0)||!Array.isArray(s.questions)||s.questions.length!==2||s.questions.some(q=>!q.correct||!(q.attempts>=1))))return bad(res,'Complete every section, knowledge check, acknowledgement and signature');
 // Stop accidental duplicate submissions. Renewals are allowed when the manager returned the previous orientation.
 const current=await portalDb('profile',workerId);
 const approval=String(current?.orientation?.approval_status||'').toLowerCase();
 if(approval==='pending')return bad(res,'Your orientation has already been submitted and is awaiting safety manager approval.',409);
 if(approval==='approved')return bad(res,'Your orientation is already approved. Ask Southwest safety staff if you need a renewal.',409);
 const time=new Date().toISOString();const attempt={id:randomUUID(),worker_id:workerId,version:'Southwest Orientation v21.3 draft',status:'completed',approval_status:'pending',started_at:txt(body.started_at,30)||time,completed_at:time,language:['en','fr','es','ar','uk','de'].includes(body.language)?body.language:'en',sections:sections.map(s=>({title:txt(s.title),language:txt(s.language,4),displayed_translation:s.displayed_translation&&typeof s.displayed_translation==='object'?s.displayed_translation:null,elapsed_seconds:Math.min(36000,Math.max(0,Math.round(s.elapsed_seconds))),questions:s.questions.map(q=>({prompt:txt(q.prompt,300),attempts:Math.min(100,Math.round(q.attempts)),correct:true})),acknowledged_at:txt(s.acknowledged_at,30)})),signature,acknowledged:true,source:'worker_portal'};
 const summary=await portalDb('orientation',workerId,{orientation:attempt});
 return res.status(200).json({worker:summary});}
 return bad(res,'Unsupported action');
 }catch(e){console.error('worker-portal',e.message);return bad(res,e.status?e.message:'The request could not be saved. Please try again.',e.status||500);}}
