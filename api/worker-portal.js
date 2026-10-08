// v21.1 PREVIEW-ONLY worker self-service intake. No staff data is returned.
// Production requires identity verification, abuse throttling, separate private file storage and security review.
import {createHmac, randomUUID, timingSafeEqual} from 'node:crypto';
const allowedOrigins = null;
const secret = () => process.env.SAFETY_PORTAL_SIGNING_KEY;
const sign = s => createHmac('sha256',secret()).update(s).digest('base64url');
const safeEq=(a,b)=>{const x=Buffer.from(String(a||'')),y=Buffer.from(String(b||''));return x.length===y.length&&timingSafeEqual(x,y)};
function createToken(workerId){const payload=Buffer.from(JSON.stringify({workerId,exp:Date.now()+1000*60*60*12})).toString('base64url');return `${payload}.${sign(payload)}`;}
function decodeToken(token){try{const [p,mac]=String(token||'').split('.');if(!p||!mac||!safeEq(mac,sign(p)))return null;const o=JSON.parse(Buffer.from(p,'base64url'));return o.exp>Date.now()&&typeof o.workerId==='string'?o.workerId:null;}catch{return null;}}
const txt=(v,max=140)=>String(v??'').trim().slice(0,max);
const isoDate=(v)=> !v||/^\d{4}-\d{2}-\d{2}$/.test(v);
async function fetchRecord(){const root=process.env.SUPABASE_URL,key=process.env.SUPABASE_SERVICE_KEY;if(!root||!key)throw new Error('Storage is not configured');const headers={apikey:key,Authorization:`Bearer ${key}`,'Content-Type':'application/json'};const r=await fetch(`${root}/rest/v1/safety_data?id=eq.1&select=payload,revision`,{headers,cache:'no-store'});if(!r.ok)throw new Error('Storage read failed');const arr=await r.json();if(arr.length!==1)throw new Error('Safety database not initialized');return {record:arr[0],root,headers};}
async function update(mutator){for(let tries=0;tries<3;tries++){const {record,root,headers}=await fetchRecord();const payload=structuredClone(record.payload||{workers:[],orientations:[]});const outcome=mutator(payload);const r=await fetch(`${root}/rest/v1/safety_data?id=eq.1&revision=eq.${record.revision}`,{method:'PATCH',headers:{...headers,Prefer:'return=representation'},body:JSON.stringify({payload,revision:Number(record.revision)+1})});if(!r.ok)throw new Error('Storage write failed');const rows=await r.json();if(rows.length===1)return outcome;}throw Object.assign(new Error('Records changed; retry'),{status:409});}
function workerSummary(w,p){const orientations=(p.orientations||[]).filter(o=>o.worker_id===w.id).sort((a,b)=>String(b.started_at).localeCompare(String(a.started_at)));return {id:w.id,name:w.name,company:w.company,trade:w.trade,occupation:w.occupation||'',tickets:(w.tickets||[]).map(t=>({id:t.id,type:t.type,expiry_date:t.expiry_date,status:t.status})),orientation:orientations[0]?{status:orientations[0].status,approval_status:orientations[0].approval_status}:null};}

async function callClaude(content,max_tokens=1400){
 const apiKey=process.env.VITE_ANTHROPIC_API_KEY;
 if(!apiKey)throw Object.assign(new Error('Document and translation service is not configured'),{status:503});
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
export default async function handler(req,res){res.setHeader('Cache-Control','no-store');if(req.method!=='POST')return bad(res,'POST required',405);if(!secret()||secret().length<32)return bad(res,'Worker portal is not configured',503);
 try{const body=req.body||{},action=body.action;
 if(action==='register'){
 const name=txt(body.name),company=txt(body.company),trade=txt(body.trade),occupation=txt(body.occupation),consent=body.consent===true;
 if(!name||!company||!trade||!consent)return bad(res,'Name, company, trade and acknowledgement are required');
 const worker={id:randomUUID(),name,company,trade,occupation,preferred_language:['en','fr','es','ar','uk','de'].includes(body.language)?body.language:'en',supervisor:'',tickets:[],created_at:new Date().toISOString(),source:'worker_portal',intake_status:'needs_staff_review'};
 await update(p=>{p.workers||=[];p.orientations||=[];if(p.workers.some(w=>w.name.toLowerCase()===name.toLowerCase()&&w.company.toLowerCase()===company.toLowerCase()))throw Object.assign(new Error('A matching worker already exists. Ask Southwest staff for a renewal link.'),{status:409});p.workers.push(worker)});
 return res.status(200).json({token:createToken(worker.id),worker:workerSummary(worker,{orientations:[]})});}
 const workerId=decodeToken(body.token);if(!workerId)return bad(res,'This worker session has expired. Ask Southwest staff for a new link.',401);
 if(action==='profile'){const {record}=await fetchRecord();const w=(record.payload?.workers||[]).find(w=>w.id===workerId);return w?res.status(200).json({worker:workerSummary(w,record.payload)}):bad(res,'Worker record not found',404);}
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
  const summary=await update(p=>{const w=p.workers.find(w=>w.id===workerId);if(!w)throw Object.assign(new Error('Worker not found'),{status:404});w.tickets||=[];w.tickets.push(t);return workerSummary(w,p)});
  return res.status(200).json({worker:summary});
 }
 if(action==='ticket'){
 const type=txt(body.type,100),issue=txt(body.issued_date,10),expiry=txt(body.expiry_date,10),notes=txt(body.notes,300);
 if(!type||!isoDate(issue)||!isoDate(expiry))return bad(res,'Check certificate type and dates');
 const attachment=body.attachment;
 if(!attachment?.data||!/^data:(image\/(jpeg|png|webp)|application\/pdf);base64,/.test(attachment.data)||attachment.data.length>1450000)return bad(res,'Upload a JPEG, PNG, WebP or PDF smaller than 1 MB');
 const t={id:randomUUID(),type,issued_date:issue,expiry_date:expiry,notes,status:'needs_review',attachment:{name:txt(attachment.name,100),type:txt(attachment.type,80),data:attachment.data},added_at:new Date().toISOString(),source:'worker_portal'};
 const summary=await update(p=>{const w=p.workers.find(w=>w.id===workerId);if(!w)throw Object.assign(new Error('Worker not found'),{status:404});w.tickets||=[];w.tickets.push(t);return workerSummary(w,p)});
 return res.status(200).json({worker:summary});}
 if(action==='orientation'){
 const sections=body.sections,signature=txt(body.signature),ack=body.acknowledged===true;
 if(!Array.isArray(sections)||sections.length!==6||!ack||!signature||sections.some(s=>!s.title||!(s.elapsed_seconds>=0)||!Array.isArray(s.questions)||s.questions.length!==2||s.questions.some(q=>!q.correct||!(q.attempts>=1))))return bad(res,'Complete every section, knowledge check, acknowledgement and signature');
 const time=new Date().toISOString();const attempt={id:randomUUID(),worker_id:workerId,version:'Southwest Orientation v21.3 draft',status:'completed',approval_status:'pending',started_at:txt(body.started_at,30)||time,completed_at:time,language:['en','fr','es','ar','uk','de'].includes(body.language)?body.language:'en',sections:sections.map(s=>({title:txt(s.title),language:txt(s.language,4),displayed_translation:s.displayed_translation&&typeof s.displayed_translation==='object'?s.displayed_translation:null,elapsed_seconds:Math.min(36000,Math.max(0,Math.round(s.elapsed_seconds))),questions:s.questions.map(q=>({prompt:txt(q.prompt,300),attempts:Math.min(100,Math.round(q.attempts)),correct:true})),acknowledged_at:txt(s.acknowledged_at,30)})),signature,acknowledged:true,source:'worker_portal'};
 const summary=await update(p=>{const w=p.workers.find(w=>w.id===workerId);if(!w)throw Object.assign(new Error('Worker not found'),{status:404});p.orientations||=[];p.orientations.push(attempt);return workerSummary(w,p)});
 return res.status(200).json({worker:summary});}
 return bad(res,'Unsupported action');
 }catch(e){console.error('worker-portal',e.message);return bad(res,e.status?e.message:'The request could not be saved. Please try again.',e.status||500);}}
