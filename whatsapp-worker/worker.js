// BUSYWEB_WHATSAPP_PORTLESS_V1 - Unofficial WhatsApp Web integration.
// No listen(), express, http server, socket, or externally available interface.
'use strict';
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const crypto = require('crypto');
const qrcode = require('qrcode-terminal');
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const root = path.resolve(__dirname, '..', 'whatsapp');
const folders = { pending: path.join(root,'pending'), processing:path.join(root,'processing'),
  completed:path.join(root,'completed'),failed:path.join(root,'failed') };
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
let ready = false, state = 'starting', stopping = false;
const atomicJSON = async (filename, obj) => {
  const tmp = filename + '.' + process.pid + '.' + crypto.randomBytes(8).toString('hex') + '.tmp';
  try {
    await fsp.writeFile(tmp, JSON.stringify(obj,null,2),{flag:'wx',mode:0o600});
    await fsp.rename(tmp,filename);
  } catch (e) {
    await fsp.unlink(tmp).catch(()=>{});
    throw e;
  }
};
let statusWriter = Promise.resolve();
const status = () => {
  // Serialize status snapshots so auth/ready/heartbeat writes cannot race.
  const snapshot = {ready,state,updatedAt:new Date().toISOString()};
  statusWriter = statusWriter.then(() => atomicJSON(path.join(root,'connection-status.json'),snapshot))
    .catch(e => console.error('[WA] status:',e.message));
  return statusWriter;
};
const saveJob = async (folder, job) => {
  job.updatedAt = new Date().toISOString();
  await atomicJSON(path.join(folders[folder],job.jobId+'.json'),job);
};
async function markInterrupted() {
  // Process died after claiming job. Delivery outcome could be unknown. Never auto-resend.
  for (const f of await fsp.readdir(folders.processing)) {
    if (!/^[0-9a-f]{32}\.json$/.test(f)) continue;
    const p=path.join(folders.processing,f);
    try {
      const job=JSON.parse(await fsp.readFile(p,'utf8'));
      job.status='unknown';job.error='Worker interrupted while sending. Check WhatsApp chat before resending.';
      await saveJob('failed',job);await fsp.unlink(p);
      await fsp.unlink(path.join(folders.pending,job.jobId+'.pdf')).catch(()=>{});
    } catch(e) { console.error('[WA] interrupted job:',e.message); }
  }
}
// WhatsApp can resolve a phone number to a privacy-preserving LID instead of @c.us.
// Trust only the ID returned by getNumberId, not a browser-supplied chat ID.
function getRecipientChatId(wid) {
  if (!wid) throw new Error('WhatsApp recipient lookup returned no ID');
  const chatId = String(wid._serialized || '');
  if (!/^[0-9]+@(c\.us|lid)$/.test(chatId)) {
    throw new Error('WhatsApp recipient lookup returned an unsupported ID type');
  }
  return chatId;
}
async function sendOne(client, name) {
  if (!/^[0-9a-f]{32}\.json$/.test(name)) return;
  const jobId = name.slice(0,-5);
  const pendingPath=path.join(folders.pending,name);
  const processingPath=path.join(folders.processing,name);
  try { await fsp.rename(pendingPath,processingPath); } catch { return; }
  let job;
  const pdf=path.join(folders.pending,jobId+'.pdf');
  try {
    job=JSON.parse(await fsp.readFile(processingPath,'utf8'));
    if (job.jobId!==jobId || !/^[1-9][0-9]{7,14}$/.test(job.phone||'')) throw new Error('Invalid job metadata');
    if (!['A4','A5','TAX'].includes(job.invoiceFormat)) throw new Error('Invalid format');
    const buf=await fsp.readFile(pdf);
    if (buf.length<100 || buf.length>8*1024*1024 || buf.subarray(0,5).toString('ascii')!=='%PDF-') throw new Error('Invalid PDF');
    const hash=crypto.createHash('sha256').update(buf).digest('hex');
    if (hash!==job.pdfSha256) throw new Error('PDF integrity check failed');
    // Avoid sending until the client reports READY. Group/broadcast messages unsupported.
    const wid=await client.getNumberId(job.phone);
    const chatId=getRecipientChatId(wid);
    console.log(`[WA] Job ${jobId}: recipient lookup resolved ${chatId.endsWith('@lid')?'@lid':'@c.us'}`);
    job.status='sending';job.attempts=1;job.updatedAt=new Date().toISOString();
    await fsp.writeFile(processingPath,JSON.stringify(job,null,2),'utf8');
    const media = new MessageMedia('application/pdf',buf.toString('base64'),
      `${String(job.vchNo).replace(/[^a-zA-Z0-9._-]/g,'-')}-${job.invoiceFormat}.pdf`);
    // Branded document caption, with the accurate BUSY voucher type.
    const documentNames = {
      2: 'Purchase Invoice',
      3: 'Sale Return',
      4: 'Goods Received Note',
      9: 'Sale Invoice',
      10: 'Purchase Return',
      11: 'Delivery Order',
      12: 'Sales Order',
      13: 'Purchase Order',
      14: 'Receipt Voucher',
      19: 'Payment Voucher',
      26: 'Sale Quotation',
      27: 'Purchase Quotation',
    };
    const documentName = documentNames[Number(job.vchType)] || 'Voucher';
    const caption=`*BusyCloud*\n\nDear Customer,\nPlease find attached your ${documentName} ${job.vchNo} (${job.invoiceFormat} PDF).\n\nThank you for choosing BusyCloud.`;
    const sent=await client.sendMessage(chatId,media,{caption,sendMediaAsDocument:true});
    // 'submitted' means WhatsApp Web accepted the send call, NOT delivered/read.
    job.status='submitted';job.messageId=String(sent?.id?._serialized||'');job.error='';
    await saveJob('completed',job);
    console.log(`[WA] Submitted ${jobId}, voucher=${job.vchNo}`);
  } catch(e) {
    console.error(`[WA] Job ${jobId} error:`,e.message);
    if (job) {job.status=job.status==='sending'?'unknown':'failed';job.error=job.status==='unknown' ?
      'Send outcome uncertain. Check recipient chat before retrying. '+e.message : e.message;
      try {await saveJob('failed',job);}catch(ex){console.error('[WA] failed to persist:',ex.message);} }
  } finally {
    await fsp.unlink(processingPath).catch(()=>{});
    await fsp.unlink(pdf).catch(()=>{}); // no invoice PDF retained after attempt
  }
}
async function run() {
  for(const dir of Object.values(folders).concat([path.join(root,'session')])) await fsp.mkdir(dir,{recursive:true});
  await markInterrupted();
  const client=new Client({authStrategy:new LocalAuth({clientId:'busyweb',dataPath:path.join(root,'session')}),
    puppeteer:{headless:true,args:['--no-first-run','--no-default-browser-check']}});
  client.on('qr',qr => {ready=false;state='pairing';console.log('\n[WA] Scan with WhatsApp > Linked Devices > Link a device');qrcode.generate(qr,{small:true});void status();});
  client.on('authenticated',()=>{state='authenticated';void status();console.log('[WA] Authenticated');});
  client.on('ready',()=>{ready=true;state='ready';void status();console.log('[WA] READY, worker listening to file queue');});
  client.on('disconnected',why=>{ready=false;state='disconnected';void status();console.error('[WA] Disconnected:',why);});
  client.on('auth_failure',why=>{ready=false;state='auth-failed';void status();console.error('[WA] Authentication failed:',why);});
  const periodic=setInterval(()=>void status(),10000);
  await status();
  const onStop=async()=>{if(stopping)return;stopping=true;ready=false;state='stopping';clearInterval(periodic);await status();
    try{await client.destroy();}catch{}process.exit(0);};
  process.on('SIGINT',()=>void onStop());process.on('SIGTERM',()=>void onStop());
  await client.initialize();
  while(!stopping){
    if(ready){
      try { const files=(await fsp.readdir(folders.pending)).filter(f=>/^[0-9a-f]{32}\.json$/.test(f)).sort();
        for(const f of files){if(!ready||stopping)break;await sendOne(client,f);} }
      catch(e){console.error('[WA] queue:',e.message);} }
    await sleep(3000);
  }
}
run().catch(e=>{console.error('[WA] Fatal:',e);process.exitCode=1});
