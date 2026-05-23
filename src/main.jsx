
import React, { useEffect, useRef, useState } from 'react'
import ReactDOM from 'react-dom/client'
import { BrowserRouter, NavLink, Route, Routes, useNavigate } from 'react-router-dom'
import { motion } from 'framer-motion'
import SignatureCanvas from 'react-signature-canvas'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'
import * as pdfjsLib from 'pdfjs-dist'
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import mammoth from 'mammoth'
import { renderAsync } from 'docx-preview'
import {
  FileSignature, FilePlus2, PenLine, Upload, FolderCog, Clock3, Settings,
  Plus, Trash2, Download, Share2, Search, Copy, Moon, Sun, Image as ImageIcon,
  RotateCcw, CheckCircle2, Move, Maximize2
} from 'lucide-react'
import './styles.css'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker

const K = {
  doc:'qspu2_doc', templates:'qspu2_templates', recents:'qspu2_recents',
  dark:'qspu2_dark', deletedDefaults:'qspu2_deleted_defaults'
}
const today = () => new Date().toISOString().slice(0,10)
const uid = () => crypto?.randomUUID ? crypto.randomUUID() : String(Date.now()+Math.random())
const load = (k, f) => { try { const v=localStorage.getItem(k); return v ? JSON.parse(v) : f } catch { return f } }
const save = (k, v) => localStorage.setItem(k, JSON.stringify(v))

const cats = [
  ['business','商業合約','合作、服務、一般協議'],
  ['nda','NDA','保密、資料使用及商務往來'],
  ['event','活動／家長同意書','活動、家長、參加者同意'],
  ['volunteer','義工服務文件','義工服務、守則及聲明'],
  ['social','社工文件','社區服務文件'],
  ['quote','報價單／收據','報價、收據、確認'],
  ['auth','授權書','授權、委託及確認'],
  ['declaration','聲明文件','聲明、確認及備註'],
  ['custom','我的範本','自行建立及匯入範本'],
  ['blank','空白文件','由零開始建立文件']
].map(([id,name,desc])=>({id,name,desc}))

const defaultTemplates = [
  ['nda-basic','nda','NDA 保密協議','合作前保密、文件傳閱及商務討論。',`本協議由 {{company_name}} 與簽署人於 {{date}} 訂立。\n\n簽署人同意對因業務、合作、會議或文件往來而取得的資料保密，不得在未獲授權情況下披露、複製或轉交予第三方。\n\n本協議適用於所有書面、口頭、電子或其他形式的資料。\n\n備註：{{notes}}`],
  ['event-consent','event','活動同意書','社區活動、戶外活動及一般參加同意。',`本人確認已閱讀並明白 {{company_name}} 所舉辦活動的內容、安排及注意事項。\n\n本人同意參與相關活動，並會遵守工作員或負責人的合理指示。\n\n如活動涉及戶外、交通或體能活動，本人明白活動可能存在一般風險，並會按自身情況量力參與。\n\n參加者／客戶：{{client_name}}\n備註：{{notes}}`],
  ['volunteer','volunteer','義工服務協議','義工服務守則、保密原則及服務確認。',`本人同意以義工身份參與 {{company_name}} 的服務或活動。\n\n本人會準時出席、尊重服務對象、遵守保密原則，並按機構指引提供協助。\n\n本人明白義工服務不構成僱傭關係。\n\n備註：{{notes}}`],
  ['quote','quote','報價單','服務報價、活動收費及商務確認。',`公司名稱：{{company_name}}\n文件日期：{{date}}\n\n服務／產品內容：\n1. 請填寫項目\n2. 請填寫項目\n\n總金額：HK$ ________\n\n備註：以上報價有效期為 14 天，確認後方可安排服務。\n\n客戶／確認人：{{client_name}}`],
  ['partnership','business','合作協議','一般合作、項目分工及合作確認。',`{{company_name}} 與簽署人同意就相關項目進行合作。\n\n雙方將按誠信、互惠及專業原則推進合作，並共同確認工作範圍、時間表及責任分工。\n\n任何修訂應以書面或電子方式確認。\n\n備註：{{notes}}`],
  ['parent','event','家長同意書','學生或兒童活動。',`本人作為參加者之家長／監護人，同意其參與 {{company_name}} 安排的活動。\n\n本人確認已知悉活動內容、日期、地點及注意事項。\n\n如有緊急情況，本人同意工作員按實際需要作出合理處理。\n\n參加者：{{client_name}}\n備註：{{notes}}`],
  ['social','social','社工活動聲明','服務參與、資料使用及活動跟進。',`本人確認已明白活動目的、內容及個人資料使用安排。\n\n本人同意 {{company_name}} 為活動安排、跟進及服務記錄目的保存必要資料。\n\n如有需要，本人同意工作員就活動或服務安排與本人聯絡。\n\n備註：{{notes}}`],
  ['auth','auth','授權書','一般授權及委託安排。',`本人 {{client_name}} 授權 {{company_name}} 就以下事項作出處理：\n\n授權事項：\n1. ________________________\n2. ________________________\n\n本授權由 {{date}} 起生效。\n\n備註：{{notes}}`],
  ['receipt','quote','收據','簡單收款確認。',`收款單位：{{company_name}}\n日期：{{date}}\n\n茲收到 {{client_name}} 款項 HK$ ________。\n\n款項用途：________________________\n付款方式：________________________\n\n備註：{{notes}}`],
  ['blank','blank','空白文件','由零開始撰寫文件。',`請在此輸入文件內容。\n\n備註：{{notes}}`]
].map(([templateId,category,title,description,content])=>({templateId,category,title,description,content,default:true,createdAt:'2026-01-01',updatedAt:'2026-01-01',uses:0}))

const freshDoc = () => ({
  company:'', title:'', content:'', clientName:'', date:today(), notes:'',
  logo:'', logoPosition:'center', selectedCategory:'business', selectedTemplateId:'',
  type:'document', status:'draft', createdAt:new Date().toISOString(), updatedAt:new Date().toISOString(),
  signers:[{id:uid(),name:'',title:'',date:today(),signature:'',status:'pending'}]
})

function availableTemplates(){
  const deleted = load(K.deletedDefaults, [])
  return [...defaultTemplates.filter(t=>!deleted.includes(t.templateId)), ...load(K.templates, [])]
}
function applyVars(text, doc){
  const s=doc.signers?.[0]||{}
  return (text||'')
    .replaceAll('{{company_name}}', doc.company||'')
    .replaceAll('{{document_title}}', doc.title||'')
    .replaceAll('{{client_name}}', doc.clientName||'')
    .replaceAll('{{date}}', doc.date||today())
    .replaceAll('{{notes}}', doc.notes||'')
    .replaceAll('{{signer_name}}', s.name||'')
    .replaceAll('{{signer_title}}', s.title||'')
    .replaceAll('{{signature}}', s.signature?'[已簽署]':'')
}
function signGridClass(count){
  if(count<=1) return 'grid-cols-1 signature-large'
  if(count===2) return 'grid-cols-1 sm:grid-cols-2 signature-medium'
  if(count<=4) return 'grid-cols-1 sm:grid-cols-2 signature-compact'
  return 'grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 signature-mini'
}

function DocumentPreview({doc}){
  const justify = doc.logoPosition==='left'?'justify-start':doc.logoPosition==='right'?'justify-end':'justify-center'
  const signers = doc.signers || []
  return <div id="doc-preview" className="document-paper">
    <div className="doc-inner">
      <header className="doc-header">
        {doc.logo && <div className={`mb-6 flex ${justify}`}><img src={doc.logo} className="doc-logo"/></div>}
        <div className="flex items-start justify-between gap-5">
          <div>
            <p className="doc-meta">Formal Document</p>
            <h1 className="doc-title">{doc.title || '文件標題'}</h1>
          </div>
          <div className="doc-date">{doc.date || today()}</div>
        </div>
        <div className="doc-line"></div>
        <div className="doc-info">
          {doc.company && <span>公司：{doc.company}</span>}
          {doc.clientName && <span>對象：{doc.clientName}</span>}
        </div>
      </header>

      <main className="doc-body">
        {applyVars(doc.content, doc) || '文件內容會顯示在這裡。'}
      </main>

      {doc.notes && <section className="doc-note"><b>備註：</b>{doc.notes}</section>}

      <section className="signature-section">
        <h2 className="signature-heading">簽署區</h2>
        <div className={`signature-grid ${signGridClass(signers.length)}`}>
          {signers.map((s,i)=><div key={s.id} className="signature-card">
            <div className="signature-image">
              {s.signature ? <img src={s.signature}/> : <span>Signature</span>}
            </div>
            <div className="signature-data">
              <p><b>姓名</b><span>{s.name || '未填寫'}</span></p>
              <p><b>職位</b><span>{s.title || '未填寫'}</span></p>
              <p><b>日期</b><span>{s.date || today()}</span></p>
            </div>
          </div>)}
        </div>
      </section>
    </div>
  </div>
}

async function exportPdf(file='document.pdf'){
  const el=document.getElementById('doc-preview')
  if(!el) return alert('找不到文件預覽')
  const canvas=await html2canvas(el,{scale:2,backgroundColor:'#fff',useCORS:true})
  const img=canvas.toDataURL('image/png')
  const pdf=new jsPDF('p','mm','a4')
  const pw=210, ph=297, ih=canvas.height*pw/canvas.width
  let left=ih, pos=0
  pdf.addImage(img,'PNG',0,pos,pw,ih)
  left-=ph
  while(left>0){pos=left-ih; pdf.addPage(); pdf.addImage(img,'PNG',0,pos,pw,ih); left-=ph}
  pdf.save(file)
}
async function exportPng(file='document.png', share=false){
  const el=document.getElementById('doc-preview')
  if(!el) return alert('找不到文件預覽')
  const canvas=await html2canvas(el,{scale:2,backgroundColor:'#fff',useCORS:true})
  canvas.toBlob(async blob=>{
    const f=new File([blob],file,{type:'image/png'})
    if(share && navigator.canShare?.({files:[f]})) await navigator.share({title:file,files:[f]})
    else{const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=file; a.click()}
  })
}
function completeDoc(doc, status='已簽署'){
  const recents=load(K.recents,[])
  const item={...doc,id:uid(),status,completedAt:new Date().toISOString(),updatedAt:new Date().toISOString()}
  save(K.recents,[item,...recents].slice(0,50))
  save(K.doc,item)
  return item
}

function Shell(){
  const [dark,setDark]=useState(load(K.dark,false))
  useEffect(()=>{document.documentElement.classList.toggle('dark',dark); save(K.dark,dark)},[dark])
  return <div className="min-h-screen bg-slate-50 text-slate-950 dark:bg-slate-950 dark:text-slate-50">
    <div className="mx-auto min-h-screen max-w-7xl pb-28">
      <header className="sticky top-0 z-30 border-b border-slate-200/70 bg-white/85 px-4 py-3 backdrop-blur-xl dark:border-slate-800 dark:bg-slate-950/85">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="grid h-11 w-11 place-items-center rounded-2xl bg-slate-950 text-white shadow-soft dark:bg-white dark:text-slate-950"><FileSignature size={23}/></div>
            <div><h1 className="text-lg font-black tracking-tight">QuickSign Pro Ultimate v2</h1><p className="text-xs text-slate-500">快速建立、匯入、簽署及分享文件</p></div>
          </div>
          <button className="iconBtn" onClick={()=>setDark(!dark)}>{dark?<Sun size={18}/>:<Moon size={18}/>}</button>
        </div>
      </header>
      <main className="px-4 pt-5"><Routes>
        <Route path="/" element={<CreatePage/>}/>
        <Route path="/sign" element={<SignPage/>}/>
        <Route path="/pdf" element={<PdfPage/>}/>
        <Route path="/word" element={<WordPage/>}/>
        <Route path="/templates" element={<TemplatesPage/>}/>
        <Route path="/recent" element={<RecentPage/>}/>
        <Route path="/settings" element={<SettingsPage/>}/>
      </Routes></main>
    </div>
    <nav className="fixed inset-x-0 bottom-0 z-40 mx-auto max-w-3xl border-t border-slate-200 bg-white/90 px-2 pb-[max(env(safe-area-inset-bottom),12px)] pt-2 backdrop-blur-xl dark:border-slate-800 dark:bg-slate-950/90">
      <div className="grid grid-cols-6 gap-1">
        <Tab to="/" icon={<FilePlus2/>} label="建立"/><Tab to="/sign" icon={<PenLine/>} label="簽署"/><Tab to="/pdf" icon={<Upload/>} label="PDF"/><Tab to="/word" icon={<Upload/>} label="Word"/><Tab to="/templates" icon={<FolderCog/>} label="範本"/><Tab to="/recent" icon={<Clock3/>} label="最近"/>
      </div>
    </nav>
  </div>
}
function Tab({to,icon,label}){return <NavLink to={to} className={({isActive})=>`rounded-2xl px-2 py-2 text-center text-[11px] font-black transition ${isActive?'bg-slate-950 text-white dark:bg-white dark:text-slate-950':'text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-900'}`}><div className="mx-auto mb-1 flex justify-center [&>svg]:h-5 [&>svg]:w-5">{icon}</div>{label}</NavLink>}
function Input({label,value,set,type='text'}){return <><label className="label">{label}</label><input className="input" type={type} value={value||''} onChange={e=>set(e.target.value)}/></>}

function CreatePage(){
  const nav=useNavigate()
  const [doc,setDoc]=useState(load(K.doc,freshDoc()))
  const [step,setStep]=useState(1)
  const [templates,setTemplates]=useState(availableTemplates())
  useEffect(()=>save(K.doc,{...doc,updatedAt:new Date().toISOString()}),[doc])
  const shown=templates.filter(t=>t.category===doc.selectedCategory || (doc.selectedCategory==='custom'&&t.custom))
  function resetContent(){
    if(!confirm('確定要清除所有輸入內容嗎？')) return
    if(!confirm('此操作無法還原，是否確認重設？')) return
    const d=freshDoc()
    save(K.doc,d)
    setDoc(d)
    setStep(1)
  }
  async function logo(file){
    if(!file) return
    if(!file.type.startsWith('image/')) return alert('請上載 PNG / JPG / WEBP')
    const r=new FileReader(); r.onload=()=>setDoc({...doc,logo:r.result}); r.readAsDataURL(file)
  }
  return <div className="grid gap-5 lg:grid-cols-[1fr_.9fr]">
    <div className="space-y-5">
      <section className="card">
        <div className="flex items-start justify-between gap-3"><div><h2 className="text-2xl font-black">建立文件</h2><p className="mt-1 text-sm text-slate-500">選類別 → 選範本 → 輸入資料，右側即時預覽。</p></div><button className="smallDanger" onClick={resetContent}><RotateCcw size={15}/> 重設內容</button></div>
        <div className="mt-4 grid grid-cols-3 gap-2">{[1,2,3].map(n=><button key={n} onClick={()=>setStep(n)} className={`rounded-2xl px-3 py-3 text-xs font-black ${step===n?'bg-slate-950 text-white dark:bg-white dark:text-slate-950':'bg-slate-100 dark:bg-slate-800'}`}>Step {n}</button>)}</div>
      </section>

      {step===1&&<motion.section initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} className="card"><h3 className="sectionTitle">Step 1：選擇範本類別</h3><div className="mt-4 grid gap-3 sm:grid-cols-2">{cats.map(c=><button key={c.id} className="choiceCard text-left" onClick={()=>{setDoc({...doc,selectedCategory:c.id});setStep(2)}}><p className="font-black">{c.name}</p><p className="mt-1 text-xs text-slate-500">{c.desc}</p><p className="mt-3 text-xs font-black text-slate-400">{templates.filter(t=>t.category===c.id || (c.id==='custom'&&t.custom)).length} 個範本</p></button>)}</div></motion.section>}

      {step===2&&<motion.section initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} className="card"><h3 className="sectionTitle">Step 2：選擇範本</h3><div className="mt-4 grid gap-3">{shown.length?shown.map(t=><button key={t.templateId} className="choiceCard text-left" onClick={()=>{setDoc({...doc,selectedTemplateId:t.templateId,title:t.title,content:t.content});setStep(3)}}><div className="flex justify-between gap-3"><p className="font-black">{t.title}</p><span className="badge">{t.default?'預設':'自訂'}</span></div><p className="mt-1 text-xs text-slate-500">{t.description}</p><p className="mt-3 line-clamp-2 text-xs text-slate-400">{t.content}</p></button>):<p className="text-sm text-slate-500">此分類未有範本。</p>}</div></motion.section>}

      {step===3&&<motion.section initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} className="card">
        <h3 className="sectionTitle">Step 3：輸入文件資料</h3>
        <label className="label">公司 Logo</label>
        <div className="rounded-3xl border border-dashed border-slate-300 p-4 dark:border-slate-700">
          {doc.logo&&<img src={doc.logo} className="mb-3 max-h-20 max-w-[220px] rounded-xl object-contain"/>}
          <div className="flex flex-wrap gap-2"><label className="smallBtn cursor-pointer"><ImageIcon size={16}/> 上載 Logo<input type="file" accept="image/*" className="hidden" onChange={e=>logo(e.target.files?.[0])}/></label>{doc.logo&&<button className="smallDanger" onClick={()=>setDoc({...doc,logo:''})}><Trash2 size={16}/> 刪除</button>}</div>
          <div className="mt-3 grid grid-cols-3 gap-2">{['left','center','right'].map(p=><button key={p} className={`rounded-xl px-3 py-2 text-xs font-black ${doc.logoPosition===p?'bg-slate-950 text-white dark:bg-white dark:text-slate-950':'bg-slate-100 dark:bg-slate-800'}`} onClick={()=>setDoc({...doc,logoPosition:p})}>{p==='left'?'左方':p==='center'?'中間':'右方'}</button>)}</div>
        </div>
        <Input label="公司名稱" value={doc.company} set={v=>setDoc({...doc,company:v})}/><Input label="文件標題" value={doc.title} set={v=>setDoc({...doc,title:v})}/><Input label="客戶名稱" value={doc.clientName} set={v=>setDoc({...doc,clientName:v})}/><Input label="日期" type="date" value={doc.date} set={v=>setDoc({...doc,date:v})}/>
        <label className="label">文件內容</label><textarea className="textarea" value={doc.content} onChange={e=>setDoc({...doc,content:e.target.value})}/>
        <label className="label">備註</label><textarea className="input min-h-[88px]" value={doc.notes} onChange={e=>setDoc({...doc,notes:e.target.value})}/>
        <div className="mt-4 flex flex-wrap gap-2"><button className="primaryBtn" onClick={()=>nav('/sign')}>下一步：電子簽署</button><button className="actionBtn" onClick={resetContent}><RotateCcw size={18}/> 重設內容</button></div>
      </motion.section>}
    </div>
    <aside className="lg:sticky lg:top-24 lg:h-fit"><DocumentPreview doc={doc}/></aside>
  </div>
}

function SignPage(){
  const [doc,setDoc]=useState(load(K.doc,freshDoc()))
  useEffect(()=>save(K.doc,doc),[doc])
  function updateSigner(i,next){const arr=[...(doc.signers||[])]; arr[i]=next; setDoc({...doc,signers:arr})}
  function addSigner(){setDoc({...doc,signers:[...(doc.signers||[]),{id:uid(),name:'',title:'',date:today(),signature:'',status:'pending'}]})}
  function finish(){
    const unsigned=(doc.signers||[]).filter(s=>!s.signature).length
    if(unsigned && !confirm(`仍有 ${unsigned} 位簽署人未簽名，是否仍然完成？`)) return
    completeDoc({...doc,status:'已簽署'},'已簽署')
    alert('已完成簽署，文件已儲存至最近頁面。')
  }
  return <div className="grid gap-5 lg:grid-cols-[1fr_.9fr]">
    <div className="space-y-5">
      <section className="card"><h2 className="text-2xl font-black">電子簽署</h2><p className="mt-1 text-sm text-slate-500">支援多位簽署人，文件底部會按數量自動調整簽署欄。</p><div className="mt-4 flex flex-wrap gap-2"><button className="primaryBtn" onClick={addSigner}><Plus size={18}/> 新增簽署人</button><button className="actionBtn" onClick={finish}><CheckCircle2 size={18}/> 完成簽署</button></div></section>
      {(doc.signers||[]).map((s,i)=><Signer key={s.id} signer={s} index={i} onChange={x=>updateSigner(i,x)} onDelete={doc.signers.length>1?()=>{if(confirm('確定刪除此簽署人？')) setDoc({...doc,signers:doc.signers.filter((_,n)=>n!==i)})}:null}/>)}
      <Toolbar doc={doc} finish={finish}/>
    </div>
    <aside className="lg:sticky lg:top-24 lg:h-fit"><DocumentPreview doc={doc}/></aside>
  </div>
}
function Signer({signer,index,onChange,onDelete}){
  const ref=useRef(null)
  return <section className="card">
    <div className="mb-3 flex justify-between"><h3 className="font-black">簽署人 {index+1}</h3>{onDelete&&<button className="smallDanger" onClick={onDelete}>刪除</button>}</div>
    <Input label="姓名" value={signer.name} set={v=>onChange({...signer,name:v})}/><Input label="職位" value={signer.title} set={v=>onChange({...signer,title:v})}/><Input label="日期" type="date" value={signer.date} set={v=>onChange({...signer,date:v})}/>
    <div className="mt-4 rounded-3xl border border-slate-200 bg-slate-50 p-3 dark:border-slate-700 dark:bg-slate-950"><div className="mb-2 flex justify-between"><p className="text-sm font-black">手寫簽名</p><button className="smallBtn" onClick={()=>ref.current?.clear()}>清除</button></div><SignatureCanvas ref={ref} penColor="black" canvasProps={{className:'signatureCanvas'}}/></div>
    <button className="primaryBtn mt-4" onClick={()=>{if(ref.current?.isEmpty()) return alert('請先簽名'); onChange({...signer,signature:ref.current.getCanvas().toDataURL('image/png'),status:'signed'})}}>確認此簽名</button>
    {signer.signature&&<p className="mt-2 text-center text-sm font-bold text-emerald-600">已儲存簽名</p>}
  </section>
}
function Toolbar({doc,finish}){
  return <section className="card"><h3 className="sectionTitle">文件工具列</h3><div className="mt-4 grid grid-cols-2 gap-3">
    <button className="actionBtn" onClick={()=>exportPdf(`${doc.title||'signed-document'}.pdf`)}><Download size={18}/> 匯出 PDF</button>
    <button className="actionBtn" onClick={()=>exportPng(`${doc.title||'signed-document'}.png`)}><Download size={18}/> 匯出 PNG</button>
    <button className="actionBtn" onClick={()=>exportPng(`${doc.title||'signed-document'}.png`,true)}><Share2 size={18}/> 分享</button>
    <button className="actionBtn" onClick={finish}><CheckCircle2 size={18}/> 完成簽署</button>
  </div></section>
}

function PdfPage(){
  const [name,setName]=useState(''), [buf,setBuf]=useState(null), [stamps,setStamps]=useState([]), [active,setActive]=useState(null), [loading,setLoading]=useState(false)
  const box=useRef(null), sig=useRef(null)
  async function loadPdf(file){
    if(!file) return
    if(file.type!=='application/pdf') return alert('請上載 PDF')
    setLoading(true); setName(file.name)
    const arr=await file.arrayBuffer(); setBuf(arr); setStamps([]); setActive(null)
    setTimeout(async()=>{
      try{
        box.current.innerHTML=''
        const pdf=await pdfjsLib.getDocument({data:arr.slice(0)}).promise
        for(let n=1;n<=pdf.numPages;n++){
          const page=await pdf.getPage(n), viewport=page.getViewport({scale:1.35})
          const wrap=document.createElement('div'); wrap.className='pdf-page-wrap'; wrap.dataset.page=String(n)
          const lab=document.createElement('div'); lab.className='pdf-page-label'; lab.textContent=`第 ${n} 頁：點擊位置新增簽署框`
          const stage=document.createElement('div'); stage.className='pdf-stage'
          const canvas=document.createElement('canvas'); canvas.className='pdf-canvas'; canvas.width=viewport.width; canvas.height=viewport.height
          await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise
          stage.appendChild(canvas); wrap.append(lab,stage); box.current.appendChild(wrap)
          stage.onclick=e=>{
            if(e.target.classList.contains('pdf-stamp-view')) return
            const r=stage.getBoundingClientRect()
            const s={id:uid(),pageNo:n,xRatio:(e.clientX-r.left)/r.width,yRatio:(e.clientY-r.top)/r.height,widthRatio:.24,heightRatio:.1,name:'',title:'',date:today(),signature:''}
            setStamps(prev=>[...prev,s]); setActive(s.id)
          }
        }
      } finally { setLoading(false) }
    },50)
  }
  const cur=stamps.find(s=>s.id===active)||stamps[0]
  function patch(id, p){setStamps(prev=>prev.map(s=>s.id===id?{...s,...p}:s))}
  function saveSig(){
    if(!cur) return alert('請先點擊 PDF 加入簽署框')
    if(sig.current?.isEmpty()) return alert('請先簽名')
    patch(cur.id,{signature:sig.current.getCanvas().toDataURL('image/png')})
    alert('簽名已加入 PDF 預覽及將會嵌入輸出檔案。')
  }
  async function exportSigned(){
    if(!buf) return alert('請先上載 PDF')
    if(!stamps.length) return alert('請先點擊 PDF 新增簽署框')
    const pdf=await PDFDocument.load(buf)
    const font=await pdf.embedFont(StandardFonts.Helvetica)
    for(const s of stamps){
      const page=pdf.getPages()[s.pageNo-1]; if(!page) continue
      const {width,height}=page.getSize()
      const x=Math.max(10,Math.min(width-40,s.xRatio*width))
      const y=Math.max(10,Math.min(height-40,height-(s.yRatio*height)))
      const boxW=(s.widthRatio||.24)*width, boxH=(s.heightRatio||.1)*height
      page.drawRectangle({x:x-3,y:y-boxH-16,width:boxW+6,height:boxH+28,borderColor:rgb(.85,.85,.85),borderWidth:.5,color:rgb(1,1,1),opacity:.05})
      if(s.signature){
        try{const png=await pdf.embedPng(s.signature); page.drawImage(png,{x,y:y-boxH+8,width:boxW,height:boxH})}catch{}
      }
      page.drawText(`${s.name||'Signer'} ${s.title?`(${s.title})`:''}`,{x,y:y-boxH,size:9,font,color:rgb(0,0,0)})
      page.drawText(s.date||today(),{x,y:y-boxH-12,size:8,font,color:rgb(.25,.25,.25)})
    }
    const bytes=await pdf.save()
    const blob=new Blob([bytes],{type:'application/pdf'})
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='signed-'+(name||'document.pdf'); a.click()
    const rec=load(K.recents,[])
    save(K.recents,[{id:uid(),title:name||'PDF 文件',company:'PDF 匯入',type:'PDF 匯入',status:'已簽署',signedAt:new Date().toISOString(),signers:stamps.map(s=>({name:s.name,title:s.title,date:s.date,signature:s.signature}))},...rec].slice(0,50))
  }
  return <div className="grid gap-5 lg:grid-cols-[1fr_380px]">
    <section className="card"><h2 className="text-2xl font-black">PDF 匯入簽署</h2><p className="mt-1 text-sm text-slate-500">真正讀取整份 PDF。點擊 PDF 任意位置新增簽署框，簽名後會顯示於預覽並嵌入原 PDF。</p>
      <label className="mt-5 flex min-h-[150px] cursor-pointer flex-col items-center justify-center rounded-[2rem] border-2 border-dashed border-slate-300 bg-slate-50 p-6 text-center dark:border-slate-700 dark:bg-slate-950"><Upload className="mb-3 text-slate-400" size={34}/><p className="font-black">點擊上載 PDF</p><input type="file" accept="application/pdf" className="hidden" onChange={e=>loadPdf(e.target.files?.[0])}/></label>
      {name&&<p className="mt-4 rounded-2xl bg-emerald-50 p-3 text-sm font-bold text-emerald-700">已匯入：{name}</p>}
      {loading&&<div className="mt-5 space-y-3"><div className="skeleton h-72"></div><div className="skeleton h-72"></div></div>}
      <div ref={box} className="mt-5 space-y-5 overflow-x-auto rounded-3xl bg-slate-100 p-3 dark:bg-slate-950">
        {stamps.map(s=><PdfStamp key={s.id} stamp={s} active={s.id===active} setActive={setActive} patch={patch}/>)}
      </div>
    </section>
    <aside className="card lg:sticky lg:top-24 lg:h-fit"><h3 className="sectionTitle">簽署框設定</h3>{cur?<div><p className="mt-3 rounded-2xl bg-slate-100 p-3 text-xs font-bold dark:bg-slate-800">第 {cur.pageNo} 頁 · 可拖曳簽署框重新定位</p><Input label="姓名" value={cur.name} set={v=>patch(cur.id,{name:v})}/><Input label="職位" value={cur.title} set={v=>patch(cur.id,{title:v})}/><Input label="日期" type="date" value={cur.date} set={v=>patch(cur.id,{date:v})}/><div className="mt-4 rounded-3xl border border-slate-200 bg-white p-3"><SignatureCanvas ref={sig} penColor="black" canvasProps={{className:'signatureCanvas'}}/></div><div className="mt-3 grid grid-cols-2 gap-2"><button className="primaryBtn" onClick={saveSig}>套用簽名</button><button className="smallDanger" onClick={()=>{setStamps(stamps.filter(s=>s.id!==cur.id));setActive(null)}}><Trash2 size={15}/> 刪除框</button></div></div>:<p className="mt-4 text-sm text-slate-500">請先上載 PDF，然後點擊文件位置新增簽署框。</p>}<button className="primaryBtn mt-5" onClick={exportSigned}><Download size={18}/> 下載已簽署 PDF</button></aside>
  </div>
}
function PdfStamp({stamp,active,setActive,patch}){
  const ref=useRef(null)
  useEffect(()=>{
    const root=document.querySelector(`.pdf-page-wrap[data-page="${stamp.pageNo}"] .pdf-stage`)
    if(root && ref.current && !root.contains(ref.current)) root.appendChild(ref.current)
  })
  return <div ref={ref} onClick={e=>{e.stopPropagation();setActive(stamp.id)}} draggable onDragEnd={e=>{
    const parent=ref.current?.parentElement?.getBoundingClientRect(); if(!parent) return
    patch(stamp.id,{xRatio:Math.max(0,Math.min(.95,(e.clientX-parent.left)/parent.width)), yRatio:Math.max(0,Math.min(.95,(e.clientY-parent.top)/parent.height))})
  }} className={`pdf-stamp-view ${active?'active':''}`} style={{left:`${stamp.xRatio*100}%`,top:`${stamp.yRatio*100}%`,width:`${(stamp.widthRatio||.24)*100}%`,minHeight:`${(stamp.heightRatio||.1)*100}%`}}>
    <div className="flex items-center gap-1 text-[10px] font-bold text-slate-500"><Move size={10}/> 簽署框</div>
    {stamp.signature?<img src={stamp.signature}/>:<span className="text-xs text-slate-400">尚未簽名</span>}
    <p>{stamp.name||'姓名'} · {stamp.title||'職位'}</p><p>{stamp.date}</p>
  </div>
}

function WordPage(){
  const nav=useNavigate(), preview=useRef(null)
  const [doc,setDoc]=useState(load(K.doc,freshDoc()))
  useEffect(()=>save(K.doc,doc),[doc])
  async function importWord(file){
    if(!file) return
    const ext=file.name.split('.').pop().toLowerCase()
    if(!['doc','docx'].includes(ext)) return alert('請上載 DOC / DOCX')
    try{
      const ab=await file.arrayBuffer()
      if(ext==='docx' && preview.current){ preview.current.innerHTML=''; await renderAsync(ab, preview.current) }
      const res=await mammoth.convertToHtml({arrayBuffer:ab})
      const text=(res.value||'').replace(/<br\s*\/?>/gi,'\n').replace(/<\/p>/gi,'\n\n').replace(/<[^>]+>/g,'').trim()
      setDoc({...doc,title:file.name.replace(/\.[^/.]+$/,''),content:text,type:'Word 匯入',status:'draft'})
      alert('已讀取 Word 內容，並轉換為可編輯文件。')
    }catch{ alert('無法讀取此 Word 文件，請嘗試 DOCX 格式。') }
  }
  return <div className="grid gap-5 lg:grid-cols-[1fr_.9fr]"><section className="space-y-5"><div className="card"><h2 className="text-2xl font-black">Word 匯入</h2><p className="mt-1 text-sm text-slate-500">支援 DOC / DOCX，匯入後可編輯、簽署及匯出 PDF。</p><label className="mt-5 flex min-h-[150px] cursor-pointer flex-col items-center justify-center rounded-[2rem] border-2 border-dashed border-slate-300 bg-slate-50 p-6 text-center dark:border-slate-700 dark:bg-slate-950"><Upload className="mb-3 text-slate-400" size={34}/><p className="font-black">點擊上載 Word 文件</p><input type="file" accept=".doc,.docx" className="hidden" onChange={e=>importWord(e.target.files?.[0])}/></label></div><div className="card"><h3 className="sectionTitle">可編輯內容</h3><Input label="文件標題" value={doc.title} set={v=>setDoc({...doc,title:v})}/><label className="label">文件內容</label><textarea className="textarea" value={doc.content} onChange={e=>setDoc({...doc,content:e.target.value})}/><button className="primaryBtn mt-4" onClick={()=>nav('/sign')}>進入電子簽署</button></div><div className="card"><h3 className="sectionTitle">Word 原格式預覽（DOCX）</h3><div ref={preview} className="docxPreview mt-3"></div></div></section><aside className="lg:sticky lg:top-24 lg:h-fit"><DocumentPreview doc={doc}/></aside></div>
}

function TemplatesPage(){
  const [custom,setCustom]=useState(load(K.templates,[])), [deleted,setDeleted]=useState(load(K.deletedDefaults,[])), [q,setQ]=useState(''), [form,setForm]=useState({title:'',category:'business',description:'',content:''})
  useEffect(()=>save(K.templates,custom),[custom]); useEffect(()=>save(K.deletedDefaults,deleted),[deleted])
  const defaults=defaultTemplates.filter(t=>!deleted.includes(t.templateId))
  const all=[...defaults,...custom]
  async function importFile(file){
    if(!file) return
    const ext=file.name.split('.').pop().toLowerCase(); let content=''
    if(ext==='txt') content=await file.text()
    else if(ext==='docx'||ext==='doc') content=(await mammoth.extractRawText({arrayBuffer:await file.arrayBuffer()})).value
    else if(ext==='pdf'){const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise, parts=[]; for(let i=1;i<=pdf.numPages;i++){const p=await pdf.getPage(i); const t=await p.getTextContent(); parts.push(t.items.map(x=>x.str).join(' '))} content=parts.join('\n\n')}
    else return alert('只支援 TXT / DOC / DOCX / PDF')
    setForm({...form,title:file.name.replace(/\.[^/.]+$/,''),content})
  }
  function add(){if(!form.title||!form.content) return alert('請輸入範本名稱及內容'); setCustom([{...form,templateId:uid(),custom:true,createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),uses:0},...custom]); setForm({title:'',category:'business',description:'',content:''})}
  const shown=all.filter(t=>(t.title+t.description+t.content).toLowerCase().includes(q.toLowerCase()))
  return <div className="space-y-5"><section className="card"><h2 className="text-2xl font-black">範本管理</h2><p className="mt-1 text-sm text-slate-500">預設範本可刪除；如需要，可一鍵恢復全部預設範本。</p><button className="smallBtn mt-3" onClick={()=>setDeleted([])}>恢復預設範本</button></section><section className="card"><h3 className="sectionTitle">建立／匯入範本</h3><Input label="範本名稱" value={form.title} set={v=>setForm({...form,title:v})}/><label className="label">分類</label><select className="input" value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{cats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select><Input label="描述" value={form.description} set={v=>setForm({...form,description:v})}/><label className="label">內容</label><textarea className="textarea" value={form.content} onChange={e=>setForm({...form,content:e.target.value})}/><div className="mt-4 flex flex-wrap gap-2"><button className="primaryBtn" onClick={add}><Plus size={18}/> 儲存範本</button><label className="actionBtn cursor-pointer"><Upload size={18}/> 匯入 TXT / DOC / DOCX / PDF<input type="file" accept=".txt,.doc,.docx,.pdf" className="hidden" onChange={e=>importFile(e.target.files?.[0])}/></label></div></section><section className="card"><div className="relative"><Search className="absolute left-3 top-3 text-slate-400" size={18}/><input className="input pl-10" value={q} onChange={e=>setQ(e.target.value)} placeholder="搜尋範本"/></div><div className="mt-4 grid gap-3">{shown.map(t=><div className="rounded-3xl border border-slate-200 p-4 dark:border-slate-800" key={t.templateId}><div className="flex justify-between gap-3"><div><h4 className="font-black">{t.title}</h4><p className="text-xs text-slate-500">{t.description}</p></div><span className="badge">{t.default?'預設':'自訂'}</span></div><p className="mt-3 line-clamp-3 text-xs leading-6 text-slate-500">{t.content}</p><div className="mt-3 flex gap-2"><button className="smallBtn" onClick={()=>setCustom([{...t,templateId:uid(),title:t.title+' 副本',default:false,custom:true},...custom])}><Copy size={14}/> 複製</button>{t.default?<button className="smallDanger" onClick={()=>confirm('刪除預設範本？可稍後恢復。')&&setDeleted([...deleted,t.templateId])}><Trash2 size={14}/> 刪除預設</button>:<button className="smallDanger" onClick={()=>confirm('確定刪除？')&&setCustom(custom.filter(x=>x.templateId!==t.templateId))}><Trash2 size={14}/> 刪除</button>}</div></div>)}</div></section></div>
}

function RecentPage(){
  const nav=useNavigate(), [items,setItems]=useState(load(K.recents,[])), [q,setQ]=useState('')
  const shown=items.filter(i=>((i.title||'')+(i.company||'')+(i.status||'')+(i.type||'')).toLowerCase().includes(q.toLowerCase()))
  function open(i,path){save(K.doc,i); nav(path)}
  return <div className="space-y-5"><section className="card"><h2 className="text-2xl font-black">最近文件</h2><p className="mt-1 text-sm text-slate-500">完成簽署後會在這裡留下紀錄，用戶可自行刪除。</p><div className="relative mt-4"><Search className="absolute left-3 top-3 text-slate-400" size={18}/><input className="input pl-10" value={q} onChange={e=>setQ(e.target.value)} placeholder="搜尋文件"/></div><button className="smallDanger mt-3" onClick={()=>confirm('確定清空所有紀錄？')&&(setItems([]),save(K.recents,[]))}>清空紀錄</button></section>{shown.map(i=><section className="card" key={i.id}><div className="flex justify-between"><div><h3 className="font-black">{i.title||'未命名文件'}</h3><p className="mt-1 text-xs text-slate-500">{i.company||''} · {i.completedAt?new Date(i.completedAt).toLocaleString():i.signedAt?new Date(i.signedAt).toLocaleString():''}</p><p className="mt-2 text-xs font-bold text-slate-400">狀態：{i.status||'草稿'} · 類型：{i.type||'document'} · 簽署人：{i.signers?.length||0}</p></div><span className="badge">{i.status}</span></div><div className="mt-4 grid grid-cols-3 gap-2"><button className="actionBtn" onClick={()=>open(i,'/')}>編輯</button><button className="actionBtn" onClick={()=>open(i,'/sign')}>簽署</button><button className="actionBtn" onClick={()=>{if(confirm('確定刪除此紀錄？')){const n=items.filter(x=>x.id!==i.id);setItems(n);save(K.recents,n)}}}>刪除</button></div></section>)}{!shown.length&&<p className="text-center text-sm text-slate-500">暫未有文件紀錄。</p>}</div>
}
function SettingsPage(){return <div className="space-y-5"><section className="card"><h2 className="text-2xl font-black">設定</h2><p className="mt-1 text-sm text-slate-500">資料儲存在你的瀏覽器 LocalStorage。</p></section><section className="card"><button className="smallDanger" onClick={()=>{if(confirm('清空所有資料？')){Object.values(K).forEach(k=>localStorage.removeItem(k));location.reload()}}}>清空所有本機資料</button></section></div>}

ReactDOM.createRoot(document.getElementById('root')).render(<BrowserRouter><Shell/></BrowserRouter>)
