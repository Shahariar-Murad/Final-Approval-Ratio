import React, { useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx-js-style'
import { REASON_MAP } from './reasonMap'
import { COUNTRY_NAMES, COUNTRY_CODES } from './countryMeta'

const PSP_ORDER = [
  'Nuvei (Safecharge)','Nuvei-AE','Nuvei-EU','Network International','Axcess','Axcess-Truevo','Axcess-Paystrax',
  'Payabl','Payabl-Global','Payabl-India','Paysafe','Unlimit','Zen (Card)','PayMaxis','PayPal','PayPal-Global',
  'Paypal-India','Trust Payment','Apple Pay via Payabl','Google Pay via Payabl'
]
const PSP_RULES = [
  ['CardPay','Unlimit','All countries','All CardPay transactions are reported under Unlimit.'],
  ['Innatech','Payabl','All countries','All Innatech transactions are reported under Payabl.'],
  ['Innatech','Payabl-Global','All except India','Country ≠ IN.'],
  ['Innatech','Payabl-India','India only','Country = IN.'],
  ['PayPal','PayPal','All countries','Overall PayPal.'],
  ['PayPal','PayPal-Global','All except India','Country ≠ IN.'],
  ['PayPal','Paypal-India','India only','Country = IN.'],
  ['SafeCharge','Nuvei (Safecharge)','All countries','Overall SafeCharge.'],
  ['SafeCharge','Nuvei-AE','MID-AE','SafeCharge-CreditCard-MID-AE.'],
  ['SafeCharge','Nuvei-EU','MID-EU','SafeCharge-CreditCard-MID-EU.'],
  ['Paystra','Axcess-Paystrax','All countries','All Paystra transactions are reported under Axcess-Paystrax.'],
  ['TrustPayments','Trust Payment','All countries','All TrustPayments transactions are reported under Trust Payment.'],
  ['Paysafe','Paysafe','All countries','All Paysafe transactions are reported under Paysafe.'],
]

const HARD_FALLBACKS = [
  ['blocked, first used or special condition','Card Issue (Stolen/ Invalid/ Wrong Card Deatils)','Hard'],
  ['temporarily blocked','Card Issue (Stolen/ Invalid/ Wrong Card Deatils)','Hard'],
  ['insufficient funds','Insufficient Fund','Hard'],
  ['insufficient fund','Insufficient Fund','Hard'],
  ['do not honor','Do Not Honor','Hard'],
  ['do not honour','Do Not Honor','Hard'],
  ['generic decline','Generic Decline','Hard'],
  ['cancelled by customer','Cancalled by Customer','Hard'],
  ['canceled by the user','Cancalled by Customer','Hard'],
  ['payment aborted by customer','Cancalled by Customer','Hard'],
  ['issuing bank','Issuing Bank Related Issue','Hard'],
  ['3dsecure verification failed','3DS / Authentication Issue','Soft'],
  ['3d secure','3DS / Authentication Issue','Soft'],
  ['suspected fraud','Fraud Issue','Soft'],
  ['fraud','Fraud Issue','Soft'],
  ['re-enter transaction','Re-enter Transaction','Soft'],
  ['risk management','Risk Management Check','Soft'],
  ['blacklisted bin','Blacklisted BIN','Soft'],
  ['banned card','Blocked / Banned Card','Soft'],
  ['security violation','Security Violation Issue','Soft'],
]

const pct = v => v == null || Number.isNaN(Number(v)) ? '—' : `${(Number(v)*100).toFixed(2)}%`
const money = v => v == null || Number.isNaN(Number(v)) ? '—' : `$${Number(v).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}`
const num = v => Number(v||0).toLocaleString()
const safe = v => v == null ? '' : String(v).trim()
const normalize = v => safe(v).toLowerCase()
const clamp = v => Math.max(0,Math.min(100,Number(v||0)*100))
const asNumber = v => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0
  const n = Number(String(v??'').replace(/[$,\s]/g,''))
  return Number.isFinite(n) ? n : 0
}

const TIMEZONE_OPTIONS = [
  {label:'GMT+0', offset:0},
  {label:'GMT+3', offset:3},
  {label:'GMT+6', offset:6},
  {label:'GMT+1', offset:1},
  {label:'GMT+2', offset:2},
  {label:'GMT+4', offset:4},
  {label:'GMT+5', offset:5},
  {label:'GMT+7', offset:7},
  {label:'GMT+8', offset:8},
  {label:'GMT+9', offset:9},
  {label:'GMT+10', offset:10},
  {label:'GMT-1', offset:-1},
  {label:'GMT-3', offset:-3},
  {label:'GMT-5', offset:-5},
  {label:'GMT-6', offset:-6},
  {label:'GMT-8', offset:-8}
]
const timezoneOffset = label => TIMEZONE_OPTIONS.find(x=>x.label===label)?.offset ?? 0
const countryCodeFromValue = value => {
  const raw=safe(value)
  if(!raw) return ''
  const upper=raw.toUpperCase()
  if(COUNTRY_NAMES[upper]) return upper
  return COUNTRY_CODES[raw.toLowerCase()] || ''
}
const countryName = value => {
  const code=countryCodeFromValue(value)
  return code ? COUNTRY_NAMES[code] : (safe(value) || 'Unknown')
}
const isIndiaCountry = value => countryCodeFromValue(value)==='IN' || normalize(value)==='india'
const dateKeyInTimezone = (value, timezone='GMT+0') => {
  const d=parseSourceDate(value)
  if(!d) return ''
  const shifted=new Date(d.getTime()+timezoneOffset(timezone)*60*60*1000)
  return shifted.toISOString().slice(0,10)
}

function reasonInfo(raw){
  const s=safe(raw)
  if(!s) return null
  if(REASON_MAP[s]) return REASON_MAP[s]
  const n=normalize(s)
  for(const [needle,reason,type] of HARD_FALLBACKS){ if(n.includes(needle)) return {reason,type} }
  return null
}
function mapSpecificPsp(psp,country,midAlias){
  const p=safe(psp), c=safe(country).toUpperCase(), mid=safe(midAlias)
  if(p==='CardPay') return 'Unlimit'
  if(p==='Innatech') return isIndiaCountry(country) ? 'Payabl-India' : 'Payabl-Global'
  if(p==='PayPal') return isIndiaCountry(country) ? 'Paypal-India' : 'PayPal-Global'
  if(p==='SafeCharge'){
    if(mid.includes('MID-AE')) return 'Nuvei-AE'
    if(mid.includes('MID-EU')) return 'Nuvei-EU'
    return 'Nuvei (Safecharge)'
  }
  if(p==='Paystra') return 'Axcess-Paystrax'
  if(p==='TrustPayments') return 'Trust Payment'
  if(p==='Paysafe') return 'Paysafe'
  return p || 'Unknown'
}
function rawRowsToObjects(rows){
  const headers=(rows[0]||[]).map(x=>safe(x))
  const index={}; headers.forEach((h,i)=>{index[normalize(h)]=i})
  const pick=(r,...names)=>{for(const n of names){const i=index[normalize(n)]; if(i!=null) return r[i]} return null}
  const required=['merchantOrderId','pspName','country','amount','status']
  if(required.filter(x=>index[normalize(x)]==null).length) throw new Error('RAW_FILE')
  return rows.slice(1).filter(r=>r && r.length).map((r,i)=>({
    _i:i,
    merchantOrderId:safe(pick(r,'merchantOrderId')) || safe(pick(r,'id')) || `ROW_${i+1}`,
    pspName:safe(pick(r,'pspName')),
    country:safe(pick(r,'country')) || 'Unknown',
    amount:asNumber(pick(r,'amount')),
    status:normalize(pick(r,'status')),
    declineReason:safe(pick(r,'declineReason','statusReason')),
    processingDate:safe(pick(r,'processing_date','processingDate')),
    midAlias:safe(pick(r,'midAlias')),
    currency:safe(pick(r,'currency'))
  }))
}

function parseSourceDate(value){
  const s=safe(value)
  if(!s) return null
  if(/^\d{2}\/\d{2}\/\d{4}/.test(s)){
    const m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/)
    if(m) return new Date(Date.UTC(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(m[4]||0),Number(m[5]||0),Number(m[6]||0)))
  }
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){
    const iso=s.includes('T')?s:s.replace(' ','T')
    return new Date(/[zZ]|[+-]\d{2}:?\d{2}$/.test(iso)?iso:`${iso}Z`)
  }
  const d=new Date(s)
  return Number.isNaN(d.getTime())?null:d
}
function gmt3DateKey(value){ return dateKeyInTimezone(value,'GMT+3') }
function matchesMulti(value, selected, mode){
  if(!selected?.length) return true
  const hit=selected.includes(value)
  return mode==='exclude' ? !hit : hit
}
function applyRawFilters(rows, filters){
  if(!filters) return rows
  return rows.filter(x=>{
    const dateKey=dateKeyInTimezone(x.processingDate,filters.timezone||'GMT+0')
    if(filters.startDate && (!dateKey || dateKey < filters.startDate)) return false
    if(filters.endDate && (!dateKey || dateKey > filters.endDate)) return false
    if(!matchesMulti(safe(x.pspName),filters.pspSelected,filters.pspMode)) return false
    if(!matchesMulti(safe(x.midAlias)||'Unknown',filters.midSelected,filters.midMode)) return false
    if(!matchesMulti(countryName(x.country),filters.countrySelected,filters.countryMode)) return false
    return true
  })
}

function buildFromRaw(rows,fileName='Transaction Data',filterConfig=null){
  const timezone=filterConfig?.timezone || 'GMT+0'
  const parsed=rawRowsToObjects(rows)
  const r=applyRawFilters(parsed,{...(filterConfig||{}),timezone})
  const approvedRows=r.filter(x=>x.status==='approved')
  const declinedRows=r.filter(x=>x.status==='declined')
  const otherRows=r.filter(x=>!['approved','declined'].includes(x.status))

  const ordersMap=new Map()
  for(const x of r){
    let o=ordersMap.get(x.merchantOrderId)
    if(!o){o={merchantOrderId:x.merchantOrderId,country:x.country||'Unknown',approved:false,revenue:0,attempts:0,firstDeclineReason:'',firstDeclineInfo:null,rows:[]};ordersMap.set(x.merchantOrderId,o)}
    o.attempts++
    o.rows.push(x)
    if(!o.country || o.country==='Unknown') o.country=x.country||'Unknown'
    if(x.status==='approved'){o.approved=true;o.revenue+=x.amount}
    if(!o.firstDeclineReason && x.declineReason){o.firstDeclineReason=x.declineReason;o.firstDeclineInfo=reasonInfo(x.declineReason)}
  }
  const orders=[...ordersMap.values()]
  const terminal=orders.filter(o=>!o.approved)
  for(const o of terminal){ if(!o.firstDeclineInfo) o.firstDeclineInfo=reasonInfo(o.firstDeclineReason) }
  const hard=terminal.filter(o=>o.firstDeclineInfo?.type==='Hard')
  const soft=terminal.filter(o=>o.firstDeclineInfo?.type==='Soft')
  const approved=orders.filter(o=>o.approved)
  const total=orders.length
  const ar=total?approved.length/total:0
  const arNoHard=(approved.length+soft.length)?approved.length/(approved.length+soft.length):0
  const hardPct=(hard.length+soft.length)?hard.length/(hard.length+soft.length):0
  const softPct=(hard.length+soft.length)?soft.length/(hard.length+soft.length):0
  const reasonCounts={}; for(const o of terminal){const key=o.firstDeclineInfo?.reason || 'Unclassified'; reasonCounts[key]=(reasonCounts[key]||0)+1}
  const topReason=Object.entries(reasonCounts).sort((a,b)=>b[1]-a[1])[0]?.[0]||'Unclassified'
  const periodDates=r.map(x=>new Date(x.processingDate)).filter(d=>!Number.isNaN(d.getTime()))
  let period='Uploaded transaction period'
  if(periodDates.length){
    const shifted=periodDates.map(d=>new Date(d.getTime()+timezoneOffset(timezone)*60*60*1000)).sort((a,b)=>a-b)
    const f=shifted[0], l=shifted[shifted.length-1]
    const fmt=d=>d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'})
    period=`${fmt(f)} – ${fmt(l)} · ${timezone}`
  }

  const makeKpi=(label,a,rev)=>({Metric:label,Count:total,Approved:approved.length,Declined:terminal.length,'Hard Declines':hard.length,'Soft Declines':soft.length,'Displayed Approval Ratio':a,'Approval Ratio without Hard Decline':arNoHard,'Hard Decline %':hardPct,'Soft Decline %':softPct,'Top Declined Reason':topReason,'Revenue (USD)':rev})
  const allRevenue=approved.reduce((s,o)=>s+o.revenue,0)
  const withoutIndiaOrders=orders.filter(o=>!isIndiaCountry(o.country))
  const wiApproved=withoutIndiaOrders.filter(o=>o.approved), wiTerminal=withoutIndiaOrders.filter(o=>!o.approved)
  const wiHard=wiTerminal.filter(o=>o.firstDeclineInfo?.type==='Hard'), wiSoft=wiTerminal.filter(o=>o.firstDeclineInfo?.type==='Soft')
  const wiAr=withoutIndiaOrders.length?wiApproved.length/withoutIndiaOrders.length:0
  const wiNoHard=(wiApproved.length+wiSoft.length)?wiApproved.length/(wiApproved.length+wiSoft.length):0
  const wiReasons={}; for(const o of wiTerminal){const k=o.firstDeclineInfo?.reason||'Unclassified';wiReasons[k]=(wiReasons[k]||0)+1}
  const wiTop=Object.entries(wiReasons).sort((a,b)=>b[1]-a[1])[0]?.[0]||'Unclassified'
  const wiRevenue=wiApproved.reduce((s,o)=>s+o.revenue,0)
  const kpis=[
    makeKpi('Cascaded approval ratio',ar,allRevenue),
    {...makeKpi('Cascaded approval ratio without Hard Decline',arNoHard,allRevenue), 'Displayed Approval Ratio':arNoHard},
    {Metric:'Cascaded approval ratio - Without India',Count:withoutIndiaOrders.length,Approved:wiApproved.length,Declined:wiTerminal.length,'Hard Declines':wiHard.length,'Soft Declines':wiSoft.length,'Displayed Approval Ratio':wiAr,'Approval Ratio without Hard Decline':wiNoHard,'Hard Decline %':(wiHard.length+wiSoft.length)?wiHard.length/(wiHard.length+wiSoft.length):0,'Soft Decline %':(wiHard.length+wiSoft.length)?wiSoft.length/(wiHard.length+wiSoft.length):0,'Top Declined Reason':wiTop,'Revenue (USD)':wiRevenue},
    {Metric:'Cascaded approval ratio - Without India and Hard Decline',Count:withoutIndiaOrders.length,Approved:wiApproved.length,Declined:wiTerminal.length,'Hard Declines':wiHard.length,'Soft Declines':wiSoft.length,'Displayed Approval Ratio':wiNoHard,'Approval Ratio without Hard Decline':wiNoHard,'Hard Decline %':(wiHard.length+wiSoft.length)?wiHard.length/(wiHard.length+wiSoft.length):0,'Soft Decline %':(wiHard.length+wiSoft.length)?wiSoft.length/(wiHard.length+wiSoft.length):0,'Top Declined Reason':wiTop,'Revenue (USD)':wiRevenue}
  ]

  const countryMap=new Map()
  for(const x of r){
    const code=countryCodeFromValue(x.country), key=code||normalize(x.country)||'unknown';
    let g=countryMap.get(key); if(!g) g={country:countryName(x.country),countryCode:code||'',count:0,approved:0,revenue:0,orders:[],attemptRows:[]}; countryMap.set(key,g); g.attemptRows.push(x)
  }
  for(const o of orders){const code=countryCodeFromValue(o.country), key=code||normalize(o.country)||'unknown'; const g=countryMap.get(key)||{country:countryName(o.country),countryCode:code||'',count:0,approved:0,revenue:0,orders:[],attemptRows:[]}; g.count++;if(o.approved)g.approved++;g.revenue+=o.revenue;g.orders.push(o);countryMap.set(key,g)}

  const countryRecords=[...countryMap.values()].map(g=>{
    const terminalOrders=g.orders.filter(o=>!o.approved)
    const h=terminalOrders.filter(o=>o.firstDeclineInfo?.type==='Hard').length
    const s=terminalOrders.filter(o=>o.firstDeclineInfo?.type==='Soft').length
    const rr={}; terminalOrders.forEach(o=>{const k=o.firstDeclineInfo?.reason||'Unclassified';rr[k]=(rr[k]||0)+1})
    const tr=Object.entries(rr).sort((a,b)=>b[1]-a[1])[0]?.[0]||'Unclassified'
    const pspGroups=new Map(); for(const x of g.attemptRows){const p=mapSpecificPsp(x.pspName,x.country,x.midAlias);let q=pspGroups.get(p);if(!q)q={psp:p,attempts:0,approved:0,declined:0};q.attempts++;if(x.status==='approved')q.approved++;if(x.status==='declined')q.declined++;pspGroups.set(p,q)}
    let candidates=[...pspGroups.values()].filter(q=>q.attempts>=20)
    if(!candidates.length)candidates=[...pspGroups.values()]
    candidates=candidates.filter(q=>q.approved+q.declined>0).map(q=>({...q,approval:q.approved/(q.approved+q.declined)})).sort((a,b)=>b.approval-a.approval||b.attempts-a.attempts)
    const top=candidates[0]||{psp:'—',approval:null,attempts:0}
    return {rank:0,country:g.country,code:g.countryCode,amount:g.revenue,count:g.count,approval:g.count?g.approved/g.count:0,approvalWithoutHard:(g.approved+h+s)?g.approved/(g.approved+s):0,hardPct:(h+s)?h/(h+s):0,softPct:(h+s)?s/(h+s):0,topReason:tr,topPsp:top.psp,topPspApproval:top.approval,topPspCount:top.attempts}
  }).sort((a,b)=>b.amount-a.amount||b.count-a.count).map((x,i)=>({...x,rank:i+1}))

  const pspDefs=[]
  const pspSubset=(label, filter, source)=>{
    const g=r.filter(filter); const at=g.length, ap=g.filter(x=>x.status==='approved').length, de=g.filter(x=>x.status==='declined').length, other=at-ap-de
    const revenue=g.filter(x=>x.status==='approved').reduce((s,x)=>s+x.amount,0)
    const hardC=g.filter(x=>x.status==='declined' && reasonInfo(x.declineReason)?.type==='Hard').length
    const softC=g.filter(x=>x.status==='declined' && reasonInfo(x.declineReason)?.type==='Soft').length
    const unC=Math.max(0,de-hardC-softC); const rc={}; g.filter(x=>x.status==='declined').forEach(x=>{const inf=reasonInfo(x.declineReason);if(inf)rc[inf.reason]=(rc[inf.reason]||0)+1})
    const tr=Object.entries(rc).sort((a,b)=>b[1]-a[1])[0]?.[0]||'Unclassified'
    return {psp:label,source,attempts:at,approved:ap,declined:de,pendingOther:other,approval:(ap+de)?ap/(ap+de):null,approvalWithoutHard:(ap+softC)?ap/(ap+softC):null,revenue,hard:hardC,soft:softC,unclassified:unC,hardPct:(hardC+softC)?hardC/(hardC+softC):null,softPct:(hardC+softC)?softC/(hardC+softC):null,topReason:tr,topCount:rc[tr]||0}
  }
  const sourceFilter=(p)=>x=>x.pspName===p
  const def=[]
  def.push(pspSubset('Nuvei (Safecharge)',sourceFilter('SafeCharge'),'SafeCharge (all MIDs)'))
  def.push(pspSubset('Nuvei-AE',x=>x.pspName==='SafeCharge'&&safe(x.midAlias).includes('MID-AE'),'SafeCharge-CreditCard-MID-AE'))
  def.push(pspSubset('Nuvei-EU',x=>x.pspName==='SafeCharge'&&safe(x.midAlias).includes('MID-EU'),'SafeCharge-CreditCard-MID-EU'))
  def.push(...['Network International','Axcess','Axcess-Truevo'].map(x=>({psp:x,source:'No matching source data',attempts:0,approved:0,declined:0,pendingOther:0,approval:null,approvalWithoutHard:null,revenue:0,hard:0,soft:0,unclassified:0,hardPct:null,softPct:null,topReason:'Unclassified',topCount:0})))
  def.push(pspSubset('Axcess-Paystrax',sourceFilter('Paystra'),'Paystra'))
  def.push(pspSubset('Payabl',sourceFilter('Innatech'),'Innatech (all countries)'))
  def.push(pspSubset('Payabl-Global',x=>x.pspName==='Innatech'&&!isIndiaCountry(x.country),'Innatech → Payabl (non-India)'))
  def.push(pspSubset('Payabl-India',x=>x.pspName==='Innatech'&&isIndiaCountry(x.country),'Innatech → Payabl (India)'))
  def.push(pspSubset('Paysafe',sourceFilter('Paysafe'),'Paysafe'))
  def.push(pspSubset('Unlimit',sourceFilter('CardPay'),'CardPay → Unlimit'))
  def.push(...['Zen (Card)','PayMaxis'].map(x=>({psp:x,source:'No matching source data',attempts:0,approved:0,declined:0,pendingOther:0,approval:null,approvalWithoutHard:null,revenue:0,hard:0,soft:0,unclassified:0,hardPct:null,softPct:null,topReason:'Unclassified',topCount:0})))
  def.push(pspSubset('PayPal',sourceFilter('PayPal'),'PayPal (all countries)'))
  def.push(pspSubset('PayPal-Global',x=>x.pspName==='PayPal'&&!isIndiaCountry(x.country),'PayPal (non-India)'))
  def.push(pspSubset('Paypal-India',x=>x.pspName==='PayPal'&&isIndiaCountry(x.country),'PayPal (India)'))
  def.push(pspSubset('Trust Payment',sourceFilter('TrustPayments'),'TrustPayments'))
  def.push(...['Apple Pay via Payabl','Google Pay via Payabl'].map(x=>({psp:x,source:'No matching source data',attempts:0,approved:0,declined:0,pendingOther:0,approval:null,approvalWithoutHard:null,revenue:0,hard:0,soft:0,unclassified:0,hardPct:null,softPct:null,topReason:'Unclassified',topCount:0})))

  const taxonomy={}
  for(const x of r.filter(x=>x.status==='declined')){const inf=reasonInfo(x.declineReason);if(inf){const k=inf.reason+'||'+inf.type;taxonomy[k]=(taxonomy[k]||0)+1}}
  const taxonomyRows=Object.entries(taxonomy).map(([k,count])=>{const [reason,type]=k.split('||');return {reason,type,count}}).sort((a,b)=>b.count-a.count)

  return {period,timezone,fileName,rawRows:rows,rawObjects:parsed,filterable:true,kpis,pspWise:def,countries:countryRecords,mappingRules:PSP_RULES.map(x=>({source:x[0],dashboard:x[1],scope:x[2],rule:x[3]})),taxonomy:taxonomyRows,methodology:[
    {metric:'Period',definition:period,note:'processing_date is interpreted as source GMT+0 and converted to the selected analysis timezone.'},
    {metric:'Unique transaction key',definition:'merchantOrderId',note:'Repeated PSP attempts are cascaded into one country-level outcome.'},
    {metric:'Cascaded outcome',definition:'Approved if any attempt is approved; otherwise Declined',note:'Approved wins on retry/cascade.'},
    {metric:'Approval Ratio',definition:'Approved unique orders / total unique orders',note:'Country/cascaded calculation.'},
    {metric:'Approval Ratio without Hard Decline',definition:'Approved / (Approved + Soft Declines)',note:'Hard terminal declines removed from denominator.'},
    {metric:'PSP Approval Ratio',definition:'Approved PSP attempts / (Approved + Declined PSP attempts)',note:'PSP calculation does NOT deduplicate merchantOrderId.'},
    {metric:'Revenue',definition:'Sum of approved transaction amounts',note:'Reported by country and PSP.'},
    {metric:'Top Country ranking',definition:'Countries ranked by approved revenue',note:'Full country list is available; no Top-50 limit.'},
    {metric:'Top Performing PSP',definition:'Highest attempt-based PSP approval ratio in the country',note:'Minimum 20 attempts where available; otherwise best available PSP.'}
  ]}
}

function parseAnalysisWorkbook(wb,fileName){
  const sheet=(name)=>wb.Sheets[name]
  if(!sheet('Executive Summary')) throw new Error('ANALYSIS_FILE')
  const rowsOf=name=>XLSX.utils.sheet_to_json(sheet(name),{header:1,defval:null})
  const findHeader=(rows,key)=>rows.findIndex(r=>r?.some(v=>normalize(v)===normalize(key)))
  const objects=(name,key)=>{const rows=rowsOf(name);const h=findHeader(rows,key);return h<0?[]:XLSX.utils.sheet_to_json(sheet(name),{range:h,defval:null})}
  const pspSheet=sheet('PSP Wise')?'PSP Wise':null
  const countrySheet=sheet('All Countries')?'All Countries':(sheet('Top 50 Countries')?'Top 50 Countries':null)
  const pspRows=pspSheet?objects(pspSheet,'PSP'):[]
  const countryRows=countrySheet?objects(countrySheet,'Country'):[]
  const mapRows=sheet('Mapping Rules')?objects('Mapping Rules','Source PSP'):[]
  const kRows=objects('Executive Summary','Metric')
  const methodRows=sheet('Methodology')?objects('Methodology','Metric'):[]
  return {
    period:'Uploaded analysis workbook',timezone:'GMT+0',fileName,filterable:false,
    kpis:kRows.filter(x=>x.Metric).map(x=>({'Metric':x.Metric,Count:Number(x.Count||0),Approved:Number(x.Approved||0),Declined:Number(x.Declined||0),'Hard Declines':Number(x['Hard Declines']||0),'Soft Declines':Number(x['Soft Declines']||0),'Displayed Approval Ratio':Number(x['Displayed Approval Ratio']||0),'Approval Ratio without Hard Decline':Number(x['Approval Ratio without Hard Decline']||0),'Hard Decline %':Number(x['Hard Decline %']||0),'Soft Decline %':Number(x['Soft Decline %']||0),'Top Declined Reason':safe(x['Top Declined Reason']),'Revenue (USD)':Number(x['Revenue (USD)']||0)})),
    pspWise:pspRows.filter(x=>x.PSP).map(x=>({psp:safe(x.PSP),source:safe(x['Source Mapping']),attempts:Number(x.Attempts||0),approved:Number(x.Approved||0),declined:Number(x.Declined||0),pendingOther:Number(x['Pending / Other']||x['Pending/Other']||0),approval:Number(x['PSP Approval Ratio']??x['Approval Ratio']),approvalWithoutHard:Number(x['Approval Ratio without Hard Decline']||0),revenue:Number(x['Revenue (USD)']||0),hard:Number(x['Hard Declines']||0),soft:Number(x['Soft Declines']||0),unclassified:Number(x.Unclassified||0),hardPct:Number(x['Hard %']),softPct:Number(x['Soft %']),topReason:safe(x['Top Declined Reason']),topCount:Number(x['Top Declined Count']||0)})),
    countries:countryRows.filter(x=>x.Rank).map(x=>({rank:Number(x.Rank),country:safe(x.Country),code:safe(x['Country Code']),amount:asNumber(x['Approved Revenue (USD)']??String(x['Amount / Duration']||'').split(' / ')[0]),count:Number(x.Count||0),approval:Number(x['Approval Ratio']||0),approvalWithoutHard:Number(x['Approval Ratio without Hard Decline']||0),hardPct:Number(x['Hard Decline %']||0),softPct:Number(x['Soft Decline %']||0),topReason:safe(x['Top Declined Reason']),topPsp:safe(x['Top Performing PSP']),topPspApproval:Number(x['Top PSP Approval']||0),topPspCount:Number(x['Top PSP Count']||0)})),
    mappingRules:mapRows.filter(x=>x['Source PSP']).map(x=>({source:safe(x['Source PSP']),dashboard:safe(x['Dashboard PSP']),scope:safe(x['Country Scope']||x['Country scope']),rule:safe(x.Rule)})),
    taxonomy:[],
    methodology:methodRows.filter(x=>x.Metric).map(x=>({metric:safe(x.Metric),definition:safe(x.Definition),note:safe(x['Rule / Note'])}))
  }
}

function readFile(file){
  return new Promise((resolve,reject)=>{const fr=new FileReader();fr.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:'array'});const first=wb.SheetNames.find(n=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:null,raw:false});const h=(rows[0]||[]).map(normalize);return h.includes('merchantorderid')&&h.includes('pspname')&&h.includes('status')})||wb.SheetNames[0];const rows=XLSX.utils.sheet_to_json(wb.Sheets[first],{header:1,defval:null,raw:false});const headers=(rows[0]||[]).map(normalize);if(headers.includes('merchantorderid')&&headers.includes('pspname')&&headers.includes('status')) resolve(buildFromRaw(rows,file.name)); else resolve(parseAnalysisWorkbook(wb,file.name));}catch(err){reject(err)}};fr.onerror=reject;fr.readAsArrayBuffer(file)})
}

function RatioBar({value,label,accent='blue'}){return <div className="ratio-row"><div className="ratio-label"><span>{label}</span><strong>{pct(value)}</strong></div><div className="bar"><span className={`bar-fill ${accent}`} style={{width:`${clamp(value)}%`}}/></div></div>}
function KpiCard({label,value,sub,tone='blue'}){return <div className={`kpi-card ${tone}`}><div className="kpi-label">{label}</div><div className="kpi-value">{value}</div>{sub&&<div className="kpi-sub">{sub}</div>}</div>}
function UploadBox({onFile}){const ref=useRef();const [drag,setDrag]=useState(false);const choose=e=>{const f=e.target.files?.[0];if(f)onFile(f)};return <div className={`upload-box ${drag?'dragging':''}`} onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files?.[0];if(f)onFile(f)}}><div className="upload-icons"><span>↥</span><span>▤</span><span>▧</span></div><h2>Upload Transaction Data</h2><p>Drag and drop your CSV or Excel file here, or click to browse</p><small>Supports .csv, .xlsx, and .xls formats</small><button className="browse-btn" onClick={()=>ref.current?.click()}>Browse Files</button><input ref={ref} type="file" accept=".csv,.xlsx,.xls" onChange={choose}/></div>}

function MultiSelectFilter({title,icon,options,selected,setSelected,mode,setMode,placeholder}){
  const [open,setOpen]=useState(false)
  const allSelected=selected.length===0
  const toggle=(value)=>setSelected(selected.includes(value)?selected.filter(x=>x!==value):[...selected,value])
  const selectAll=()=>setSelected([])
  return <div className="filter-group">
    <div className="filter-title"><span className="filter-icon">{icon}</span><span>{title} ({selected.length})</span></div>
    <div className="filter-mode"><button className={mode==='include'?'active':''} onClick={()=>setMode('include')}>Include</button><button className={mode==='exclude'?'active':''} onClick={()=>setMode('exclude')}>Exclude</button></div>
    <div className={`multi-select ${open?'open':''}`}>
      <button className="multi-trigger" onClick={()=>setOpen(v=>!v)}><span>{allSelected?placeholder:`${selected.length} selected`}</span><span>⌄</span></button>
      {open&&<div className="multi-menu">
        <button className={`multi-option select-all ${allSelected?'checked':''}`} onClick={selectAll}><span className="check">{allSelected?'✓':''}</span>Select All</button>
        <div className="multi-divider"/>
        {options.map(v=><button key={v} className={`multi-option ${selected.includes(v)?'checked':''}`} onClick={()=>toggle(v)}><span className="check">{selected.includes(v)?'✓':''}</span>{v}</button>)}
      </div>}
    </div>
    {selected.length>0&&<div className="selection-chips">{selected.map(v=><span className="selection-chip" key={v}>{v}<button onClick={()=>toggle(v)}>×</button></span>)}</div>}
  </div>
}

function DateFilter({startDate,setStartDate,endDate,setEndDate,timezone,setTimezone}){
  return <div className="filter-group">
    <div className="filter-title"><span className="filter-icon">▣</span><span>Date Range</span></div>
    <label className="date-label">Analysis Timezone<select className="timezone-select" value={timezone} onChange={e=>setTimezone(e.target.value)}>{TIMEZONE_OPTIONS.map(t=><option key={t.label} value={t.label}>{t.label}</option>)}</select></label>
    <label className="date-label">Start Date<input type="date" value={startDate} onChange={e=>setStartDate(e.target.value)} /></label>
    <label className="date-label">End Date<input type="date" value={endDate} onChange={e=>setEndDate(e.target.value)} /></label>
  </div>
}

function App(){
 const [source,setSource]=useState(null),[tab,setTab]=useState('overview'),[approvalMode,setApprovalMode]=useState('with'),[countryQuery,setCountryQuery]=useState(''),[countryPsp,setCountryPsp]=useState('All'),[pspQuery,setPspQuery]=useState(''),[pspMin,setPspMin]=useState(20),[declineType,setDeclineType]=useState('All'),[declineQuery,setDeclineQuery]=useState(''),[notice,setNotice]=useState(''),[loading,setLoading]=useState(false)
 const [startDate,setStartDate]=useState(''),[endDate,setEndDate]=useState(''),[timezone,setTimezone]=useState('GMT+0')
 const [pspSelected,setPspSelected]=useState([]),[pspMode,setPspMode]=useState('include')
 const [midSelected,setMidSelected]=useState([]),[midMode,setMidMode]=useState('include')
 const [countrySelected,setCountrySelected]=useState([]),[countryMode,setCountryMode]=useState('include')
 const filters=useMemo(()=>({startDate,endDate,timezone,pspSelected,pspMode,midSelected,midMode,countrySelected,countryMode}),[startDate,endDate,timezone,pspSelected,pspMode,midSelected,midMode,countrySelected,countryMode])
 const data=useMemo(()=>{if(!source)return null;if(!source.filterable)return source;return buildFromRaw(source.rawRows,source.fileName,filters)},[source,filters])
 const options=useMemo(()=>{
   if(!source?.filterable)return {psps:[],mids:[],countries:[]}
   const rows=source.rawObjects||[]
   return {
     psps:Array.from(new Set(rows.map(x=>safe(x.pspName)).filter(Boolean))).sort(),
     mids:Array.from(new Set(rows.map(x=>safe(x.midAlias)||'Unknown'))).sort(),
     countries:Array.from(new Set(rows.map(x=>countryName(x.country)||'Unknown'))).sort()
   }
 },[source])
 const activeFilterCount=[startDate,endDate,pspSelected.length,midSelected.length,countrySelected.length].filter(Boolean).length
 const clearFilters=()=>{setStartDate('');setEndDate('');setTimezone('GMT+0');setPspSelected([]);setMidSelected([]);setCountrySelected([]);setPspMode('include');setMidMode('include');setCountryMode('include')}
 const handleFile=async file=>{setLoading(true);try{const d=await readFile(file);setSource(d);clearFilters();setTab('overview');setCountryQuery('');setPspQuery('');setNotice(`Loaded ${file.name}`)}catch(e){setNotice(e.message==='RAW_FILE'?'This file does not contain the required transaction columns (merchantOrderId, pspName, status, amount).':'Could not read this file. Please upload a transaction CSV/Excel report.')}finally{setLoading(false);setTimeout(()=>setNotice(''),4500)}}
 const currentKpi=data?.kpis?.[approvalMode==='with'?0:1],withoutIndia=data?.kpis?.[approvalMode==='with'?2:3]
 const countries=useMemo(()=>{if(!data)return[];let a=[...data.countries];if(countryQuery){const q=countryQuery.toLowerCase();a=a.filter(x=>`${x.country} ${x.code} ${x.topReason} ${x.topPsp}`.toLowerCase().includes(q))}if(countryPsp!=='All')a=a.filter(x=>x.topPsp===countryPsp);return a},[data,countryQuery,countryPsp])
 const pspRows=useMemo(()=>{if(!data)return[];let a=data.pspWise.filter(x=>x.attempts>=pspMin);if(pspQuery){const q=pspQuery.toLowerCase();a=a.filter(x=>`${x.psp} ${x.source} ${x.topReason}`.toLowerCase().includes(q))}return [...a].sort((a,b)=>(b.approval??-1)-(a.approval??-1))},[data,pspQuery,pspMin])
 const declineRows=useMemo(()=>{if(!data)return[];let a=[...data.taxonomy];if(declineType!=='All')a=a.filter(x=>x.type===declineType);if(declineQuery)a=a.filter(x=>x.reason.toLowerCase().includes(declineQuery.toLowerCase()));return a.sort((a,b)=>b.count-a.count)},[data,declineType,declineQuery])
 const topCountries=useMemo(()=>data?[...data.countries].slice(0,10):[],[data])
 const topPsps=useMemo(()=>data?[...data.pspWise].filter(x=>x.attempts>=pspMin).sort((a,b)=>(b.approval??-1)-(a.approval??-1)).slice(0,10):[],[data,pspMin])
 const pspOptions=data?['All',...Array.from(new Set(data.countries.map(x=>x.topPsp))).filter(Boolean).sort()]:[]
 const downloadFullExcel=()=>{
   if(!data)return
   const wb=XLSX.utils.book_new()
   const moneyCols=new Set(['Revenue (USD)','Revenue','Approved Revenue (USD)','Amount / Duration (USD)'])
   const pctCols=new Set(['Approval Ratio','Approval Ratio without Hard Decline','Displayed Approval Ratio','Hard Decline %','Soft Decline %','Top PSP Approval','PSP Approval Ratio','Hard %','Soft %'])
   const addSheet=(name,title,subtitle,rows)=>{
     const keys=rows.length?Object.keys(rows[0]):[]
     const aoa=[[title],[subtitle],[],keys,...rows.map(r=>keys.map(k=>r[k] ?? ''))]
     const ws=XLSX.utils.aoa_to_sheet(aoa)
     const lastCol=Math.max(keys.length,1), lastRow=aoa.length
     ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:lastCol-1}},{s:{r:1,c:0},e:{r:1,c:lastCol-1}}]
     ws['!cols']=keys.map(k=>({wch:Math.min(Math.max(String(k).length+4,14),32)}))
     if(lastCol>0) ws['!autofilter']={ref:`A4:${XLSX.utils.encode_col(lastCol-1)}${lastRow}`}
     const titleStyle={font:{name:'Aptos Display',sz:16,bold:true,color:{rgb:'FFFFFFFF'}},fill:{patternType:'solid',fgColor:{rgb:'FF163A5F'}},alignment:{vertical:'center'}}
     const subtitleStyle={font:{name:'Aptos',sz:10,italic:true,color:{rgb:'FF5B6B7F'}},fill:{patternType:'solid',fgColor:{rgb:'FFEAF0F6'}}}
     const headerStyle={font:{name:'Aptos',sz:10,bold:true,color:{rgb:'FFFFFFFF'}},fill:{patternType:'solid',fgColor:{rgb:'FF2F6FED'}},alignment:{horizontal:'center',vertical:'center',wrapText:true}}
     const bodyStyle={font:{name:'Aptos',sz:10,color:{rgb:'FF172033'}},alignment:{vertical:'center'}}
     const altStyle={font:{name:'Aptos',sz:10,color:{rgb:'FF172033'}},fill:{patternType:'solid',fgColor:{rgb:'FFF6F8FB'}},alignment:{vertical:'center'}}
     ws['A1'].s=titleStyle; ws['A2'].s=subtitleStyle
     for(let c=0;c<lastCol;c++){
       const hc=ws[XLSX.utils.encode_cell({r:3,c})]; if(hc)hc.s=headerStyle
     }
     for(let r=4;r<aoa.length;r++){
       for(let c=0;c<lastCol;c++){
         const cell=ws[XLSX.utils.encode_cell({r,c})]
         if(!cell)continue
         cell.s=(r%2===0)?altStyle:bodyStyle
       }
     }
     for(let r=4;r<aoa.length;r++){
       for(let c=0;c<keys.length;c++){
         const key=keys[c], cell=ws[XLSX.utils.encode_cell({r,c})]
         if(!cell)continue
         if(moneyCols.has(key))cell.z='$#,##0.00'
         if(pctCols.has(key))cell.z='0.00%'
       }
     }
     ws['!rows']=[{hpt:26},{hpt:20},{hpt:8},{hpt:30}]
     XLSX.utils.book_append_sheet(wb,ws,name)
   }
   addSheet('Executive Summary','FundedNext Approval Ratio Analysis',`Analysis timezone: ${timezone} | ${data.period}`,data.kpis.map(x=>({...x,'Analysis Timezone':timezone})))
   addSheet('All Countries','All Countries — Revenue Ranked',`Full country list | Revenue highest to lowest | ${timezone}`,data.countries.map(x=>({Rank:x.rank,Country:x.country,'Country Code':x.code,'Approved Revenue (USD)':x.amount,Count:x.count,'Approval Ratio':x.approval,'Approval Ratio without Hard Decline':x.approvalWithoutHard,'Hard Decline %':x.hardPct,'Soft Decline %':x.softPct,'Top Declined Reason':x.topReason,'Top Performing PSP':x.topPsp,'Top PSP Approval':x.topPspApproval,'Top PSP Count':x.topPspCount})))
   addSheet('PSP Wise','PSP Wise — Attempt Level',`PSP approval does not deduplicate merchantOrderId | ${timezone}`,data.pspWise.map(x=>({PSP:x.psp,'Source Mapping':x.source,Attempts:x.attempts,Approved:x.approved,Declined:x.declined,'Pending / Other':x.pendingOther,'PSP Approval Ratio':x.approval,'Approval Ratio without Hard Decline':x.approvalWithoutHard,'Revenue (USD)':x.revenue,'Hard Declines':x.hard,'Soft Declines':x.soft,Unclassified:x.unclassified,'Hard %':x.hardPct,'Soft %':x.softPct,'Top Declined Reason':x.topReason,'Top Declined Count':x.topCount})))
   addSheet('Hard Soft Declines','Hard & Soft Decline List','Standardized decline taxonomy from the uploaded transaction data',data.taxonomy.map(x=>({'Decline Reason':x.reason,Type:x.type,Count:x.count})))
   addSheet('Mapping Rules','PSP Mapping Rules','Dashboard naming and country scope',data.mappingRules.map(x=>({'Source PSP':x.source,'Dashboard PSP':x.dashboard,'Country Scope':x.scope,Rule:x.rule})))
   addSheet('Methodology','Calculation Methodology',`Selected timezone: ${timezone}`,data.methodology.map(x=>({Metric:x.metric,Definition:x.definition,'Rule / Note':x.note})))
   addSheet('Applied Filters','Applied Filters','Filters used for this exported analysis',[
     {Filter:'Analysis Timezone',Value:timezone},
     {Filter:'Start Date',Value:startDate||'All'},
     {Filter:'End Date',Value:endDate||'All'},
     {Filter:'PSP Mode',Value:pspMode},
     {Filter:'PSP Selection',Value:pspSelected.length?pspSelected.join(', '):'All'},
     {Filter:'MID Alias Mode',Value:midMode},
     {Filter:'MID Alias Selection',Value:midSelected.length?midSelected.join(', '):'All'},
     {Filter:'Country Mode',Value:countryMode},
     {Filter:'Country Selection',Value:countrySelected.length?countrySelected.join(', '):'All'}
   ])
   addSheet('Report Notes','Report Notes','Core calculation rules',[
     {Item:'Country / cascaded approval',Rule:'Unique merchantOrderId; approved wins on retry/cascade.'},
     {Item:'PSP approval',Rule:'Individual PSP attempts; merchantOrderId is NOT deduplicated.'},
     {Item:'Country ranking',Rule:'All countries ranked by approved revenue; no Top-50 limit.'},
     {Item:'Country naming',Rule:'Full country name is shown; ISO alpha-2 code is kept in a separate Country Code column.'},
     {Item:'Timezone',Rule:`Source transaction timestamps are interpreted as GMT+0 and converted to ${timezone} before date filtering.`},
     {Item:'Source file',Rule:data.fileName}
   ])
   XLSX.writeFile(wb,`FundedNext_Approval_Analysis_${timezone.replace('+','plus').replace('-','minus')}_${new Date().toISOString().slice(0,10)}.xlsx`)
 }

 if(!data)return <div className="landing"><div className="landing-top"><div className="brand"><div className="brand-mark">FN</div><div><div className="brand-title">FundedNext</div><div className="brand-sub">Approval Ratio Analysis</div></div></div></div><main className="landing-content"><div className="welcome-icon">▤</div><h1>Welcome to Approval Ratio Analytics</h1><p>Upload your transaction data to unlock powerful approval, revenue, country and PSP insights.</p><UploadBox onFile={handleFile}/>{loading&&<div className="loading">Processing your transaction file…</div>}<div className="landing-note"><b>Calculation rules</b><span>Country approval uses unique merchantOrderId.</span><span>PSP approval uses individual attempts — no merchantOrderId deduplication.</span><span>Countries are ranked by approved revenue and the full country list is available.</span><span>Analysis timezone defaults to GMT+0 and can be changed after upload.</span></div></main>{notice&&<div className="notice">{notice}</div>}</div>

 return <div className="app-shell"><header className="topbar"><div className="brand"><div className="brand-mark">FN</div><div><div className="brand-title">FundedNext</div><div className="brand-sub">Approval Ratio Analysis</div></div></div><div className="header-actions"><span className="period-chip">{data.period}</span><button className="download-btn" onClick={downloadFullExcel}>Download Full Excel</button><label className="upload-btn">Upload New Data<input type="file" accept=".csv,.xlsx,.xls" onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f)}}/></label></div></header><div className="layout"><aside className="sidebar">
 {source?.filterable&&<div className="filters-panel"><div className="filters-head"><div><div className="eyebrow light">FILTERS</div><h3>Transaction Filters</h3></div>{activeFilterCount>0&&<button className="clear-btn" onClick={clearFilters}>Clear</button>}</div>
 <DateFilter startDate={startDate} setStartDate={setStartDate} endDate={endDate} setEndDate={setEndDate} timezone={timezone} setTimezone={setTimezone}/>
 <MultiSelectFilter title="PSP" icon="▥" options={options.psps} selected={pspSelected} setSelected={setPspSelected} mode={pspMode} setMode={setPspMode} placeholder="Select PSP"/>
 <MultiSelectFilter title="MID Alias" icon="#" options={options.mids} selected={midSelected} setSelected={setMidSelected} mode={midMode} setMode={setMidMode} placeholder="Select MID Alias"/>
 <MultiSelectFilter title="Country" icon="◉" options={options.countries} selected={countrySelected} setSelected={setCountrySelected} mode={countryMode} setMode={setCountryMode} placeholder="Select Country"/>
 </div>}
 {source?.filterable&&<div className="filter-summary"><span>{activeFilterCount?`${activeFilterCount} filter${activeFilterCount>1?'s':''} active`:'All transaction data'} · {timezone}</span><b>{num(data?.rawObjects?.length||0)} rows loaded</b></div>}
 {['overview','countries','psps','declines','mapping','methodology'].map((id,i)=>{const labels=[['overview','Overview','▦'],['countries','All Countries','◉'],['psps','PSP Performance','◈'],['declines','Hard / Soft Declines','!'],['mapping','PSP Mapping','↔'],['methodology','Methodology','i']][i];return <button key={id} className={`nav-item ${tab===id?'active':''}`} onClick={()=>setTab(id)}><span>{labels[2]}</span>{labels[1]}</button>})}
 <div className="sidebar-note"><b>Calculation rule</b><br/>Country / cascaded KPIs use unique <code>merchantOrderId</code>.<br/><br/>PSP KPIs use individual attempts — <b>no merchantOrderId deduplication</b>.<br/><br/><b>Country ranking</b><br/>Approved revenue, highest to lowest.</div></aside><main className="content">{notice&&<div className="notice">{notice}</div>}
 {tab==='overview'&&<><section className="hero"><div><div className="eyebrow">EXECUTIVE VIEW</div><h1>Approval & Decline Performance</h1><p>Country metrics are cascaded; PSP metrics are attempt-based. Countries are ranked by approved revenue.</p></div><div className="segmented"><button className={approvalMode==='with'?'selected':''} onClick={()=>setApprovalMode('with')}>With Hard Decline</button><button className={approvalMode==='without'?'selected':''} onClick={()=>setApprovalMode('without')}>Without Hard Decline</button></div></section><div className="kpi-grid"><KpiCard label={approvalMode==='with'?'Cascaded Approval Ratio':'Approval Ratio without Hard Decline'} value={pct(currentKpi['Displayed Approval Ratio'])} sub={`${num(currentKpi.Count)} unique merchant orders`} tone="blue"/><KpiCard label="Without India" value={pct(withoutIndia['Displayed Approval Ratio'])} sub="Cascaded country scenario" tone="purple"/><KpiCard label="Revenue" value={money(currentKpi['Revenue (USD)'])} sub="Approved transaction revenue" tone="green"/><KpiCard label="Approved" value={num(currentKpi.Approved)} sub={`${num(currentKpi.Declined)} declined`} tone="amber"/><KpiCard label="Hard Declines" value={num(currentKpi['Hard Declines'])} sub={pct(currentKpi['Hard Decline %'])+' of classified declines'} tone="red"/><KpiCard label="Top Decline Reason" value={currentKpi['Top Declined Reason']} sub="Highest-volume classified reason" tone="slate"/></div><section className="grid-2"><div className="panel"><div className="panel-head"><div><h2>Top Countries by Revenue</h2><p>Top 10 from the full country list</p></div></div>{topCountries.map(c=><RatioBar key={c.code} label={`${c.rank}. ${c.country}`} value={c.amount/(data.countries[0]?.amount||1)} accent="green"/>)} </div><div className="panel"><div className="panel-head"><div><h2>PSP Approval</h2><p>Attempt-based · minimum {pspMin} attempts</p></div></div>{topPsps.map(p=><RatioBar key={p.psp} label={p.psp} value={p.approval} accent={(p.approval??0)>=.75?'green':(p.approval??0)>=.5?'blue':'red'}/>)}</div></section><section className="grid-2"><div className="panel"><div className="panel-head"><div><h2>Decline Mix</h2><p>Terminal declined unique orders</p></div></div><div className="donut-wrap"><div className="donut" style={{'--hard':`${clamp(currentKpi['Hard Decline %'])}%`}}><div><b>{pct(currentKpi['Hard Decline %'])}</b><span>hard</span></div></div><div className="legend"><div><i className="dot hard"></i>Hard declines <b>{num(currentKpi['Hard Declines'])}</b></div><div><i className="dot soft"></i>Soft declines <b>{num(currentKpi['Soft Declines'])}</b></div></div></div></div><div className="panel"><div className="panel-head"><div><h2>Scenario Comparison</h2><p>All countries and without India</p></div></div><RatioBar label="All countries · with hard decline" value={data.kpis[0]['Displayed Approval Ratio']}/><RatioBar label="All countries · without hard decline" value={data.kpis[1]['Displayed Approval Ratio']} accent="green"/><RatioBar label="Without India · with hard decline" value={data.kpis[2]['Displayed Approval Ratio']} accent="purple"/><RatioBar label="Without India · without hard decline" value={data.kpis[3]['Displayed Approval Ratio']} accent="green"/></div></section></>}
 {tab==='countries'&&<section><div className="hero compact"><div><div className="eyebrow">COUNTRY PERFORMANCE</div><h1>All Countries — Revenue Ranked</h1><p>Full country list, ranked by approved revenue. Approval is cascaded using unique merchantOrderId.</p></div></div><div className="toolbar"><div className="segmented"><button className={approvalMode==='with'?'selected':''} onClick={()=>setApprovalMode('with')}>With Hard Decline</button><button className={approvalMode==='without'?'selected':''} onClick={()=>setApprovalMode('without')}>Without Hard Decline</button></div><input className="search" placeholder="Search country, code, PSP or decline reason..." value={countryQuery} onChange={e=>setCountryQuery(e.target.value)}/><select value={countryPsp} onChange={e=>setCountryPsp(e.target.value)}>{pspOptions.map(x=><option key={x}>{x}</option>)}</select></div><div className="panel table-panel"><table><thead><tr><th>Revenue Rank</th><th>Country</th><th>Revenue</th><th>Orders</th><th>Approval</th><th>Without Hard</th><th>Hard %</th><th>Soft %</th><th>Top Decline</th><th>Top PSP</th><th>PSP AR</th><th>PSP Count</th></tr></thead><tbody>{countries.map(c=><tr key={`${c.code}-${c.rank}`}><td><b>{c.rank}</b></td><td><b>{c.country}</b><span className="muted">{c.code}</span></td><td><b>{money(c.amount)}</b></td><td>{num(c.count)}</td><td><span className={`pill ${c.approval>=.8?'good':c.approval>=.6?'mid':'bad'}`}>{pct(approvalMode==='with'?c.approval:c.approvalWithoutHard)}</span></td><td>{pct(c.approvalWithoutHard)}</td><td>{pct(c.hardPct)}</td><td>{pct(c.softPct)}</td><td>{c.topReason}</td><td>{c.topPsp}</td><td>{pct(c.topPspApproval)}</td><td>{num(c.topPspCount)}</td></tr>)}</tbody></table></div></section>}
 {tab==='psps'&&<section><div className="hero compact"><div><div className="eyebrow">PSP PERFORMANCE</div><h1>Attempt-Level PSP Analysis</h1><p>PSP approval ratio does <b>not</b> deduplicate merchantOrderId.</p></div></div><div className="toolbar"><input className="search" placeholder="Search PSP, source mapping or decline reason..." value={pspQuery} onChange={e=>setPspQuery(e.target.value)}/><select value={pspMin} onChange={e=>setPspMin(Number(e.target.value))}><option value="0">All attempts</option><option value="20">20+ attempts</option><option value="100">100+ attempts</option><option value="500">500+ attempts</option></select></div><div className="panel table-panel"><table><thead><tr><th>PSP</th><th>Source Mapping</th><th>Attempts</th><th>Approved</th><th>Declined</th><th>Approval</th><th>Without Hard</th><th>Revenue</th><th>Hard</th><th>Soft</th><th>Unclassified</th><th>Top Reason</th></tr></thead><tbody>{pspRows.map(p=><tr key={p.psp}><td><b>{p.psp}</b></td><td className="muted">{p.source}</td><td>{num(p.attempts)}</td><td>{num(p.approved)}</td><td>{num(p.declined)}</td><td><span className={`pill ${(p.approval??0)>=.75?'good':(p.approval??0)>=.5?'mid':'bad'}`}>{pct(p.approval)}</span></td><td>{pct(p.approvalWithoutHard)}</td><td>{money(p.revenue)}</td><td>{num(p.hard)}</td><td>{num(p.soft)}</td><td>{num(p.unclassified)}</td><td>{p.topReason}</td></tr>)}</tbody></table></div></section>}
 {tab==='declines'&&<section><div className="hero compact"><div><div className="eyebrow">DECLINE TAXONOMY</div><h1>Hard & Soft Declines</h1><p>Standardized list derived from the uploaded transaction decline reasons.</p></div></div><div className="toolbar"><div className="segmented"><button className={declineType==='All'?'selected':''} onClick={()=>setDeclineType('All')}>All</button><button className={declineType==='Hard'?'selected':''} onClick={()=>setDeclineType('Hard')}>Hard</button><button className={declineType==='Soft'?'selected':''} onClick={()=>setDeclineType('Soft')}>Soft</button></div><input className="search" placeholder="Search decline reason..." value={declineQuery} onChange={e=>setDeclineQuery(e.target.value)}/></div><div className="decline-grid">{['Hard','Soft'].filter(t=>declineType==='All'||declineType===t).map(type=>{const list=declineRows.filter(x=>x.type===type);const total=list.reduce((s,x)=>s+x.count,0);return <div className="panel" key={type}><div className="panel-head"><div><h2>{type} Declines</h2><p>{list.length} reasons · {num(total)} mapped declined attempts</p></div></div>{list.map((x,i)=><div className="decline-row" key={`${type}-${x.reason}-${i}`}><div><b>{x.reason}</b><span>{x.type}</span></div><strong>{num(x.count)}</strong></div>)}</div>})}</div></section>}
 {tab==='mapping'&&<section><div className="hero compact"><div><div className="eyebrow">DATA NORMALIZATION</div><h1>PSP Mapping Rules</h1><p>Labels used in the dashboard and Excel export.</p></div></div><div className="panel table-panel"><table><thead><tr><th>Source PSP</th><th>Dashboard PSP</th><th>Country Scope</th><th>Rule</th></tr></thead><tbody>{data.mappingRules.map((m,i)=><tr key={i}><td><b>{m.source}</b></td><td>{m.dashboard}</td><td>{m.scope}</td><td>{m.rule}</td></tr>)}</tbody></table></div></section>}
 {tab==='methodology'&&<section><div className="hero compact"><div><div className="eyebrow">DEFINITIONS</div><h1>Methodology</h1><p>Country/cascaded and PSP calculations intentionally use different transaction keys.</p></div></div><div className="method-grid">{data.methodology.map((m,i)=><div className="panel method-card" key={i}><div className="eyebrow">{m.metric}</div><h3>{m.definition}</h3><p>{m.note}</p></div>)}</div></section>}
 </main></div></div>
}

export default App
