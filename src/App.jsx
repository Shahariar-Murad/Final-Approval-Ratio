import React, { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { DEFAULT_DATA } from './data'

const pct = (v) => v == null || Number.isNaN(Number(v)) ? '—' : `${(Number(v) * 100).toFixed(2)}%`
const money = (v) => v == null || Number.isNaN(Number(v)) ? '—' : `$${Number(v).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`
const num = (v) => Number(v || 0).toLocaleString()
const clamp = (v) => Math.max(0, Math.min(100, Number(v || 0) * 100))

function headerRow(rows, label) {
  return rows.findIndex(r => String(r?.[0] ?? '').trim() === label)
}
function parseSheet(wb, name) {
  const ws = wb.Sheets[name]
  if (!ws) return []
  return XLSX.utils.sheet_to_json(ws, {header: 1, defval: null})
}
function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, {type: 'array'})
        const rows = parseSheet(wb, 'Executive Summary')
        const pspRows = parseSheet(wb, 'PSP Wise')
        const countryRows = parseSheet(wb, 'Top 50 Countries')
        const mapRows = parseSheet(wb, 'Mapping Rules')
        const methRows = parseSheet(wb, 'Methodology')

        const pspH = headerRow(pspRows, 'PSP')
        const countryH = headerRow(countryRows, 'Rank')
        const mapH = headerRow(mapRows, 'Source PSP')
        const methH = headerRow(methRows, 'Metric')

        const psp = pspH >= 0 ? pspRows.slice(pspH + 1).filter(r => r?.[0]).map(r => {
          const attempts = Number(r[2] || 0), approved = Number(r[3] || 0), soft = Number(r[9] || 0), unclass = Number(r[10] || 0)
          return {
            psp: String(r[0]), source: String(r[1] ?? ''), attempts, approved,
            declined: Number(r[4] || 0), pendingOther: Number(r[5] || 0),
            approval: Number.isFinite(Number(r[6])) ? Number(r[6]) : null,
            approvalWithoutHard: (approved + soft + unclass) ? approved / (approved + soft + unclass) : null,
            revenue: Number(r[7] || 0), hard: Number(r[8] || 0), soft,
            unclassified: unclass, hardPct: Number(r[11]), softPct: Number(r[12]),
            topReason: String(r[13] ?? 'Unclassified'), topCount: Number(r[14] || 0)
          }
        }) : []

        const countries = countryH >= 0 ? countryRows.slice(countryH + 1).filter(r => r?.[0]).map(r => ({
          rank: Number(r[0]), country: String(r[1]), code: String(r[2]),
          amount: Number(String(r[3] || '').split(' / ')[0].replace(/[^0-9.-]/g, '')) || 0,
          count: Number(r[4] || 0), approval: Number(r[5] || 0),
          approvalWithoutHard: Number(r[6] || 0), hardPct: Number(r[7] || 0),
          softPct: Number(r[8] || 0), topReason: String(r[9] ?? ''),
          topPsp: String(r[10] ?? ''), topPspApproval: Number(r[11] || 0), topPspCount: Number(r[12] || 0)
        })) : []

        const mappingRules = mapH >= 0 ? mapRows.slice(mapH + 1).filter(r => r?.[0]).map(r => ({
          source: String(r[0]), dashboard: String(r[1]), scope: String(r[2]), rule: String(r[3])
        })) : []

        const methodology = methH >= 0 ? methRows.slice(methH + 1).filter(r => r?.[0]).map(r => ({
          metric: String(r[0]), definition: String(r[1]), note: String(r[2])
        })) : []

        const metricHeader = rows.findIndex(r => String(r?.[0] ?? '') === 'Metric')
        const metrics = metricHeader >= 0 ? rows.slice(metricHeader + 1).filter(r => r?.[0]).slice(0, 4) : []
        const kpis = metrics.map(r => ({
          Metric: String(r[0]), Count: Number(r[1] || 0), Approved: Number(r[2] || 0),
          Declined: Number(r[3] || 0), 'Hard Declines': Number(r[4] || 0),
          'Soft Declines': Number(r[5] || 0), 'Displayed Approval Ratio': Number(r[6] || 0),
          'Approval Ratio without Hard Decline': Number(r[7] || 0),
          'Hard Decline %': Number(r[8] || 0), 'Soft Decline %': Number(r[9] || 0),
          'Top Declined Reason': String(r[10] ?? ''), 'Revenue (USD)': Number(r[11] || 0)
        }))

        resolve({...DEFAULT_DATA, kpis: kpis.length ? kpis : DEFAULT_DATA.kpis, pspWise: psp.length ? psp : DEFAULT_DATA.pspWise,
          countries: countries.length ? countries : DEFAULT_DATA.countries,
          mappingRules: mappingRules.length ? mappingRules : DEFAULT_DATA.mappingRules,
          methodology: methodology.length ? methodology : DEFAULT_DATA.methodology})
      } catch (err) { reject(err) }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

function RatioBar({value, label, accent='blue'}) {
  return <div className="ratio-row">
    <div className="ratio-label"><span>{label}</span><strong>{pct(value)}</strong></div>
    <div className="bar"><span className={`bar-fill ${accent}`} style={{width:`${clamp(value)}%`}} /></div>
  </div>
}

function KpiCard({label, value, sub, tone='blue'}) {
  return <div className={`kpi-card ${tone}`}>
    <div className="kpi-label">{label}</div>
    <div className="kpi-value">{value}</div>
    {sub && <div className="kpi-sub">{sub}</div>}
  </div>
}

function App() {
  const [data, setData] = useState(DEFAULT_DATA)
  const [tab, setTab] = useState('overview')
  const [approvalMode, setApprovalMode] = useState('with')
  const [countryQuery, setCountryQuery] = useState('')
  const [countryPsp, setCountryPsp] = useState('All')
  const [pspQuery, setPspQuery] = useState('')
  const [pspMin, setPspMin] = useState(20)
  const [declineType, setDeclineType] = useState('All')
  const [declineQuery, setDeclineQuery] = useState('')
  const [notice, setNotice] = useState('')

  const currentKpi = data.kpis[approvalMode === 'with' ? 0 : 1] || data.kpis[0]
  const withoutIndia = data.kpis[approvalMode === 'with' ? 2 : 3] || data.kpis[2]

  const countries = useMemo(() => {
    let arr = [...data.countries]
    if (countryQuery) {
      const q = countryQuery.toLowerCase()
      arr = arr.filter(x => `${x.country} ${x.code} ${x.topReason} ${x.topPsp}`.toLowerCase().includes(q))
    }
    if (countryPsp !== 'All') arr = arr.filter(x => x.topPsp === countryPsp)
    return arr
  }, [data.countries, countryQuery, countryPsp])

  const pspRows = useMemo(() => {
    let arr = data.pspWise.filter(x => x.attempts >= pspMin)
    if (pspQuery) {
      const q = pspQuery.toLowerCase()
      arr = arr.filter(x => `${x.psp} ${x.source} ${x.topReason}`.toLowerCase().includes(q))
    }
    return [...arr].sort((a,b) => (b.approval ?? -1) - (a.approval ?? -1))
  }, [data.pspWise, pspQuery, pspMin])

  const declineRows = useMemo(() => {
    let arr = [...data.taxonomy]
    if (declineType !== 'All') arr = arr.filter(x => x.type === declineType)
    if (declineQuery) arr = arr.filter(x => x.reason.toLowerCase().includes(declineQuery.toLowerCase()))
    return arr.sort((a,b) => b.count - a.count)
  }, [data.taxonomy, declineType, declineQuery])

  const topCountries = useMemo(() => {
    return [...data.countries].sort((a,b) => {
      const av = approvalMode === 'with' ? a.approval : a.approvalWithoutHard
      const bv = approvalMode === 'with' ? b.approval : b.approvalWithoutHard
      return bv - av
    }).slice(0, 10)
  }, [data.countries, approvalMode])

  const topPsps = useMemo(() => [...data.pspWise].filter(x => x.attempts >= pspMin).sort((a,b)=>(b.approval??-1)-(a.approval??-1)).slice(0,10), [data.pspWise, pspMin])

  const exportCsv = (rows, filename) => {
    if (!rows.length) return
    const ws = XLSX.utils.json_to_sheet(rows)
    const csv = XLSX.utils.sheet_to_csv(ws)
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'})
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a'); a.href=url; a.download=filename; a.click(); URL.revokeObjectURL(url)
  }

  const downloadFullExcel = () => {
    const wb = XLSX.utils.book_new()

    const kpiRows = data.kpis.map(x => ({
      Metric: x.Metric,
      Count: x.Count,
      Approved: x.Approved,
      Declined: x.Declined,
      'Hard Declines': x['Hard Declines'],
      'Soft Declines': x['Soft Declines'],
      'Displayed Approval Ratio': x['Displayed Approval Ratio'],
      'Approval Ratio without Hard Decline': x['Approval Ratio without Hard Decline'],
      'Hard Decline %': x['Hard Decline %'],
      'Soft Decline %': x['Soft Decline %'],
      'Top Declined Reason': x['Top Declined Reason'],
      'Revenue (USD)': x['Revenue (USD)']
    }))

    const pspRows = data.pspWise.map(x => ({
      PSP: x.psp,
      'Source Mapping': x.source,
      Attempts: x.attempts,
      Approved: x.approved,
      Declined: x.declined,
      'Pending / Other': x.pendingOther,
      'PSP Approval Ratio': x.approval,
      'Approval Ratio without Hard Decline': x.approvalWithoutHard,
      'Revenue (USD)': x.revenue,
      'Hard Declines': x.hard,
      'Soft Declines': x.soft,
      Unclassified: x.unclassified,
      'Hard %': x.hardPct,
      'Soft %': x.softPct,
      'Top Declined Reason': x.topReason,
      'Top Declined Count': x.topCount
    }))

    const countryRows = data.countries.map(x => ({
      Rank: x.rank,
      Country: x.country,
      'Country Code': x.code,
      'Amount / Duration (USD)': x.amount,
      Count: x.count,
      'Approval Ratio': x.approval,
      'Approval Ratio without Hard Decline': x.approvalWithoutHard,
      'Hard Decline %': x.hardPct,
      'Soft Decline %': x.softPct,
      'Top Declined Reason': x.topReason,
      'Top Performing PSP': x.topPsp,
      'Top PSP Approval': x.topPspApproval,
      'Top PSP Count': x.topPspCount
    }))

    const declineRows = data.taxonomy.map(x => ({
      'Decline Reason': x.reason,
      Type: x.type,
      Count: x.count
    }))

    const mappingRows = data.mappingRules.map(x => ({
      'Source PSP': x.source,
      'Dashboard PSP': x.dashboard,
      'Country Scope': x.scope,
      Rule: x.rule
    }))

    const methodRows = data.methodology.map(x => ({
      Metric: x.metric,
      Definition: x.definition,
      'Rule / Note': x.note
    }))

    const readmeRows = [
      ['Report', 'FundedNext Approval Ratio Analysis'],
      ['Period', data.period],
      ['Timezone', data.timezone],
      ['Country / Cascaded Approval Ratio', 'Unique merchantOrderId; approved wins on retry/cascade'],
      ['PSP Approval Ratio', 'Individual PSP attempts; merchantOrderId is NOT deduplicated'],
      ['PSP Approval Formula', 'Approved attempts / (Approved + Declined attempts)'],
      ['Top Performing PSP', 'Attempt-based PSP approval ratio; dashboard applies the selected minimum-attempt filter'],
      ['Hard / Soft Decline List', 'Uses the supplied standardized taxonomy; unmatched reasons are not silently reclassified'],
      ['PSP Mapping', 'CardPay → Unlimit; Innatech → Payabl; Global / India mappings follow the Mapping Rules sheet']
    ]

    const add = (name, rows) => {
      const ws = XLSX.utils.aoa_to_sheet(rows)
      ws['!freeze'] = {xSplit: 0, ySplit: 1}
      ws['!cols'] = [{wch: 34}, {wch: 24}, {wch: 24}, {wch: 24}, {wch: 20}, {wch: 20}, {wch: 20}, {wch: 20}]
      XLSX.utils.book_append_sheet(wb, ws, name)
    }

    add('Executive Summary', [
      Object.keys(kpiRows[0] || {Metric:''}),
      ...kpiRows.map(r => Object.values(r))
    ])
    add('Top 50 Countries', [
      Object.keys(countryRows[0] || {Rank:''}),
      ...countryRows.map(r => Object.values(r))
    ])
    add('PSP Wise', [
      Object.keys(pspRows[0] || {PSP:''}),
      ...pspRows.map(r => Object.values(r))
    ])
    add('Hard Soft Declines', [
      Object.keys(declineRows[0] || {'Decline Reason':''}),
      ...declineRows.map(r => Object.values(r))
    ])
    add('Mapping Rules', [
      Object.keys(mappingRows[0] || {'Source PSP':''}),
      ...mappingRows.map(r => Object.values(r))
    ])
    add('Methodology', [
      Object.keys(methodRows[0] || {Metric:''}),
      ...methodRows.map(r => Object.values(r))
    ])
    add('Report Notes', readmeRows)

    const filename = `FundedNext_Approval_Analysis_${data.period.replace(/[^0-9A-Za-z]+/g,'_')}.xlsx`
    XLSX.writeFile(wb, filename)
  }

  const upload = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return
    try {
      const next = await parseWorkbook(file)
      setData(next)
      setNotice(`Loaded ${file.name}`)
      setTimeout(()=>setNotice(''), 3500)
    } catch (err) {
      setNotice('Could not read this workbook. Please use the same analysis workbook structure.')
      setTimeout(()=>setNotice(''), 5000)
    }
  }

  const pspOptions = ['All', ...Array.from(new Set(data.countries.map(x=>x.topPsp))).sort()]

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark">FN</div>
          <div><div className="brand-title">FundedNext</div><div className="brand-sub">Approval Ratio Analysis</div></div>
        </div>
        <div className="header-actions">
          <span className="period-chip">{data.period} · {data.timezone}</span>
          <button className="download-btn" onClick={downloadFullExcel}>Download Full Excel</button>
          <label className="upload-btn">Upload Analysis XLSX<input type="file" accept=".xlsx,.xls" onChange={upload}/></label>
        </div>
      </header>

      <div className="layout">
        <aside className="sidebar">
          {[
            ['overview','Overview','▦'],['countries','Top 50 Countries','◉'],['psps','PSP Performance','◈'],
            ['declines','Hard / Soft Declines','!'],['mapping','PSP Mapping','↔'],['methodology','Methodology','i']
          ].map(([id,label,icon])=><button key={id} className={`nav-item ${tab===id?'active':''}`} onClick={()=>setTab(id)}><span>{icon}</span>{label}</button>)}
          <div className="sidebar-note"><b>Calculation rule</b><br/>Country / cascaded KPIs use unique <code>merchantOrderId</code>.<br/><br/>PSP KPIs use individual attempts — <b>no merchantOrderId deduplication</b>.</div>
        </aside>

        <main className="content">
          {notice && <div className="notice">{notice}</div>}

          {tab==='overview' && <>
            <section className="hero">
              <div><div className="eyebrow">EXECUTIVE VIEW</div><h1>Approval & Decline Performance</h1><p>Country-level metrics are cascaded; PSP metrics are attempt-based. Hard/soft taxonomy follows the supplied analysis.</p></div>
              <div className="segmented"><button className={approvalMode==='with'?'selected':''} onClick={()=>setApprovalMode('with')}>With Hard Decline</button><button className={approvalMode==='without'?'selected':''} onClick={()=>setApprovalMode('without')}>Without Hard Decline</button></div>
            </section>
            <div className="kpi-grid">
              <KpiCard label={approvalMode==='with'?'Cascaded Approval Ratio':'Approval Ratio without Hard Decline'} value={pct(currentKpi['Displayed Approval Ratio'])} sub={`${num(currentKpi.Count)} unique merchant orders`} tone="blue"/>
              <KpiCard label="Without India" value={pct(withoutIndia['Displayed Approval Ratio'])} sub="Cascaded country scenario" tone="purple"/>
              <KpiCard label="Revenue" value={money(currentKpi['Revenue (USD)'])} sub="Approved attempts / cascaded scenario" tone="green"/>
              <KpiCard label="Approved" value={num(currentKpi.Approved)} sub={`${num(currentKpi.Declined)} declined`} tone="amber"/>
              <KpiCard label="Hard Declines" value={num(currentKpi['Hard Declines'])} sub={pct(currentKpi['Hard Decline %']) + ' of classified declines'} tone="red"/>
              <KpiCard label="Top Decline Reason" value={currentKpi['Top Declined Reason']} sub="Highest-volume classified reason" tone="slate"/>
            </div>

            <section className="grid-2">
              <div className="panel">
                <div className="panel-head"><div><h2>Top Countries by Approval</h2><p>Top 10 within the supplied Top 50 list</p></div></div>
                {topCountries.map(c=><RatioBar key={c.code} label={`${c.rank}. ${c.country}`} value={approvalMode==='with'?c.approval:c.approvalWithoutHard}/>)}
              </div>
              <div className="panel">
                <div className="panel-head"><div><h2>PSP Approval</h2><p>Attempt-based · minimum {pspMin} attempts</p></div></div>
                {topPsps.map(p=><RatioBar key={p.psp} label={p.psp} value={p.approval} accent={p.approval>=.75?'green':p.approval>=.5?'blue':'red'}/>)}
              </div>
            </section>

            <section className="grid-2">
              <div className="panel">
                <div className="panel-head"><div><h2>Decline Mix</h2><p>From the cascaded executive scenario</p></div></div>
                <div className="donut-wrap">
                  <div className="donut" style={{'--hard':`${clamp(currentKpi['Hard Decline %'])}%`}}><div><b>{pct(currentKpi['Hard Decline %'])}</b><span>hard</span></div></div>
                  <div className="legend"><div><i className="dot hard"></i>Hard declines <b>{num(currentKpi['Hard Declines'])}</b></div><div><i className="dot soft"></i>Soft declines <b>{num(currentKpi['Soft Declines'])}</b></div></div>
                </div>
              </div>
              <div className="panel">
                <div className="panel-head"><div><h2>Scenario Comparison</h2><p>Same data, different hard-decline treatment</p></div></div>
                <RatioBar label="All countries · with hard decline" value={data.kpis[0]['Displayed Approval Ratio']}/>
                <RatioBar label="All countries · without hard decline" value={data.kpis[1]['Displayed Approval Ratio']} accent="green"/>
                <RatioBar label="Without India · with hard decline" value={data.kpis[2]['Displayed Approval Ratio']} accent="purple"/>
                <RatioBar label="Without India · without hard decline" value={data.kpis[3]['Displayed Approval Ratio']} accent="green"/>
              </div>
            </section>
          </>}

          {tab==='countries' && <section>
            <div className="hero compact"><div><div className="eyebrow">COUNTRY PERFORMANCE</div><h1>Top 50 Countries</h1><p>Approval is cascaded using unique merchantOrderId. Top PSP is attempt-based.</p></div><button className="secondary-btn" onClick={()=>exportCsv(countries,'fundednext-top50-countries.csv')}>Export CSV</button></div>
            <div className="toolbar">
              <div className="segmented"><button className={approvalMode==='with'?'selected':''} onClick={()=>setApprovalMode('with')}>With Hard Decline</button><button className={approvalMode==='without'?'selected':''} onClick={()=>setApprovalMode('without')}>Without Hard Decline</button></div>
              <input className="search" placeholder="Search country, code, PSP or decline reason..." value={countryQuery} onChange={e=>setCountryQuery(e.target.value)}/>
              <select value={countryPsp} onChange={e=>setCountryPsp(e.target.value)}>{pspOptions.map(x=><option key={x}>{x}</option>)}</select>
            </div>
            <div className="panel table-panel"><table><thead><tr><th>Rank</th><th>Country</th><th>Volume</th><th>Orders</th><th>Approval</th><th>Without Hard</th><th>Hard %</th><th>Soft %</th><th>Top Decline</th><th>Top PSP</th><th>PSP AR</th><th>PSP Count</th></tr></thead><tbody>
              {countries.map(c=><tr key={c.code}><td>{c.rank}</td><td><b>{c.country}</b><span className="muted">{c.code}</span></td><td>{money(c.amount)}</td><td>{num(c.count)}</td><td><span className={`pill ${c.approval>=.8?'good':c.approval>=.6?'mid':'bad'}`}>{pct(approvalMode==='with'?c.approval:c.approvalWithoutHard)}</span></td><td>{pct(c.approvalWithoutHard)}</td><td>{pct(c.hardPct)}</td><td>{pct(c.softPct)}</td><td>{c.topReason}</td><td>{c.topPsp}</td><td>{pct(c.topPspApproval)}</td><td>{num(c.topPspCount)}</td></tr>)}
            </tbody></table></div>
          </section>}

          {tab==='psps' && <section>
            <div className="hero compact"><div><div className="eyebrow">PSP PERFORMANCE</div><h1>Attempt-Level PSP Analysis</h1><p>PSP approval ratio does <b>not</b> deduplicate merchantOrderId.</p></div><button className="secondary-btn" onClick={()=>exportCsv(pspRows,'fundednext-psp-performance.csv')}>Export CSV</button></div>
            <div className="toolbar"><input className="search" placeholder="Search PSP, source mapping or decline reason..." value={pspQuery} onChange={e=>setPspQuery(e.target.value)}/><select value={pspMin} onChange={e=>setPspMin(Number(e.target.value))}><option value="0">All attempts</option><option value="20">20+ attempts</option><option value="100">100+ attempts</option><option value="500">500+ attempts</option></select></div>
            <div className="panel table-panel"><table><thead><tr><th>PSP</th><th>Source Mapping</th><th>Attempts</th><th>Approved</th><th>Declined</th><th>Approval</th><th>Without Hard</th><th>Revenue</th><th>Hard</th><th>Soft</th><th>Unclassified</th><th>Top Reason</th></tr></thead><tbody>
              {pspRows.map(p=><tr key={p.psp}><td><b>{p.psp}</b></td><td className="muted">{p.source}</td><td>{num(p.attempts)}</td><td>{num(p.approved)}</td><td>{num(p.declined)}</td><td><span className={`pill ${(p.approval??0)>=.75?'good':(p.approval??0)>=.5?'mid':'bad'}`}>{pct(p.approval)}</span></td><td>{pct(p.approvalWithoutHard)}</td><td>{money(p.revenue)}</td><td>{num(p.hard)}</td><td>{num(p.soft)}</td><td>{num(p.unclassified)}</td><td>{p.topReason}</td></tr>)}
            </tbody></table></div>
          </section>}

          {tab==='declines' && <section>
            <div className="hero compact"><div><div className="eyebrow">DECLINE TAXONOMY</div><h1>Hard & Soft Declines</h1><p>Exact standardized taxonomy supplied with the analysis. Unmatched reasons are not silently reclassified.</p></div></div>
            <div className="toolbar"><div className="segmented"><button className={declineType==='All'?'selected':''} onClick={()=>setDeclineType('All')}>All</button><button className={declineType==='Hard'?'selected':''} onClick={()=>setDeclineType('Hard')}>Hard</button><button className={declineType==='Soft'?'selected':''} onClick={()=>setDeclineType('Soft')}>Soft</button></div><input className="search" placeholder="Search decline reason..." value={declineQuery} onChange={e=>setDeclineQuery(e.target.value)}/></div>
            <div className="decline-grid">{['Hard','Soft'].filter(t=>declineType==='All'||declineType===t).map(type=>{
              const list=declineRows.filter(x=>x.type===type)
              const total=list.reduce((s,x)=>s+x.count,0)
              return <div className="panel" key={type}><div className="panel-head"><div><h2>{type} Declines</h2><p>{list.length} taxonomy entries · {num(total)} mapped occurrences</p></div></div>{list.map((x,i)=><div className="decline-row" key={`${type}-${x.reason}-${i}`}><div><b>{x.reason}</b><span>{x.type}</span></div><strong>{num(x.count)}</strong></div>)}</div>
            })}</div>
          </section>}

          {tab==='mapping' && <section>
            <div className="hero compact"><div><div className="eyebrow">DATA NORMALIZATION</div><h1>PSP Mapping Rules</h1><p>Dashboard labels used in the analysis.</p></div></div>
            <div className="panel table-panel"><table><thead><tr><th>Source PSP</th><th>Dashboard PSP</th><th>Country Scope</th><th>Rule</th></tr></thead><tbody>{data.mappingRules.map((m,i)=><tr key={i}><td><b>{m.source}</b></td><td>{m.dashboard}</td><td>{m.scope}</td><td>{m.rule}</td></tr>)}</tbody></table></div>
          </section>}

          {tab==='methodology' && <section>
            <div className="hero compact"><div><div className="eyebrow">DEFINITIONS</div><h1>Methodology</h1><p>Keep these two approval-ratio logics separate.</p></div></div>
            <div className="method-grid">{data.methodology.map((m,i)=><div className="panel method-card" key={i}><div className="eyebrow">{m.metric}</div><h3>{m.definition}</h3><p>{m.note}</p></div>)}</div>
          </section>}
        </main>
      </div>
    </div>
  )
}

export default App
