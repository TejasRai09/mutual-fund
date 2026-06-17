import { useState, useRef, useCallback } from 'react'
import {
  FileText, Table2, Upload, X, Lock, Settings2, ChevronRight,
  CheckCircle2, AlertTriangle, XCircle, Download, Sparkles,
  TrendingUp, FileCheck2, AlertCircle, ArrowLeft
} from 'lucide-react'

function fmtBytes(b) {
  if (b < 1024) return `${b}B`
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(0)}KB`
  return `${(b / 1024 / 1024).toFixed(1)}MB`
}

function FileChip({ file, onRemove, variant }) {
  const isSource = variant === 'source'
  return (
    <div
      className={`flex items-center gap-2 rounded-lg px-3 py-2 border group transition-all
        ${isSource ? 'bg-violet-50 border-violet-100 text-violet-800' : 'bg-emerald-50 border-emerald-100 text-emerald-800'}`}
      onClick={e => e.stopPropagation()}
    >
      <FileText className="w-3.5 h-3.5 shrink-0 opacity-70" />
      <span className="truncate max-w-[180px] font-medium text-xs">{file.name}</span>
      <span className="text-[10px] opacity-50 shrink-0">{fmtBytes(file.size)}</span>
      <button onClick={onRemove} className="ml-auto opacity-0 group-hover:opacity-100 transition-opacity hover:text-red-500">
        <X className="w-3 h-3" />
      </button>
    </div>
  )
}

function DropZone({ label, sublabel, icon: Icon, accept, files, onFiles, variant }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef(null)
  const isSource = variant === 'source'

  const addFiles = useCallback(newFiles => {
    const filtered = newFiles.filter(f => accept.includes(f.name.split('.').pop().toLowerCase()))
    if (filtered.length) onFiles(prev => [...prev, ...filtered])
  }, [accept, onFiles])

  const handleDrop = useCallback(e => {
    e.preventDefault(); setDragging(false); addFiles([...e.dataTransfer.files])
  }, [addFiles])

  const borderClass = dragging
    ? (isSource ? 'border-violet-400 bg-violet-50' : 'border-emerald-400 bg-emerald-50')
    : 'border-slate-200 bg-slate-50/50 hover:border-slate-300 hover:bg-white'

  return (
    <div className="flex-1 min-w-0">
      <div className="flex items-center justify-between mb-2">
        <div>
          <h3 className="text-sm font-semibold text-slate-800">{label}</h3>
          <p className="text-xs text-slate-400 mt-0.5">{sublabel}</p>
        </div>
        {files.length > 0 && (
          <button onClick={() => onFiles([])} className="text-xs text-slate-400 hover:text-red-500 transition-colors">
            Clear all
          </button>
        )}
      </div>
      <div
        onDragOver={e => { e.preventDefault(); setDragging(true) }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current?.click()}
        className={`relative border-2 border-dashed rounded-2xl cursor-pointer transition-all duration-200 min-h-[200px] ${borderClass} ${dragging ? 'scale-[1.01]' : ''}`}
      >
        <input ref={inputRef} type="file" multiple accept={accept.map(e => `.${e}`).join(',')} className="hidden"
          onChange={e => { addFiles([...e.target.files]); e.target.value = '' }} />
        {files.length === 0 ? (
          <div className="flex flex-col items-center justify-center gap-3 p-8 text-center h-full min-h-[200px]">
            <div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${isSource ? 'bg-violet-100 text-violet-600' : 'bg-emerald-100 text-emerald-600'}`}>
              <Icon className="w-7 h-7" />
            </div>
            <div>
              <p className="text-sm font-semibold text-slate-600">Drop {accept[0].toUpperCase()} files here</p>
              <p className="text-xs text-slate-400 mt-1">or click to browse</p>
            </div>
            <span className={`text-[10px] font-semibold tracking-wider uppercase px-3 py-1 rounded-full ${isSource ? 'bg-violet-100 text-violet-600' : 'bg-emerald-100 text-emerald-600'}`}>
              .{accept.join(' / .')}
            </span>
          </div>
        ) : (
          <div className="p-3 space-y-1.5" onClick={e => e.stopPropagation()}>
            {files.map((f, i) => (
              <FileChip key={`${f.name}-${i}`} file={f} variant={variant}
                onRemove={() => onFiles(prev => prev.filter((_, j) => j !== i))} />
            ))}
            <button onClick={e => { e.stopPropagation(); inputRef.current?.click() }}
              className={`w-full text-xs font-medium py-2 rounded-lg border-2 border-dashed transition-colors mt-1
                ${isSource ? 'border-violet-200 text-violet-500 hover:border-violet-400 hover:bg-violet-50'
                : 'border-emerald-200 text-emerald-500 hover:border-emerald-400 hover:bg-emerald-50'}`}>
              + Add more {accept[0].toUpperCase()} files
            </button>
          </div>
        )}
      </div>
    </div>
  )
}

function StatusIcon({ icon }) {
  if (icon === 'ok')   return <CheckCircle2 className="w-4 h-4 text-emerald-500 shrink-0" />
  if (icon === 'warn') return <AlertTriangle className="w-4 h-4 text-amber-500 shrink-0" />
  if (icon === 'err')  return <XCircle className="w-4 h-4 text-red-500 shrink-0" />
  return <AlertCircle className="w-4 h-4 text-slate-400 shrink-0" />
}

function LogTable({ title, icon: Icon, rows, columns }) {
  const [expanded, setExpanded] = useState(true)
  if (!rows?.length) return null
  return (
    <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden">
      <button onClick={() => setExpanded(e => !e)}
        className="w-full flex items-center justify-between px-5 py-4 hover:bg-slate-50 transition-colors">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-xl bg-slate-100 flex items-center justify-center">
            <Icon className="w-4 h-4 text-slate-600" />
          </div>
          <span className="font-semibold text-slate-800 text-sm">{title}</span>
          <span className="text-xs font-medium px-2 py-0.5 bg-slate-100 text-slate-500 rounded-full">{rows.length}</span>
        </div>
        <ChevronRight className={`w-4 h-4 text-slate-400 transition-transform duration-200 ${expanded ? 'rotate-90' : ''}`} />
      </button>
      {expanded && (
        <div className="overflow-x-auto border-t border-slate-100">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-50">
                <th className="text-left px-5 py-2.5 text-xs font-semibold text-slate-500 w-8"></th>
                {columns.map(c => <th key={c.key} className="text-left px-5 py-2.5 text-xs font-semibold text-slate-500">{c.label}</th>)}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {rows.map((row, i) => (
                <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                  <td className="pl-5 py-3"><StatusIcon icon={row.icon} /></td>
                  {columns.map(c => (
                    <td key={c.key} className="px-5 py-3 text-slate-600">
                      {c.key === 'file' ? (
                        <span className="font-medium text-slate-800 truncate max-w-[240px] block">{row[c.key]}</span>
                      ) : c.key === 'filled' ? (
                        <span className={`font-mono text-xs px-2 py-0.5 rounded-md font-semibold
                          ${row.status === 'ok' ? 'bg-emerald-50 text-emerald-700' :
                            row.status === 'warn' ? 'bg-amber-50 text-amber-700' : 'bg-red-50 text-red-700'}`}>
                          {row[c.key]}
                        </span>
                      ) : <span className="text-xs">{row[c.key]}</span>}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

function StatCard({ label, value, icon: Icon, colorClass }) {
  return (
    <div className="flex items-center gap-3 bg-white rounded-2xl border border-slate-200 px-5 py-4 flex-1 min-w-0">
      <div className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${colorClass}`}>
        <Icon className="w-5 h-5" />
      </div>
      <div className="min-w-0">
        <p className="text-2xl font-bold text-slate-800 leading-tight">{value}</p>
        <p className="text-xs text-slate-500 font-medium mt-0.5 truncate">{label}</p>
      </div>
    </div>
  )
}

export default function Tool({ onBack }) {
  const [sourcePdfs, setSourcePdfs] = useState([])
  const [targetXlsx, setTargetXlsx] = useState([])
  const [password, setPassword]     = useState('')
  const [showPwd, setShowPwd]       = useState(false)
  const [format, setFormat]         = useState('auto')
  const [loading, setLoading]       = useState(false)
  const [result, setResult]         = useState(null)
  const [error, setError]           = useState('')

  const handleProcess = async () => {
    if (!sourcePdfs.length && !targetXlsx.length) {
      setError('Please upload at least one source PDF and one target Excel file.')
      return
    }
    setError(''); setLoading(true); setResult(null)
    try {
      const fd = new FormData()
      sourcePdfs.forEach(f => fd.append('source_files', f))
      targetXlsx.forEach(f => fd.append('target_files', f))
      fd.append('pdf_password', password)
      fd.append('vendor_fmt', format)
      const res = await fetch('/api/process', { method: 'POST', body: fd })
      if (!res.ok) throw new Error(`Server error ${res.status}: ${await res.text()}`)
      setResult(await res.json())
    } catch (e) {
      setError(e.message || 'Unknown error')
    } finally {
      setLoading(false)
    }
  }

  const canProcess = (sourcePdfs.length > 0 || targetXlsx.length > 0) && !loading

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-indigo-50/30">
      <div className="fixed inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-32 -right-32 w-[500px] h-[500px] bg-indigo-100/40 rounded-full blur-3xl" />
        <div className="absolute -bottom-32 -left-32 w-[500px] h-[500px] bg-violet-100/30 rounded-full blur-3xl" />
      </div>

      <div className="relative max-w-5xl mx-auto px-4 py-8">
        {/* Topbar */}
        <div className="flex items-center justify-between mb-8">
          <button onClick={onBack}
            className="flex items-center gap-1.5 text-sm text-slate-500 hover:text-slate-800 transition-colors group">
            <ArrowLeft className="w-4 h-4 group-hover:-translate-x-0.5 transition-transform" />
            Back to home
          </button>
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-xl bg-gradient-to-br from-indigo-500 to-violet-600 flex items-center justify-center">
              <Sparkles className="w-3.5 h-3.5 text-white" />
            </div>
            <span className="font-bold text-slate-800 text-sm">BrokerageAI</span>
          </div>
        </div>

        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-slate-900 tracking-tight">
            Process Brokerage <span className="bg-gradient-to-r from-indigo-600 to-violet-600 bg-clip-text text-transparent">Files</span>
          </h1>
          <p className="mt-2 text-slate-500 text-sm">Upload AMC rate PDFs and Excel templates — get filled files in seconds.</p>
        </div>

        {/* Upload */}
        <div className="bg-white/80 backdrop-blur-sm rounded-3xl border border-slate-200/80 shadow-sm p-6 mb-4">
          <div className="flex gap-4 flex-col sm:flex-row">
            <DropZone label="Source Files" sublabel="AMC brokerage PDFs (or Excel files)" icon={FileText}
              accept={['pdf','xlsx','xls']} files={sourcePdfs} onFiles={setSourcePdfs} variant="source" />
            <div className="w-px bg-slate-200 hidden sm:block my-2" />
            <DropZone label="Target Templates" sublabel="Excel templates to fill" icon={Table2}
              accept={['xlsx','xls']} files={targetXlsx} onFiles={setTargetXlsx} variant="target" />
          </div>
        </div>

        {/* Settings */}
        <div className="bg-white/80 backdrop-blur-sm rounded-3xl border border-slate-200/80 shadow-sm px-6 py-4 mb-4">
          <div className="flex items-center gap-2 mb-4">
            <Settings2 className="w-4 h-4 text-slate-400" />
            <span className="text-sm font-semibold text-slate-700">Settings</span>
          </div>
          <div className="flex flex-col sm:flex-row gap-4">
            <div className="flex-1">
              <label className="block text-xs font-semibold text-slate-500 mb-1.5 uppercase tracking-wide">PDF Password</label>
              <div className="relative">
                <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                <input type={showPwd ? 'text' : 'password'} value={password} onChange={e => setPassword(e.target.value)}
                  placeholder="Leave blank if not required"
                  className="w-full pl-9 pr-10 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-sm text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-300 focus:border-indigo-400 transition" />
                <button type="button" onClick={() => setShowPwd(s => !s)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-slate-400 hover:text-slate-600 transition-colors">
                  {showPwd ? 'Hide' : 'Show'}
                </button>
              </div>
            </div>
            <div className="sm:w-64">
              <label className="block text-xs font-semibold text-slate-500 mb-1.5 uppercase tracking-wide">Template Format</label>
              <select value={format} onChange={e => setFormat(e.target.value)}
                className="w-full px-3 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-300 focus:border-indigo-400 transition appearance-none cursor-pointer">
                <option value="auto">Auto-detect format</option>
                <option value="redos">Standard (Redos / column-based)</option>
                <option value="vijay">Vijayinfotech (row-based)</option>
              </select>
            </div>
          </div>
        </div>

        {/* Error */}
        {error && (
          <div className="flex items-start gap-3 bg-red-50 border border-red-200 rounded-2xl px-4 py-3 mb-4 anim-slide-up">
            <XCircle className="w-4 h-4 text-red-500 mt-0.5 shrink-0" />
            <p className="text-sm text-red-700">{error}</p>
            <button onClick={() => setError('')} className="ml-auto text-red-400 hover:text-red-600"><X className="w-4 h-4" /></button>
          </div>
        )}

        {/* Process button */}
        <button onClick={handleProcess} disabled={!canProcess}
          className={`w-full py-3.5 rounded-2xl text-sm font-semibold flex items-center justify-center gap-2.5 transition-all duration-200
            ${canProcess
              ? 'bg-gradient-to-r from-indigo-600 to-violet-600 text-white hover:from-indigo-500 hover:to-violet-500 shadow-lg shadow-indigo-200 hover:shadow-xl hover:-translate-y-0.5 active:translate-y-0'
              : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}>
          {loading ? (<><span className="spinner" />Processing files…</>)
            : (<><Upload className="w-4 h-4" />Process &amp; Fill Brokerage Data<ChevronRight className="w-4 h-4 opacity-70" /></>)}
        </button>

        {/* Results */}
        {result && (
          <div className="mt-6 space-y-4 anim-slide-up">
            <div className="flex gap-3 flex-wrap">
              <StatCard label="Files Filled" value={result.filled_count ?? 0} icon={FileCheck2} colorClass="bg-indigo-100 text-indigo-600" />
              <StatCard label="Schemes Matched" value={result.total_matched ?? 0} icon={TrendingUp} colorClass="bg-emerald-100 text-emerald-600" />
              <StatCard label="Sources Parsed" value={result.source_log?.filter(r => r.icon === 'ok').length ?? 0} icon={FileText} colorClass="bg-violet-100 text-violet-600" />
            </div>
            <LogTable title="Source Files" icon={FileText} rows={result.source_log}
              columns={[{key:'file',label:'File'},{key:'amc',label:'AMC'},{key:'msg',label:'Status'}]} />
            <LogTable title="Target Templates" icon={Table2} rows={result.target_log}
              columns={[{key:'file',label:'File'},{key:'amc',label:'AMC'},{key:'filled',label:'Filled'},{key:'notes',label:'Notes'}]} />
            {result.token && (
              <a href={`/api/download/${result.token}`} download="brokerage_filled.zip"
                className="flex items-center justify-center gap-2.5 w-full py-3.5 rounded-2xl text-sm font-semibold
                  bg-gradient-to-r from-emerald-500 to-teal-500 text-white hover:from-emerald-400 hover:to-teal-400
                  shadow-lg shadow-emerald-200 hover:shadow-xl hover:-translate-y-0.5 active:translate-y-0 transition-all duration-200">
                <Download className="w-4 h-4" />
                Download Filled Excel Files (.zip)
              </a>
            )}
          </div>
        )}

        <p className="text-center text-xs text-slate-400 mt-8">BrokerageAI · EX-GST values · Supports 25+ AMCs</p>
      </div>
    </div>
  )
}
