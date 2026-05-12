'use client'

import { useState, useRef } from 'react'

const LOGO = "https://i.imgur.com/szjzoxt.png"

export default function Home() {
  const [authenticated, setAuthenticated] = useState(false)
  const [passwordInput, setPasswordInput] = useState('')
  const [passwordError, setPasswordError] = useState(false)
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [status, setStatus] = useState<'idle' | 'loading' | 'done' | 'error'>('idle')
  const [message, setMessage] = useState('')
  const [previewData, setPreviewData] = useState<Record<string, string>[] | null>(null)
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'pdf' | 'both'>('both')
  const excelRef = useRef<HTMLInputElement>(null)

  const handlePasswordSubmit = () => {
    if (passwordInput === 'BOP2026') {
      setAuthenticated(true)
      setPasswordError(false)
    } else {
      setPasswordError(true)
    }
  }

  const handleExcelChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setExcelFile(file)
    const formData = new FormData()
    formData.append('excel', file)
    try {
      const res = await fetch('/api/preview', { method: 'POST', body: formData })
      const data = await res.json()
      setPreviewData(data.rows)
    } catch {
      // preview is optional
    }
  }

  const handleGenerate = async () => {
    if (!excelFile) {
      setMessage('Please upload an Excel file.')
      setStatus('error')
      return
    }
    setStatus('loading')
    setMessage('Generating invoices...')

    const formData = new FormData()
    formData.append('excel', excelFile)
    formData.append('format', outputFormat)

    try {
      const res = await fetch('/api/generate', { method: 'POST', body: formData })
      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.error || 'Generation failed')
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `bop-invoices-${Date.now()}.zip`
      a.click()
      URL.revokeObjectURL(url)
      setStatus('done')
      setMessage('✅ Invoices generated and downloaded!')
    } catch (err: unknown) {
      setStatus('error')
      setMessage(err instanceof Error ? err.message : 'Something went wrong.')
    }
  }

  if (!authenticated) {
    return (
      <main style={{
        minHeight: '100vh',
        background: '#0f1117',
        fontFamily: "'Georgia', serif",
        color: '#e8e0d0',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}>
        <div style={{
          background: '#0d0f14',
          border: '1px solid #2a2a3a',
          borderRadius: 12,
          padding: '48px 40px',
          width: '100%',
          maxWidth: 400,
          textAlign: 'center',
        }}>
          <img
            src={LOGO}
            alt="BOP Acquisition Logo"
            style={{ width: 140, height: 140, objectFit: 'contain', margin: '0 auto 24px', display: 'block' }}
          />
          <div style={{ fontSize: 20, fontWeight: 600, color: '#c8a96e', marginBottom: 4 }}>
            BOP ACQUISITION
          </div>
          <div style={{ fontSize: 12, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase', marginBottom: 32 }}>
            Invoice Generator
          </div>
          <input
            type="password"
            placeholder="Enter password"
            value={passwordInput}
            onChange={e => { setPasswordInput(e.target.value); setPasswordError(false) }}
            onKeyDown={e => e.key === 'Enter' && handlePasswordSubmit()}
            style={{
              width: '100%',
              padding: '12px 16px',
              background: '#0f1117',
              border: `1px solid ${passwordError ? '#8b2020' : '#2a2a3a'}`,
              borderRadius: 6,
              color: '#e8e0d0',
              fontSize: 15,
              fontFamily: "'Georgia', serif",
              boxSizing: 'border-box',
              marginBottom: 12,
              outline: 'none',
            }}
          />
          {passwordError && (
            <div style={{ color: '#e07070', fontSize: 13, marginBottom: 12 }}>
              Incorrect password. Please try again.
            </div>
          )}
          <button
            onClick={handlePasswordSubmit}
            style={{
              width: '100%',
              padding: '12px 32px',
              background: 'linear-gradient(135deg, #c8a96e, #8b6914)',
              color: '#fff',
              border: 'none',
              borderRadius: 6,
              fontSize: 15,
              fontFamily: "'Georgia', serif",
              cursor: 'pointer',
              letterSpacing: '0.04em',
            }}
          >
            Enter
          </button>
        </div>
      </main>
    )
  }

  return (
    <main style={{
      minHeight: '100vh',
      background: '#0f1117',
      fontFamily: "'Georgia', serif",
      color: '#e8e0d0',
      padding: '0',
    }}>
      <header style={{
        borderBottom: '1px solid #2a2a3a',
        padding: '16px 48px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        background: '#0d0f14',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          <img
            src={LOGO}
            alt="BOP Acquisition Logo"
            style={{ width: 52, height: 52, objectFit: 'contain', flexShrink: 0 }}
          />
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, letterSpacing: '0.04em', color: '#c8a96e' }}>
              BOP ACQUISITION
            </div>
            <div style={{ fontSize: 11, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase' }}>
              Invoice Generator
            </div>
          </div>
        </div>
        <a href="/help" style={{
          color: '#c8a96e', fontSize: 13, textDecoration: 'none',
          border: '1px solid #333', padding: '6px 14px', borderRadius: 4,
          letterSpacing: '0.04em',
        }}>User Guide</a>
      </header>

      <div style={{ maxWidth: 820, margin: '0 auto', padding: '48px 32px' }}>
        <div style={{ marginBottom: 48 }}>
          <h1 style={{
            fontSize: 36, fontWeight: 400, color: '#e8e0d0',
            margin: '0 0 12px 0', letterSpacing: '-0.01em', lineHeight: 1.2,
          }}>
            Generate Invoices
          </h1>
          <p style={{ color: '#888', fontSize: 15, margin: 0, lineHeight: 1.6 }}>
            Upload your Excel spreadsheet. The app will generate one invoice per row
            as Excel (.xlsx) and/or PDF, then package them into a ZIP for download.
          </p>
        </div>

        {/* Step 1: Upload Excel */}
        <Section number="1" title="Upload Excel Spreadsheet">
          <UploadBox
            label="Drop your .xlsx file here or click to browse"
            accept=".xlsx,.xls"
            file={excelFile}
            onChange={handleExcelChange}
            inputRef={excelRef}
          />
          {previewData && previewData.length > 0 && (
            <div style={{ marginTop: 16 }}>
              <div style={{ fontSize: 12, color: '#888', marginBottom: 8, letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                Preview — {previewData.length} invoice{previewData.length !== 1 ? 's' : ''} will be generated
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr>
                      {['Invoice #', 'File #', 'PID', 'Unit', 'Period', 'Flat Rate'].map(k => (
                        <th key={k} style={{
                          textAlign: 'left', padding: '6px 10px',
                          borderBottom: '1px solid #2a2a3a',
                          color: '#c8a96e', fontWeight: 500, whiteSpace: 'nowrap',
                        }}>{k}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewData.slice(0, 5).map((row, i) => (
                      <tr key={i} style={{ borderBottom: '1px solid #1a1a2a' }}>
                        {['Invoice Number', 'File #', 'PID', 'Unit', 'Period', 'Flat Rate'].map((k, j) => (
                          <td key={j} style={{ padding: '6px 10px', color: '#bbb', whiteSpace: 'nowrap' }}>
                            {String(row[k] ?? '—')}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </Section>

        {/* Step 2: Output Format */}
        <Section number="2" title="Select Output Format">
          <div style={{ display: 'flex', gap: 12 }}>
            {(['xlsx', 'pdf', 'both'] as const).map(fmt => (
              <button
                key={fmt}
                onClick={() => setOutputFormat(fmt)}
                style={{
                  flex: 1,
                  padding: '14px 12px',
                  borderRadius: 6,
                  border: `1px solid ${outputFormat === fmt ? '#c8a96e' : '#2a2a3a'}`,
                  background: outputFormat === fmt ? 'rgba(200,169,110,0.1)' : '#0d0f14',
                  color: outputFormat === fmt ? '#c8a96e' : '#666',
                  cursor: 'pointer',
                  fontSize: 14,
                  fontFamily: "'Georgia', serif",
                  letterSpacing: '0.04em',
                  transition: 'all 0.2s',
                }}
              >
                {fmt === 'xlsx' ? '📊 Excel (.xlsx)' : fmt === 'pdf' ? '📄 PDF' : '📦 Both'}
              </button>
            ))}
          </div>
        </Section>

        {/* Step 3: Generate */}
        <Section number="3" title="Generate Invoices">
          <button
            onClick={handleGenerate}
            disabled={status === 'loading'}
            style={{
              width: '100%', padding: '16px 32px',
              background: status === 'loading'
                ? '#2a2a3a'
                : 'linear-gradient(135deg, #c8a96e, #8b6914)',
              color: status === 'loading' ? '#666' : '#fff',
              border: 'none', borderRadius: 6, fontSize: 16,
              fontFamily: "'Georgia', serif", letterSpacing: '0.04em',
              cursor: status === 'loading' ? 'not-allowed' : 'pointer',
              transition: 'all 0.2s',
            }}
          >
            {status === 'loading' ? '⏳ Generating...' : '⬇ Generate & Download ZIP'}
          </button>

          {message && (
            <div style={{
              marginTop: 16, padding: '12px 16px', borderRadius: 6,
              background: status === 'error' ? 'rgba(200,60,60,0.1)' : 'rgba(60,180,100,0.1)',
              border: `1px solid ${status === 'error' ? '#8b2020' : '#2a6640'}`,
              color: status === 'error' ? '#e07070' : '#70c090',
              fontSize: 14,
            }}>
              {message}
            </div>
          )}
        </Section>

        {/* Reference Card */}
        <div style={{ marginTop: 48, padding: 24, background: '#0d0f14', borderRadius: 8, border: '1px solid #1e1e2e' }}>
          <div style={{ fontSize: 11, color: '#c8a96e', letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 16 }}>
            Expected Excel Columns
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px 24px' }}>
            {[
              'Year', 'Lessor 1', 'Lessor 2', 'Status', 'Address',
              'Township', 'County', 'State', 'Gross Acres', 'Date',
              'Invoice Number', 'Period', 'File #', 'PID', 'Unit', 'Flat Rate', 'Total'
            ].map(col => (
              <span key={col} style={{ fontSize: 12, color: '#666', fontFamily: 'monospace' }}>{col}</span>
            ))}
          </div>
        </div>
      </div>
    </main>
  )
}

function Section({ number, title, children }: { number: string; title: string; children: React.ReactNode }) {
  return (
    <div style={{ marginBottom: 40 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
        <div style={{
          width: 28, height: 28, borderRadius: '50%',
          background: 'rgba(200,169,110,0.15)',
          border: '1px solid #c8a96e',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: 13, color: '#c8a96e', fontWeight: 600, flexShrink: 0,
        }}>{number}</div>
        <h2 style={{ margin: 0, fontSize: 17, fontWeight: 500, color: '#e8e0d0', letterSpacing: '0.01em' }}>{title}</h2>
      </div>
      {children}
    </div>
  )
}

function UploadBox({
  label, accept, file, onChange, inputRef,
}: {
  label: string; accept: string; file: File | null;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  inputRef: React.RefObject<HTMLInputElement>;
}) {
  return (
    <div
      onClick={() => inputRef.current?.click()}
      style={{
        border: `2px dashed ${file ? '#c8a96e' : '#2a2a3a'}`,
        borderRadius: 8, padding: '28px 24px', textAlign: 'center',
        cursor: 'pointer',
        background: file ? 'rgba(200,169,110,0.04)' : '#0d0f14',
        transition: 'all 0.2s',
      }}
    >
      <input ref={inputRef} type="file" accept={accept} onChange={onChange} style={{ display: 'none' }} />
      {!file ? (
        <>
          <div style={{ fontSize: 28, marginBottom: 8 }}>📊</div>
          <div style={{ color: '#888', fontSize: 14 }}>{label}</div>
          <div style={{ color: '#555', fontSize: 12, marginTop: 4 }}>XLSX / XLS</div>
        </>
      ) : (
        <div style={{ textAlign: 'left' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <span style={{ fontSize: 16 }}>✅</span>
            <span style={{ color: '#c8a96e', fontSize: 14 }}>{file.name}</span>
            <span style={{ color: '#555', fontSize: 12 }}>({(file.size / 1024).toFixed(1)} KB)</span>
          </div>
          <div style={{ color: '#555', fontSize: 12, marginTop: 8 }}>Click to change</div>
        </div>
      )}
    </div>
  )
}
