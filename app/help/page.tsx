'use client'

const LOGO = "https://i.imgur.com/szjzoxt.png"

export default function HelpPage() {
  return (
    <main style={{
      minHeight: '100vh',
      background: '#0f1117',
      fontFamily: "'Georgia', serif",
      color: '#e8e0d0',
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
          <img src={LOGO} alt="BOP Logo" style={{ width: 52, height: 52, objectFit: 'contain' }} />
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, letterSpacing: '0.04em', color: '#c8a96e' }}>BOP ACQUISITION</div>
            <div style={{ fontSize: 11, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase' }}>Invoice Generator</div>
          </div>
        </div>
        <a href="/" style={{ color: '#c8a96e', fontSize: 13, textDecoration: 'none', border: '1px solid #333', padding: '6px 14px', borderRadius: 4 }}>
          ← Back to App
        </a>
      </header>

      <div style={{ maxWidth: 820, margin: '0 auto', padding: '48px 32px' }}>
        <h1 style={{ fontSize: 32, fontWeight: 400, color: '#e8e0d0', marginBottom: 8 }}>User Guide</h1>
        <p style={{ color: '#888', fontSize: 15, marginBottom: 40 }}>How to use the BOP Invoice Generator</p>

        {[
          {
            title: '1. Prepare Your Excel File',
            content: `Your Excel spreadsheet must have the following columns in the first sheet (Sheet1). Each row will become one invoice.\n\nRequired columns:\n• Year\n• Lessor 1, Lessor 2\n• Status\n• Address, Township, County, State\n• Gross Acres\n• Date\n• Invoice Number\n• Period\n• File #\n• PID\n• Unit\n• Flat Rate\n• Total`
          },
          {
            title: '2. Upload the File',
            content: 'Click the upload box or drag and drop your .xlsx file. A preview table will appear showing the first 5 rows — verify the data looks correct before proceeding.'
          },
          {
            title: '3. Choose Output Format',
            content: 'Select Excel (.xlsx), PDF, or Both. Both is recommended — you\'ll get one Excel and one PDF invoice per row, all inside a single ZIP file.'
          },
          {
            title: '4. Generate & Download',
            content: 'Click "Generate & Download ZIP". The app will create one invoice per row and package them into a ZIP file that downloads automatically to your computer.'
          },
          {
            title: 'File Naming',
            content: 'Each invoice is named: [Invoice Number]_[File #]_[PID].xlsx / .pdf\n\nExample: 13970_TT-198450_240-005-00-01-0005-00.xlsx'
          },
          {
            title: 'Invoice Contents',
            content: 'Each invoice includes:\n• BOP Abstract LLC header with address and phone\n• Bill To: EQT Production Company\n• Invoice number, date, period\n• Line item table with File #, PID, Unit, County, Work Type, Flat Rate, and Total\n• "DUE UPON RECEIPT" notation'
          },
        ].map((section, i) => (
          <div key={i} style={{ marginBottom: 32, padding: 24, background: '#0d0f14', borderRadius: 8, border: '1px solid #1e1e2e' }}>
            <div style={{ fontSize: 15, fontWeight: 600, color: '#c8a96e', marginBottom: 12 }}>{section.title}</div>
            <div style={{ fontSize: 14, color: '#aaa', lineHeight: 1.8, whiteSpace: 'pre-line' }}>{section.content}</div>
          </div>
        ))}
      </div>
    </main>
  )
}
