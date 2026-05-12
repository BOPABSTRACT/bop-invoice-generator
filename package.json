import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const excelFile = formData.get('excel') as File
    if (!excelFile) return NextResponse.json({ rows: [] })

    const buffer = Buffer.from(await excelFile.arrayBuffer())
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' }) as Record<string, unknown>[]

    const formatted = rows.map(row => {
      const out: Record<string, string> = {}
      for (const [k, v] of Object.entries(row)) {
        if (v instanceof Date) {
          out[k] = v.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
        } else {
          out[k] = String(v ?? '')
        }
      }
      return out
    })

    return NextResponse.json({ rows: formatted })
  } catch {
    return NextResponse.json({ rows: [] })
  }
}
