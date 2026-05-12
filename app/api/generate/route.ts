import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'

const BILL_TO = {
  company: 'EQT Production Company',
  attn: 'Attn: Zita Lammay',
  address: 'P.O. Box 23425',
  city: 'Pittsburgh, Pennsylvania 15222',
  project: 'EQT Targeted Scope',
}

const BOP = {
  name: 'BOP Abstract, LLC',
  address: '2547 Washington Rd., Bldg. 700, Ste. 720',
  city: 'Pittsburgh, Pennsylvania, 15241',
  phone: '724-747-1594',
}

function formatDate(val: unknown): string {
  if (!val) return ''
  if (val instanceof Date) {
    return val.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  }
  const d = new Date(String(val))
  if (!isNaN(d.getTime())) {
    return d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  }
  return String(val)
}

function sanitize(name: string): string {
  return name.replace(/[^a-zA-Z0-9_\-. ]/g, '_').trim()
}

async function buildExcelInvoice(row: Record<string, unknown>): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Summary')

  ws.getColumn(1).width = 18
  ws.getColumn(2).width = 28
  ws.getColumn(3).width = 14
  ws.getColumn(4).width = 14
  ws.getColumn(5).width = 18
  ws.getColumn(6).width = 14
  ws.getColumn(7).width = 14

  const gold = 'FFC8A96E'
  const darkBg = 'FF0D0F14'
  const headerBg = 'FF1a1a2a'
  const white = 'FFE8E0D0'
  const lightGray = 'FFbbbbbb'

  function setCell(ws: ExcelJS.Worksheet, addr: string, value: ExcelJS.CellValue, opts: {
    bold?: boolean, size?: number, color?: string, bg?: string,
    align?: ExcelJS.Alignment['horizontal'], border?: boolean, italic?: boolean
  } = {}) {
    const cell = ws.getCell(addr)
    cell.value = value
    cell.font = {
      name: 'Georgia',
      size: opts.size ?? 10,
      bold: opts.bold ?? false,
      italic: opts.italic ?? false,
      color: { argb: opts.color ?? white },
    }
    if (opts.bg) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.bg } }
    }
    if (opts.align) {
      cell.alignment = { horizontal: opts.align, vertical: 'middle' }
    }
    if (opts.border) {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF2a2a3a' } },
        bottom: { style: 'thin', color: { argb: 'FF2a2a3a' } },
        left: { style: 'thin', color: { argb: 'FF2a2a3a' } },
        right: { style: 'thin', color: { argb: 'FF2a2a3a' } },
      }
    }
  }

  ws.getRow(1).height = 18
  setCell(ws, 'A1', BOP.name, { bold: true, size: 13, color: gold })
  setCell(ws, 'A2', BOP.address, { size: 10, color: lightGray })
  setCell(ws, 'A3', BOP.city, { size: 10, color: lightGray })
  setCell(ws, 'A4', BOP.phone, { size: 10, color: lightGray })

  ws.getRow(6).height = 22
  setCell(ws, 'A6', 'INVOICE', { bold: true, size: 20, color: gold })

  ws.getRow(8).height = 16
  setCell(ws, 'A8', 'Date:', { bold: true, color: gold })
  setCell(ws, 'B8', formatDate(row['Date']))
  setCell(ws, 'E8', 'Invoice #:', { bold: true, color: gold })
  setCell(ws, 'F8', String(row['Invoice Number'] ?? ''))

  ws.getRow(10).height = 16
  setCell(ws, 'A10', 'Bill To:', { bold: true, color: gold })
  setCell(ws, 'B10', BILL_TO.company, { bold: true })
  setCell(ws, 'E10', 'Project:', { bold: true, color: gold })
  setCell(ws, 'F10', BILL_TO.project)

  setCell(ws, 'B11', BILL_TO.attn, { color: lightGray })
  const county = String(row['County'] ?? '')
  const state = String(row['State'] ?? 'PA')
  setCell(ws, 'E11', 'County:', { bold: true, color: gold })
  setCell(ws, 'F11', `${county}, ${state}`)

  setCell(ws, 'B12', BILL_TO.address, { color: lightGray })
  setCell(ws, 'B13', BILL_TO.city, { color: lightGray })

  ws.getRow(16).height = 16
  setCell(ws, 'A16', 'Period:', { bold: true, color: gold })
  setCell(ws, 'B16', String(row['Period'] ?? ''))
  setCell(ws, 'G17', 'DUE UPON RECEIPT', { bold: true, color: gold, align: 'right' })

  ws.getRow(18).height = 18
  const headers = ['File #', 'PID', 'Unit', 'County', 'Work Type', 'Flat Rate', 'Total']
  const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
  headers.forEach((h, i) => {
    setCell(ws, `${cols[i]}18`, h, { bold: true, color: gold, bg: headerBg, border: true, align: 'center' })
  })

  ws.getRow(19).height = 16
  const flatRate = Number(row['Flat Rate'] ?? row['Total'] ?? 0)
  setCell(ws, 'A19', String(row['File #'] ?? ''), { border: true })
  setCell(ws, 'B19', String(row['PID'] ?? ''), { border: true })
  setCell(ws, 'C19', String(row['Unit'] ?? ''), { border: true })
  setCell(ws, 'D19', county, { border: true })
  setCell(ws, 'E19', 'Targeted Scope', { border: true })

  const flatCell = ws.getCell('F19')
  flatCell.value = flatRate
  flatCell.numFmt = '$#,##0.00'
  flatCell.font = { name: 'Georgia', size: 10, color: { argb: white } }
  flatCell.border = { top: { style: 'thin', color: { argb: 'FF2a2a3a' } }, bottom: { style: 'thin', color: { argb: 'FF2a2a3a' } }, left: { style: 'thin', color: { argb: 'FF2a2a3a' } }, right: { style: 'thin', color: { argb: 'FF2a2a3a' } } }

  const totalCell = ws.getCell('G19')
  totalCell.value = { formula: '=F19' }
  totalCell.numFmt = '$#,##0.00'
  totalCell.font = { name: 'Georgia', size: 10, color: { argb: white } }
  totalCell.border = { top: { style: 'thin', color: { argb: 'FF2a2a3a' } }, bottom: { style: 'thin', color: { argb: 'FF2a2a3a' } }, left: { style: 'thin', color: { argb: 'FF2a2a3a' } }, right: { style: 'thin', color: { argb: 'FF2a2a3a' } } }

  ws.getRow(20).height = 18
  setCell(ws, 'B20', '1 File', { color: lightGray, italic: true })
  setCell(ws, 'E20', 'Total', { bold: true, color: gold, align: 'right' })

  const sumFlat = ws.getCell('F20')
  sumFlat.value = { formula: '=SUM(F19:F19)' }
  sumFlat.numFmt = '$#,##0.00'
  sumFlat.font = { name: 'Georgia', size: 10, bold: true, color: { argb: gold } }

  const sumTotal = ws.getCell('G20')
  sumTotal.value = { formula: '=SUM(G19:G19)' }
  sumTotal.numFmt = '$#,##0.00'
  sumTotal.font = { name: 'Georgia', size: 10, bold: true, color: { argb: gold } }

  for (let r = 1; r <= 25; r++) {
    for (let c = 1; c <= 7; c++) {
      const cell = ws.getCell(r, c)
      if (!cell.fill || (cell.fill as ExcelJS.FillPattern).fgColor?.argb === undefined) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: darkBg } }
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}

async function buildPdfInvoice(row: Record<string, unknown>): Promise<Buffer> {
  const { jsPDF } = await import('jspdf')
  const autoTable = (await import('jspdf-autotable')).default

  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'letter' })

  const gold = [200, 169, 110] as [number, number, number]
  const dark = [13, 15, 20] as [number, number, number]
  const lightText = [180, 170, 155] as [number, number, number]
  const white = [232, 224, 208] as [number, number, number]

  doc.setFillColor(...dark)
  doc.rect(0, 0, 612, 792, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(16)
  doc.setTextColor(...gold)
  doc.text(BOP.name, 40, 55)

  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.setTextColor(...lightText)
  doc.text(BOP.address, 40, 70)
  doc.text(BOP.city, 40, 82)
  doc.text(BOP.phone, 40, 94)

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(28)
  doc.setTextColor(...gold)
  doc.text('INVOICE', 40, 135)

  doc.setDrawColor(...gold)
  doc.setLineWidth(0.5)
  doc.line(40, 145, 572, 145)

  doc.setFontSize(9)
  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('Date:', 40, 162)
  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...white)
  doc.text(formatDate(row['Date']), 90, 162)

  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('Invoice #:', 380, 162)
  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...white)
  doc.text(String(row['Invoice Number'] ?? ''), 440, 162)

  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('Bill To:', 40, 185)
  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...white)
  doc.text(BILL_TO.company, 90, 185)

  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('Project:', 380, 185)
  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...white)
  doc.text(BILL_TO.project, 430, 185)

  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...lightText)
  doc.text(BILL_TO.attn, 90, 198)

  const county = String(row['County'] ?? '')
  const state = String(row['State'] ?? 'PA')
  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('County:', 380, 198)
  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...white)
  doc.text(`${county}, ${state}`, 430, 198)

  doc.setTextColor(...lightText)
  doc.text(BILL_TO.address, 90, 211)
  doc.text(BILL_TO.city, 90, 224)

  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('Period:', 40, 248)
  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...white)
  doc.text(String(row['Period'] ?? ''), 90, 248)

  doc.setFont('helvetica', 'bold')
  doc.setTextColor(...gold)
  doc.text('DUE UPON RECEIPT', 572, 260, { align: 'right' })

  const flatRate = Number(row['Flat Rate'] ?? row['Total'] ?? 0)
  const fmt = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' })

  autoTable(doc, {
    startY: 275,
    head: [['File #', 'PID', 'Unit', 'County', 'Work Type', 'Flat Rate', 'Total']],
    body: [
      [
        String(row['File #'] ?? ''),
        String(row['PID'] ?? ''),
        String(row['Unit'] ?? ''),
        county,
        'Targeted Scope',
        fmt.format(flatRate),
        fmt.format(flatRate),
      ],
      ['', '1 File', '', '', 'Total', fmt.format(flatRate), fmt.format(flatRate)],
    ],
    theme: 'plain',
    styles: {
      font: 'helvetica',
      fontSize: 9,
      textColor: white,
      fillColor: dark,
      cellPadding: 6,
    },
    headStyles: {
      textColor: gold,
      fillColor: [26, 26, 42],
      fontStyle: 'bold',
      lineColor: [42, 42, 58],
      lineWidth: 0.5,
    },
    bodyStyles: {
      lineColor: [42, 42, 58],
      lineWidth: 0.3,
    },
    columnStyles: {
      5: { halign: 'right' },
      6: { halign: 'right' },
    },
    margin: { left: 40, right: 40 },
  })

  return Buffer.from(doc.output('arraybuffer'))
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const excelFile = formData.get('excel') as File
    const format = (formData.get('format') as string) || 'both'

    if (!excelFile) {
      return NextResponse.json({ error: 'Missing excel file' }, { status: 400 })
    }

    const buffer = Buffer.from(await excelFile.arrayBuffer())
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' }) as Record<string, unknown>[]

    if (rows.length === 0) {
      return NextResponse.json({ error: 'No data rows found in spreadsheet' }, { status: 400 })
    }

    const zip = new JSZip()

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]
      const invoiceNum = String(row['Invoice Number'] ?? `INV_${i + 1}`)
      const fileNum = sanitize(String(row['File #'] ?? ''))
      const pid = sanitize(String(row['PID'] ?? ''))
      const baseName = `${invoiceNum}_${fileNum}_${pid}`

      if (format === 'xlsx' || format === 'both') {
        try {
          const xlsxBuf = await buildExcelInvoice(row)
          zip.file(`${baseName}.xlsx`, xlsxBuf)
        } catch (err) {
          return NextResponse.json({ error: `Excel failed for row ${i + 1}: ${String(err)}` }, { status: 500 })
        }
      }

      if (format === 'pdf' || format === 'both') {
        try {
          const pdfBuf = await buildPdfInvoice(row)
          zip.file(`${baseName}.pdf`, pdfBuf)
        } catch (err) {
          return NextResponse.json({ error: `PDF failed for row ${i + 1}: ${String(err)}` }, { status: 500 })
        }
      }
    }

    const zipBuffer = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })

    return new NextResponse(zipBuffer, {
      headers: {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="bop-invoices.zip"`,
      },
    })
  } catch (err) {
    console.error(err)
    return NextResponse.json({ error: String(err) }, { status: 500 })
  }
}
