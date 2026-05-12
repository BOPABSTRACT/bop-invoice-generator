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
  if (!val) return new Date().toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  if (val instanceof Date) return val.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  const d = new Date(String(val))
  if (!isNaN(d.getTime())) return d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  return String(val)
}

function sanitize(name: string): string {
  return name.replace(/[^a-zA-Z0-9_\-. ]/g, '_').trim()
}

const thinBorder: ExcelJS.Borders = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
}

async function buildExcelInvoice(row: Record<string, unknown>): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Summary')

  ws.getColumn(1).width = 14
  ws.getColumn(2).width = 26
  ws.getColumn(3).width = 14
  ws.getColumn(4).width = 14
  ws.getColumn(5).width = 16
  ws.getColumn(6).width = 12
  ws.getColumn(7).width = 12

  const county = String(row['County'] ?? '')
  const flatRate = Number(row['Flat Rate'] ?? row['Total'] ?? 0)

  function c(addr: string) { return ws.getCell(addr) }

  function label(addr: string, val: string, align: ExcelJS.Alignment['horizontal'] = 'right') {
    const cell = c(addr)
    cell.value = val
    cell.font = { name: 'Calibri', size: 9 }
    cell.alignment = { horizontal: align }
  }

  function value(addr: string, val: ExcelJS.CellValue, align: ExcelJS.Alignment['horizontal'] = 'left') {
    const cell = c(addr)
    cell.value = val
    cell.font = { name: 'Calibri', size: 9 }
    cell.alignment = { horizontal: align }
  }

  ws.mergeCells('A1:G1')
  const nameCell = c('A1')
  nameCell.value = BOP.name
  nameCell.font = { name: 'Calibri', size: 10, bold: true }
  nameCell.alignment = { horizontal: 'center' }

  ws.mergeCells('A2:G2')
  const addrCell = c('A2')
  addrCell.value = BOP.address
  addrCell.font = { name: 'Calibri', size: 10 }
  addrCell.alignment = { horizontal: 'center' }

  ws.mergeCells('A3:G3')
  const cityCell = c('A3')
  cityCell.value = BOP.city
  cityCell.font = { name: 'Calibri', size: 10 }
  cityCell.alignment = { horizontal: 'center' }

  ws.mergeCells('A4:G4')
  const phoneCell = c('A4')
  phoneCell.value = BOP.phone
  phoneCell.font = { name: 'Calibri', size: 10 }
  phoneCell.alignment = { horizontal: 'center' }

  ws.mergeCells('A6:G6')
  const invTitle = c('A6')
  invTitle.value = 'INVOICE'
  invTitle.font = { name: 'Calibri', size: 10, bold: true }
  invTitle.alignment = { horizontal: 'center' }

  label('A8', 'Date:')
  value('B8', formatDate(row['Date'] ?? new Date()))
  label('E8', 'Invoice #:')
  value('F8', String(row['Invoice Number'] ?? ''))

  label('A10', 'Bill To:')
  value('B10', BILL_TO.company)
  label('E10', 'Project:')
  value('F10', BILL_TO.project)

  value('B11', BILL_TO.attn)
  label('E11', 'County:')
  value('F11', `${county}, PA`)

  value('B12', BILL_TO.address)
  value('B13', BILL_TO.city)

  label('A16', 'Period:')
  value('B16', String(row['Period'] ?? ''))

  const dueCell = c('G17')
  dueCell.value = 'DUE UPON RECEIPT'
  dueCell.font = { name: 'Calibri', size: 9, color: { argb: 'FFFF0000' } }
  dueCell.alignment = { horizontal: 'right' }

  const headers = ['File #', 'PID', 'Unit', 'County', 'Work Type', 'Flat Rate', 'Total']
  const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
  ws.getRow(18).height = 16
  headers.forEach((h, i) => {
    const cell = c(`${cols[i]}18`)
    cell.value = h
    cell.font = { name: 'Calibri', size: 9, bold: true }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } }
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
    cell.border = thinBorder
  })

  const dataVals = [
    String(row['File #'] ?? ''),
    String(row['PID'] ?? ''),
    String(row['Unit'] ?? ''),
    county,
    'Targeted Scope',
  ]
  dataVals.forEach((v, i) => {
    const cell = c(`${cols[i]}19`)
    cell.value = v
    cell.font = { name: 'Calibri', size: 9 }
    cell.alignment = { horizontal: 'center' }
    cell.border = thinBorder
  })

  const flatCell = c('F19')
  flatCell.value = flatRate
  flatCell.numFmt = '$#,##0.00'
  flatCell.font = { name: 'Calibri', size: 9 }
  flatCell.border = thinBorder

  const totalCell = c('G19')
  totalCell.value = { formula: '=F19' }
  totalCell.numFmt = '$#,##0.00'
  totalCell.font = { name: 'Calibri', size: 9 }
  totalCell.border = thinBorder

  const filesCell = c('B20')
  filesCell.value = { formula: '=COUNTA(B19:B19)&" Files"' }
  filesCell.font = { name: 'Calibri', size: 9, bold: true }
  filesCell.alignment = { horizontal: 'right' }

  const totalLbl = c('E20')
  totalLbl.value = 'Total'
  totalLbl.font = { name: 'Calibri', size: 9, bold: true }
  totalLbl.alignment = { horizontal: 'right' }

  const sumFlat = c('F20')
  sumFlat.value = { formula: '=SUM(F19:F19)' }
  sumFlat.numFmt = '$#,##0.00'
  sumFlat.font = { name: 'Calibri', size: 9, bold: true }
  sumFlat.alignment = { horizontal: 'right' }

  const sumTotal = c('G20')
  sumTotal.value = { formula: '=SUM(G19:G19)' }
  sumTotal.numFmt = '$#,##0.00'
  sumTotal.font = { name: 'Calibri', size: 9, bold: true }
  sumTotal.alignment = { horizontal: 'right' }
  sumTotal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }

  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}

async function buildPdfInvoice(row: Record<string, unknown>): Promise<Buffer> {
  const { jsPDF } = await import('jspdf')
  const autoTable = (await import('jspdf-autotable')).default

  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'letter' })

  const black = [0, 0, 0] as [number, number, number]
  const red = [255, 0, 0] as [number, number, number]
  const headerBg = [242, 220, 219] as [number, number, number]
  const white = [255, 255, 255] as [number, number, number]

  const county = String(row['County'] ?? '')

  doc.setFillColor(...white)
  doc.rect(0, 0, 612, 792, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.setTextColor(...black)
  doc.text(BOP.name, 306, 50, { align: 'center' })

  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.text(BOP.address, 306, 63, { align: 'center' })
  doc.text(BOP.city, 306, 75, { align: 'center' })
  doc.text(BOP.phone, 306, 87, { align: 'center' })

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.text('INVOICE', 306, 115, { align: 'center' })

  doc.setDrawColor(...black)
  doc.setLineWidth(0.5)
  doc.line(40, 122, 572, 122)

  doc.setFontSize(9)
  doc.setFont('helvetica', 'normal')
  doc.setTextColor(...black)
  doc.text('Date:', 100, 140, { align: 'right' })
  doc.text(formatDate(row['Date'] ?? new Date()), 105, 140)

  doc.text('Invoice #:', 420, 140, { align: 'right' })
  doc.text(String(row['Invoice Number'] ?? ''), 425, 140)

  doc.text('Bill To:', 100, 158, { align: 'right' })
  doc.text(BILL_TO.company, 105, 158)
  doc.text('Project:', 420, 158, { align: 'right' })
  doc.text(BILL_TO.project, 425, 158)

  doc.text(BILL_TO.attn, 105, 170)
  doc.text('County:', 420, 170, { align: 'right' })
  doc.text(`${county}, PA`, 425, 170)

  doc.text(BILL_TO.address, 105, 182)
  doc.text(BILL_TO.city, 105, 194)

  doc.text('Period:', 100, 215, { align: 'right' })
  doc.text(String(row['Period'] ?? ''), 105, 215)

  doc.setTextColor(...red)
  doc.text('DUE UPON RECEIPT', 572, 227, { align: 'right' })
  doc.setTextColor(...black)

  const flatRate = Number(row['Flat Rate'] ?? row['Total'] ?? 0)
  const fmt = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' })

  autoTable(doc, {
    startY: 240,
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
      ['', '1 Files', '', '', 'Total', fmt.format(flatRate), fmt.format(flatRate)],
    ],
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: 9,
      textColor: black,
      fillColor: white,
      cellPadding: 5,
      lineColor: black,
      lineWidth: 0.3,
    },
    headStyles: {
      textColor: black,
      fillColor: headerBg,
      fontStyle: 'bold',
      halign: 'center',
    },
    bodyStyles: {
      halign: 'center',
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
      const invoiceNum = sanitize(String(row['Invoice Number'] ?? `INV_${i + 1}`))
      const fileNum = sanitize(String(row['File #'] ?? ''))
      const pid = sanitize(String(row['PID'] ?? ''))
      const unit = sanitize(String(row['Unit'] ?? ''))
      const county = sanitize(String(row['County'] ?? ''))
      const baseName = `${invoiceNum} - ${fileNum} - ${pid} - ${unit} - ${county}`

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
