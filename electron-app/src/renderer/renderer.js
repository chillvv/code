const $ = sel => document.querySelector(sel)
let df = []
let columns = []

function setStatus(msg) { $('#status').textContent = msg }

window.api.onStatus((msg) => setStatus(msg))
window.api.onQR(({ qrcode, status }) => {
  setStatus(`扫码登录：${status}`)
  try {
    const url = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(qrcode)}`
    $('#qr').src = url
  } catch {}
})

function inferColumns() {
  const selectIds = ['#colProduct', '#colPay', '#colStatus', '#colAddr', '#colNote']
  selectIds.forEach(id => {
    const el = $(id)
    el.innerHTML = ''
    columns.forEach(c => {
      const opt = document.createElement('option')
      opt.value = c
      opt.textContent = c
      el.appendChild(opt)
    })
  })
  const setIfExists = (id, name) => { if (columns.includes(name)) $(id).value = name }
  setIfExists('#colProduct', '商品信息')
  setIfExists('#colPay', '支付状态')
  setIfExists('#colStatus', '订单状态')
  setIfExists('#colAddr', '收货地址')
  setIfExists('#colNote', '用户备注')
}

function detectHeader(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1')
  let bestRow = range.s.r
  let bestCount = -1
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 10); r++) {
    let count = 0
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })]
      if (cell && cell.v !== undefined && cell.v !== '') count++
    }
    if (count > bestCount) { bestCount = count; bestRow = r }
  }
  return bestRow
}

function loadFile(file) {
  const reader = new FileReader()
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result)
    const wb = XLSX.read(data, { type: 'array' })
    const sn = wb.SheetNames[0]
    const sheet = wb.Sheets[sn]
    const headerRow = detectHeader(sheet)
    const opts = { header: 1, raw: false }
    const rows = XLSX.utils.sheet_to_json(sheet, opts)
    const headers = rows[headerRow] || []
    columns = headers.map(h => String(h).trim())
    df = rows.slice(headerRow + 1).map(r => {
      const obj = {}
      columns.forEach((c, i) => obj[c] = r[i] ?? '')
      return obj
    })
    inferColumns()
    setStatus(`已加载：${file.name || file}`)
    $('#fileName').textContent = file.name || file
  }
  reader.readAsArrayBuffer(file)
}

function splitAddress(text) {
  const t = String(text || '').trim()
  if (!t) return ' -  - '
  const parts = t.split(/\s*[-\-\–\—\－]\s*/)
  if (parts.length >= 3) return `${parts[0]} - ${parts[1]} - ${parts.slice(2).join(' - ')}`
  const m = t.match(/(\d[\d\s-]{5,19}\d)/)
  if (m) {
    const s = m.index, e = s + m[0].length
    const name = t.slice(0, s).replace(/[ -]+$/, '')
    const phone = t.slice(s, e)
    const addr = t.slice(e).replace(/^[ -]+/, '')
    return `${name || ''} - ${phone} - ${addr || ''}`
  }
  return ` -  - ${t}`
}

function filterAndOrder(mapping) {
  const paid = df.filter(r => String(r[mapping.pay]).trim() === '已支付')
  const effective = paid.filter(r => {
    const st = String(r[mapping.status]).trim()
    const pay = String(r[mapping.pay]).trim()
    if (pay === '未支付') return false
    if (st === '已取消' || st === '用户申请退款') return false
    if (pay === '已退款') return false
    return true
  })
  const withIndex = effective.map((r, i) => ({ ...r, __row__: i + 1 }))
  const lunch = withIndex.filter(r => String(r[mapping.product]).trim() === '明日午餐 x1').sort((a,b) => b.__row__ - a.__row__)
  const dinner = withIndex.filter(r => String(r[mapping.product]).trim() === '明日晚餐 x1').sort((a,b) => b.__row__ - a.__row__)
  return { lunch, dinner }
}

function buildOutput(rows, mapping, start, title, productLabel) {
  const lines = []
  lines.push(`### ${title}（商品信息：${productLabel}，编号从${start}开始）`)
  let curr = start
  for (const r of rows) {
    const addr = splitAddress(r[mapping.addr])
    const note = String(r[mapping.note] ?? '').trim()
    lines.push(String(curr))
    lines.push(addr)
    if (note) lines.push(`（用户备注：${note}）`)
    curr++
  }
  return lines.join('\n')
}

function buildPreview() {
  const mapping = {
    product: $('#colProduct').value,
    pay: $('#colPay').value,
    status: $('#colStatus').value,
    addr: $('#colAddr').value,
    note: $('#colNote').value,
  }
  const { lunch, dinner } = filterAndOrder(mapping)
  const lunchText = buildOutput(lunch, mapping, Number($('#startLunch').value), '一、午餐', '明日午餐 x1')
  const dinnerText = buildOutput(dinner, mapping, Number($('#startDinner').value), '二、晚餐', '明日晚餐 x1')
  return `${lunchText}\n\n${dinnerText}`
}

$('#btnPreview').addEventListener('click', () => {
  try {
    $('#preview').value = buildPreview()
    setStatus('预览已生成')
  } catch (e) {
    setStatus(String(e))
  }
})

$('#btnSend').addEventListener('click', async () => {
  const mapping = {
    product: $('#colProduct').value,
    pay: $('#colPay').value,
    status: $('#colStatus').value,
    addr: $('#colAddr').value,
    note: $('#colNote').value,
  }
  const { lunch, dinner } = filterAndOrder(mapping)
  const lunchText = buildOutput(lunch, mapping, Number($('#startLunch').value), '一、午餐', '明日午餐 x1')
  const dinnerText = buildOutput(dinner, mapping, Number($('#startDinner').value), '二、晚餐', '明日晚餐 x1')
  const text = `${lunchText}\n\n${dinnerText}`
  
  const test = $('#testMode').checked
  const items = []
  const groupLunch = test ? '末' : $('#groupLunch').value.trim()
  const groupDinner = test ? '末' : $('#groupDinner').value.trim()
  if (groupLunch) items.push({ group: groupLunch, text: lunchText })
  if (groupDinner) items.push({ group: groupDinner, text: dinnerText })
  const mi = Number($('#intMin').value), ma = Number($('#intMax').value)
  const res = await window.api.sendOrders(items, mi, ma)
  if (!res.ok) setStatus(res.error)
  else setStatus('发送任务已提交')
})

$('#btnStop').addEventListener('click', async () => {
  await window.api.stopSend()
  setStatus('已请求停止')
})

$('#btnConnect').addEventListener('click', async () => {
  const mode = $('#mode').value
  const res = await window.api.connectBot(mode)
  if (!res.ok) setStatus(res.error)
  else {
    $('#btnConnect').disabled = true
    $('#btnDisconnect').disabled = false
    setStatus('已连接，请扫码登录')
  }
})

$('#btnDisconnect').addEventListener('click', async () => {
  await window.api.disconnectBot()
  $('#btnConnect').disabled = false
  $('#btnDisconnect').disabled = true
  setStatus('已断开')
})

$('#pick').addEventListener('click', async () => {
  const fp = await window.api.chooseFile()
  if (!fp) return
  const resp = await fetch(`file:///${fp}`)
})

const drop = $('#drop')
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('hover') })
drop.addEventListener('dragleave', e => { drop.classList.remove('hover') })
drop.addEventListener('drop', e => {
  e.preventDefault(); drop.classList.remove('hover')
  const file = e.dataTransfer.files[0]
  if (file) loadFile(file)
})

