import { app, BrowserWindow, ipcMain, dialog, globalShortcut } from 'electron'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { WechatyBuilder, log } from 'wechaty'

let mainWindow = null
let bot = null
let sending = false
let stopFlag = false

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1080,
    height: 760,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      devTools: true,
    }
  })
  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'))
  mainWindow.webContents.on('render-process-gone', (_e, details) => {
    console.error('renderer crashed', details)
  })
  mainWindow.webContents.on('did-fail-load', (_e, errCode, errDesc) => {
    console.error('did-fail-load', errCode, errDesc)
  })

  globalShortcut.register('Control+Shift+S', () => {
    stopFlag = true
    mainWindow.webContents.send('status', '收到停止指令')
  })
}

app.whenReady().then(() => {
  // Compatibility flags and logging
  app.commandLine.appendSwitch('no-sandbox')
  app.commandLine.appendSwitch('disable-gpu')
  app.commandLine.appendSwitch('ignore-gpu-blocklist')
  app.commandLine.appendSwitch('disable-software-rasterizer')
  process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = '1'
  process.env.ELECTRON_ENABLE_LOGGING = '1'
  process.on('uncaughtException', (err) => {
    try { mainWindow?.webContents.send('status', String(err)) } catch {}
    console.error('uncaughtException', err)
  })
  process.on('unhandledRejection', (reason) => {
    try { mainWindow?.webContents.send('status', String(reason)) } catch {}
    console.error('unhandledRejection', reason)
  })
  createWindow()
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('choose-file', async () => {
  const res = await dialog.showOpenDialog(mainWindow, {
    filters: [
      { name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'csv'] },
      { name: 'All Files', extensions: ['*'] }
    ],
    properties: ['openFile']
  })
  if (res.canceled || res.filePaths.length === 0) return null
  return res.filePaths[0]
})

ipcMain.handle('connect-bot', async (e, { mode }) => {
  try {
    if (bot) return { ok: true }
    const options = {}
    if (mode === 'wechat') {
      options.puppet = 'wechaty-puppet-wechat'
    } else {
      options.puppet = 'wechaty-puppet-wechat'
    }
    bot = WechatyBuilder.build(options)
    bot.on('scan', (qrcode, status) => {
      log.info('BOT', `Scan QR Code: ${status}`)
      mainWindow.webContents.send('qr', { qrcode, status })
    })
    bot.on('login', user => mainWindow.webContents.send('status', `${user.name()} 已登录`))
    bot.on('logout', user => mainWindow.webContents.send('status', `${user.name()} 已登出`))
    await bot.start()
    return { ok: true }
  } catch (e) {
    return { ok: false, error: String(e) }
  }
})

ipcMain.handle('disconnect-bot', async () => {
  try {
    if (bot) { await bot.stop(); bot = null }
    return { ok: true }
  } catch (e) {
    return { ok: false, error: String(e) }
  }
})

function chunkText(text, maxLen = 3500) {
  if (text.length <= maxLen) return [text]
  const lines = text.split(/\r?\n/)
  const parts = []
  let current = ''
  for (const line of lines) {
    const add = (current ? '\n' : '') + line
    if ((current + add).length > maxLen) {
      if (current) parts.push(current)
      current = line
    } else {
      current += add
    }
  }
  if (current) parts.push(current)
  return parts
}

ipcMain.handle('send-orders', async (e, { items, intervalMin, intervalMax }) => {
  try {
    if (!bot) throw new Error('Bot 未连接')
    if (sending) throw new Error('发送中')
    sending = true
    stopFlag = false
    for (let i = 0; i < items.length; i++) {
      if (stopFlag) break
      const { group, text } = items[i]
      const room = await bot.Room.find({ topic: group })
      if (!room) {
        mainWindow.webContents.send('status', `未找到群聊：${group}`)
        continue
      }
      const chunks = chunkText(text)
      for (let n = 0; n < chunks.length; n++) {
        if (stopFlag) break
        await room.say(chunks[n])
        mainWindow.webContents.send('status', `已发送 ${group} 第 ${n + 1} 段`)
        if (n < chunks.length - 1) {
          const d = intervalMin + Math.random() * (intervalMax - intervalMin)
          await new Promise(r => setTimeout(r, Math.floor(d * 1000)))
        }
      }
      if (i < items.length - 1) {
        const d = intervalMin + Math.random() * (intervalMax - intervalMin)
        await new Promise(r => setTimeout(r, Math.floor(d * 1000)))
      }
    }
    sending = false
    return { ok: true }
  } catch (e) {
    sending = false
    return { ok: false, error: String(e) }
  }
})

ipcMain.handle('stop-send', async () => {
  stopFlag = true
  return { ok: true }
})

