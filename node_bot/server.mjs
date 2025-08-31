import express from 'express'
import cors from 'cors'
import bodyParser from 'body-parser'
import { WechatyBuilder, log } from 'wechaty'
import { PuppetWeChat } from 'wechaty-puppet-wechat'

const app = express()
app.use(cors())
app.use(bodyParser.json({ limit: '5mb' }))

const PORT = process.env.WX_BOT_PORT ? Number(process.env.WX_BOT_PORT) : 8788

let bot = null
let sending = false
let stopFlag = false

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

async function ensureBot(puppet, token) {
  if (bot) return bot
  const options = {}
  if (puppet === 'service') {
    options.puppet = 'wechaty-puppet-service'
    options.puppetOptions = { token }
  } else {
    options.puppet = 'wechaty-puppet-wechat'
  }
  bot = WechatyBuilder.build(options)
  bot.on('scan', (qrcode, status) => {
    log.info('BOT', `Scan QR Code to login: ${status} ${qrcode}`)
  })
  bot.on('login', (user) => log.info('BOT', `${user.name()} login`))
  bot.on('logout', (user) => log.info('BOT', `${user.name()} logout`))
  await bot.start()
  return bot
}

app.post('/connect', async (req, res) => {
  try {
    const { puppet = 'wechat', token } = req.body || {}
    await ensureBot(puppet, token)
    return res.json({ ok: true })
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e) })
  }
})

app.post('/disconnect', async (req, res) => {
  try {
    if (bot) {
      await bot.stop()
      bot = null
    }
    return res.json({ ok: true })
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e) })
  }
})

app.get('/status', async (req, res) => {
  return res.json({ ok: true, hasBot: !!bot, sending, stopFlag })
})

app.post('/stop', async (req, res) => {
  stopFlag = true
  return res.json({ ok: true })
})

app.post('/send', async (req, res) => {
  try {
    if (!bot) return res.status(400).json({ ok: false, error: 'bot not connected' })
    const { items = [], intervalMin = 1.0, intervalMax = 1.5 } = req.body || {}
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ ok: false, error: 'empty items' })
    }
    if (sending) return res.status(409).json({ ok: false, error: 'sending in progress' })
    sending = true
    stopFlag = false
    ;(async () => {
      try {
        for (let i = 0; i < items.length; i++) {
          if (stopFlag) break
          const { group, text } = items[i]
          const room = await bot.Room.find({ topic: group })
          if (!room) continue
          const chunks = chunkText(String(text))
          for (let n = 0; n < chunks.length; n++) {
            if (stopFlag) break
            await room.say(chunks[n])
            if (n < chunks.length - 1) {
              const d = intervalMin + Math.random() * (intervalMax - intervalMin)
              await new Promise((r) => setTimeout(r, Math.floor(d * 1000)))
            }
          }
          if (i < items.length - 1) {
            const d = intervalMin + Math.random() * (intervalMax - intervalMin)
            await new Promise((r) => setTimeout(r, Math.floor(d * 1000)))
          }
        }
      } catch (e) {
        log.error('SEND', String(e))
      } finally {
        sending = false
        stopFlag = false
      }
    })()
    return res.json({ ok: true })
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e) })
  }
})

app.listen(PORT, () => {
  console.log(`bot server listening at http://127.0.0.1:${PORT}`)
})

