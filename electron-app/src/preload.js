import { contextBridge, ipcRenderer } from 'electron'

contextBridge.exposeInMainWorld('api', {
  chooseFile: () => ipcRenderer.invoke('choose-file'),
  connectBot: (mode) => ipcRenderer.invoke('connect-bot', { mode }),
  disconnectBot: () => ipcRenderer.invoke('disconnect-bot'),
  sendOrders: (items, intervalMin, intervalMax) => ipcRenderer.invoke('send-orders', { items, intervalMin, intervalMax }),
  stopSend: () => ipcRenderer.invoke('stop-send'),
  onStatus: (cb) => ipcRenderer.on('status', (_, msg) => cb(msg)),
  onQR: (cb) => ipcRenderer.on('qr', (_, payload) => cb(payload)),
})

