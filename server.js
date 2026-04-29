import express from 'express'
import fs from 'node:fs/promises'
import fsSync from 'node:fs'
import path from 'node:path'

const app = express()
const PORT = Number(process.env.PORT || 8787)
/** Слушать все интерфейсы — иначе с других машин по LAN-IP не достучаться */
const HOST = process.env.HOST || '0.0.0.0'
const DB_PATH = path.resolve(process.cwd(), 'server-data', 'project.json')
const DB_HISTORY_DIR = path.resolve(process.cwd(), 'server-data', 'history')
const DIST_PATH = path.resolve(process.cwd(), 'dist')

app.use(express.json({ limit: '10mb' }))

const ensureDbFile = async () => {
  const dir = path.dirname(DB_PATH)
  await fs.mkdir(dir, { recursive: true })
  await fs.mkdir(DB_HISTORY_DIR, { recursive: true })
  try {
    await fs.access(DB_PATH)
  } catch {
    await fs.writeFile(DB_PATH, JSON.stringify({ updatedAt: null, data: null }, null, 2), 'utf8')
  }
}

const readDb = async () => {
  await ensureDbFile()
  const raw = await fs.readFile(DB_PATH, 'utf8')
  return JSON.parse(raw)
}

const writeDb = async (data) => {
  await ensureDbFile()
  const payload = {
    updatedAt: new Date().toISOString(),
    data,
  }
  const serialized = JSON.stringify(payload, null, 2)

  // 1) snapshot history (чтобы можно было вернуть состояние)
  const stamp = new Date().toISOString().replace(/[:.]/g, '-')
  const snapshotPath = path.join(DB_HISTORY_DIR, `project-${stamp}.json`)
  await fs.writeFile(snapshotPath, serialized, 'utf8')

  // 2) atomic write текущей базы
  const tmpPath = `${DB_PATH}.tmp`
  await fs.writeFile(tmpPath, serialized, 'utf8')
  await fs.rename(tmpPath, DB_PATH)

  // 3) keep only latest 200 snapshots
  const files = (await fs.readdir(DB_HISTORY_DIR))
    .filter((name) => name.startsWith('project-') && name.endsWith('.json'))
    .sort()
  const overflow = files.length - 200
  if (overflow > 0) {
    await Promise.all(
      files.slice(0, overflow).map((name) => fs.unlink(path.join(DB_HISTORY_DIR, name))),
    )
  }
}

app.get('/api/health', (_req, res) => {
  res.json({ ok: true })
})

app.get('/api/project', async (_req, res) => {
  try {
    const payload = await readDb()
    res.json(payload)
  } catch (error) {
    res.status(500).json({
      message: 'Не удалось прочитать серверную базу проекта',
      error: error instanceof Error ? error.message : 'Unknown error',
    })
  }
})

app.put('/api/project', async (req, res) => {
  try {
    const nextProject = req.body
    if (!nextProject || typeof nextProject !== 'object') {
      res.status(400).json({ message: 'Тело запроса должно содержать объект проекта' })
      return
    }
    await writeDb(nextProject)
    res.json({ ok: true, updatedAt: new Date().toISOString() })
  } catch (error) {
    res.status(500).json({
      message: 'Не удалось сохранить серверную базу проекта',
      error: error instanceof Error ? error.message : 'Unknown error',
    })
  }
})

// Stable mode: serve built frontend without Vite/HMR.
if (fsSync.existsSync(DIST_PATH)) {
  app.use(express.static(DIST_PATH))
  app.use((_req, res) => {
    res.sendFile(path.join(DIST_PATH, 'index.html'))
  })
} else {
  app.get('/', (_req, res) => {
    res
      .status(503)
      .type('html')
      .send(
        '<p>Сборка не найдена. Выполните в каталоге проекта: <code>npm run build</code>, затем перезапустите сервер.</p>',
      )
  })
}

app.listen(PORT, HOST, () => {
  const mode = fsSync.existsSync(DIST_PATH) ? 'статика из dist' : 'только API — нет папки dist'
  console.log(`Сервер: http://127.0.0.1:${PORT} и по LAN (HOST=${HOST}) — ${mode}`)
})
