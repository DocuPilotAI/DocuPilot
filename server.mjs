import { createServer } from 'https'
import { readFileSync } from 'fs'
import { parse } from 'url'
import next from 'next'
import { homedir } from 'os'
import { join } from 'path'

const dev = process.env.NODE_ENV !== 'production'
const hostname = 'localhost'
const port = 3000

// 准备 Next.js 应用
const app = next({ dev, hostname, port })
const handle = app.getRequestHandler()

app.prepare().then(() => {
  // 读取证书
  const httpsOptions = {
    key: readFileSync(join(homedir(), '.office-addin-dev-certs', 'localhost.key')),
    cert: readFileSync(join(homedir(), '.office-addin-dev-certs', 'localhost.crt')),
  }

  // 创建 HTTPS 服务器
  createServer(httpsOptions, async (req, res) => {
    try {
      const parsedUrl = parse(req.url, true)
      await handle(req, res, parsedUrl)
    } catch (err) {
      console.error('Error occurred handling', req.url, err)
      res.statusCode = 500
      res.end('internal server error')
    }
  })
    .once('error', (err) => {
      console.error(err)
      process.exit(1)
    })
    .listen(port, () => {
      console.log(`> Ready on https://${hostname}:${port}`)
    })
})
