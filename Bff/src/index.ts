import express from 'express'
import cors from 'cors'
import dotenv from 'dotenv'
import contactsRouter from './routes/contacts.routes'

dotenv.config()

const app = express()
const PORT = process.env.PORT || 3000

app.use(cors())
app.use(express.json())

app.get('/health', (_req, res) => res.json({ status: 'ok' }))
app.use('/api/contacts', contactsRouter)

app.listen(PORT, () => {
  console.log(`BFF running on http://localhost:${PORT}`)
})
