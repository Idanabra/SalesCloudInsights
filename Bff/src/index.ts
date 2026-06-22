import 'dotenv/config'
import express from 'express'
import cors from 'cors'
import opportunitiesRouter from './routes/opportunities.routes'
import saveEmailRouter     from './routes/saveEmail.routes'

const app  = express()
const PORT = process.env.PORT || 3000

app.use(cors())
app.use(express.json({ limit: '10mb' }))

app.use('/api/opportunities', opportunitiesRouter)
app.use('/api/save-email',    saveEmailRouter)

app.listen(PORT, () => console.log(`[BFF] running on http://localhost:${PORT}`))
