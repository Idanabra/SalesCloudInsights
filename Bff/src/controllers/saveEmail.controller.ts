import { Request, Response } from 'express'
import { saveEmailToSAP } from '../services/saveEmail.service'

export async function saveEmailController(req: Request, res: Response) {
  const { emailData, oppId, oppDisplayId, oppName } = req.body
  if (!emailData || !oppId) {
    return res.status(400).json({ error: 'emailData and oppId are required' })
  }
  try {
    const result = await saveEmailToSAP(emailData, oppId, oppDisplayId, oppName)
    res.json(result)
  } catch (err: any) {
    res.status(500).json({ error: err.message })
  }
}
