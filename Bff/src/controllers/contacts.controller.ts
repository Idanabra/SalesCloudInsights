import { Request, Response } from 'express'
import { getContacts } from '../services/contacts.service'

export async function getContactsController(req: Request, res: Response) {
  const { accountId } = req.query
  if (!accountId) {
    return res.status(400).json({ error: 'accountId is required' })
  }
  try {
    const contacts = await getContacts(accountId as string)
    res.json(contacts)
  } catch (err: any) {
    console.error('Error fetching contacts:', err.message)
    res.status(500).json({ error: err.message })
  }
}
