import { Request, Response } from 'express'
import { getOpportunities } from '../services/opportunities.service'

export async function getOpportunitiesController(req: Request, res: Response) {
  const search       = String(req.query.search       ?? '')
  const salesCycleId = String(req.query.salesCycleId ?? '')
  try {
    res.json(await getOpportunities(search, salesCycleId))
  } catch (err: any) {
    res.status(500).json({ error: err.message })
  }
}
