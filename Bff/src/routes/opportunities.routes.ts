import { Router } from 'express'
import { getOpportunitiesController } from '../controllers/opportunities.controller'

const router = Router()
router.get('/', getOpportunitiesController)
export default router
