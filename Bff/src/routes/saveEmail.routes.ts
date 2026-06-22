import { Router } from 'express'
import { saveEmailController } from '../controllers/saveEmail.controller'

const router = Router()
router.post('/', saveEmailController)
export default router
