import { Router } from 'express'
import { getContactsController } from '../controllers/contacts.controller'

const router = Router()
router.get('/', getContactsController)

export default router
