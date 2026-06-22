import axios from 'axios'
import dotenv from 'dotenv'
dotenv.config()

const sapClient = axios.create({
  baseURL: process.env.SAP_BASE_URL,
  headers: {
    Accept:         'application/json',
    'Content-Type': 'application/json',
  },
  auth: {
    username: process.env.SAP_USER || '',
    password: process.env.SAP_PASS || '',
  },
})

export default sapClient
