import axios from 'axios'

const apiBase = import.meta.env.VITE_API_BASE_URL || '/api'

export function getAccountId(): string {
  const p = new URLSearchParams(location.search)
  return (
    p.get('accountId') ||
    p.get('account_id') ||
    p.get('AccountID') ||
    p.get('accountid') ||
    ''
  )
}

export interface Contact {
  id: string
  firstName: string
  lastName: string
  phone: string
  functionCode: string
}

export async function fetchContacts(accountId: string): Promise<Contact[]> {
  const res = await axios.get(`${apiBase}/contacts`, {
    params: { accountId }
  })
  return res.data
}
