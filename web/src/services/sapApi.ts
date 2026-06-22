import axios from 'axios'

const api = axios.create({
  baseURL: import.meta.env.VITE_API_BASE_URL,
})

export interface Opportunity {
  id:        string
  displayId: string
  name:      string
  owner:     string
}

export async function fetchOpportunities(
  search      = '',
  salesCycleId = ''
): Promise<Opportunity[]> {
  const res = await api.get<Opportunity[]>('/opportunities', {
    params: { search, salesCycleId },
  })
  return res.data
}

export async function saveEmail(payload: {
  emailData:    any
  oppId:        string
  oppDisplayId: string
  oppName:      string
}): Promise<{ emailId: string }> {
  const res = await api.post<{ emailId: string }>('/save-email', payload)
  return res.data
}
