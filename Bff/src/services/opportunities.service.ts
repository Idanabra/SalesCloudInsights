import sapClient from '../sapClient'

export interface Opportunity {
  id:        string
  displayId: string
  name:      string
  owner:     string
}

function oppName(o: any)      { return o.name ?? o.Name ?? o.subject ?? '(Unnamed)' }
function oppDisplayId(o: any) { return String(o.displayId ?? o.DisplayID ?? o.id ?? '') }
function oppUUID(o: any)      { return String(o.id ?? o.ObjectID ?? '') }
function oppOwner(o: any): string {
  const raw = o.OwnerName ?? o.ownerName ?? o.owner
  if (!raw) return ''
  if (typeof raw === 'object') return raw.content ?? raw.name ?? raw.displayName ?? ''
  return String(raw)
}

export async function getOpportunities(search: string, salesCycleId: string): Promise<Opportunity[]> {
  const params: Record<string, string> = {
    '$top':    '30',
    '$select': 'id,displayId,name,OwnerName,ownerName,owner,salesCycleCode',
    '$orderby': 'name asc',
  }

  if (salesCycleId.trim()) {
    params['$filter'] = `salesCycleCode eq '${salesCycleId.trim()}'`
  }
  if (search.trim()) {
    params['$search'] = search.trim()
  }

  const { data } = await sapClient.get('/sap/c4c/api/v1/opportunity-service/opportunities', { params })

  const items: any[] = data?.value ?? (Array.isArray(data) ? data : [])

  const mapped: Opportunity[] = items.map(o => ({
    id:        oppUUID(o),
    displayId: oppDisplayId(o),
    name:      oppName(o),
    owner:     oppOwner(o),
  }))

  return mapped.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }))
}
