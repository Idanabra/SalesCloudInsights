import sapClient from '../sapClient'

export interface ContactPerson {
  id: string
  firstName: string
  lastName: string
  phone: string
  functionCode: string
}

function extractPhone(phones: any[]): string {
  if (!phones || !phones.length) return ''
  const mobile = phones.find((p: any) =>
    p.phoneUsage === 'MOBILE' || p.usageCode === '3' || p.usage === 'MOBILE'
  )
  const work = phones.find((p: any) =>
    p.phoneUsage === 'WORK' || p.usageCode === '1' || p.usage === 'WORK'
  )
  const any = phones[0]
  const entry = mobile || work || any
  return entry?.phoneNumber || entry?.number || entry?.phone || ''
}

function mapContact(raw: any): ContactPerson {
  const person = raw.person || raw.individual || raw
  const firstName =
    person.firstName || person.givenName || raw.firstName || raw.givenName || ''
  const lastName =
    person.lastName || person.familyName || raw.lastName || raw.familyName || ''

  const phones: any[] = raw.phones || raw.phoneNumbers || person.phones || person.phoneNumbers || []
  const phone = typeof phones === 'string' ? phones : extractPhone(phones)

  const functionCode =
    raw.functionCode || raw.contactPersonFunctionCode || raw.function || ''

  const id =
    raw.contactPersonID || raw.id || raw.contactId || raw.objectID || ''

  return { id, firstName, lastName, phone, functionCode }
}

async function getContactsMethod1(accountId: string): Promise<ContactPerson[]> {
  const url = `/sap/c4c/api/v1/account-service/accounts/${accountId}/contact-persons`
  const res = await sapClient.get(url, {
    params: { $top: 200, $expand: 'person,phones' }
  })
  const items: any[] = res.data?.value || res.data?.items || res.data || []
  return items.map(mapContact)
}

async function getContactsMethod2(accountId: string): Promise<ContactPerson[]> {
  const url = '/sap/c4c/api/v1/contact-service/contact-persons'
  const res = await sapClient.get(url, {
    params: {
      $filter: `accountId eq '${accountId}'`,
      $top: 200,
      $expand: 'person,phones'
    }
  })
  const items: any[] = res.data?.value || res.data?.items || res.data || []
  return items.map(mapContact)
}

export async function getContacts(accountId: string): Promise<ContactPerson[]> {
  try {
    return await getContactsMethod1(accountId)
  } catch (err1: any) {
    console.warn('Method 1 failed, trying Method 2:', err1.message)
    try {
      return await getContactsMethod2(accountId)
    } catch (err2: any) {
      console.error('Both methods failed. Method 2 error:', err2.message)
      throw new Error(`Failed to fetch contacts: ${err2.message}`)
    }
  }
}
