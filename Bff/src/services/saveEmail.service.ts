import axios from 'axios'
import sapClient from '../sapClient'

function wrapHtmlDocument(body: string): string {
  if (/<!DOCTYPE/i.test(body)) return body
  return `<!DOCTYPE html><html><head><meta charset="utf-8"/></head><body>${body}</body></html>`
}

function plainToHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .split('\n')
    .map(line => `<p>${line || '&nbsp;'}</p>`)
    .join('')
}

export async function saveEmailToSAP(
  emailData:    any,
  oppId:        string,
  oppDisplayId: string,
  oppName:      string
): Promise<{ emailId: string }> {

  const htmlDoc = wrapHtmlDocument(emailData.richTextBody || plainToHtml(emailData.plainContent || ''))

  // Step 1: Create document record → get S3 upload URL
  const { data: docMeta } = await sapClient.post('/sap/c4c/api/v1/document-service/documents', {
    fileName: '__OriginalContent.html',
    category: 'DOCUMENT',
    type:     '11005',
  })

  const docId     = docMeta?.id     ?? docMeta?.value?.id
  const uploadUrl = docMeta?.uploadUrl ?? docMeta?.value?.uploadUrl
  if (!docId || !uploadUrl) throw new Error('document-service: missing id or uploadUrl')

  // Step 2: Upload HTML to pre-signed S3 URL (no auth)
  await axios.put(uploadUrl, htmlDoc, {
    headers: { 'Content-Type': 'text/html; charset=utf-8' },
  })

  // Step 3: POST email WITH richTextDocumentId in the initial body (must be included from start)
  const emailPayload = {
    subject:            emailData.subject     ?? '',
    messageId:          emailData.messageId   ?? null,
    transmissionStatus: 'CREATE',
    direction:          emailData.direction   ?? 'INBOUND',
    dataOrigin:         'MANUAL',
    isDraft:            false,
    isAutoReply:        false,
    isBounce:           false,
    sentOn:             emailData.sentOn      ?? new Date().toISOString(),
    from:               emailData.from        ?? '',
    toRecipients:       emailData.to          ?? [],
    ccRecipients:       emailData.cc          ?? [],
    bccRecipients:      [],
    plainContent:       emailData.plainContent ?? '',
    richTextDocumentId: docId,
    accounts:           [],
    contacts:           [],
    individualCustomers:[],
    employees:          [],
    attachments:        [],
    relatedData: { interactionNumber: null, interactionOutboundNumber: null },
    relatedObjects: [{
      objectId:  oppId,
      displayId: oppDisplayId,
      type:      '72',
      role:      'PREDECESSOR',
    }],
  }

  const { data: emailResult } = await sapClient.post(
    '/sap/c4c/api/v1/email-service/emails',
    emailPayload
  )

  const emailId: string = emailResult?.id ?? emailResult?.value?.id ?? ''
  if (!emailId) throw new Error('email-service: no id in POST response')

  // Step 4: Back-link document to email
  if (docId && emailId) {
    try {
      const getRes = await sapClient.get(`/sap/c4c/api/v1/document-service/documents/${docId}`)
      const etag   = getRes.headers['etag'] || getRes.headers['ETag'] || '*'
      await sapClient.patch(
        `/sap/c4c/api/v1/document-service/documents/${docId}`,
        { hostObjectId: emailId, hostObjectType: '39' },
        { headers: { 'If-Match': etag } }
      )
    } catch (linkErr: any) {
      console.warn('[BFF] document back-link failed (non-fatal):', linkErr.message)
    }
  }

  return { emailId }
}
