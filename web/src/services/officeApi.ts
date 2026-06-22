export interface EmailData {
  subject:      string
  messageId:    string | null
  from:         string
  to:           string[]
  cc:           string[]
  plainContent: string
  richTextBody: string | null
  sentOn:       string
  direction:    'INBOUND' | 'OUTBOUND'
}

export async function readOutlookEmail(): Promise<EmailData> {
  const item = Office.context.mailbox.item

  const [plainContent, rawHtml] = await Promise.all([
    getBodyAsync('text'),
    getBodyAsync('html'),
  ])

  const bodyOnly   = extractBodyContent(rawHtml)
  const resolvedHtml = await resolveCidImages(bodyOnly)

  const fromEmail    = item.from?.emailAddress ?? ''
  const mailboxEmail = Office.context.mailbox.userProfile?.emailAddress ?? ''
  const direction: 'INBOUND' | 'OUTBOUND' =
    fromEmail.toLowerCase() === mailboxEmail.toLowerCase() ? 'OUTBOUND' : 'INBOUND'

  return {
    subject:      item.subject ?? '',
    messageId:    item.internetMessageId ?? null,
    from:         fromEmail,
    to:           emailList(item.to),
    cc:           emailList(item.cc),
    plainContent: cleanPlainText(plainContent),
    richTextBody: resolvedHtml || null,
    sentOn:       item.dateTimeCreated
      ? new Date(item.dateTimeCreated).toISOString()
      : new Date().toISOString(),
    direction,
  }
}

function getBodyAsync(type: 'text' | 'html'): Promise<string> {
  return new Promise(resolve => {
    const coercion = type === 'text'
      ? Office.CoercionType.Text
      : Office.CoercionType.Html
    Office.context.mailbox.item.body.getAsync(coercion, {}, (r: any) => {
      resolve(r.status === 'succeeded' ? r.value ?? '' : '')
    })
  })
}

function extractBodyContent(html: string): string {
  const match = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i)
  return match ? match[1] : html
}

async function resolveCidImages(html: string): Promise<string> {
  if (!html || !html.includes('cid:')) return html
  const item = Office.context.mailbox.item
  if (typeof item.getAttachmentContentAsync !== 'function') return html

  const inlineAtts = (item.attachments ?? []).filter((a: any) => a.isInline)
  if (!inlineAtts.length) return html

  const fetched = await Promise.all(
    inlineAtts.map((att: any) => new Promise(resolve => {
      item.getAttachmentContentAsync(att.id, (r: any) => {
        resolve(r.status === 'succeeded' ? { att, value: r.value } : null)
      })
    }))
  )

  let out = html
  for (const r of fetched as any[]) {
    if (!r?.value?.content) continue
    const dataUri  = `data:${r.att.contentType};base64,${r.value.content}`
    const safeName = r.att.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    out = out.replace(
      new RegExp(`src=["']cid:[^"']*${safeName}[^"']*["']`, 'gi'),
      () => `src="${dataUri}"`
    )
  }
  return out
}

function cleanPlainText(text: string): string {
  return text
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/[ \t]*<(?:https?|mailto|tel)[^>]*>/g, '')
    .replace(/<(?:https?|mailto|tel)[^>]*>/g, '')
    .replace(/\r\n/g, '\n')
    .replace(/^[ \t]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function emailList(arr: any[]): string[] {
  if (!Array.isArray(arr)) return []
  return arr.map(r => r?.emailAddress ?? '').filter(Boolean)
}
