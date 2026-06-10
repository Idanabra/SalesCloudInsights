import { createI18n } from 'vue-i18n'

const messages = {
  he: {
    title: 'אנשי קשר',
    roleHolders: 'בעלי תפקיד',
    tenants: 'דיירים',
    firstName: 'שם פרטי',
    lastName: 'שם משפחה',
    phone: 'טלפון',
    noContacts: 'אין אנשי קשר',
    loading: 'טוען...',
    error: 'שגיאה בטעינת נתונים',
    noAccountId: 'לא סופק מזהה חשבון',
  },
  en: {
    title: 'Contacts',
    roleHolders: 'Role Holders',
    tenants: 'Tenants',
    firstName: 'First Name',
    lastName: 'Last Name',
    phone: 'Phone',
    noContacts: 'No contacts found',
    loading: 'Loading...',
    error: 'Error loading data',
    noAccountId: 'No account ID provided',
  }
}

export default createI18n({
  legacy: false,
  locale: 'he',
  fallbackLocale: 'en',
  messages,
})
