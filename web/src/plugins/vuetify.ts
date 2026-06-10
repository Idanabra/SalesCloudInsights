import 'vuetify/styles'
import '@mdi/font/css/materialdesignicons.css'
import { createVuetify } from 'vuetify'
import * as components from 'vuetify/components'
import * as directives from 'vuetify/directives'

export default createVuetify({
  components,
  directives,
  theme: {
    defaultTheme: 'light',
    themes: {
      light: {
        colors: {
          primary: '#0070F2',
          secondary: '#354A5E',
          surface: '#FFFFFF',
          background: '#F5F6F7',
          error: '#BB0000',
          success: '#107E3E',
          warning: '#E9730C',
        }
      }
    }
  },
  defaults: {
    VCard: { elevation: 1 },
    VBtn: { variant: 'flat' },
    VTextField: { variant: 'outlined', density: 'compact' }
  }
})
