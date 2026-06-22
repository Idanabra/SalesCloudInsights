import 'vuetify/styles'
import '@mdi/font/css/materialdesignicons.css'
import { createVuetify } from 'vuetify'
import * as components from 'vuetify/components'
import * as directives from 'vuetify/directives'

export default createVuetify({
  components,
  directives,
  theme: {
    defaultTheme: 'sapLight',
    themes: {
      sapLight: {
        dark: false,
        colors: {
          primary:    '#0070C3',
          secondary:  '#005A9E',
          success:    '#107C10',
          error:      '#A80000',
          warning:    '#E9730C',
          background: '#F3F2F1',
          surface:    '#FFFFFF',
        }
      }
    }
  },
  defaults: {
    VBtn:       { style: 'text-transform:none; font-weight:600; letter-spacing:0;' },
    VTextField: { variant: 'outlined', density: 'compact' },
    VSelect:    { variant: 'outlined', density: 'compact' },
    VTextarea:  { variant: 'outlined', density: 'compact' },
  }
})
