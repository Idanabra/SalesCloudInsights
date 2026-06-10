import { createApp } from 'vue'
import { createVuetify } from 'vuetify'
import { createI18n } from 'vue-i18n'
import App from './App.vue'
import vuetify from './plugins/vuetify'
import i18n from './plugins/i18n'

const app = createApp(App)
app.use(vuetify)
app.use(i18n)
app.mount('#app')
