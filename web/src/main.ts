import { createApp } from 'vue'
import App from './App.vue'
import vuetify from './plugins/vuetify'
import i18n from './plugins/i18n'

function mountApp() {
  createApp(App).use(vuetify).use(i18n).mount('#app')
}

if (typeof Office !== 'undefined') {
  Office.onReady(() => mountApp())
} else {
  mountApp()
}
