<template>
  <v-app :rtl="isRtl">
    <OutlookView />
  </v-app>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import OutlookView from './views/OutlookView.vue'

const { locale } = useI18n()
const isRtl = computed(() => locale.value === 'he')
</script>
