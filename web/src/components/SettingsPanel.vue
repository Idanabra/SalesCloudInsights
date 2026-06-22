<template>
  <v-card flat class="pa-4">
    <v-card-title class="text-subtitle-1 font-weight-bold mb-2">
      {{ $t('settings') }}
    </v-card-title>

    <v-text-field
      v-model="form.salesCycleId"
      :label="$t('salesCycleId')"
      :hint="$t('salesCycleHint')"
      persistent-hint
      clearable
      class="mb-4"
      style="direction: ltr"
    />

    <div class="d-flex gap-2 justify-end">
      <v-btn variant="text" @click="$emit('cancel')">{{ $t('cancel') }}</v-btn>
      <v-btn color="primary" @click="save">{{ $t('saveSettings') }}</v-btn>
    </div>
  </v-card>
</template>

<script setup lang="ts">
import { reactive, onMounted } from 'vue'

const STORAGE_KEY = 'outlook-addin-settings'

interface Settings { salesCycleId: string }

const emit = defineEmits<{ (e: 'cancel'): void; (e: 'saved', s: Settings): void }>()

const form = reactive<Settings>({ salesCycleId: '' })

onMounted(() => {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? '{}') as Partial<Settings>
    form.salesCycleId = saved.salesCycleId ?? ''
  } catch { /* ignore */ }
})

function save() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ salesCycleId: form.salesCycleId }))
  emit('saved', { salesCycleId: form.salesCycleId })
}
</script>
