<template>
  <div>
    <!-- App bar -->
    <v-app-bar color="primary" density="compact" flat>
      <v-app-bar-title class="text-white font-weight-bold text-body-1">
        <v-icon size="18" class="me-1">mdi-cloud</v-icon>
        {{ $t('title') }}
      </v-app-bar-title>
      <template #append>
        <v-btn icon size="small" @click="refresh" :disabled="loading" title="רענן">
          <v-icon color="white">mdi-refresh</v-icon>
        </v-btn>
        <v-btn icon size="small" @click="showSettings = true" title="הגדרות">
          <v-icon color="white">mdi-cog</v-icon>
        </v-btn>
      </template>
    </v-app-bar>

    <!-- Settings panel -->
    <v-dialog v-model="showSettings" max-width="360">
      <SettingsPanel
        @cancel="showSettings = false"
        @saved="onSettingsSaved"
      />
    </v-dialog>

    <v-main>
      <v-container fluid class="pa-2">

        <!-- Search -->
        <v-text-field
          v-model="search"
          :placeholder="$t('search')"
          prepend-inner-icon="mdi-magnify"
          clearable
          hide-details
          class="mb-2"
          @update:model-value="onSearch"
        />

        <!-- Loading -->
        <div v-if="loading" class="text-center py-4">
          <v-progress-circular indeterminate color="primary" size="24" />
          <div class="text-caption mt-1">{{ $t('loading') }}</div>
        </div>

        <!-- Empty -->
        <div v-else-if="filtered.length === 0" class="text-center text-caption text-medium-emphasis py-4">
          {{ $t('noOpps') }}
        </div>

        <!-- Opportunity list -->
        <v-card
          v-for="opp in filtered"
          :key="opp.id"
          class="mb-2 opp-card"
          :class="{ 'opp-selected': selected?.id === opp.id }"
          variant="outlined"
          rounded="lg"
          @click="select(opp)"
          role="button"
          :aria-selected="selected?.id === opp.id"
        >
          <v-card-text class="pa-3">
            <div class="d-flex justify-space-between align-start">
              <span class="font-weight-medium text-body-2">{{ opp.name }}</span>
              <span class="text-caption text-medium-emphasis ms-2 text-no-wrap">
                {{ $t('id') }}: {{ opp.displayId }}
              </span>
            </div>
            <div v-if="opp.owner" class="text-caption text-medium-emphasis mt-1">
              <v-icon size="12" class="me-1">mdi-account</v-icon>{{ opp.owner }}
            </div>
          </v-card-text>
        </v-card>

        <!-- Status banner -->
        <v-alert
          v-if="status"
          :type="status.type"
          density="compact"
          class="mt-2 text-caption"
          closable
          @click:close="status = null"
        >
          <div>{{ status.title }}</div>
          <div v-if="status.detail" class="mt-1 text-caption">{{ status.detail }}</div>
        </v-alert>

      </v-container>
    </v-main>

    <!-- Save button -->
    <v-footer app class="pa-2">
      <v-btn
        color="primary"
        block
        :disabled="!selected || saving"
        :loading="saving"
        prepend-icon="mdi-email-arrow-right"
        @click="saveEmail"
      >
        {{ $t('save') }}
      </v-btn>
    </v-footer>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from 'vue'
import { useI18n } from 'vue-i18n'
import SettingsPanel from '../components/SettingsPanel.vue'
import { fetchOpportunities, saveEmail as saveEmailApi, type Opportunity } from '../services/sapApi'
import { readOutlookEmail } from '../services/officeApi'

const { t } = useI18n()

const opps        = ref<Opportunity[]>([])
const search      = ref('')
const selected    = ref<Opportunity | null>(null)
const loading     = ref(false)
const saving      = ref(false)
const showSettings = ref(false)
const status      = ref<{ type: 'success' | 'error' | 'info'; title: string; detail?: string } | null>(null)

const STORAGE_KEY = 'outlook-addin-settings'

function getSettings() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) ?? '{}') } catch { return {} }
}

const filtered = computed(() => {
  const q = search.value.toLowerCase().trim()
  if (!q) return opps.value
  return opps.value.filter(o =>
    o.name.toLowerCase().includes(q) || o.displayId.toLowerCase().includes(q)
  )
})

async function load() {
  loading.value = true
  status.value  = null
  try {
    const settings   = getSettings()
    opps.value = await fetchOpportunities(search.value, settings.salesCycleId ?? '')
  } catch (err: any) {
    status.value = { type: 'error', title: t('saveFailed'), detail: err.message }
  } finally {
    loading.value = false
  }
}

function refresh() { load() }
function select(opp: Opportunity) { selected.value = selected.value?.id === opp.id ? null : opp }
function onSearch() { load() }
function onSettingsSaved() { showSettings.value = false; load() }

async function saveEmail() {
  if (!selected.value) return
  saving.value  = true
  status.value  = { type: 'info', title: t('savingToSap') }
  try {
    const emailData = await readOutlookEmail()
    const res = await saveEmailApi({
      emailData,
      oppId:        selected.value.id,
      oppDisplayId: selected.value.displayId,
      oppName:      selected.value.name,
    })
    status.value  = {
      type:   'success',
      title:  t('savedSuccess'),
      detail: `Email ID: ${res.emailId} | ${selected.value.name}`,
    }
    selected.value = null
  } catch (err: any) {
    status.value = { type: 'error', title: t('saveFailed'), detail: err.message }
  } finally {
    saving.value = false
  }
}

onMounted(() => load())
</script>

<style scoped>
.opp-card { cursor: pointer; transition: border-color .15s; }
.opp-card:hover { border-color: #0070C3 !important; }
.opp-selected { border-color: #0070C3 !important; border-width: 2px !important; background: #EBF5FE; }
</style>
