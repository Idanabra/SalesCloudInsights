<template>
  <v-container fluid class="pa-3" dir="rtl" style="font-family: 'Segoe UI', Arial, sans-serif;">
    <!-- Header -->
    <v-row class="mb-2" align="center">
      <v-col>
        <h2 class="text-h6 font-weight-bold" style="color: #354A5E;">
          {{ t('title') }}
        </h2>
      </v-col>
    </v-row>

    <!-- Loading -->
    <v-row v-if="loading" justify="center" class="mt-6">
      <v-col cols="auto">
        <v-progress-circular indeterminate color="primary" size="40" />
        <div class="text-body-2 mt-2 text-center">{{ t('loading') }}</div>
      </v-col>
    </v-row>

    <!-- Error -->
    <v-alert
      v-else-if="error"
      type="error"
      variant="tonal"
      class="mb-3"
      :text="error"
    />

    <!-- No account ID -->
    <v-alert
      v-else-if="!accountId"
      type="warning"
      variant="tonal"
      class="mb-3"
      :text="t('noAccountId')"
    />

    <!-- Content -->
    <template v-else>
      <!-- בעלי תפקיד section -->
      <v-card class="mb-4" rounded="lg">
        <v-card-title
          class="text-subtitle-1 font-weight-bold py-2 px-3"
          style="background-color: #0070F2; color: white;"
        >
          {{ t('roleHolders') }}
          <v-chip size="x-small" class="mr-2" color="white" variant="outlined">
            {{ roleHolders.length }}
          </v-chip>
        </v-card-title>

        <v-table density="compact" v-if="roleHolders.length > 0">
          <thead>
            <tr style="background-color: #f0f4f8;">
              <th class="text-right">{{ t('firstName') }}</th>
              <th class="text-right">{{ t('lastName') }}</th>
              <th class="text-right">{{ t('phone') }}</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="c in roleHolders" :key="c.id">
              <td class="text-right">{{ c.firstName }}</td>
              <td class="text-right">{{ c.lastName }}</td>
              <td class="text-right">
                <a v-if="c.phone" :href="`tel:${c.phone}`" style="color: #0070F2; text-decoration: none;">
                  {{ c.phone }}
                </a>
                <span v-else style="color: #999;">—</span>
              </td>
            </tr>
          </tbody>
        </v-table>

        <v-card-text v-else class="text-body-2 text-medium-emphasis text-center py-3">
          {{ t('noContacts') }}
        </v-card-text>
      </v-card>

      <!-- דיירים section -->
      <v-card rounded="lg">
        <v-card-title
          class="text-subtitle-1 font-weight-bold py-2 px-3"
          style="background-color: #354A5E; color: white;"
        >
          {{ t('tenants') }}
          <v-chip size="x-small" class="mr-2" color="white" variant="outlined">
            {{ tenants.length }}
          </v-chip>
        </v-card-title>

        <v-table density="compact" v-if="tenants.length > 0">
          <thead>
            <tr style="background-color: #f0f4f8;">
              <th class="text-right">{{ t('firstName') }}</th>
              <th class="text-right">{{ t('lastName') }}</th>
              <th class="text-right">{{ t('phone') }}</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="c in tenants" :key="c.id">
              <td class="text-right">{{ c.firstName }}</td>
              <td class="text-right">{{ c.lastName }}</td>
              <td class="text-right">
                <a v-if="c.phone" :href="`tel:${c.phone}`" style="color: #0070F2; text-decoration: none;">
                  {{ c.phone }}
                </a>
                <span v-else style="color: #999;">—</span>
              </td>
            </tr>
          </tbody>
        </v-table>

        <v-card-text v-else class="text-body-2 text-medium-emphasis text-center py-3">
          {{ t('noContacts') }}
        </v-card-text>
      </v-card>
    </template>
  </v-container>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from 'vue'
import { useI18n } from 'vue-i18n'
import { fetchContacts, getAccountId, type Contact } from '../services/sapApi'

const { t } = useI18n()

const accountId = ref('')
const contacts = ref<Contact[]>([])
const loading = ref(false)
const error = ref('')

const TENANT_FUNCTION_CODE = '007'

const roleHolders = computed(() =>
  contacts.value.filter(c => c.functionCode !== TENANT_FUNCTION_CODE)
)

const tenants = computed(() =>
  contacts.value.filter(c => c.functionCode === TENANT_FUNCTION_CODE)
)

onMounted(async () => {
  accountId.value = getAccountId()
  if (!accountId.value) return

  loading.value = true
  error.value = ''
  try {
    contacts.value = await fetchContacts(accountId.value)
  } catch (e: any) {
    error.value = e?.response?.data?.error || e?.message || t('error')
  } finally {
    loading.value = false
  }
})
</script>
