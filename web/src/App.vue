<template>
  <v-app :dir="'rtl'" :theme="'light'">
    <v-main>
      <ContactsView />
    </v-main>
  </v-app>
</template>

<script setup lang="ts">
import ContactsView from './views/ContactsView.vue'
</script>
