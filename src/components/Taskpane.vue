<script setup lang="ts">
import { inject, computed, type ShallowRef } from 'vue'
import { type VueExcelGlobalState } from '@/state'

const vueExcel = inject('vueExcel') as VueExcelGlobalState
const worksheet = inject('vueExcel.scope.worksheet') as ShallowRef<Excel.Worksheet>

const isActive = computed(() => {
  if (!worksheet.value) return false
  if (!vueExcel.activeWorksheet.value) return false

  return worksheet.value.name == vueExcel.activeWorksheet.value.name
})
</script>

<template>
  <div v-if="isActive">
    <slot></slot>
  </div>
</template>
