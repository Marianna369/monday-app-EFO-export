<script lang="ts" setup>
import { ref } from 'vue';

const isLoadingExport = ref(false);
const errorMessage = ref<string | null>(null);

async function exportToExcel() {
  isLoadingExport.value = true;
  errorMessage.value = null;

  try {
    const response = await fetch('/api/excel_export', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        boardId: Number(import.meta.env.VITE_TABLE_ID),
        statusColumnId: import.meta.env.VITE_COLUMN_ID_STATUS,
        allowedStatus: 'Új belépő',
        targetStatus: 'Adatbázis',
        columnIds: [
          {
            id: import.meta.env.VITE_COLUMN_ID_SZULETESI_IDO,
            label: 'Születési dátum'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_SZULETESI_HELY,
            label: 'Születési hely'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_SZULETESI_NEV,
            label: 'Születési név'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_ANYJA_NEVE,
            label: 'Anyja neve'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_LAKCIM,
            label: 'Lakcím'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_ALLAMPOLGARSAG,
            label: 'Állampolgárság'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_TAJSZAM,
            label: 'Tajszám'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_ADOAZONOSITO,
            label: 'Adóazonosító'
          },
          {
            id: import.meta.env.VITE_COLUMN_ID_BANKSZAMLASZAM,
            label: 'Bankszámlaszám'
          }
        ]
      })
    });

    if (!response.ok) {
      throw new Error('Az exportálás sikertelen.');
    }

    // ⬇️ IFRAME-BIZTOS EXCEL LETÖLTÉS
    const arrayBuffer = await response.arrayBuffer();
    const blob = new Blob([arrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // ❗ NEM adjuk meg a fájlnevet → backend dönti el
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

  } catch (err) {
    console.error(err);
    errorMessage.value = 'Hiba történt az export során.';
  } finally {
    isLoadingExport.value = false;
  }
}
</script>

<template>
  <div class="fill-height d-flex align-center justify-center flex-column">
    <v-btn
      color="primary"
      size="x-large"
      :loading="isLoadingExport"
      @click="exportToExcel"
    >
      Excel export
    </v-btn>

    <v-alert
      v-if="errorMessage"
      type="error"
      class="mt-4"
      density="compact"
    >
      {{ errorMessage }}
    </v-alert>
  </div>
</template>
