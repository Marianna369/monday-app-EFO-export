<script lang="ts" setup>
import MondayApi from '@/plugins/MondayApi';
import { ColumnType, MainBoardItem, MainBoardStructure } from '@/plugins/types';
import { onMounted, reactive } from 'vue';
import ExcelExport from "./ExcelExport.vue";

const state = reactive({
    structure: null as MainBoardStructure,
    Item: null as MainBoardItem
});

MondayApi.init.setSearchColumn(import.meta.env.VITE_COLUMN_ID_NAME);

onMounted(async () => {
    state.structure = await MondayApi.getMainBoardStructure(import.meta.env.VITE_TABLE_ID);
    state.Item = {
        Id: null as number,
        Name: {ColumnId: 'VITE_COLUMN_ID_NAME', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Adoazonosito: {ColumnId: 'VITE_COLUMN_ID_ADOAZONOSITO', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Szuletesi_ido: {ColumnId: 'VITE_COLUMN_ID_SZULETESI_IDO', ColumnValue: null as Date, ColumnType: ColumnType.Date, ColumnValid: false},
        Szuletesi_hely: {ColumnId: 'VITE_COLUMN_ID_SZULETESI_HELY', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        EFO_igenylo: {ColumnId: 'VITE_COLUMN_ID_EFO_IGENYLO', ColumnValue:  {caption: "", value: ""}, ColumnType: ColumnType.User, ColumnValid: false},
        Munkakor: {ColumnId: 'VITE_COLUMN_ID_MUNKAKOR', ColumnValue:  {caption: "", value: ""}, ColumnType: ColumnType.Dropdown, ColumnValid: false},
        Koltseghely: {ColumnId: 'VITE_COLUMN_ID_KOLTSEGHELY', ColumnValue:  {caption: "", value: ""}, ColumnType: ColumnType.Dropdown, ColumnValid: false},
        Szuletesi_nev: {ColumnId: 'VITE_COLUMN_ID_SZULETESI_NEV', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Anyja_neve: {ColumnId: 'VITE_COLUMN_ID_ANYJA_NEVE', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Lakcim: {ColumnId: 'VITE_COLUMN_ID_LAKCIM', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Allampolgarsag: {ColumnId: 'VITE_COLUMN_ID_ALLAMPOLGARSAG', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Tajszam: {ColumnId: 'VITE_COLUMN_ID_TAJSZAM', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
        Bankszamlaszam: {ColumnId: 'VITE_COLUMN_ID_BANKSZAMLASZAM', ColumnValue: "", ColumnType: ColumnType.String, ColumnValid: false},
    } as MainBoardItem
});
</script>

<template>
  <div class="fill-height">
    <div v-if="state.structure == null" class="fill-height d-flex align-center justify-center">
      <v-progress-circular
        color="purple"
        indeterminate
        :size="100"
        :width="7"
      />
    </div>

    <div v-else class="d-flex justify-center">
      <!-- Itt jelenik meg az export gomb -->
      <ExcelExport />
    </div>
  </div>
</template>
