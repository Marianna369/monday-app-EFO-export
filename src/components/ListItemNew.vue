<script setup lang="ts">
import { reactive } from 'vue';
// import MondayApi from '@/plugins/MondayApi';

const props = defineProps<{
    structure: any
}>();

const state = reactive({
    isLoading: false,
    dialog: false,
    dialogTitle: '',
    dialogText: ''
});

const onExportClick = async () => {
    try {
        state.isLoading = true;

        // boardId a structure-ből
        const boardId = props.structure.BoardId || import.meta.env.VITE_TABLE_ID;

        // Backend hívás — Excel letöltése
        const response = await fetch(
            `${import.meta.env.VITE_EXPORT_API}/api/export?boardId=${boardId}`
        );

        if (!response.ok) {
            throw new Error("Az exportálás sikertelen.");
        }

        // Excel blob letöltése
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = `export_${boardId}.xlsx`;
        a.click();

        state.dialog = true;
        state.dialogTitle = "Sikeres export";
        state.dialogText = "Az Excel fájl letöltése megtörtént.";

    } catch (err) {
        state.dialog = true;
        state.dialogTitle = "Hiba";
        state.dialogText = err instanceof Error ? err.message : "Ismeretlen hiba";

    } finally {
        state.isLoading = false;
    }
};
</script>


<template>
    <div class="pa-6">

        <v-row>
            <v-col cols="6">
                <v-btn color="blue" size="x-large" :loading="state.isLoading"
                       prepend-icon="mdi-file-excel" @click="onExportClick">
                    Excel export
                </v-btn>

                <v-dialog v-model="state.dialog" max-width="500">
                    <v-card>
                        <v-card-title>{{ state.dialogTitle }}</v-card-title>
                        <v-card-text>{{ state.dialogText }}</v-card-text>
                        <v-card-actions>
                            <v-spacer />
                            <v-btn color="primary" @click="state.dialog = false">OK</v-btn>
                        </v-card-actions>
                    </v-card>
                </v-dialog>
            </v-col>
        </v-row>

    </div>
</template>
