<script setup>
import { ref, onMounted, reactive } from 'vue';
import { read, utils } from "https://cdn.sheetjs.com/xlsx-0.18.11/package/xlsx.mjs"

const fileInput = ref();
const result = ref('');

function readWorkbookFromLocalFile(file, callback) {
  var reader = new FileReader();
  reader.onload = function (e) {
    var data = e.target.result;
    var workbook = read(data, { type: 'binary' });
    if (callback) callback(workbook);
  };
  reader.readAsBinaryString(file);
}

const state = reactive({ row: 100, column: 'B' });

onMounted(() => {
  fileInput.value.addEventListener('change', (e) => {
    const file = e.target.files[0];
    readWorkbookFromLocalFile(file, (workbook) => {
      const worksheet = workbook['Sheets'];
      const content = worksheet['Sheet1'];
      let str = '';
      for (let i = 2; i < state.row; i++) {
        str += `'${content['A' + i].v.trim()}': '${content[state.column + '' + i].v.trim()}', \n`;
      }
      result.value = str;
    });
  });
});
</script>

<template>
  <input type='file' ref='fileInput' />
  <input v-model="state.row" />
  <input v-model="state.column" />
  <br />
  <textarea :style="{width: '1000px', height: '700px'}">{{result}}</textarea>
</template>
