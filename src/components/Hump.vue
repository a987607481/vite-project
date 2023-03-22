<template>
  <div
    style="
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      width: 400px;
      margin: 250px auto;
    "
  >
    <input type="file" accept=".xlsx,.xls" @change="importExcelFile" /> <br />
    <div>
      需要处理的字段名：<input v-model="inputValue" placeholder="将_改成驼峰命名" style="padding: 5px 0" />
    </div>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from "vue";

const inputValue = ref();

onMounted(() => {});
const convertToCamelCase = (str) => {
  let words = str.toLowerCase().split("_");
  for (let i = 1; i < words.length; i++) {
    words[i] = words[i].charAt(0).toUpperCase() + words[i].slice(1);
  }
  return words.join("");
};

const importExcelFile = (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = (event) => {
    const data = event.target.result;
    const workbook = XLSX.read(data, { type: "binary" });
    workbook.SheetNames.forEach(function (sheetName) {
      var XL_row_object = XLSX.utils.sheet_to_row_object_array(
        workbook.Sheets[sheetName]
      );

      if (inputValue.value) {
        XL_row_object.forEach((item) => {
          if (item[inputValue.value]) {
            item[inputValue.value] = convertToCamelCase(item[inputValue.value]);
          }
        });
      }

      downloadJson(XL_row_object, sheetName);
    });
  };

  reader.readAsBinaryString(file);
};
const downloadJson = (jsonData, fileName) => {
  const dataStr = JSON.stringify(jsonData);
  const dataUri =
    "data:application/json;charset=utf-8," + encodeURIComponent(dataStr);
  const downloadLink = document.createElement("a");
  downloadLink.setAttribute("href", dataUri);
  downloadLink.setAttribute("download", fileName);
  document.body.appendChild(downloadLink);
  downloadLink.click();
  document.body.removeChild(downloadLink);
};
</script>

<style scoped lang='less'>
</style>