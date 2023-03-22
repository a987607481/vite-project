<template >
  <div class="box">
    <div v-for="(item, index) in list" :key="index" class="item">
      <router-link :to="{ path: item.router }">
        {{ item.name }}
      </router-link>
    </div>
  </div>
</template>

<script setup>
import { ref } from "vue";

const list = ref([
  { name: "导出excel", router: "/excel" },
  { name: "炫酷的按钮", router: "/btnWrapper" },
  { name: "粘性气泡效果", router: "/codePen" },
  { name: "拖拽效果", router: "/dragstart" },
  { name: "JSON数据转驼峰", router: "/hump" },
]);
</script>


<style scoped lang="less">
.box {
  width: 80%;
  height: 600px;
  margin: 150px auto;
  border: 1px solid #ccc;
  border-radius: 8px;
  display: flex;
  .item {
    padding: 0 20px;
    height: 32px;
    line-height: 32px;
    border: 1px solid #ccc;
    border-radius: 6px;
    margin: 6px;
    &:hover {
      background-color: #f9cc9d;
      cursor: pointer;
    }
  }
}
</style>
