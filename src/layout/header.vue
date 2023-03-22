<template>
  <div class="header">
    <span class="headerbtn" @click="goback">返回主页</span>
  </div>
</template>

<script setup>
import router from "../router";
const goback = () => {
  let body = document.body;
  body.className = "";
  router.push({ path: "/" });
};
</script>

<style  lang="less">
.header {
  // display: flex;
  width: 100%;
  height: 50px;
  line-height: 50px;
  background-color: #569cd6;
  .headerbtn {
    color: #fff;
    float: right;
    margin-right: 36px;
    cursor: pointer;
  }
}
</style>