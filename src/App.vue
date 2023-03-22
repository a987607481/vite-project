<template>
  <Header />
  <router-view :key="this.$route.path" />
</template>
<script setup>
import Header from "./layout/header.vue";
</script>
<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  /* background-color: #edf1f4; */
}
body {
  margin: 0;
}
</style>
