<template>
  <div class="g-container">
    <div class="g-wrap">
      <div class="g-footer">
        <div class="g-bubble" v-for="(item, index) in 200" :key="index"></div>
      </div>
    </div>
  </div>
</template>

<script setup>
</script>

<style  lang='scss'>
@use "sass:math";
@use "sass:list";
:root {
  --color: #26b4f5;
}
.g-container {
  width: 800px;
  height: 600px;
  margin: 150px auto 0;
  overflow: hidden;
}

.g-wrap {
  width: 100%;
  height: 100%;
  background: #fff;
  filter: contrast(8);
}

.g-footer {
  position: absolute;
  bottom: 0;
  left: 0;
  height: 86px;
  width: 100%;
  background: var(--color);
  // filter: blur(5px);

  &:before {
    content: "";
    position: absolute;
    top: -300px;
    left: 0;
    right: 0;
    bottom: 0;
    z-index: 1;
    backdrop-filter: blur(5px);
  }
}

.g-bubble {
  position: absolute;
  border-radius: 50%;
  background-color: var(--color);
}

@for $i from 0 through 200 {
  $random1: random(2500) + 1500;
  $random2: random(5000);
  .g-bubble:nth-child(#{$i}) {
    $width: 30 + random(100) + px;
    left: #{(random(100)) + "%"};
    top: #{(30 + random(100))}px;
    width: $width;
    height: $width;
    animation: moveToTop
      #{math.div($random1, 1000)}s
      ease-in-out -#{math.div($random2, 1000)}s
      infinite;
  }
}

@keyframes moveToTop {
  90% {
    opacity: 1;
  }
  100% {
    opacity: 0.08;
    transform: translate(-50%, -180px) scale(0.3);
  }
}
</style>