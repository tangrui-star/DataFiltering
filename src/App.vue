<script setup>
import HelloWorld from './components/HelloWorld.vue'
</script>

<template>
  <header>
    <div class="wrapper">
      <HelloWorld msg="You did it!" />
    </div>
  </header>
  <main>
    <!-- 左侧导航栏 -->
    <div>左侧导航栏</div>
    <!-- 主内容区 -->
    <div>主内容区</div>
  </main>
</template>

<style scoped>
header {
  line-height: 1.5;
}
</style>
