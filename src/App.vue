<script setup>
import HelloWorld from './components/HelloWorld.vue'
import Header from './components/Header.vue';
import LeftAside from './components/LeftAside.vue';
import './assets/styles/comman.css'
</script>

<template>
  <header>
    <div class="wrapper">    
      <Header></Header>
    </div>
  </header>
  <main>
    <!-- 左侧导航栏 -->
    <div> 
      <LeftAside></LeftAside>
    </div>
    <!-- 主内容区 -->
    <div>
      <HelloWorld msg="You did it!" />
    </div>
  </main>
</template>

<style scoped>
header {
  line-height: 1.5;
}
</style>
