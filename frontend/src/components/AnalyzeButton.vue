<template>
  <div>
    <button @click="startAnalysis">Запустить анализ</button>
    <div v-if="taskId">
      <a :href="`http://127.0.0.1:8000/download/${taskId}`" target="_blank">
        Скачать результат
      </a>
    </div>
  </div>
</template>

<script>
export default {
  data() {
    return { taskId: null };
  },
  methods: {
    async startAnalysis() {
      const res = await fetch("http://127.0.0.1:8000/analyze", { method: "POST" });
      const data = await res.json();
      this.taskId = data.task_id;
    }
  }
};
</script>

