<template>
  <div class="container">
    <h1 class="text-center">关于PDF转Word工具</h1>
    
    <el-card class="about-card">
      <h2>项目介绍</h2>
      <p>PDF转Word工具是一个简单易用的在线转换工具，可以帮助您快速将PDF文件转换为可编辑的Word文档。</p>
      
      <h2>主要功能</h2>
      <ul>
        <li>支持拖拽上传PDF文件</li>
        <li>保留原PDF中的文本和图片</li>
        <li>实时显示转换进度</li>
        <li>转换完成后提供下载链接</li>
      </ul>
      
      <h2>技术栈</h2>
      <h3>前端</h3>
      <ul>
        <li>Vue 3</li>
        <li>Element Plus</li>
        <li>Axios</li>
        <li>FileSaver.js</li>
      </ul>
      
      <h3>后端</h3>
      <ul>
        <li>Python 3.13.2</li>
        <li>FastAPI</li>
        <li>pdfplumber</li>
        <li>python-docx</li>
        <li>pdf2image</li>
      </ul>
      
      <h2>使用说明</h2>
      <ol>
        <li>在首页拖拽或点击上传PDF文件</li>
        <li>点击"开始转换"按钮</li>
        <li>等待转换完成</li>
        <li>点击"下载Word文档"按钮保存文件</li>
      </ol>
      
      <h2>注意事项</h2>
      <ul>
        <li>文件大小限制为500MB</li>
        <li>仅支持PDF格式文件</li>
        <li>转换后的文件将保留24小时，请及时下载</li>
      </ul>
    </el-card>
    
    <div class="text-center mt-20">
      <el-button type="primary" @click="$router.push('/')">返回首页</el-button>
    </div>
  </div>
</template>

<script setup>
// 组件逻辑
</script>

<style scoped>
.about-card {
  margin: 20px 0;
  max-width: 800px;
  margin-left: auto;
  margin-right: auto;
}

h2 {
  margin-top: 20px;
  margin-bottom: 10px;
  color: #409eff;
}

h3 {
  margin-top: 15px;
  margin-bottom: 10px;
}

ul, ol {
  padding-left: 20px;
}

li {
  margin-bottom: 5px;
}
</style> 