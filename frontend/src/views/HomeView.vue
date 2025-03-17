<template>
  <div class="app-container">
    <!-- 顶部导航栏 -->
    <div class="nav-header">
      <div class="logo">
        <span class="logo-text">文档转换工具</span>
      </div>
      <div class="theme-toggle">
        <el-switch
          v-model="darkMode"
          @change="toggleTheme"
          active-icon="MoonNight"
          inactive-icon="Sunny"
          inline-prompt
        />
      </div>
    </div>

    <!-- 主内容区 -->
    <div class="main-content">
      <!-- 转换类型选择器 -->
      <div class="conversion-selector">
        <el-radio-group v-model="conversionType" size="large" @change="resetState">
          <el-radio-button label="pdf-to-word" class="conversion-type-button">
            <el-icon><document /></el-icon>
            <span>PDF转Word</span>
          </el-radio-button>
          <el-radio-button label="markdown-to-word" class="conversion-type-button">
            <el-icon><edit-pen /></el-icon>
            <span>Markdown转Word</span>
          </el-radio-button>
          <el-radio-button label="word-to-pdf" class="conversion-type-button">
            <el-icon><files /></el-icon>
            <span>Word转PDF</span>
          </el-radio-button>
          <el-radio-button label="pdf-to-markdown" class="conversion-type-button">
            <el-icon><document-copy /></el-icon>
            <span>PDF转Markdown</span>
          </el-radio-button>
        </el-radio-group>
      </div>

      <!-- 转换面板 -->
      <div class="conversion-panel">
        <el-card class="panel-card" :body-style="{ padding: '0px' }">
          <div class="panel-header" :class="conversionTypeClass">
            <div class="header-content">
              <el-icon class="header-icon" :size="24">
                <component :is="conversionTypeIcon" />
              </el-icon>
              <div class="header-text">
                <h2>{{ conversionTypeTitle }}</h2>
                <p>{{ conversionTypeDescription }}</p>
              </div>
            </div>
          </div>

          <div class="panel-body">
            <!-- 文件上传区域 -->
            <div class="upload-area" 
                 :class="{ 'has-file': selectedFile, 'is-dragging': isDragging }" 
                 @dragenter="isDragging = true"
                 @dragleave="isDragging = false"
                 @dragover.prevent
                 @drop="isDragging = false">
              <el-upload
                class="upload-component"
                drag
                :auto-upload="false"
                :show-file-list="false"
                :on-change="handleFileChange"
                :before-upload="beforeUpload"
                :accept="acceptFileTypes"
              >
                <div class="drop-zone" v-if="!selectedFile">
                  <div class="upload-icon-container">
                    <el-icon class="upload-icon"><upload-filled /></el-icon>
                    <div class="upload-pulse"></div>
                  </div>
                  <div class="upload-text">
                    拖拽文件到此处，或 <em>点击上传</em>
                  </div>
                  <div class="upload-tip">
                    {{ uploadTip }}
                  </div>
                </div>
                
                <div v-else class="file-preview">
                  <div class="file-info">
                    <div class="file-icon-container">
                      <el-icon class="file-icon"><component :is="fileTypeIcon" /></el-icon>
                    </div>
                    <div class="file-details">
                      <h3 class="file-name">{{ selectedFile.name }}</h3>
                      <p class="file-size">{{ formatFileSize(selectedFile.size) }}</p>
                    </div>
                    <el-button 
                      type="danger" 
                      circle 
                      size="small" 
                      @click.stop="removeFile"
                      class="remove-file-btn"
                    >
                      <el-icon><delete /></el-icon>
                    </el-button>
                  </div>
                </div>
              </el-upload>
            </div>

            <!-- 转换选项 -->
            <div class="conversion-options" v-if="selectedFile">
              <div class="options-row">
                <div class="action-buttons">
                  <el-button 
                    type="primary" 
                    class="convert-button"
                    :icon="isUploading ? 'Loading' : 'Right'" 
                    :loading="isUploading"
                    @click="uploadFile" 
                    :disabled="isUploading || isCompleted"
                  >
                    <el-icon class="button-icon"><component :is="isUploading ? 'Loading' : 'Right'" /></el-icon>
                    <span>{{ isUploading ? '转换中...' : '开始转换' }}</span>
                  </el-button>
                  
                  <el-button 
                    type="success" 
                    class="download-button"
                    v-if="isCompleted && !hasError" 
                    @click="downloadFile"
                  >
                    <el-icon class="button-icon"><download /></el-icon>
                    <span>下载文件</span>
                  </el-button>
                  
                  <el-button 
                    type="info" 
                    class="refresh-button"
                    v-if="isCompleted || hasError" 
                    @click="refreshCurrentView"
                  >
                    <el-icon class="button-icon"><refresh /></el-icon>
                    <span>刷新</span>
                  </el-button>
                </div>
              </div>
            </div>

            <!-- 转换进度 -->
            <div v-if="isUploading || isCompleted" class="conversion-progress">
              <el-progress 
                :percentage="conversionProgress" 
                :status="progressStatus"
                :stroke-width="8"
                :format="progressFormat"
                class="progress-bar"
              />
              <p class="status-message" v-if="hasError">
                <el-icon class="error-icon"><circle-close /></el-icon>
                {{ errorMessage }}
              </p>
            </div>
          </div>
        </el-card>
      </div>
    </div>

    <!-- 底部信息 -->
    <div class="footer">
      <p>© 2025 文档转换工具 | 所有文件将在下载完成后自动删除</p>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import { UploadFilled, Document, Loading, CircleCheck, CircleClose, Download, Delete, Right, EditPen, Files, DocumentCopy, Refresh } from '@element-plus/icons-vue'
import axios from 'axios'
import { saveAs } from 'file-saver'

// 状态变量
const darkMode = ref(false)
const isDragging = ref(false)
const styleConfig = ref(null)

// 转换类型
const conversionType = ref('pdf-to-word')
const useAdvancedMode = ref(true) // 默认始终使用高级模式，但不显示选项

// 每个转换类型的状态
const conversionStates = ref({
  'pdf-to-word': {
    selectedFile: null,
    isUploading: false,
    isCompleted: false,
    hasError: false,
    errorMessage: '',
    conversionProgress: 0,
    taskId: null,
    progressInterval: null,
    taskInfo: null,
    isDownloaded: false,
    isFileCleanedUp: false
  },
  'markdown-to-word': {
    selectedFile: null,
    isUploading: false,
    isCompleted: false,
    hasError: false,
    errorMessage: '',
    conversionProgress: 0,
    taskId: null,
    progressInterval: null,
    taskInfo: null,
    isDownloaded: false,
    isFileCleanedUp: false
  },
  'word-to-pdf': {
    selectedFile: null,
    isUploading: false,
    isCompleted: false,
    hasError: false,
    errorMessage: '',
    conversionProgress: 0,
    taskId: null,
    progressInterval: null,
    taskInfo: null,
    isDownloaded: false,
    isFileCleanedUp: false
  },
  'pdf-to-markdown': {
    selectedFile: null,
    isUploading: false,
    isCompleted: false,
    hasError: false,
    errorMessage: '',
    conversionProgress: 0,
    taskId: null,
    progressInterval: null,
    taskInfo: null,
    isDownloaded: false,
    isFileCleanedUp: false
  }
})

// 当前转换类型的状态的计算属性
const selectedFile = computed({
  get: () => conversionStates.value[conversionType.value].selectedFile,
  set: (val) => { conversionStates.value[conversionType.value].selectedFile = val }
})
const isUploading = computed({
  get: () => conversionStates.value[conversionType.value].isUploading,
  set: (val) => { conversionStates.value[conversionType.value].isUploading = val }
})
const isCompleted = computed({
  get: () => conversionStates.value[conversionType.value].isCompleted,
  set: (val) => { conversionStates.value[conversionType.value].isCompleted = val }
})
const hasError = computed({
  get: () => conversionStates.value[conversionType.value].hasError,
  set: (val) => { conversionStates.value[conversionType.value].hasError = val }
})
const errorMessage = computed({
  get: () => conversionStates.value[conversionType.value].errorMessage,
  set: (val) => { conversionStates.value[conversionType.value].errorMessage = val }
})
const conversionProgress = computed({
  get: () => conversionStates.value[conversionType.value].conversionProgress,
  set: (val) => { conversionStates.value[conversionType.value].conversionProgress = val }
})
const taskId = computed({
  get: () => conversionStates.value[conversionType.value].taskId,
  set: (val) => { conversionStates.value[conversionType.value].taskId = val }
})
const progressInterval = computed({
  get: () => conversionStates.value[conversionType.value].progressInterval,
  set: (val) => { conversionStates.value[conversionType.value].progressInterval = val }
})
const taskInfo = computed({
  get: () => conversionStates.value[conversionType.value].taskInfo,
  set: (val) => { conversionStates.value[conversionType.value].taskInfo = val }
})
const isDownloaded = computed({
  get: () => conversionStates.value[conversionType.value].isDownloaded,
  set: (val) => { conversionStates.value[conversionType.value].isDownloaded = val }
})
const isFileCleanedUp = computed({
  get: () => conversionStates.value[conversionType.value].isFileCleanedUp,
  set: (val) => { conversionStates.value[conversionType.value].isFileCleanedUp = val }
})

// API基础URL
const apiBaseUrl = import.meta.env.VITE_API_BASE_URL || ''

// 计算属性
const progressStatus = computed(() => {
  if (hasError.value) return 'exception'
  if (isCompleted.value) return 'success'
  return ''
})

const conversionTypeTitle = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word': return 'PDF转Word'
    case 'markdown-to-word': return 'Markdown转Word'
    case 'word-to-pdf': return 'Word转PDF'
    case 'pdf-to-markdown': return 'PDF转Markdown'
    default: return 'PDF转Word'
  }
})

const conversionTypeDescription = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word':
      return '将PDF文件转换为可编辑的Word文档，完整保留原始格式、图片和表格'
    case 'markdown-to-word':
      return '将Markdown文件转换为格式化的Word文档，支持代码高亮和数学公式'
    case 'word-to-pdf':
      return '将Word文档转换为PDF文件，保持布局一致性，适合打印和分享'
    case 'pdf-to-markdown':
      return '将PDF文件转换为Markdown格式，便于在线编辑和版本控制'
    default:
      return '将PDF文件转换为可编辑的Word文档，完整保留原始格式、图片和表格'
  }
})

const acceptFileTypes = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word':
      return '.pdf'
    case 'markdown-to-word':
      return '.md,.markdown'
    case 'word-to-pdf':
      return '.docx,.doc'
    case 'pdf-to-markdown':
      return '.pdf'
    default:
      return '.pdf'
  }
})

const uploadTip = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word':
      return '支持上传PDF文件，最大20MB'
    case 'markdown-to-word':
      return '支持上传Markdown文件，最大20MB'
    case 'word-to-pdf':
      return '支持上传Word文档，最大20MB'
    case 'pdf-to-markdown':
      return '支持上传PDF文件，最大20MB'
    default:
      return '支持上传PDF文件，最大20MB'
  }
})

const conversionTypeIcon = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word': return Document
    case 'markdown-to-word': return EditPen
    case 'word-to-pdf': return Files
    case 'pdf-to-markdown': return DocumentCopy
    default: return Document
  }
})

const fileTypeIcon = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word': return Document
    case 'markdown-to-word': return EditPen
    case 'word-to-pdf': return Files
    case 'pdf-to-markdown': return DocumentCopy
    default: return Document
  }
})

const conversionTypeClass = computed(() => {
  switch (conversionType.value) {
    case 'pdf-to-word':
      return 'header-pdf-word'
    case 'markdown-to-word':
      return 'header-md-word'
    case 'word-to-pdf':
      return 'header-word-pdf'
    case 'pdf-to-markdown':
      return 'header-pdf-md'
    default:
      return 'header-pdf-word'
  }
})

// 进度格式化
const progressFormat = (percentage) => {
  if (isCompleted.value && !hasError.value) {
    return '完成'
  }
  if (hasError.value) {
    return '失败'
  }
  return `${percentage}%`
}

// 加载样式配置
onMounted(async () => {
  try {
    const response = await axios.get(`${apiBaseUrl}/style-config`)
    // 处理样式配置
  } catch (error) {
    console.error('加载样式配置失败:', error)
  }
  
  // 检查系统主题偏好
  if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
    darkMode.value = true
    document.body.classList.add('dark-theme')
  }
})

// 切换主题
const toggleTheme = () => {
  if (darkMode.value) {
    document.body.classList.add('dark-theme')
  } else {
    document.body.classList.remove('dark-theme')
  }
}

// 重置状态
const resetState = () => {
  // 清除轮询定时器
  if (progressInterval.value) {
    clearInterval(progressInterval.value)
    progressInterval.value = null
  }
  
  // 重置当前转换类型的状态
  conversionStates.value[conversionType.value] = {
    selectedFile: null,
    isUploading: false,
    isCompleted: false,
    hasError: false,
    errorMessage: '',
    conversionProgress: 0,
    taskId: null,
    progressInterval: null,
    taskInfo: null,
    isDownloaded: false,
    isFileCleanedUp: false
  }
}

// 文件大小格式化
const formatFileSize = (size) => {
  if (size < 1024) {
    return size + ' B'
  } else if (size < 1024 * 1024) {
    return (size / 1024).toFixed(2) + ' KB'
  } else if (size < 1024 * 1024 * 1024) {
    return (size / (1024 * 1024)).toFixed(2) + ' MB'
  }
  return (size / (1024 * 1024 * 1024)).toFixed(2) + ' GB'
}

// 文件选择处理
const handleFileChange = (file) => {
  selectedFile.value = file.raw
  // 重置状态
  isUploading.value = false
  isCompleted.value = false
  hasError.value = false
  errorMessage.value = ''
  conversionProgress.value = 0
  taskId.value = null
  taskInfo.value = null
  isDownloaded.value = false
  isFileCleanedUp.value = false
  
  if (progressInterval.value) {
    clearInterval(progressInterval.value)
    progressInterval.value = null
  }
}

// 移除文件
const removeFile = (e) => {
  e.stopPropagation()
  resetState()
}

// 文件上传前验证
const beforeUpload = (file) => {
  let isValidType = false
  const maxSize = 20 * 1024 * 1024 // 统一设置为20MB
  
  switch (conversionType.value) {
    case 'pdf-to-word':
      isValidType = file.type === 'application/pdf'
      break
    case 'markdown-to-word':
      isValidType = file.name.endsWith('.md') || file.name.endsWith('.markdown')
      break
    case 'word-to-pdf':
      isValidType = file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
                    file.type === 'application/msword'
      break
    case 'pdf-to-markdown':
      isValidType = file.type === 'application/pdf'
      break
    default:
      isValidType = file.type === 'application/pdf'
  }
  
  const isLtMaxSize = file.size <= maxSize

  if (!isValidType) {
    ElMessage.error(`只能上传${conversionTypeTitle.value.split('转')[0]}文件!`)
    return false
  }
  
  if (!isLtMaxSize) {
    ElMessage.error('文件大小不能超过20MB!')
    return false
  }
  
  return isValidType && isLtMaxSize
}

// 上传文件
const uploadFile = async () => {
  if (!selectedFile.value) {
    ElMessage.warning('请先选择文件')
    return
  }
  
  // 重置状态
  isUploading.value = true
  isCompleted.value = false
  hasError.value = false
  errorMessage.value = ''
  conversionProgress.value = 0
  taskId.value = null
  isDownloaded.value = false
  isFileCleanedUp.value = false
  
  try {
    console.log('开始上传文件:', selectedFile.value.name)
    
    // 创建FormData对象
    const formData = new FormData()
    formData.append('file', selectedFile.value)
    formData.append('conversion_type', conversionType.value)  // 添加转换类型参数
    
    // 添加样式配置（如果有）
    if (styleConfig.value) {
      formData.append('style_config', JSON.stringify(styleConfig.value))
    }
    
    // 发送上传请求
    const response = await axios.post(`${apiBaseUrl}/upload`, formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      }
    })
    
    console.log('上传响应:', response.data)
    
    // 保存任务ID
    taskId.value = response.data.task_id
    
    // 开始轮询任务状态
    startPollingTaskStatus()
    
    ElMessage.success('文件上传成功，开始转换')
  } catch (error) {
    console.error('上传失败:', error)
    isUploading.value = false
    hasError.value = true
    errorMessage.value = error.response?.data?.detail || '上传失败，请重试'
    ElMessage.error(`上传失败: ${errorMessage.value}`)
  }
}

// 轮询任务状态
const startPollingTaskStatus = () => {
  if (progressInterval.value) {
    clearInterval(progressInterval.value)
  }
  
  progressInterval.value = setInterval(async () => {
    try {
      // 如果任务ID为空或文件已清理，停止轮询
      if (!taskId.value || isFileCleanedUp.value) {
        clearInterval(progressInterval.value)
        return
      }
      
      console.log(`获取任务状态: ${taskId.value}`)
      const response = await axios.get(`${apiBaseUrl}/status/${taskId.value}`)
      const taskStatus = response.data
      
      // 保存任务信息
      taskInfo.value = taskStatus
      console.log('任务状态:', taskStatus)
      
      // 更新进度
      conversionProgress.value = taskStatus.progress || 0
      
      // 检查任务状态
      if (taskStatus.status === 'completed') {
        isCompleted.value = true
        isUploading.value = false
        clearInterval(progressInterval.value)
        ElMessage.success('文件转换成功，请点击下载按钮获取文件')
      } else if (taskStatus.status === 'failed') {
        hasError.value = true
        errorMessage.value = taskStatus.error || '转换失败，请重试'
        isUploading.value = false
        clearInterval(progressInterval.value)
      }
    } catch (error) {
      console.error('获取任务状态失败:', error)
      
      // 如果是404错误，说明任务可能已被清理
      if (error.response && error.response.status === 404) {
        console.log('任务不存在，可能已被清理')
        isFileCleanedUp.value = true
        clearInterval(progressInterval.value)
      } else {
        hasError.value = true
        errorMessage.value = '获取任务状态失败，请重试'
        isUploading.value = false
        clearInterval(progressInterval.value)
      }
    }
  }, 1000) // 每秒轮询一次
}

// 下载转换后的文件
const downloadFile = async () => {
  if (!taskId.value) {
    ElMessage.error('无法下载文件，任务ID不存在')
    return
  }
  
  try {
    console.log(`开始下载文件，任务ID: ${taskId.value}`)
    
    // 构建下载URL
    const downloadUrl = `${apiBaseUrl}/download/${taskId.value}`
    console.log('下载URL:', downloadUrl)
    
    const response = await axios.get(downloadUrl, {
      responseType: 'blob'
    })
    
    console.log('下载响应:', response)
    
    // 检查响应内容
    if (response.data.size === 0) {
      throw new Error('下载的文件为空')
    }
    
    // 确定文件扩展名
    let fileExtension = '.docx'
    switch (conversionType.value) {
      case 'pdf-to-word': fileExtension = '.docx'; break
      case 'markdown-to-word': fileExtension = '.docx'; break
      case 'word-to-pdf': fileExtension = '.pdf'; break
      case 'pdf-to-markdown': fileExtension = '.md'; break
    }
    
    // 使用file-saver保存文件
    const originalExt = selectedFile.value.name.substring(selectedFile.value.name.lastIndexOf('.'))
    const filename = selectedFile.value.name.replace(originalExt, fileExtension)
    saveAs(new Blob([response.data]), filename)
    
    // 标记为已下载
    isDownloaded.value = true
    
    // 下载成功后，延迟一段时间再检查任务状态
    setTimeout(async () => {
      try {
        // 尝试获取任务状态，检查任务是否还存在
        await axios.get(`${apiBaseUrl}/status/${taskId.value}`)
      } catch (error) {
        // 如果任务不存在，说明已经被自动清理
        console.log('任务已被自动清理')
        isFileCleanedUp.value = true
        taskId.value = null
        taskInfo.value = null
      }
    }, 4000) // 等待4秒，比后端清理的3秒多一点
    
    ElMessage.success('下载成功，文件将在3秒后自动清理')
  } catch (error) {
    console.error('下载失败:', error)
    
    // 如果是404错误，说明任务可能已被清理
    if (error.response && error.response.status === 404) {
      ElMessage.warning('文件已被清理，请重新上传')
      isFileCleanedUp.value = true
      taskId.value = null
      taskInfo.value = null
    } else {
      ElMessage.error(`下载失败: ${error.message || '请重试'}`)
    }
  }
}

// 刷新当前视图
const refreshCurrentView = () => {
  // 清除轮询定时器
  if (progressInterval.value) {
    clearInterval(progressInterval.value)
    progressInterval.value = null
  }
  
  // 重置当前转换类型的状态
  conversionStates.value[conversionType.value] = {
    selectedFile: null,
    isUploading: false,
    isCompleted: false,
    hasError: false,
    errorMessage: '',
    conversionProgress: 0,
    taskId: null,
    progressInterval: null,
    taskInfo: null,
    isDownloaded: false,
    isFileCleanedUp: false
  }
  
  ElMessage.success('已刷新，可以重新上传文件')
}
</script>

<style scoped>
.app-container {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
  background: linear-gradient(135deg, #1a237e 0%, #0d47a1 50%, #01579b 100%);
  position: relative;
  overflow: hidden;
  z-index: 0;
}

.app-container::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: 
    radial-gradient(circle at 20% 20%, rgba(124, 77, 255, 0.2) 0%, transparent 25%),
    radial-gradient(circle at 80% 80%, rgba(0, 229, 255, 0.2) 0%, transparent 25%),
    url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.05'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
  z-index: -1;
  opacity: 0.8;
  animation: backgroundShift 20s ease-in-out infinite;
}

.app-container::after {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: 
    linear-gradient(45deg, 
      rgba(63, 81, 181, 0.15) 0%,
      rgba(3, 169, 244, 0.15) 100%);
  backdrop-filter: blur(100px);
  z-index: -1;
}

@keyframes backgroundShift {
  0%, 100% {
    background-position: 0% 0%;
  }
  50% {
    background-position: 100% 100%;
  }
}

.nav-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 12px 30px;
  background: rgba(255, 255, 255, 0.1);
  backdrop-filter: blur(20px);
  box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
  position: sticky;
  top: 0;
  z-index: 100;
  border-bottom: 1px solid rgba(255, 255, 255, 0.1);
}

.logo-text {
  font-size: 1.6rem;
  font-weight: 800;
  background: linear-gradient(45deg, #4481eb, #04befe);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  letter-spacing: -0.5px;
  text-shadow: 0 0 20px rgba(0, 229, 255, 0.5);
  animation: glowPulse 2s ease-in-out infinite;
}

@keyframes glowPulse {
  0%, 100% {
    text-shadow: 0 0 20px rgba(0, 229, 255, 0.5);
  }
  50% {
    text-shadow: 0 0 30px rgba(124, 77, 255, 0.8);
  }
}

.main-content {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 12px;
  max-width: 900px;
  margin: 0 auto;
  width: 100%;
}

.conversion-selector {
  margin: 16px 0;
  width: 100%;
  display: flex;
  justify-content: center;
  padding: 0 12px;
}

.conversion-panel {
  width: 100%;
}

.panel-card {
  border-radius: 12px;
  overflow: hidden;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.1);
  transition: all 0.3s ease;
  background: rgba(255, 255, 255, 0.1);
  backdrop-filter: blur(20px);
  border: 1px solid rgba(255, 255, 255, 0.2);
}

.panel-card:hover {
  transform: translateY(-5px);
  box-shadow: 0 12px 30px rgba(0, 0, 0, 0.15);
}

.panel-header {
  background: linear-gradient(45deg, rgba(124, 77, 255, 0.8), rgba(0, 229, 255, 0.8));
  color: white;
  padding: 16px;
}

.panel-header h2 {
  margin: 0;
  font-size: 1.3rem;
}

.panel-header p {
  margin: 6px 0 0;
  opacity: 0.9;
  font-size: 0.85rem;
}

.panel-body {
  padding: 16px;
  background-color: white;
}

.upload-area {
  border: 2px dashed #dcdfe6;
  border-radius: 16px;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  margin-bottom: 20px;
  background: rgba(255, 255, 255, 0.8);
  backdrop-filter: blur(10px);
  box-shadow: 0 8px 30px rgba(0, 0, 0, 0.05);
  position: relative;
  overflow: hidden;
}

.upload-area::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  border-radius: 17px;
  background: linear-gradient(45deg, #4481eb10, #04befe10);
  opacity: 0;
  transition: opacity 0.3s ease;
}

.upload-area:hover::before,
.upload-area.is-dragging::before {
  opacity: 1;
}

.upload-area:hover,
.upload-area.is-dragging {
  border-color: #409eff;
  transform: translateY(-2px);
  box-shadow: 0 12px 40px rgba(64, 158, 255, 0.15);
}

.upload-area.has-file {
  border-style: solid;
  border-color: #67c23a;
  background: rgba(103, 194, 58, 0.05);
}

.upload-component {
  width: 100%;
}

.drop-zone {
  padding: 30px 20px;
  text-align: center;
  cursor: pointer;
  position: relative;
}

.upload-icon-container {
  position: relative;
  display: inline-block;
  margin-bottom: 16px;
}

.upload-icon {
  font-size: 50px;
  color: #409eff;
  filter: drop-shadow(0 4px 6px rgba(64, 158, 255, 0.2));
  animation: float 3s ease-in-out infinite;
}

@keyframes float {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-10px); }
}

.upload-pulse {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  width: 60px;
  height: 60px;
  border-radius: 50%;
  background: rgba(64, 158, 255, 0.1);
  animation: pulse 2s infinite;
}

@keyframes pulse {
  0% {
    transform: translate(-50%, -50%) scale(0.95);
    opacity: 0.6;
  }
  50% {
    transform: translate(-50%, -50%) scale(1.05);
    opacity: 0.3;
  }
  100% {
    transform: translate(-50%, -50%) scale(0.95);
    opacity: 0.6;
  }
}

.upload-text {
  font-size: 16px;
  color: #606266;
  margin-bottom: 12px;
  font-weight: 500;
}

.upload-text em {
  color: #409eff;
  font-style: normal;
  font-weight: 600;
  text-decoration: underline;
  cursor: pointer;
  position: relative;
  padding: 0 4px;
}

.upload-text em::before {
  content: '';
  position: absolute;
  bottom: 0;
  left: 0;
  width: 100%;
  height: 2px;
  background: #409eff;
  transform: scaleX(0);
  transition: transform 0.3s ease;
  transform-origin: right;
}

.upload-text em:hover::before {
  transform: scaleX(1);
  transform-origin: left;
}

.upload-tip {
  font-size: 13px;
  color: #909399;
  background: rgba(144, 147, 153, 0.1);
  padding: 8px 16px;
  border-radius: 8px;
  display: inline-block;
  backdrop-filter: blur(5px);
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.05);
  border: 1px solid rgba(144, 147, 153, 0.2);
}

.file-preview {
  padding: 12px;
}

.file-info {
  display: flex;
  align-items: center;
  position: relative;
  background: rgba(64, 158, 255, 0.05);
  padding: 16px;
  border-radius: 12px;
  backdrop-filter: blur(10px);
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
}

.file-icon-container {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 40px;
  height: 40px;
  background: rgba(64, 158, 255, 0.1);
  border-radius: 8px;
  margin-right: 16px;
}

.file-icon {
  font-size: 20px;
  color: #409eff;
}

.file-details {
  flex: 1;
}

.file-name {
  margin: 0;
  font-size: 14px;
  font-weight: 600;
  color: #303133;
  word-break: break-all;
}

.file-size {
  margin: 4px 0 0;
  font-size: 12px;
  color: #909399;
}

.file-progress {
  margin-top: 12px;
}

.remove-file-btn {
  position: absolute;
  right: 0;
  top: 50%;
  transform: translateY(-50%);
}

.conversion-options {
  margin-bottom: 16px;
}

.options-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
}

.option-item {
  display: flex;
  align-items: center;
  gap: 8px;
  background: rgba(64, 158, 255, 0.05);
  padding: 8px 16px;
  border-radius: 8px;
}

.option-label {
  font-size: 14px;
  font-weight: 500;
  color: #606266;
}

.action-buttons {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}

.conversion-progress {
  margin-top: 20px;
  padding: 16px;
  background: rgba(255, 255, 255, 0.8);
  border-radius: 12px;
  backdrop-filter: blur(10px);
  box-shadow: 0 8px 30px rgba(0, 0, 0, 0.05);
  border: 1px solid rgba(64, 158, 255, 0.1);
}

.progress-bar {
  margin-bottom: 12px;
}

.status-message {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 12px;
  border-radius: 8px;
  background: rgba(245, 108, 108, 0.1);
  color: #f56c6c;
  font-size: 13px;
  font-weight: 500;
  backdrop-filter: blur(5px);
  border: 1px solid rgba(245, 108, 108, 0.2);
}

.error-icon {
  font-size: 18px;
}

.footer {
  text-align: center;
  padding: 10px;
  font-size: 12px;
  color: #909399;
  background-color: rgba(255, 255, 255, 0.8);
  backdrop-filter: blur(10px);
}

/* 暗黑模式 */
:global(.dark-theme) .app-container {
  background: linear-gradient(135deg, #0a0f2c 0%, #1a1f3c 50%, #141e3c 100%);
}

:global(.dark-theme) .app-container::before {
  opacity: 0.4;
}

:global(.dark-theme) .nav-header {
  background: rgba(0, 0, 0, 0.2);
}

/* 响应式设计优化 */
@media (max-width: 768px) {
  .nav-header {
    padding: 8px 16px;
  }

  .logo-text {
    font-size: 1.3rem;
  }

  .conversion-type-button {
    padding: 8px 16px;
  }

  .upload-icon {
    font-size: 40px;
  }

  .upload-text {
    font-size: 14px;
  }
}

@media (max-width: 480px) {
  .conversion-selector {
    margin: 12px 0;
  }

  .panel-header {
    padding: 12px;
  }

  .panel-header h2 {
    font-size: 1.1rem;
  }

  .upload-area {
    margin: 12px;
  }

  .drop-zone {
    padding: 20px 16px;
  }
}

/* 新增和修改的样式 */
.conversion-type-button {
  position: relative;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 10px 20px;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  overflow: hidden;
  border-radius: 8px;
  margin: 0 4px;
  backdrop-filter: blur(10px);
}

.conversion-type-button.is-active {
  box-shadow: 0 8px 16px rgba(0, 0, 0, 0.1);
  transform: translateY(-1px);
}

.conversion-type-button:not(.is-active) {
  opacity: 0.8;
}

.conversion-type-button:hover {
  transform: translateY(-2px);
  box-shadow: 0 12px 20px rgba(0, 0, 0, 0.15);
}

.conversion-type-button .el-icon {
  font-size: 20px;
  transition: transform 0.3s ease;
}

.conversion-type-button:hover .el-icon {
  transform: scale(1.1);
}

.conversion-type-button span {
  font-weight: 500;
}

.header-content {
  display: flex;
  align-items: center;
  gap: 16px;
}

.header-icon {
  color: white;
  background: rgba(255, 255, 255, 0.2);
  padding: 12px;
  border-radius: 12px;
}

.header-text {
  flex: 1;
}

.convert-button,
.download-button,
.cleanup-button {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 10px 20px;
  border-radius: 8px;
  font-weight: 600;
  letter-spacing: 0.5px;
  font-size: 14px;
}

.convert-button {
  background: linear-gradient(45deg, #4481eb, #04befe);
}

.convert-button:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 25px rgba(68, 129, 235, 0.4);
}

.download-button {
  background: linear-gradient(45deg, #67c23a, #95d475);
}

.download-button:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 25px rgba(103, 194, 58, 0.4);
}

.cleanup-button {
  background: linear-gradient(45deg, #909399, #c0c4cc);
}

.cleanup-button:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 25px rgba(144, 147, 153, 0.4);
}

.button-icon {
  font-size: 20px;
  transition: transform 0.3s ease;
}

.convert-button:hover .button-icon,
.download-button:hover .button-icon,
.cleanup-button:hover .button-icon {
  transform: scale(1.1);
}

.advanced-mode-switch {
  transform: scale(1.2);
}

/* 添加新的样式 */
.header-pdf-word {
  background: linear-gradient(45deg, rgba(124, 77, 255, 0.8), rgba(0, 229, 255, 0.8));
}

.header-md-word {
  background: linear-gradient(45deg, rgba(52, 168, 83, 0.8), rgba(156, 204, 101, 0.8));
}

.header-word-pdf {
  background: linear-gradient(45deg, rgba(66, 133, 244, 0.8), rgba(52, 168, 83, 0.8));
}

.header-pdf-md {
  background: linear-gradient(45deg, rgba(234, 67, 53, 0.8), rgba(251, 188, 4, 0.8));
}

/* 更新文件图标样式 */
.file-icon-container {
  background: linear-gradient(45deg, rgba(64, 158, 255, 0.2), rgba(0, 229, 255, 0.2));
}

[data-conversion-type="markdown-to-word"] .file-icon-container {
  background: linear-gradient(45deg, rgba(52, 168, 83, 0.2), rgba(156, 204, 101, 0.2));
}

[data-conversion-type="word-to-pdf"] .file-icon-container {
  background: linear-gradient(45deg, rgba(66, 133, 244, 0.2), rgba(52, 168, 83, 0.2));
}

[data-conversion-type="pdf-to-markdown"] .file-icon-container {
  background: linear-gradient(45deg, rgba(234, 67, 53, 0.2), rgba(251, 188, 4, 0.2));
}
</style>