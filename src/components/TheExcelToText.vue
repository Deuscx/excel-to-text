<script setup lang="ts">
import * as XLSX from 'xlsx'

const fileInput = ref<HTMLInputElement>()
const convertedText = ref('')
const isProcessing = ref(false)
const fileName = ref('')
const errorMessage = ref('')

// 处理文件上传
async function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  if (!target.files || target.files.length === 0)
    return

  const file = target.files[0]
  fileName.value = file.name
  errorMessage.value = ''

  try {
    isProcessing.value = true
    const text = await convertExcelToText(file)
    convertedText.value = text

    // 重置文件输入，以便可以再次选择同一个文件
    if (fileInput.value) {
      fileInput.value.value = ''
    }
  }
  catch (error) {
    errorMessage.value = error instanceof Error ? error.message : '处理文件时出错'
    convertedText.value = ''
  }
  finally {
    isProcessing.value = false
  }
}

// 转换Excel为文本
function convertExcelToText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const data = e.target?.result
        if (!data)
          throw new Error('无法读取文件数据')

        const workbook = XLSX.read(data, { type: 'array' })
        let text = ''

        // 遍历所有工作表
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName]
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][]

          // 添加工作表名称
          text += `===== ${sheetName} =====\n`

          // 转换为文本格式
          jsonData.forEach((row) => {
            // 过滤掉空行
            if (row.some(cell => cell !== undefined && cell !== null && cell !== '')) {
              text += row.map((cell) => {
                // 处理undefined/null值
                if (cell === undefined || cell === null)
                  return ''
                // 确保是字符串
                return String(cell)
              }).join('\t') // 使用制表符分隔单元格
              text += '\n'
            }
          })

          text += '\n'
        })

        resolve(text)
      }
      catch (error) {
        reject(error)
      }
    }

    reader.onerror = () => {
      reject(new Error('读取文件时出错'))
    }

    reader.readAsArrayBuffer(file)
  })
}

// 下载文本文件
function downloadTextFile() {
  if (!convertedText.value)
    return

  const blob = new Blob([convertedText.value], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')

  // 使用原始文件名但更改扩展名
  const baseName = fileName.value.replace(/\.[^/.]+$/, '')
  a.download = `${baseName}.txt`
  a.href = url

  document.body.appendChild(a)
  a.click()

  // 清理
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

// 复制到剪贴板
async function copyToClipboard() {
  if (!convertedText.value)
    return

  try {
    await navigator.clipboard.writeText(convertedText.value)
  }
  catch {
    errorMessage.value = '复制到剪贴板失败'
  }
}

// 清除结果
function clearResult() {
  convertedText.value = ''
  fileName.value = ''
  errorMessage.value = ''
}
</script>

<template>
  <div class="excel-to-text-container mx-auto p-4 max-w-4xl">
    <h2 class="text-2xl font-bold mb-6 text-center">
      Excel转文本工具
    </h2>

    <!-- 文件上传区域 -->
    <div class="upload-section mb-6">
      <label for="excel-file" class="border-2 border-gray-300 rounded-lg border-dashed bg-gray-50 flex flex-col h-32 w-full cursor-pointer transition-colors items-center justify-center dark:bg-gray-800 hover:bg-gray-100 dark:hover:bg-gray-700">
        <div class="pb-6 pt-5 flex flex-col items-center justify-center">
          <svg class="text-gray-400 mb-3 h-12 w-12" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
          </svg>
          <p class="text-sm text-gray-500 mb-2 dark:text-gray-400">点击或拖拽文件到此处上传</p>
          <p class="text-xs text-gray-500 dark:text-gray-400">支持 .xlsx, .xls 等Excel格式</p>
        </div>
        <input
          id="excel-file"
          ref="fileInput"
          type="file"
          class="hidden"
          accept=".xlsx,.xls,.csv"
          @change="handleFileUpload"
        >
      </label>

      <!-- 显示已选择的文件名 -->
      <div v-if="fileName" class="text-sm text-gray-600 mt-3 dark:text-gray-300">
        已选择文件: {{ fileName }}
      </div>

      <!-- 错误提示 -->
      <div v-if="errorMessage" class="text-sm text-red-500 mt-3">
        {{ errorMessage }}
      </div>
    </div>

    <!-- 处理状态 -->
    <div v-if="isProcessing" class="text-gray-600 mb-6 text-center dark:text-gray-300">
      <svg class="h-6 w-6 inline-block animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <circle class="opacity-25" cx="12" cy="12" r="10" stroke-width="4" />
        <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
      </svg>
      <span class="ml-2">正在处理文件...</span>
    </div>

    <!-- 结果区域 -->
    <div v-if="convertedText" class="result-section">
      <div class="mb-3 flex items-center justify-between">
        <h3 class="text-lg font-semibold">
          转换结果
        </h3>
        <div class="flex gap-2">
          <button
            class="text-sm px-3 py-1 rounded bg-gray-200 dark:bg-gray-700 hover:bg-gray-300 dark:hover:bg-gray-600"
            title="复制到剪贴板"
            @click="copyToClipboard"
          >
            复制
          </button>
          <button
            class="text-sm text-white px-3 py-1 rounded bg-blue-500 hover:bg-blue-600"
            title="下载文本文件"
            @click="downloadTextFile"
          >
            下载
          </button>
          <button
            class="text-sm text-white px-3 py-1 rounded bg-red-500 hover:bg-red-600"
            title="清除结果"
            @click="clearResult"
          >
            清除
          </button>
        </div>
      </div>

      <!-- 预览区域 -->
      <div class="preview-container">
        <pre class="text-sm p-4 rounded-lg bg-gray-100 max-h-96 overflow-auto dark:bg-gray-800">
          {{ convertedText }}
        </pre>
      </div>
    </div>

    <!-- 说明文本 -->
    <div class="instructions mt-8 p-4 rounded-lg bg-gray-50 dark:bg-gray-800">
      <h4 class="font-semibold mb-2">
        使用说明
      </h4>
      <ul class="text-sm text-gray-600 space-y-1 dark:text-gray-300">
        <li>1. 点击或拖拽Excel文件到上传区域</li>
        <li>2. 系统会自动处理并转换Excel内容为文本格式</li>
        <li>3. 转换完成后，可以在预览区域查看结果</li>
        <li>4. 使用提供的按钮可以复制文本、下载为.txt文件或清除结果</li>
        <li>5. 支持多工作表转换，每个工作表会单独显示</li>
      </ul>
    </div>
  </div>
</template>

<style scoped>
.excel-to-text-container {
  min-height: calc(100vh - 200px);
}

.preview-container {
  font-family: monospace;
}

/* 确保在深色模式下有良好的对比度 */
.dark .preview-container pre {
  color: #e2e8f0;
}
</style>
