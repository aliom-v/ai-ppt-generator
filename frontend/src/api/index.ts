import axios from 'axios';
import type {
  GenerateRequest,
  GenerateResult,
  PreviewRequest,
  PreviewResult,
  TemplatesResponse,
  TestConnectionResult,
  UploadResult,
  HistoryRecord,
  HistoryStats,
} from '@/types';

const apiClient = axios.create({
  baseURL: '',
  timeout: 180000, // 3 分钟
  headers: {
    'Content-Type': 'application/json',
  },
});

// 响应拦截器
apiClient.interceptors.response.use(
  (response) => response.data,
  (error) => {
    const message = error.response?.data?.error || error.message;
    return Promise.reject(new Error(message));
  }
);

// 生成 PPT
export async function generatePPT(data: GenerateRequest): Promise<GenerateResult> {
  return apiClient.post('/api/generate', data);
}

// 预览结构
export async function previewStructure(data: PreviewRequest): Promise<PreviewResult> {
  return apiClient.post('/api/preview', data);
}

// 获取模板列表
export async function getTemplates(): Promise<TemplatesResponse> {
  return apiClient.get('/api/templates');
}

// 测试 API 连接
export async function testConnection(data: {
  api_key: string;
  api_base: string;
  model_name: string;
}): Promise<TestConnectionResult> {
  return apiClient.post('/api/test-connection', data);
}

// 上传文件
export async function uploadFile(file: File): Promise<UploadResult> {
  const formData = new FormData();
  formData.append('file', file);
  return apiClient.post('/api/upload', formData, {
    headers: {
      'Content-Type': 'multipart/form-data',
    },
  });
}

// 获取历史记录
export async function getHistory(params?: {
  limit?: number;
  offset?: number;
}): Promise<{ success: boolean; records: HistoryRecord[]; count: number }> {
  return apiClient.get('/api/history', { params });
}

// 获取历史统计
export async function getHistoryStats(): Promise<{ success: boolean } & HistoryStats> {
  return apiClient.get('/api/history/stats');
}

// 搜索历史
export async function searchHistory(params: {
  keyword?: string;
  status?: string;
  start_date?: string;
  end_date?: string;
  limit?: number;
}): Promise<{ success: boolean; records: HistoryRecord[]; count: number }> {
  return apiClient.get('/api/history/search', { params });
}

// 获取配置状态
export async function getConfig(): Promise<{
  ai_configured: boolean;
  image_search_available: boolean;
}> {
  return apiClient.get('/api/config');
}

// 获取下载 URL
export function getDownloadUrl(filename: string): string {
  return `/api/download/${encodeURIComponent(filename)}`;
}

export default apiClient;
