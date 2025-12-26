// API 响应类型
export interface ApiResponse<T = unknown> {
  success: boolean;
  error?: string;
  data?: T;
}

// 生成请求参数
export interface GenerateRequest {
  topic: string;
  audience: string;
  page_count: number;
  description: string;
  auto_page_count: boolean;
  auto_download: boolean;
  template_id: string;
  api_key: string;
  api_base: string;
  model_name: string;
  unsplash_key: string;
}

// 生成结果
export interface GenerateResult {
  success: boolean;
  filename: string;
  title: string;
  subtitle: string;
  slide_count: number;
  duration_ms: number;
  download_url: string;
}

// 预览请求
export interface PreviewRequest {
  topic: string;
  audience: string;
  page_count: number;
  api_key: string;
  api_base: string;
  model_name: string;
}

// 幻灯片预览
export interface SlidePreview {
  index: number;
  type: string;
  title: string;
  bullets: string[];
  text: string;
  image_keyword: string;
  subtitle: string;
}

// 预览结果
export interface PreviewResult {
  success: boolean;
  title: string;
  subtitle: string;
  slides: SlidePreview[];
}

// 模板
export interface Template {
  id: string;
  name: string;
  description: string;
  category?: string;
}

// 模板列表响应
export interface TemplatesResponse {
  success: boolean;
  templates: Template[];
  count: number;
}

// 测试连接结果
export interface TestConnectionResult {
  success: boolean;
  message: string;
  model: string;
  api_base: string;
  response_time: number;
  response?: string;
}

// 文件上传结果
export interface UploadResult {
  success: boolean;
  content: string;
  summary: {
    length: number;
    lines: number;
    estimated_tokens: number;
    is_truncated: boolean;
  };
  filename: string;
  error?: string;
}

// 历史记录
export interface HistoryRecord {
  id: number;
  topic: string;
  audience: string;
  page_count: number;
  model_name: string;
  template_id: string;
  filename: string;
  file_size: number;
  slide_count: number;
  duration_ms: number;
  status: string;
  created_at: string;
}

// 历史统计
export interface HistoryStats {
  total: number;
  success_count: number;
  error_count: number;
  avg_duration_ms: number;
  total_slides: number;
}

// 设置
export interface Settings {
  apiKey: string;
  apiBase: string;
  modelName: string;
  unsplashKey: string;
}
