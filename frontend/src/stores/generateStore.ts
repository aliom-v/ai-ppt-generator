import { create } from 'zustand';
import type { GenerateResult, SlidePreview } from '@/types';

export type GenerateStatus = 'idle' | 'generating' | 'previewing' | 'success' | 'error';

interface GenerateState {
  // 表单数据
  topic: string;
  description: string;
  pageCount: number;
  autoPageCount: boolean;
  templateId: string;
  autoDownload: boolean;

  // 上传文件信息
  uploadedFileName: string | null;
  uploadedContent: string | null;

  // 生成状态
  status: GenerateStatus;
  progress: number;
  progressMessage: string;
  result: GenerateResult | null;
  error: string | null;

  // 预览数据
  previewData: {
    title: string;
    subtitle: string;
    slides: SlidePreview[];
  } | null;

  // Actions
  setTopic: (topic: string) => void;
  setDescription: (description: string) => void;
  setPageCount: (count: number) => void;
  setAutoPageCount: (auto: boolean) => void;
  setTemplateId: (id: string) => void;
  setAutoDownload: (auto: boolean) => void;
  setUploadedFile: (fileName: string | null, content: string | null) => void;
  appendToDescription: (content: string) => void;

  setStatus: (status: GenerateStatus) => void;
  setProgress: (progress: number, message?: string) => void;
  setResult: (result: GenerateResult | null) => void;
  setError: (error: string | null) => void;
  setPreviewData: (data: GenerateState['previewData']) => void;

  reset: () => void;
  resetResult: () => void;
}

const initialState = {
  topic: '',
  description: '',
  pageCount: 5,
  autoPageCount: false,
  templateId: 'mckinsey_consulting',
  autoDownload: false,
  uploadedFileName: null,
  uploadedContent: null,
  status: 'idle' as GenerateStatus,
  progress: 0,
  progressMessage: '',
  result: null,
  error: null,
  previewData: null,
};

export const useGenerateStore = create<GenerateState>()((set) => ({
  ...initialState,

  setTopic: (topic) => set({ topic }),
  setDescription: (description) => set({ description }),
  setPageCount: (pageCount) => set({ pageCount }),
  setAutoPageCount: (autoPageCount) => set({ autoPageCount }),
  setTemplateId: (templateId) => set({ templateId }),
  setAutoDownload: (autoDownload) => set({ autoDownload }),

  setUploadedFile: (uploadedFileName, uploadedContent) =>
    set({ uploadedFileName, uploadedContent }),

  appendToDescription: (content) =>
    set((state) => ({
      description: state.description
        ? `${state.description}\n\n【参考资料】\n${content}`
        : content,
    })),

  setStatus: (status) => set({ status }),
  setProgress: (progress, message) =>
    set({ progress, progressMessage: message || '' }),
  setResult: (result) => set({ result }),
  setError: (error) => set({ error }),
  setPreviewData: (previewData) => set({ previewData }),

  reset: () => set(initialState),
  resetResult: () =>
    set({
      status: 'idle',
      progress: 0,
      progressMessage: '',
      result: null,
      error: null,
      previewData: null,
    }),
}));
