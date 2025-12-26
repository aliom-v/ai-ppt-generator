import { create } from 'zustand';

interface UIState {
  // 弹窗状态
  settingsModalOpen: boolean;
  previewModalOpen: boolean;

  // Actions
  openSettingsModal: () => void;
  closeSettingsModal: () => void;
  openPreviewModal: () => void;
  closePreviewModal: () => void;
}

export const useUIStore = create<UIState>()((set) => ({
  settingsModalOpen: false,
  previewModalOpen: false,

  openSettingsModal: () => set({ settingsModalOpen: true }),
  closeSettingsModal: () => set({ settingsModalOpen: false }),
  openPreviewModal: () => set({ previewModalOpen: true }),
  closePreviewModal: () => set({ previewModalOpen: false }),
}));
