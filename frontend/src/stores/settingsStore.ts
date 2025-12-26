import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import type { Settings } from '@/types';

interface SettingsState extends Settings {
  // Actions
  setApiKey: (key: string) => void;
  setApiBase: (base: string) => void;
  setModelName: (name: string) => void;
  setUnsplashKey: (key: string) => void;
  updateSettings: (settings: Partial<Settings>) => void;
  resetSettings: () => void;
}

const defaultSettings: Settings = {
  apiKey: '',
  apiBase: 'https://api.openai.com/v1',
  modelName: 'gpt-4o-mini',
  unsplashKey: '',
};

export const useSettingsStore = create<SettingsState>()(
  persist(
    (set) => ({
      ...defaultSettings,

      setApiKey: (apiKey) => set({ apiKey }),
      setApiBase: (apiBase) => set({ apiBase }),
      setModelName: (modelName) => set({ modelName }),
      setUnsplashKey: (unsplashKey) => set({ unsplashKey }),

      updateSettings: (settings) => set((state) => ({ ...state, ...settings })),

      resetSettings: () => set(defaultSettings),
    }),
    {
      name: 'ppt-settings',
    }
  )
);
