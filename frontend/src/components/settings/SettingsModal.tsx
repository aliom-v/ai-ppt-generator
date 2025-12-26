import { useState, useEffect } from 'react';
import {
  Modal,
  TextInput,
  PasswordInput,
  Button,
  Stack,
  Group,
  Text,
  Divider,
  Alert,
} from '@mantine/core';
import { IconSettings, IconPlugConnected, IconPhoto } from '@tabler/icons-react';
import { useSettingsStore, useUIStore } from '@/stores';
import { testConnection } from '@/api';

export function SettingsModal() {
  const { settingsModalOpen, closeSettingsModal } = useUIStore();
  const settings = useSettingsStore();

  const [localSettings, setLocalSettings] = useState({
    apiKey: settings.apiKey,
    apiBase: settings.apiBase,
    modelName: settings.modelName,
    unsplashKey: settings.unsplashKey,
  });

  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{
    success: boolean;
    message: string;
  } | null>(null);

  // 当弹窗打开时，同步设置
  useEffect(() => {
    if (settingsModalOpen) {
      setLocalSettings({
        apiKey: settings.apiKey,
        apiBase: settings.apiBase,
        modelName: settings.modelName,
        unsplashKey: settings.unsplashKey,
      });
      setTestResult(null);
    }
  }, [settingsModalOpen, settings.apiKey, settings.apiBase, settings.modelName, settings.unsplashKey]);

  const handleSave = () => {
    settings.updateSettings(localSettings);
    closeSettingsModal();
  };

  const handleTest = async () => {
    if (!localSettings.apiKey) {
      setTestResult({ success: false, message: '请先输入 API Key' });
      return;
    }

    setTesting(true);
    setTestResult(null);

    try {
      const result = await testConnection({
        api_key: localSettings.apiKey,
        api_base: localSettings.apiBase || 'https://api.openai.com/v1',
        model_name: localSettings.modelName || 'gpt-4o-mini',
      });

      if (result.success) {
        setTestResult({
          success: true,
          message: `连接成功！响应时间: ${result.response_time}ms`,
        });
      } else {
        setTestResult({ success: false, message: result.message });
      }
    } catch (error) {
      setTestResult({
        success: false,
        message: error instanceof Error ? error.message : '测试失败',
      });
    } finally {
      setTesting(false);
    }
  };

  return (
    <Modal
      opened={settingsModalOpen}
      onClose={closeSettingsModal}
      title={
        <Group gap="xs">
          <IconSettings size={20} />
          <Text fw={600}>配置设置</Text>
        </Group>
      }
      size="lg"
    >
      <Stack gap="lg">
        {/* AI 配置 */}
        <div>
          <Group gap="xs" mb="sm">
            <IconPlugConnected size={18} />
            <Text fw={500}>AI 模型配置</Text>
          </Group>

          <Stack gap="sm">
            <PasswordInput
              label="API Key"
              placeholder="输入你的 API Key"
              value={localSettings.apiKey}
              onChange={(e) =>
                setLocalSettings({ ...localSettings, apiKey: e.target.value })
              }
              description="支持 OpenAI、Claude、国内大模型等"
            />

            <TextInput
              label="API Base URL"
              placeholder="https://api.openai.com/v1"
              value={localSettings.apiBase}
              onChange={(e) =>
                setLocalSettings({ ...localSettings, apiBase: e.target.value })
              }
              description="API 端点地址"
            />

            <TextInput
              label="模型名称"
              placeholder="gpt-4o-mini"
              value={localSettings.modelName}
              onChange={(e) =>
                setLocalSettings({ ...localSettings, modelName: e.target.value })
              }
              description="如 gpt-4o-mini, claude-3-5-sonnet, qwen-plus 等"
            />
          </Stack>
        </div>

        <Divider />

        {/* 图片配置 */}
        <div>
          <Group gap="xs" mb="sm">
            <IconPhoto size={18} />
            <Text fw={500}>图片搜索配置</Text>
          </Group>

          <PasswordInput
            label="Unsplash Access Key"
            placeholder="输入 Unsplash API Key"
            value={localSettings.unsplashKey}
            onChange={(e) =>
              setLocalSettings({ ...localSettings, unsplashKey: e.target.value })
            }
            description={
              <>
                获取免费 API Key:{' '}
                <a
                  href="https://unsplash.com/developers"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  unsplash.com/developers
                </a>
              </>
            }
          />
        </div>

        {/* 测试结果 */}
        {testResult && (
          <Alert
            color={testResult.success ? 'green' : 'red'}
            title={testResult.success ? '连接成功' : '连接失败'}
          >
            {testResult.message}
          </Alert>
        )}

        {/* 操作按钮 */}
        <Group justify="flex-end" mt="md">
          <Button variant="default" onClick={closeSettingsModal}>
            取消
          </Button>
          <Button variant="light" onClick={handleTest} loading={testing}>
            测试连接
          </Button>
          <Button onClick={handleSave}>保存配置</Button>
        </Group>
      </Stack>
    </Modal>
  );
}
