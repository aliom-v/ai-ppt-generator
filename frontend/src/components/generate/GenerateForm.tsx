import { useState, useEffect } from 'react';
import {
  Card,
  TextInput,
  Textarea,
  Select,
  NumberInput,
  Button,
  Group,
  Stack,
  Text,
  SegmentedControl,
  Checkbox,
  FileButton,
  Progress,
  Alert,
  Badge,
  Paper,
  Loader,
  Divider,
} from '@mantine/core';
import {
  IconSparkles,
  IconUpload,
  IconEye,
  IconSettings,
  IconCheck,
  IconX,
  IconDownload,
  IconRefresh,
} from '@tabler/icons-react';
import { useGenerateStore, useSettingsStore, useUIStore } from '@/stores';
import { generatePPT, previewStructure, uploadFile, getTemplates, getDownloadUrl } from '@/api';
import type { Template } from '@/types';

const loadingMessages = [
  'AI 正在分析主题',
  'AI 正在构思内容',
  'AI 正在生成结构',
  'AI 正在优化排版',
  '即将完成',
];

export function GenerateForm() {
  const store = useGenerateStore();
  const settings = useSettingsStore();
  const { openSettingsModal } = useUIStore();

  const [templates, setTemplates] = useState<Template[]>([]);
  const [loadingTemplates, setLoadingTemplates] = useState(true);
  const [messageIndex, setMessageIndex] = useState(0);

  // 加载模板列表
  useEffect(() => {
    const loadTemplates = async () => {
      try {
        const res = await getTemplates();
        if (res.success && res.templates) {
          setTemplates(res.templates);
        }
      } catch (error) {
        console.error('加载模板失败:', error);
      } finally {
        setLoadingTemplates(false);
      }
    };
    loadTemplates();
  }, []);

  // 加载动画
  useEffect(() => {
    if (store.status === 'generating' || store.status === 'previewing') {
      const interval = setInterval(() => {
        setMessageIndex((i) => (i + 1) % loadingMessages.length);
      }, 2000);
      return () => clearInterval(interval);
    }
  }, [store.status]);

  // 处理文件上传
  const handleFileUpload = async (file: File | null) => {
    if (!file) return;

    if (file.size > 5 * 1024 * 1024) {
      store.setError('文件过大，最大支持 5 MB');
      return;
    }

    try {
      const result = await uploadFile(file);
      if (result.success) {
        store.setUploadedFile(result.filename, result.content);
        store.appendToDescription(result.content);

        if (result.summary.is_truncated) {
          store.setError('文件内容过长，已自动截断到 5 万字');
        }
      } else {
        store.setError(result.error || '文件解析失败');
      }
    } catch (error) {
      store.setError(error instanceof Error ? error.message : '上传失败');
    }
  };

  // 生成 PPT
  const handleGenerate = async () => {
    if (!settings.apiKey) {
      store.setError('请先配置 AI API Key');
      openSettingsModal();
      return;
    }

    if (!store.topic.trim()) {
      store.setError('请输入 PPT 主题');
      return;
    }

    store.setStatus('generating');
    store.setError(null);
    store.setProgress(0);

    try {
      const result = await generatePPT({
        topic: store.topic,
        audience: '通用受众',
        page_count: store.autoPageCount ? 0 : store.pageCount,
        description: store.description,
        auto_page_count: store.autoPageCount,
        auto_download: store.autoDownload,
        template_id: store.templateId,
        api_key: settings.apiKey,
        api_base: settings.apiBase || 'https://api.openai.com/v1',
        model_name: settings.modelName || 'gpt-4o-mini',
        unsplash_key: settings.unsplashKey || '',
      });

      if (result.success) {
        store.setResult(result);
        store.setStatus('success');
      } else {
        throw new Error('生成失败');
      }
    } catch (error) {
      store.setError(error instanceof Error ? error.message : '生成失败');
      store.setStatus('error');
    }
  };

  // 预览结构
  const handlePreview = async () => {
    if (!settings.apiKey) {
      store.setError('请先配置 AI API Key');
      openSettingsModal();
      return;
    }

    if (!store.topic.trim()) {
      store.setError('请输入 PPT 主题');
      return;
    }

    store.setStatus('previewing');
    store.setError(null);

    try {
      const result = await previewStructure({
        topic: store.topic,
        audience: '通用受众',
        page_count: store.pageCount,
        api_key: settings.apiKey,
        api_base: settings.apiBase || 'https://api.openai.com/v1',
        model_name: settings.modelName || 'gpt-4o-mini',
      });

      if (result.success) {
        store.setPreviewData({
          title: result.title,
          subtitle: result.subtitle,
          slides: result.slides,
        });
        store.setStatus('idle');
      } else {
        throw new Error('预览失败');
      }
    } catch (error) {
      store.setError(error instanceof Error ? error.message : '预览失败');
      store.setStatus('error');
    }
  };

  const isLoading = store.status === 'generating' || store.status === 'previewing';
  const hasUnsplashKey = !!settings.unsplashKey;

  const templateOptions = templates.map((t) => ({
    value: t.id,
    label: `${t.name} - ${t.description}`,
  }));

  return (
    <Stack gap="lg">
      {/* 配置状态提示 */}
      <Paper p="md" withBorder>
        <Group justify="space-between">
          <div>
            <Text fw={500} mb={4}>配置状态</Text>
            <Group gap="xs">
              <Badge
                color={settings.apiKey ? 'green' : 'yellow'}
                variant="light"
                leftSection={settings.apiKey ? <IconCheck size={12} /> : <IconX size={12} />}
              >
                {settings.apiKey ? 'AI 已配置' : 'AI 未配置'}
              </Badge>
              <Badge
                color={hasUnsplashKey ? 'green' : 'gray'}
                variant="light"
                leftSection={hasUnsplashKey ? <IconCheck size={12} /> : null}
              >
                {hasUnsplashKey ? '图片搜索已启用' : '图片搜索未启用'}
              </Badge>
            </Group>
          </div>
          <Button
            variant="light"
            leftSection={<IconSettings size={16} />}
            onClick={openSettingsModal}
          >
            配置设置
          </Button>
        </Group>
      </Paper>

      {/* 主表单 */}
      <Card shadow="sm" padding="lg" radius="lg" withBorder>
        <Stack gap="md">
          {/* 主题输入 */}
          <TextInput
            label="PPT 主题"
            placeholder="例如：AI 技术发展趋势"
            required
            value={store.topic}
            onChange={(e) => store.setTopic(e.target.value)}
            description="输入你想要生成的 PPT 主题"
            disabled={isLoading}
          />

          {/* 详细描述 */}
          <Textarea
            label="详细描述（可选）"
            placeholder="可以输入要点、大纲或粘贴参考资料"
            minRows={4}
            value={store.description}
            onChange={(e) => store.setDescription(e.target.value)}
            description="提供更多细节，让 AI 生成更精准的内容"
            disabled={isLoading}
          />

          {/* 文件上传 */}
          <div>
            <Text size="sm" fw={500} mb={4}>
              或上传参考资料
            </Text>
            <Group>
              <FileButton
                onChange={handleFileUpload}
                accept=".txt,.md,.docx,.pdf"
                disabled={isLoading}
              >
                {(props) => (
                  <Button
                    {...props}
                    variant="light"
                    leftSection={<IconUpload size={16} />}
                  >
                    {store.uploadedFileName || '选择文件'}
                  </Button>
                )}
              </FileButton>
              <Text size="xs" c="dimmed">
                支持 TXT、MD、DOCX、PDF，最大 5 MB
              </Text>
            </Group>
          </div>

          <Divider />

          {/* 页数设置 */}
          <div>
            <Text size="sm" fw={500} mb="xs">
              页数设置
            </Text>
            <SegmentedControl
              value={store.autoPageCount ? 'auto' : 'manual'}
              onChange={(value) => store.setAutoPageCount(value === 'auto')}
              data={[
                { label: '手动指定', value: 'manual' },
                { label: 'AI 智能判断', value: 'auto' },
              ]}
              disabled={isLoading}
              fullWidth
              mb="sm"
            />
            {!store.autoPageCount && (
              <NumberInput
                value={store.pageCount}
                onChange={(value) => store.setPageCount(Number(value) || 5)}
                min={1}
                max={100}
                description="不包括封面页，建议 5-10 页"
                disabled={isLoading}
              />
            )}
          </div>

          {/* 模板选择 */}
          <Select
            label="选择模板"
            placeholder={loadingTemplates ? '加载中...' : '选择模板'}
            data={templateOptions}
            value={store.templateId}
            onChange={(value) => store.setTemplateId(value || '')}
            disabled={isLoading || loadingTemplates}
            description={`${templates.length} 个模板可用`}
            searchable
          />

          {/* 自动下载图片 */}
          <Checkbox
            label="自动搜索下载图片"
            checked={store.autoDownload}
            onChange={(e) => store.setAutoDownload(e.currentTarget.checked)}
            disabled={isLoading || !hasUnsplashKey}
            description={
              hasUnsplashKey
                ? '已配置 Unsplash API Key'
                : '需要配置 Unsplash API Key'
            }
          />

          {/* 错误提示 */}
          {store.error && (
            <Alert color="red" title="错误" onClose={() => store.setError(null)} withCloseButton>
              {store.error}
            </Alert>
          )}

          {/* 加载状态 */}
          {isLoading && (
            <Paper p="xl" withBorder>
              <Stack align="center" gap="md">
                <Loader size="lg" />
                <Text fw={500}>{loadingMessages[messageIndex]}...</Text>
                <Text size="sm" c="dimmed">
                  这通常需要 5-10 秒
                </Text>
                <Progress value={30} animated w="100%" />
              </Stack>
            </Paper>
          )}

          {/* 成功结果 */}
          {store.status === 'success' && store.result && (
            <Alert color="green" title="🎉 PPT 生成成功！">
              <Stack gap="sm">
                <Text>
                  <strong>标题：</strong>
                  {store.result.title}
                </Text>
                <Text>
                  <strong>副标题：</strong>
                  {store.result.subtitle}
                </Text>
                <Text>
                  <strong>页数：</strong>
                  {store.result.slide_count + 1} 页（含封面）
                </Text>
                <Group mt="sm">
                  <Button
                    component="a"
                    href={getDownloadUrl(store.result.filename)}
                    download={store.result.filename}
                    leftSection={<IconDownload size={16} />}
                  >
                    立即下载 PPT
                  </Button>
                  <Button
                    variant="light"
                    leftSection={<IconRefresh size={16} />}
                    onClick={() => store.resetResult()}
                  >
                    再生成一个
                  </Button>
                </Group>
              </Stack>
            </Alert>
          )}

          {/* 预览结果 */}
          {store.previewData && (
            <Paper p="md" withBorder>
              <Text fw={600} size="lg" mb="md">
                内容预览
              </Text>
              <Paper p="sm" bg="gray.1" mb="md">
                <Text fw={600}>{store.previewData.title}</Text>
                <Text c="dimmed">{store.previewData.subtitle}</Text>
              </Paper>
              <Stack gap="xs">
                {store.previewData.slides.map((slide) => (
                  <Paper key={slide.index} p="sm" withBorder>
                    <Group gap="xs" mb="xs">
                      <Badge size="sm" variant="gradient" gradient={{ from: 'pink', to: 'orange' }}>
                        {getSlideTypeLabel(slide.type)}
                      </Badge>
                      <Text fw={500}>
                        {slide.index}. {slide.title}
                      </Text>
                    </Group>
                    {slide.bullets && slide.bullets.length > 0 && (
                      <ul style={{ margin: 0, paddingLeft: 20 }}>
                        {slide.bullets.map((bullet, i) => (
                          <li key={i}>
                            <Text size="sm" c="dimmed">
                              {bullet}
                            </Text>
                          </li>
                        ))}
                      </ul>
                    )}
                  </Paper>
                ))}
              </Stack>
            </Paper>
          )}

          {/* 操作按钮 */}
          <Group grow>
            <Button
              variant="light"
              leftSection={<IconEye size={16} />}
              onClick={handlePreview}
              disabled={isLoading}
            >
              预览结构
            </Button>
            <Button
              leftSection={<IconSparkles size={16} />}
              onClick={handleGenerate}
              loading={store.status === 'generating'}
              disabled={isLoading}
            >
              生成 PPT
            </Button>
          </Group>
        </Stack>
      </Card>
    </Stack>
  );
}

function getSlideTypeLabel(type: string): string {
  const labels: Record<string, string> = {
    bullets: '要点页',
    image_with_text: '图文页',
    two_column: '双栏页',
    timeline: '时间线',
    comparison: '对比页',
    quote: '引用页',
    ending: '结束页',
  };
  return labels[type] || type;
}
