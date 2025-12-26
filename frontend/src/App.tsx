import '@mantine/core/styles.css';
import '@mantine/notifications/styles.css';

import { MantineProvider, Container, Title, Text, Stack, Group, Box } from '@mantine/core';
import { Notifications } from '@mantine/notifications';
import { theme } from './theme';
import { ThemeToggle } from './components/common/ThemeToggle';
import { SettingsModal } from './components/settings/SettingsModal';
import { GenerateForm } from './components/generate/GenerateForm';

function App() {
  return (
    <MantineProvider theme={theme} defaultColorScheme="auto">
      <Notifications position="top-right" />

      {/* 主题切换按钮 */}
      <Box pos="fixed" top={20} right={20} style={{ zIndex: 1000 }}>
        <ThemeToggle />
      </Box>

      <Container size="md" py="xl">
        <Stack gap="xl">
          {/* 头部 */}
          <Stack align="center" gap="sm" ta="center">
            <Title order={1} size="3rem" fw={800}>
              AI PPT 生成器
            </Title>
            <Text size="lg" c="dimmed">
              输入主题，AI 自动生成专业 PPT
            </Text>
            <Group gap="sm" mt="sm">
              <FeatureTag>AI 智能生成</FeatureTag>
              <FeatureTag>自动配图</FeatureTag>
              <FeatureTag>一键下载</FeatureTag>
            </Group>
          </Stack>

          {/* 主表单 */}
          <GenerateForm />

          {/* 页脚 */}
          <Text ta="center" c="dimmed" size="sm">
            Powered by AI | 支持多种页面类型 | 自动图片搜索
          </Text>
        </Stack>
      </Container>

      {/* 设置弹窗 */}
      <SettingsModal />
    </MantineProvider>
  );
}

function FeatureTag({ children }: { children: React.ReactNode }) {
  return (
    <Box
      px="md"
      py="xs"
      style={(theme) => ({
        borderRadius: theme.radius.xl,
        border: `1px solid ${theme.colors.gray[3]}`,
        fontSize: theme.fontSizes.sm,
        fontWeight: 500,
      })}
    >
      {children}
    </Box>
  );
}

export default App;
