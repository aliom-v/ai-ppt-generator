import type { MantineColorsTuple } from '@mantine/core';
import { createTheme } from '@mantine/core';

// 自定义主色调 - 青色系
const primary: MantineColorsTuple = [
  '#e0f7fa',
  '#b2ebf2',
  '#80deea',
  '#4dd0e1',
  '#26c6da',
  '#00bcd4',
  '#00acc1',
  '#0097a7',
  '#00838f',
  '#006064',
];

// 粉色系（用于强调）
const pink: MantineColorsTuple = [
  '#fce4ec',
  '#f8bbd9',
  '#f48fb1',
  '#f06292',
  '#ec4899',
  '#e91e63',
  '#d81b60',
  '#c2185b',
  '#ad1457',
  '#880e4f',
];

export const theme = createTheme({
  primaryColor: 'primary',
  colors: {
    primary,
    pink,
  },
  fontFamily: 'Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
  defaultRadius: 'md',
  components: {
    Button: {
      defaultProps: {
        radius: 'md',
      },
    },
    Card: {
      defaultProps: {
        radius: 'lg',
        shadow: 'sm',
      },
    },
    TextInput: {
      defaultProps: {
        radius: 'md',
      },
    },
    Textarea: {
      defaultProps: {
        radius: 'md',
      },
    },
    Select: {
      defaultProps: {
        radius: 'md',
      },
    },
    Modal: {
      defaultProps: {
        radius: 'lg',
      },
    },
    Paper: {
      defaultProps: {
        radius: 'md',
      },
    },
  },
});
