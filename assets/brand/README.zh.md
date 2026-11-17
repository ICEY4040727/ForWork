# ForWork 品牌资产

[English](README.md) | 中文

ForWork 的标记是一扇半开的门：方形门框、向内打开的门扇，以及圆形门把手。它表示进入工作空间并开始工作，不借用字母、动物或 AI 星光。

## 构成

标准标记使用 24 单位方形。门框占 `x=4..20, y=3..21`：左边柱、右边柱与顶梁均为 2.5 单位。门扇铰链在 `(7,7)`，依次连到 `(16,9)`、`(16,19)`、`(7,21)`。门把手是 `(14,14)` 处半径 `0.85` 的圆形；它是门扇上的洞，而不是第二层墨色。

- 标记：`forwork-mark.svg`（单色 `currentColor`，不含把手；背景已知时使用底板版本）
- 底板：`forwork-tile.svg`（1024 单位，`#F4F2EC` 底色，墨色门形，把手镂空）
- 组合标：`forwork-lockup.svg`（182 x 24；去掉标记时为 156 x 24）
- 安装器：`installer-brand-light.svg`、`installer-brand-dark.svg`（600 x 196），`uninstaller-sidebar.svg`（164 x 314）

## 色彩

| 用途 | 值 |
| --- | --- |
| 底板底色 | `#F4F2EC` |
| 底板描边 | `#D8D3C8` |
| 标记墨色 | `#141414` |
| 深色背景字标 | `#F9FAFB` |

标记是纯平的：没有渐变、发光或第二主色。唯一的层次来自桌面图标源中的底板与标记投影。

## 尺寸与安全区

- 标记最小 16 px；侧栏与托盘分别使用 24 px 和 16 px。
- 底板母版为 1024 单位。Web maskable 图标将图形缩放到 `0.95`，使门形保持在 `0.4` 安全半径内。
- 桌面源保留 1104 单位画布和 40/32 底板内边距：`apps/desktop/resources/icon.svg` 为方形裁切源，`icon-windows.svg` 使用 192 圆角，`icon-macos.svg` 使用 230。
- 托盘源把门形暴露为 `id="tray-glyph"`，并复用底板背景。

## 字体

字标用平台 UI 字体栈（`Segoe UI`、`Helvetica Neue`、Arial）以 600 字重排 `ForWork`。文字没有转轮廓；`textLength="148"` 让 182/156 画板宽度在不同已安装字体下保持稳定。

## 动效

在空白首页，悬停标记会让门扇绕铰链旋转 `-7deg`，持续 700 ms。只有在支持悬停且未请求减少动态效果时才安装悬停状态。

## 派生资产

- 桌面图标：`apps/desktop/resources/icon{,-windows,-macos}.{svg,png}`、`tray-windows.ico`
- Web 安装图标：`apps/web/public/icon-{192,512}.png`、`icon-maskable-512.png`、`apple-touch-icon.png`
- Favicon：`apps/web/public/favicon{,-dark}.svg`、`website/public/favicon.svg`
- 官网字标：`website/public/wordmark.svg`
- 欢迎页与安装器：`apps/desktop/renderer/assets/welcome-brand.svg`、`apps/desktop/installer/assets/brand{,-2x,-dark,-dark-2x}.png`

用 `corepack pnpm --filter @forwork/web-frontend run render:app-icons` 重新生成 Web 图标；在 `apps/desktop` 中运行 `corepack pnpm run render:tray-icon` 重新生成托盘；安装器 PNG 由 `installer-brand-{light,dark}.svg` 通过 sharp 渲染。

## 首页 frame 层

官方 `@forwork/client-ui-brand-official` 包不再占用 `conversation.hero.frame`；DeepSeek 鲸鱼图集及其绘制器已随 T136 移除。该 slot 仍为其他 bundle 声明，空白首页保留静态门形标记与悬停门扇。
