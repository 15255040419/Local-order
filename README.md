# FTH 订单自动化助手 (FTH Order Automation Assistant) v3.0

一个基于 React + Vite 开发的高效订单处理工具，专为物流订单自动化设计。支持福鹿家系列订单处理及商颂下单转换，具备智能查重、规则匹配及一键导出功能。

![Version](https://img.shields.io/badge/version-3.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## 核心特性

- **📦 福鹿家自动化工作流**：
  - **第一步：提取门店信息** - 自动过滤历史重复订单，精准提取合格门店。
  - **第二步：生成 ERP 订单** - 自动关联货品模板，批量生成符合吉客云导入标准的 Excel。
  - **第三步：回填快递单号** - 智能匹配吉客云与菜单屏物流单号，实现 OMS 单号批量回填。
- **🚀 商颂下单转换**：
  - 支持复杂界面的订单文本解析。
  - 智能补纸逻辑：根据打印机类型自动匹配对应的纸张规格。
  - 深度物流匹配：内置快递价格明细规则，自动根据省份、平台及重量计算最佳物流。
- **🎨 极致交互体验**：
  - **现代审美**：采用 Apple 风格设计语言，支持深色/浅色模式无缝切换。
  - **实时反馈**：内置模拟终端记录执行日志，关键错误实时弹窗提醒。
  - **表格预览**：在导出前支持实时数据预览及红色标注查重项。

## 技术栈

- **前端框架**: React (TypeScript)
- **构建工具**: Vite
- **UI 组件**: Tailwind CSS + Lucide Icons
- **数据处理**: XLSX (xlsx-js-style)
- **状态提示**: Sonner

## 快速开始

### 安装依赖
```powershell
npm install
```

### 本地开发
```powershell
npm run dev
```

### 构建发布
```powershell
npm run build
```

## 开发者说明

项目核心逻辑位于 `src/app/App.tsx`，采用 React Hooks 进行状态管理，并使用 IndexedDB (IndexedDB API) 实现文件和配置的本地持久化。

---
由 Antigravity 协助优化与构建。