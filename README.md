
# Word 中英双语图名自动插入脚本

📑 专为中文学位论文设计的Word自动化脚本，支持自动编号、章节匹配和智能翻译建议
![WINWORD_YCWAoBdaFW200701](https://github.com/user-attachments/assets/9b35513f-25aa-454f-9692-11e6cc6edb48)


## 功能特性

- **自动化编号**  
  自动生成符合"图X.Y"和"Fig X.Y"格式的连续编号（X=章节号，Y=图序号）

- **双语支持**  
  单次操作同时插入中英文图名，中英文上下自动对齐

- **智能章节号提取**  
  自动识别最近的"标题 1"样式段落，提取"第X章"中的数字编号

- **翻译建议系统**  
  内置专业术语词典（可扩展），自动提供英文翻译建议

- **错误防御机制**  
  智能检测标题样式异常，提供中文错误指引

## 使用说明

### 安装方法
1. 打开Word文档，按 `Alt+F11` 进入VBA编辑器
2. 右击项目资源管理器 → 导入 → 选择`.bas`文件
3. 关闭VBA编辑器，将文档另存为`.docm`格式

### 操作演示
1. 将光标定位到需要插入图名的位置
2. 按 `Alt+F8` 打开宏对话框 → 选择 `InsertFigureCaption`
3. 按提示输入中英文图名：
   ![输入示例](https://via.placeholder.com/400x200?text=输入中文图名→自动建议英文翻译)

4. 生成效果示例：
   ```
   图2.1 膨胀土击实曲线
   Fig 2.1 Expansive soil compaction curve
   ```

### 参数配置
- **扩展翻译词典**  
  修改 `TranslateToEnglish` 函数中的字典数据：
  ```vba
  dict.Add "你的术语", "your translation"
  ```

- **调整编号格式**  
  修改 `InsertFigureCaption` 中的格式字符串：
  ```vba
  .TypeText text:="图" & chapNum & "."
  ' 改为其他格式如："Figure " & chapNum & "-"
  ```

## 注意事项

⚠️ **必看提示**  
1. 确保章节标题使用"标题 1"样式
2. 章节标题必须包含"第X章"格式（如：`第三章 实验结果`）
3. 首次使用需在Word信任中心启用宏
4. 英文翻译建议需自行扩展术语词典

## 贡献指南

欢迎通过 Issue 或 PR 提交：
- 发现BUG报告时请附上：
  - Word版本信息
  - 触发问题的操作步骤
  - 相关章节标题内容

- 新增翻译术语请按格式提交：
  ```vba
  dict.Add "新术语", "new_translation"
  ```
