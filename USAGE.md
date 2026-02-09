# AAT.ECF Cross-Validation - 使用说明

本项目提供两种运行方式，根据需求选择：

---

## 方式 1: 使用项目结构（推荐用于多模块）

适合需要运行多个模块、管理复杂工作流的场景。

### 使用步骤

1. **设置日期**

   打开 `main.py`，修改第 32 行：
   ```python
   DATE_STR = '20251130'  # 改成你要处理的日期
   ```

2. **选择要运行的模块**

   注释掉不需要运行的模块：
   ```python
   # ===== Run Modules =====
   run_cross_validation(DATE_STR)           # 运行
   # run_module2(DATE_STR)                  # 不运行（已注释）
   # run_module3(DATE_STR)                  # 不运行（已注释）
   ```

3. **运行**
   ```bash
   python main.py
   ```

### 优点
- ✅ 一次设置日期，所有模块使用同一日期
- ✅ 方便添加新模块
- ✅ 统一的入口和错误处理
- ✅ 清晰的项目结构

---

## 方式 2: 独立脚本（快速单次运行）

适合只需要运行交叉验证报告的场景。

### 使用步骤

1. **设置日期**

   打开 `Cross-validation.py`，修改第 18 行：
   ```python
   DATE_STR = '20251130'  # 改成你要处理的日期
   ```

2. **运行**
   ```bash
   python Cross-validation.py
   ```

### 优点
- ✅ 简单直接，单个文件
- ✅ 不依赖项目结构
- ✅ 可以快速测试

---

## 配置文件说明

### config.py
全局常量配置（路径、阈值等），两种方式都会使用：
```python
BASE_PATH = 'S:/Audrey/Audrey/AAT.DCF'
SIGNIFICANT_MV_THRESHOLD = 25_000_000
IRR_DIFF_THRESHOLD = 0.05
DURATION_DIFF_THRESHOLD = 0.5
```

### utils.py
工具函数，两种方式都会使用。

---

## 文件对应关系

| 方式 1（项目结构） | 方式 2（独立脚本） |
|------------------|------------------|
| main.py          | Cross-validation.py |
| modules/cross_validation.py | - |
| config.py        | config.py |
| utils.py         | utils.py |

---

## 选择建议

- **首次使用** → 方式 2（独立脚本）更简单
- **日常使用多个模块** → 方式 1（项目结构）更方便
- **快速测试** → 两种方式都可以

---

## 常见问题

**Q: 两种方式可以同时使用吗？**
A: 可以！它们互不影响。

**Q: 配置文件在哪里改？**
A:
- 日期：`main.py` 或 `Cross-validation.py`
- 路径/阈值：`config.py`

**Q: 如何添加新模块？**
A: 参考 `README.md` 中的"添加新模块"部分（仅适用于方式1）
