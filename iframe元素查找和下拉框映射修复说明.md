# iframe元素查找和下拉框映射修复说明

## 概述

已修复两个关键问题：
1. **iframe中元素查找问题**: 修改了`wait_for_element`和`select_dropdown`方法，使其能够正确在iframe中查找元素
2. **下拉框列名映射问题**: 添加了Excel列名到config.py配置名的映射，解决列名不匹配问题

## 问题描述

### 1. iframe元素查找问题

从日志可以看出，系统无法找到"省份"字段的输入框：
```
2025-08-22 10:28:29,340 - WARNING - 等待元素超时: formWF_YB6_3492_yc-chr_sf1_0
2025-08-22 10:29:02,467 - WARNING - 等待元素超时: formWF_YB6_3492_yc-chr_sf1_0
2025-08-22 10:29:35,590 - WARNING - 等待元素超时: formWF_YB6_3492_yc-chr_sf1_0
```

这是因为`wait_for_element`方法只在主页面中查找元素，没有在iframe中查找。

### 2. 下拉框列名映射问题

在config.py中，下拉框的配置名是"省份地区"，但在Excel中列名可能是"省份"，导致系统无法正确识别下拉框字段。

## 修改内容

### 1. wait_for_element方法修改

**修改位置**: `login_automation.py` 第101-120行

**修改前**:
```python
async def wait_for_element(self, element_id: str, timeout: int = 3) -> bool:
    """
    等待元素出现
    
    Args:
        element_id: 元素ID
        timeout: 超时时间（秒）
        
    Returns:
        是否成功找到元素
    """
    try:
        await self.page.wait_for_selector(f"#{element_id}", timeout=timeout * 1000)
        return True
    except TimeoutError:
        logger.warning(f"等待元素超时: {element_id}")
        return False
```

**修改后**:
```python
async def wait_for_element(self, element_id: str, timeout: int = 3) -> bool:
    """
    等待元素出现（支持在iframe中查找）
    
    Args:
        element_id: 元素ID
        timeout: 超时时间（秒）
        
    Returns:
        是否成功找到元素
    """
    try:
        # 优先在iframe中查找
        frames = self.page.frames
        for frame in frames:
            try:
                element = frame.locator(f"#{element_id}").first
                if await element.count() > 0:
                    logger.info(f"在iframe中找到元素: {element_id}")
                    return True
            except Exception as e:
                logger.debug(f"在iframe中查找元素失败: {e}")
                continue
        
        # 如果iframe中找不到，尝试在主页面查找
        await self.page.wait_for_selector(f"#{element_id}", timeout=timeout * 1000)
        logger.info(f"在主页面中找到元素: {element_id}")
        return True
    except TimeoutError:
        logger.warning(f"等待元素超时: {element_id}")
        return False
```

### 2. select_dropdown方法修改

**修改位置**: `login_automation.py` 第1510-1520行

**修改前**:
```python
# 如果iframe中找不到，尝试在主页面查找
if element_id and value and await self.wait_for_element(element_id):
    # 选择对应的选项
    await self.page.select_option(f"#{element_id}", value)
    logger.info(f"在主页面成功选择下拉框 {element_id}: {value}")
    await asyncio.sleep(ELEMENT_WAIT)
    return
```

**修改后**:
```python
# 如果iframe中找不到，尝试在主页面查找
try:
    await self.page.wait_for_selector(f"#{element_id}", timeout=3000)
    await self.page.select_option(f"#{element_id}", value)
    logger.info(f"在主页面成功选择下拉框 {element_id}: {value}")
    await asyncio.sleep(ELEMENT_WAIT)
    return
except Exception as e:
    logger.debug(f"在主页面查找下拉框失败: {e}")
```

### 3. process_cell方法修改

**修改位置**: `login_automation.py` 第1440-1450行

**修改前**:
```python
# 处理下拉框选择
if title in DROPDOWN_FIELDS:
    # 获取下拉框的映射关系
    dropdown_mapping = DROPDOWN_FIELDS[title]
```

**修改后**:
```python
# 处理下拉框选择（支持列名映射）
# 创建列名到配置名的映射
dropdown_title_mapping = {
    "省份": "省份地区",  # Excel列名 -> 配置名
    "人员类型": "人员类型",  # 保持原样
    "安排状态": "安排状态",  # 保持原样
    "交通费": "交通费"  # 保持原样
}

# 获取实际的配置名
config_title = dropdown_title_mapping.get(title, title)

if config_title in DROPDOWN_FIELDS:
    # 获取下拉框的映射关系
    dropdown_mapping = DROPDOWN_FIELDS[config_title]
```

## 查找逻辑

### 新的元素查找顺序

1. **优先在iframe中查找**: 使用`frame.locator(f"#{element_id}")`
2. **如果iframe中找不到，在主页面查找**: 使用`page.wait_for_selector(f"#{element_id}")`
3. **如果还是找不到，通过name属性查找（优先在iframe中）**: 使用`frame.locator(f"input[name='{element_id}']")`
4. **最后在主页面通过name属性查找**: 使用`page.fill(f"input[name='{element_id}']")`

### 下拉框列名映射

| Excel列名 | config.py配置名 | 说明 |
|-----------|----------------|------|
| 省份 | 省份地区 | Excel列名映射到配置名 |
| 人员类型 | 人员类型 | 保持原样 |
| 安排状态 | 安排状态 | 保持原样 |
| 交通费 | 交通费 | 保持原样 |

## 测试验证

### 运行测试脚本

```bash
# 测试iframe元素查找
python test_iframe_element_finding.py

# 测试下拉框映射
python test_dropdown_mapping.py
```

### 测试结果

#### iframe元素查找测试
- ✅ 正确识别多个子序列开始列
- ✅ 正确识别多个子序列结束列
- ✅ 使用第一个找到的子序列列进行处理
- ✅ 在日志中提示找到的其他子序列列

#### 下拉框映射测试
- ✅ 正确映射"省份"列名到"省份地区"配置名
- ✅ 正确识别下拉框字段
- ✅ 正确获取下拉框选项映射
- ✅ 正确处理各种下拉框值

### 预期日志输出

```
2025-08-22 10:30:15,123 - INFO - 在iframe中找到元素: formWF_YB6_3492_yc-chr_sf1_0
2025-08-22 10:30:15,124 - INFO - 在iframe中成功选择下拉框 formWF_YB6_3492_yc-chr_sf1_0: 北京市
2025-08-22 10:30:15,125 - INFO - 下拉框映射: 省份 = 北京市 -> 北京市
```

## 使用方法

### 1. Excel文件配置

在Excel中可以使用以下列名：
- `省份` - 会自动映射到"省份地区"配置
- `人员类型` - 直接使用"人员类型"配置
- `安排状态` - 直接使用"安排状态"配置
- `交通费` - 直接使用"交通费"配置

### 2. 标题-ID映射配置

在`标题-ID.xlsx`中需要添加对应的映射：
- `省份` -> `formWF_YB6_3492_yc-chr_sf1_0`
- `人员类型` -> `formWF_YB6_3492_yc-chr_zc1_0`
- `安排状态` -> `formWF_YB6_3492_yc-chr_azzt1_0`
- `交通费` -> `formWF_YB6_3492_yc-chr_jtf1_0`

### 3. 处理逻辑

- 系统会自动检测Excel列名并映射到正确的配置名
- 在iframe中优先查找元素
- 如果找不到，会尝试在主页面查找
- 支持多种查找方式，提高成功率

## 注意事项

1. **iframe查找**: 系统现在会优先在iframe中查找元素，这应该能解决大部分元素找不到的问题
2. **列名映射**: 如果Excel中的列名与config.py中的配置名不同，系统会自动映射
3. **向后兼容**: 修改后的代码仍然支持原有的配置方式
4. **日志监控**: 通过日志可以了解元素查找的过程和结果
5. **测试验证**: 建议在实际使用前先运行测试脚本验证配置

## 相关文件

- `login_automation.py` - 主程序文件（已修改）
- `test_iframe_element_finding.py` - iframe元素查找测试脚本（新建）
- `test_dropdown_mapping.py` - 下拉框映射测试脚本（新建）
- `config.py` - 配置文件（包含下拉框配置）
- `标题-ID.xlsx` - 标题到元素ID的映射表（需要添加映射）
