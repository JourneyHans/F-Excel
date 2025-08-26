# F-Excel 架构设计说明

## 架构概述

F-Excel 采用现代化的软件架构设计，遵循SOLID原则和多种设计模式，实现了高内聚、低耦合的模块化系统。整个系统采用分层架构，各层职责明确，便于维护和扩展。

## 整体架构图

```
┌─────────────────────────────────────────────────────────────┐
│                        表现层 (Presentation Layer)           │
├─────────────────────────────────────────────────────────────┤
│  main.py (FExcelApp)                                       │
│  ├── 主界面管理                                            │
│  ├── 模块调度                                              │
│  └── 用户交互处理                                          │
└─────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────┐
│                      业务逻辑层 (Business Logic Layer)        │
├─────────────────────────────────────────────────────────────┤
│  modules/                                                   │
│  ├── base_module.py (BaseModule)                           │
│  ├── id_converter.py (IDConverterModule)                   │
│  ├── excel_translator.py (ExcelTranslatorModule)           │
│  └── data_processors.py (DataProcessor)                    │
└─────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────┐
│                      数据访问层 (Data Access Layer)           │
├─────────────────────────────────────────────────────────────┤
│  pandas, openpyxl, file I/O                                │
│  ├── Excel文件读写                                          │
│  ├── 文本文件处理                                          │
│  └── 数据格式转换                                          │
└─────────────────────────────────────────────────────────────┘
```

## 核心设计模式

### 1. 分层架构模式 (Layered Architecture)

#### 表现层 (Presentation Layer)
- **职责**: 用户界面展示、用户交互处理
- **组件**: `FExcelApp` 类、tkinter界面组件
- **特点**: 
  - 只负责界面展示和用户交互
  - 不包含业务逻辑
  - 通过接口与业务逻辑层通信

#### 业务逻辑层 (Business Logic Layer)
- **职责**: 核心业务逻辑、数据处理、模块管理
- **组件**: 各种模块类、数据处理器
- **特点**:
  - 包含所有业务规则和逻辑
  - 独立于界面和数据访问
  - 可重用和测试

#### 数据访问层 (Data Access Layer)
- **职责**: 数据读写、文件操作、格式转换
- **组件**: pandas、openpyxl、文件I/O操作
- **特点**:
  - 封装数据访问细节
  - 提供统一的数据接口
  - 支持多种数据源

### 2. 模块化设计模式 (Modular Design)

#### 模块注册表模式
```python
# modules/__init__.py
AVAILABLE_MODULES = {
    'id_converter': {
        'name': 'ID值转换器',
        'description': '将数字=值格式转换为Excel文件',
        'icon': '📊',
        'class': IDConverterModule,
        'version': '1.0.0'
    },
    # ... 其他模块
}
```

#### 模块工厂模式
```python
def create_module(module_name: str):
    """创建指定模块的实例"""
    module_info = get_module_info(module_name)
    if module_info:
        return module_info['class']()
    return None
```

#### 模块接口统一
```python
class BaseModule(ABC):
    @abstractmethod
    def get_module_config(self) -> Dict[str, Any]:
        pass
    
    @abstractmethod
    def create_interface(self) -> None:
        pass
    
    @abstractmethod
    def process_data(self, input_data: str) -> List[Dict[str, Any]]:
        pass
```

### 3. 策略模式 (Strategy Pattern)

#### 数据处理策略
```python
class DataProcessor(ABC):
    @abstractmethod
    def process(self, data: str) -> List[Dict[str, Any]]:
        pass
    
    @abstractmethod
    def validate(self, data: str) -> bool:
        pass

class IDValueProcessor(DataProcessor):
    def process(self, data: str) -> List[Dict[str, Any]]:
        # ID值转换逻辑
        pass

class ExcelTranslationProcessor(DataProcessor):
    def process(self, data: str) -> List[Dict[str, Any]]:
        # Excel翻译逻辑
        pass
```

#### 策略工厂
```python
class DataProcessorFactory:
    _processors = {
        'id_value': IDValueProcessor,
        'excel_translation': ExcelTranslationProcessor
    }
    
    @classmethod
    def create_processor(cls, processor_type: str) -> Optional[DataProcessor]:
        processor_class = cls._processors.get(processor_type)
        if processor_class:
            return processor_class()
        return None
```

### 4. 模板方法模式 (Template Method Pattern)

#### 基础模块模板
```python
class BaseModule(ABC):
    def show(self) -> None:
        """显示模块窗口的模板方法"""
        if self.window is None or not self.window.winfo_exists():
            self.create_window()
        else:
            self.window.lift()
            self.window.focus()
    
    def create_window(self) -> None:
        """创建窗口的模板方法"""
        self.window = tk.Toplevel()
        self.window.title(f"{self.module_config['name']} - F-Excel")
        self.window.geometry(self.module_config.get('window_size', '900x700'))
        
        # 创建界面（子类实现）
        self.create_interface()
        
        # 窗口居中
        self.center_window()
    
    @abstractmethod
    def create_interface(self) -> None:
        """子类必须实现的抽象方法"""
        pass
```

### 5. 观察者模式 (Observer Pattern)

#### 进度更新观察者
```python
class BatchDataProcessor:
    def process_batch(self, data: str, progress_callback=None) -> List[Dict[str, Any]]:
        # ... 处理逻辑
        
        # 更新进度（通知观察者）
        if progress_callback:
            progress = min(100, (batch_end / total_lines) * 100)
            progress_callback(progress, f"正在处理第 {i + 1}-{batch_end}/{total_lines} 行...")
```

#### 状态更新观察者
```python
class BaseModule:
    def update_status(self, message: str) -> None:
        """更新状态栏信息"""
        if self.status_bar:
            self.status_bar.config(text=message)
    
    def update_progress(self, progress: float, status: str) -> None:
        """更新进度条和状态"""
        if self.progress_var:
            self.progress_var.set(progress)
        if self.status_label:
            self.status_label.config(text=status)
```

### 6. 单例模式 (Singleton Pattern)

#### 应用实例单例
```python
class FExcelApp:
    _instance = None
    
    def __new__(cls, *args, **kwargs):
        """单例模式实现"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, root: tk.Tk):
        """初始化主应用程序"""
        if hasattr(self, '_initialized'):
            return
        # ... 初始化逻辑
        self._initialized = True
```

## 依赖关系管理

### 1. 依赖注入

#### 模块依赖注入
```python
class IDConverterModule(BaseModule):
    def __init__(self):
        super().__init__()
        
        # 依赖注入：数据处理器
        self.data_processor = DataProcessorFactory.create_processor('id_value')
        
        # 依赖注入：转换后的数据
        self.converted_data: List[Dict[str, Any]] = []
```

#### 配置依赖注入
```python
class BaseModule:
    def __init__(self):
        # 依赖注入：模块配置
        self.module_config = self.get_module_config()
        
        # 依赖注入：UI组件
        self.window: Optional[tk.Toplevel] = None
        self.input_text: Optional[tk.Text] = None
        self.output_text: Optional[tk.Text] = None
```

### 2. 依赖倒置

#### 高层模块不依赖低层模块
```python
# 高层模块：BaseModule
class BaseModule(ABC):
    def load_file_content(self, file_path: str) -> None:
        """加载文件内容到输入框"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                if self.input_text:
                    self.input_text.delete('1.0', tk.END)
                    self.input_text.insert('1.0', content)
            self.update_status(f"已加载文件: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")

# 低层模块：具体文件操作
# 通过抽象接口进行交互，不直接依赖具体实现
```

## 错误处理架构

### 1. 分层错误处理

#### 表现层错误处理
```python
class FExcelApp:
    def start_module(self, module_name: str) -> None:
        try:
            if module_name in self.modules:
                self.modules[module_name].show()
                self.update_status(f"已启动模块: {module_name}")
            else:
                self.show_error("错误", f"模块 {module_name} 不存在")
        except Exception as e:
            self.show_error("错误", f"启动模块时发生错误: {str(e)}")
            self.update_status("模块启动失败")
```

#### 业务逻辑层错误处理
```python
class BaseModule:
    def show_error(self, title: str, message: str) -> None:
        """显示错误消息"""
        messagebox.showerror(title, message)
    
    def show_warning(self, title: str, message: str) -> None:
        """显示警告消息"""
        messagebox.showwarning(title, message)
    
    def show_info(self, title: str, message: str) -> None:
        """显示信息消息"""
        messagebox.showinfo(title, message)
```

#### 数据访问层错误处理
```python
def load_file_content(self, file_path: str) -> None:
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            if self.input_text:
                self.input_text.delete('1.0', tk.END)
                self.input_text.insert('1.0', content)
        self.update_status(f"已加载文件: {os.path.basename(file_path)}")
    except FileNotFoundError:
        self.show_error("错误", f"文件不存在: {file_path}")
    except PermissionError:
        self.show_error("错误", f"没有权限读取文件: {file_path}")
    except UnicodeDecodeError:
        self.show_error("错误", f"文件编码不支持: {file_path}")
    except Exception as e:
        self.show_error("错误", f"读取文件失败: {str(e)}")
```

### 2. 错误传播机制

#### 异常包装
```python
class ModuleError(Exception):
    """模块相关错误的基类"""
    pass

class DataProcessingError(ModuleError):
    """数据处理错误"""
    pass

class FileOperationError(ModuleError):
    """文件操作错误"""
    pass
```

#### 错误恢复机制
```python
def process_data_with_retry(self, data: str, max_retries: int = 3) -> List[Dict[str, Any]]:
    """带重试机制的数据处理"""
    for attempt in range(max_retries):
        try:
            return self.data_processor.process(data)
        except Exception as e:
            if attempt == max_retries - 1:
                raise DataProcessingError(f"数据处理失败，已重试{max_retries}次: {str(e)}")
            time.sleep(1)  # 等待1秒后重试
```

## 性能优化架构

### 1. 异步处理架构

#### 异步转换处理
```python
def start_async_conversion(self, input_data: str) -> None:
    """启动异步转换"""
    self.processing = True
    self.update_progress(0, "正在启动异步转换...")
    
    # 启用取消按钮
    self.cancel_button.config(state='normal')
    
    # 在新线程中执行转换
    thread = threading.Thread(target=self.convert_async, args=(input_data,))
    thread.daemon = True
    thread.start()

def convert_async(self, input_data: str) -> None:
    """异步转换数据"""
    try:
        if self.batch_processor:
            self.converted_data = self.batch_processor.process_batch(
                input_data, 
                self.update_progress
            )
        # ... 处理逻辑
    except Exception as e:
        self.window.after(0, self.conversion_failed, str(e))
```

#### 批处理架构
```python
class BatchDataProcessor:
    def process_batch(self, data: str, progress_callback=None) -> List[Dict[str, Any]]:
        lines = [line.strip() for line in data.strip().split('\n') 
                if line.strip() and not line.startswith('示例格式')]
        
        total_lines = len(lines)
        if total_lines == 0:
            return []
        
        self.processing = True
        self.cancelled = False
        processed_data = []
        
        try:
            for i in range(0, total_lines, self.batch_size):
                if self.cancelled:
                    break
                
                batch_end = min(i + self.batch_size, total_lines)
                batch_lines = lines[i:batch_end]
                
                # 处理当前批次
                batch_data = self.process_batch_lines(batch_lines)
                processed_data.extend(batch_data)
                
                # 更新进度
                if progress_callback:
                    progress = min(100, (batch_end / total_lines) * 100)
                    progress_callback(progress, f"正在处理第 {i + 1}-{batch_end}/{total_lines} 行...")
                
                # 短暂休息，避免界面冻结
                time.sleep(0.01)
        
        finally:
            self.processing = False
        
        return processed_data
```

### 2. 内存管理架构

#### 大文件分块处理
```python
def load_large_excel_file(self, file_path: str) -> None:
    """加载大型Excel文件（分块读取）"""
    try:
        self.update_progress(10, "检测到大文件，正在分块读取...")
        
        # 先读取前几行确定列结构
        df_sample = pd.read_excel(file_path, dtype={0: str}, nrows=1000)
        
        if len(df_sample.columns) < 3:
            self.show_warning("警告", "Excel文件至少需要3列（ID、中文、韩文）")
            return
        
        # 获取总行数
        df_full = pd.read_excel(file_path, dtype={0: str})
        total_rows = len(df_full)
        
        if total_rows > 100000:  # 超过10万行时只保留最后的部分
            df_full = df_full.tail(100000)
            total_rows = 100000
            self.show_warning("警告", f"文件过大，只保留了最后 {total_rows} 行数据")
        
        # 分批处理数据
        text_content = ""
        batch_size = 5000
        
        for i in range(0, total_rows, batch_size):
            batch_end = min(i + batch_size, total_rows)
            batch_df = df_subset.iloc[i:batch_end]
            
            for _, row in batch_df.iterrows():
                id_value = str(row.iloc[0])
                chinese_value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                korean_value = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                
                text_content += f"{id_value}\t{chinese_value}\t{korean_value}\n"
            
            # 更新进度
            progress = 40 + (batch_end / total_rows) * 60
            self.update_progress(progress, f"正在处理第 {i + 1}-{batch_end}/{total_rows} 行...")
            
            # 短暂休息，避免界面冻结
            time.sleep(0.01)
```

## 扩展性架构

### 1. 插件化架构

#### 模块注册机制
```python
# 新模块只需继承BaseModule并实现必要方法
class NewModule(BaseModule):
    def get_module_config(self) -> Dict[str, Any]:
        return {
            'name': '新模块',
            'description': '新模块描述',
            'icon': '🔧',
            'window_size': '800x600'
        }
    
    def create_interface(self) -> None:
        # 实现界面创建逻辑
        pass
    
    def process_data(self, input_data: str) -> List[Dict[str, Any]]:
        # 实现数据处理逻辑
        pass

# 在modules/__init__.py中注册
AVAILABLE_MODULES['new_module'] = {
    'name': '新模块',
    'description': '新模块描述',
    'icon': '🔧',
    'class': NewModule,
    'version': '1.0.0'
}
```

#### 处理器扩展机制
```python
# 新处理器只需继承DataProcessor
class NewDataProcessor(DataProcessor):
    def process(self, data: str) -> List[Dict[str, Any]]:
        # 实现数据处理逻辑
        pass
    
    def validate(self, data: str) -> bool:
        # 实现数据验证逻辑
        pass
    
    def get_supported_formats(self) -> List[str]:
        return ['新格式1', '新格式2']

# 在DataProcessorFactory中注册
DataProcessorFactory.register_processor('new_format', NewDataProcessor)
```

### 2. 配置驱动架构

#### 配置热重载
```python
class ConfigManager:
    def __init__(self):
        self.config = self.load_config()
        self.watcher = None
    
    def load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        # 实现配置加载逻辑
        pass
    
    def watch_config_changes(self):
        """监听配置文件变化"""
        # 实现文件变化监听
        pass
    
    def reload_config(self):
        """重新加载配置"""
        self.config = self.load_config()
        self.notify_config_changed()
```

#### 环境变量覆盖
```python
import os

class ConfigManager:
    def get_config_value(self, key: str, default=None):
        """获取配置值，支持环境变量覆盖"""
        # 优先从环境变量获取
        env_key = f"FEXCEL_{key.upper()}"
        if env_key in os.environ:
            return os.environ[env_key]
        
        # 从配置文件获取
        return self.config.get(key, default)
```

## 测试架构

### 1. 单元测试架构

#### 模块测试
```python
import unittest
from unittest.mock import Mock, patch

class TestIDConverterModule(unittest.TestCase):
    def setUp(self):
        """测试前准备"""
        self.module = IDConverterModule()
    
    def test_module_config(self):
        """测试模块配置"""
        config = self.module.get_module_config()
        self.assertEqual(config['name'], 'ID值转换器')
        self.assertEqual(config['icon'], '📊')
    
    def test_data_processing(self):
        """测试数据处理"""
        test_data = "123=测试值\n456=另一个值"
        result = self.module.process_data(test_data)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]['ID'], '123')
        self.assertEqual(result[0]['Value'], '测试值')
```

#### 处理器测试
```python
class TestDataProcessors(unittest.TestCase):
    def test_id_value_processor(self):
        """测试ID值处理器"""
        processor = IDValueProcessor()
        
        # 测试有效数据
        valid_data = "123=测试值"
        result = processor.process(valid_data)
        self.assertEqual(len(result), 1)
        
        # 测试无效数据
        invalid_data = "无效数据"
        result = processor.process(invalid_data)
        self.assertEqual(len(result), 0)
```

### 2. 集成测试架构

#### 端到端测试
```python
class TestEndToEnd(unittest.TestCase):
    def test_complete_workflow(self):
        """测试完整工作流程"""
        # 创建应用实例
        app = FExcelApp(Mock())
        
        # 启动模块
        app.start_module('id_converter')
        
        # 验证模块已启动
        self.assertIn('id_converter', app.modules)
        self.assertIsNotNone(app.modules['id_converter'].window)
```

## 部署架构

### 1. 打包架构

#### PyInstaller打包
```bash
# 创建spec文件
pyi-makespec main.py --onefile --windowed --name F-Excel

# 打包应用
pyinstaller F-Excel.spec
```

#### 依赖管理
```python
# requirements.txt
tkinter
pandas>=1.3.0
openpyxl>=3.0.0
pillow>=8.0.0

# 开发依赖
pytest>=6.0.0
pytest-cov>=2.10.0
black>=21.0.0
flake8>=3.8.0
```

### 2. 跨平台支持

#### 平台检测
```python
import platform
import sys

class PlatformManager:
    @staticmethod
    def get_platform():
        """获取当前平台信息"""
        return {
            'system': platform.system(),
            'release': platform.release(),
            'version': platform.version(),
            'machine': platform.machine(),
            'python_version': sys.version
        }
    
    @staticmethod
    def is_windows():
        """是否为Windows系统"""
        return platform.system() == 'Windows'
    
    @staticmethod
    def is_linux():
        """是否为Linux系统"""
        return platform.system() == 'Linux'
    
    @staticmethod
    def is_macos():
        """是否为macOS系统"""
        return platform.system() == 'Darwin'
```

#### 平台特定配置
```python
class PlatformConfig:
    def __init__(self):
        self.platform = PlatformManager.get_platform()
        self.config = self.get_platform_specific_config()
    
    def get_platform_specific_config(self):
        """获取平台特定配置"""
        if PlatformManager.is_windows():
            return {
                'font_family': 'Microsoft YaHei',
                'file_separator': '\\',
                'startup_script': 'start.bat'
            }
        elif PlatformManager.is_linux():
            return {
                'font_family': 'DejaVu Sans',
                'file_separator': '/',
                'startup_script': 'start.sh'
            }
        elif PlatformManager.is_macos():
            return {
                'font_family': 'Helvetica',
                'file_separator': '/',
                'startup_script': 'start.sh'
            }
        else:
            return {
                'font_family': 'Arial',
                'file_separator': '/',
                'startup_script': 'start.sh'
            }
```

## 总结

F-Excel 的架构设计充分体现了现代软件工程的最佳实践：

1. **分层架构**: 清晰的职责分离，便于维护和测试
2. **设计模式**: 合理运用多种设计模式，提高代码质量
3. **模块化设计**: 高度模块化，易于扩展和维护
4. **错误处理**: 完善的错误处理机制，提高系统稳定性
5. **性能优化**: 异步处理、批处理等优化策略
6. **扩展性**: 插件化架构，支持功能扩展
7. **跨平台**: 良好的跨平台支持

这种架构设计使得F-Excel不仅具有强大的功能，还具备了良好的可维护性、可扩展性和可测试性，为未来的功能扩展和性能优化奠定了坚实的基础。
