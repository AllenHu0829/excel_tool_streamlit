# Excel工具 - Streamlit部署指南

## 概述

这是Excel文件拆分与合并工具的Streamlit Web版本，可以通过浏览器访问使用。

## 本地运行

### 1. 安装依赖

```bash
pip install -r requirements_streamlit.txt
```

### 2. 运行应用

```bash
streamlit run streamlit_app.py
```

应用将在浏览器中自动打开，默认地址为 `http://localhost:8501`

## 部署到Streamlit Cloud

### 方法一：通过GitHub部署（推荐）

1. **创建GitHub仓库**
   - 在GitHub上创建一个新仓库
   - 将以下文件上传到仓库：
     - `streamlit_app.py`
     - `requirements_streamlit.txt`
     - `README.md`（可选）

2. **部署到Streamlit Cloud**
   - 访问 [Streamlit Cloud](https://streamlit.io/cloud)
   - 使用GitHub账号登录
   - 点击 "New app"
   - 选择你的GitHub仓库
   - 设置：
     - Main file path: `streamlit_app.py`
     - Python version: 3.8 或更高
   - 点击 "Deploy"

3. **等待部署完成**
   - Streamlit会自动安装依赖并部署应用
   - 部署完成后会提供一个公开URL

### 方法二：使用Streamlit Sharing

1. **准备文件**
   - 确保 `streamlit_app.py` 和 `requirements_streamlit.txt` 在GitHub仓库中
   - 确保仓库是公开的（或使用Streamlit Sharing的私有仓库功能）

2. **申请Streamlit Sharing**
   - 访问 https://share.streamlit.io
   - 使用GitHub账号登录
   - 填写申请表单

3. **部署**
   - 在Streamlit Sharing控制台中点击 "New app"
   - 选择仓库和文件路径
   - 点击部署

## 部署到其他平台

### 部署到Heroku

1. **创建Procfile**
   ```
   web: streamlit run streamlit_app.py --server.port=$PORT --server.address=0.0.0.0
   ```

2. **创建setup.sh**（可选）
   ```bash
   mkdir -p ~/.streamlit/
   echo "\
   [server]\n\
   headless = true\n\
   port = $PORT\n\
   enableCORS = false\n\
   " > ~/.streamlit/config.toml
   ```

3. **部署**
   ```bash
   heroku create your-app-name
   git push heroku main
   ```

### 部署到Docker

1. **创建Dockerfile**
   ```dockerfile
   FROM python:3.9-slim
   
   WORKDIR /app
   
   COPY requirements_streamlit.txt .
   RUN pip install --no-cache-dir -r requirements_streamlit.txt
   
   COPY streamlit_app.py .
   
   EXPOSE 8501
   
   HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health
   
   ENTRYPOINT ["streamlit", "run", "streamlit_app.py", "--server.port=8501", "--server.address=0.0.0.0"]
   ```

2. **构建和运行**
   ```bash
   docker build -t excel-tool-streamlit .
   docker run -p 8501:8501 excel-tool-streamlit
   ```

### 部署到AWS/Azure/GCP

1. **使用容器服务**（如AWS ECS、Azure Container Instances、GCP Cloud Run）
   - 按照Docker方式构建镜像
   - 上传到容器注册表
   - 部署容器服务

2. **使用虚拟机**
   - 在虚拟机上安装Python和依赖
   - 运行 `streamlit run streamlit_app.py`
   - 配置防火墙和反向代理（如Nginx）

## 配置说明

### 环境变量

如果需要配置，可以在部署平台设置环境变量：

```bash
# Streamlit配置可以通过.streamlit/config.toml或环境变量设置
STREAMLIT_SERVER_PORT=8501
STREAMLIT_SERVER_ADDRESS=0.0.0.0
STREAMLIT_SERVER_HEADLESS=true
```

### 配置文件

创建 `.streamlit/config.toml`：

```toml
[server]
port = 8501
address = "0.0.0.0"
headless = true
enableCORS = false

[browser]
gatherUsageStats = false
```

## 功能说明

### 拆分功能
- 上传一个Excel文件
- 按行拆分成多个文件
- 每个文件包含表头和一行数据
- 下载ZIP压缩包包含所有拆分文件

### 合并功能
- 上传多个Excel文件（可多选）
- 合并所有文件的数据
- 添加"源文件"列追踪数据来源
- 下载合并后的Excel文件

## 注意事项

1. **文件大小限制**
   - Streamlit默认文件上传限制为200MB
   - 可以在`.streamlit/config.toml`中调整：
     ```toml
     [server]
     maxUploadSize = 500
     ```

2. **临时文件**
   - 应用使用临时文件处理上传的文件
   - 文件处理完成后会自动清理
   - 不会在服务器上永久存储用户文件

3. **并发处理**
   - 多用户同时使用时可能影响性能
   - 建议限制并发连接数或使用负载均衡

4. **安全性**
   - 部署到公开平台时注意文件内容安全
   - 建议添加身份验证（Streamlit支持多种认证方式）

## 故障排除

### 常见问题

1. **导入错误**
   - 确保所有依赖都已安装
   - 检查Python版本（需要3.8+）

2. **文件上传失败**
   - 检查文件大小是否超过限制
   - 检查文件格式是否正确

3. **部署失败**
   - 检查`requirements_streamlit.txt`中的依赖版本
   - 查看部署日志获取详细错误信息

## 支持

如有问题，请查看：
- [Streamlit文档](https://docs.streamlit.io/)
- [Streamlit Cloud文档](https://docs.streamlit.io/streamlit-cloud)

