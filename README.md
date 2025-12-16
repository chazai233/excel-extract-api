# Excel数据提取API服务

为Dify工作流提供Excel文件处理能力的FastAPI服务。

## 功能
- 智能识别Excel格式（边坡检查表 / 测量成果表）
- 自动提取指定列数据
- 生成新的Excel文件（Base64格式）

## 部署到 Railway

[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app/template)

1. 点击上方按钮或访问 [railway.app](https://railway.app)
2. 创建新项目 → 从GitHub部署
3. 连接您的GitHub仓库
4. Railway会自动检测并部署

## 部署到 Render

1. 访问 [render.com](https://render.com)
2. 创建新的 Web Service
3. 连接GitHub仓库
4. 设置：
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `uvicorn excel_api_service:app --host 0.0.0.0 --port $PORT`

## API使用

### 健康检查
```
GET /
```

### 提取Excel数据
```
POST /extract
Content-Type: application/json

{
  "file_base64": "base64编码的Excel文件内容"
}
```

### 响应
```json
{
  "success": true,
  "format_type": "边坡检查表",
  "row_count": 100,
  "file_base64": "提取后的Excel文件base64",
  "filename": "提取数据_边坡检查表_20241216.xlsx",
  "message": "识别为【边坡检查表】，成功提取 100 行数据"
}
```

## Dify工作流配置

导入工作流后，将API地址设置为您的部署URL，例如：
- Railway: `https://your-app.up.railway.app`
- Render: `https://your-app.onrender.com`
