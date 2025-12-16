"""
Excel数据提取FastAPI服务
用于Dify工作流调用，处理Excel文件并提取指定列数据
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import pandas as pd
from io import BytesIO
import base64
import requests
from datetime import datetime
from typing import Optional

app = FastAPI(
    title="Excel数据提取服务",
    description="为Dify工作流提供Excel文件处理能力",
    version="1.0.0"
)

# 添加CORS支持
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ExtractRequest(BaseModel):
    """提取请求模型"""
    file_url: str  # 文件下载URL
    file_base64: Optional[str] = None  # 或者直接传base64内容


class ExtractResponse(BaseModel):
    """提取响应模型"""
    success: bool
    format_type: str
    row_count: int
    file_base64: str
    filename: str
    message: str
    column_names: list = []


@app.get("/")
async def root():
    """健康检查"""
    return {"status": "ok", "service": "Excel数据提取服务"}


@app.post("/extract", response_model=ExtractResponse)
async def extract_excel(request: ExtractRequest):
    """
    智能识别Excel格式并提取指定列数据
    支持两种格式：
    - 格式1: 边坡检查表 (标准表头在第1行)
    - 格式2: 测量成果表 (表头包含Survey results)
    """
    try:
        # 获取文件内容
        if request.file_base64:
            file_content = base64.b64decode(request.file_base64)
        elif request.file_url:
            response = requests.get(request.file_url, timeout=30)
            if response.status_code != 200:
                raise HTTPException(
                    status_code=400,
                    detail=f"下载文件失败: HTTP {response.status_code}"
                )
            file_content = response.content
        else:
            raise HTTPException(status_code=400, detail="请提供file_url或file_base64")

        # 先读取原始数据判断格式
        df_raw = pd.read_excel(BytesIO(file_content), header=None)
        first_cell = str(df_raw.iloc[0, 0]) if len(df_raw) > 0 else ""

        # ======== 格式2: 测量成果表 ========
        if "Survey results" in first_cell or "测量成果表" in first_cell:
            df = pd.read_excel(BytesIO(file_content), header=None, skiprows=7)
            df = df[pd.to_numeric(df.iloc[:, 2], errors='coerce').notna()]
            target_columns = [1, 2, 3, 4]
            extracted_df = df.iloc[:, target_columns].copy()
            extracted_df.columns = ['测点编号', 'X坐标', 'Y坐标', '高程H']
            format_type = "测量成果表"

        # ======== 格式1: 边坡检查表 ========
        else:
            df = pd.read_excel(BytesIO(file_content), header=0)
            columns = df.columns.tolist()
            # 索引: 0(线路名), 5(超欠挖), 8(实测X), 9(实测Y), 10(实测Z), 11(里程), 12(偏距), 13(设计标高)
            target_columns = [0, 5, 8, 9, 10, 11, 12, 13]
            valid_indices = [i for i in target_columns if i < len(columns)]
            extracted_df = df.iloc[:, valid_indices].copy()
            new_column_names = ['线路名', '超欠挖', '实测X或里程', '实测Y或偏距',
                               '实测Z坐标', '里程', '偏距', '设计标高']
            extracted_df.columns = new_column_names[:len(valid_indices)]
            format_type = "边坡检查表"

        if len(extracted_df) == 0:
            return ExtractResponse(
                success=False,
                format_type=format_type,
                row_count=0,
                file_base64="",
                filename="",
                message="未能提取到任何数据",
                column_names=[]
            )

        # 生成新Excel文件
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            extracted_df.to_excel(writer, index=False, sheet_name='提取数据')

        excel_bytes = output.getvalue()
        file_base64 = base64.b64encode(excel_bytes).decode('utf-8')

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"提取数据_{format_type}_{timestamp}.xlsx"

        return ExtractResponse(
            success=True,
            format_type=format_type,
            row_count=len(extracted_df),
            file_base64=file_base64,
            filename=filename,
            message=f"识别为【{format_type}】，成功提取 {len(extracted_df)} 行数据",
            column_names=extracted_df.columns.tolist()
        )

    except Exception as e:
        return ExtractResponse(
            success=False,
            format_type="未知",
            row_count=0,
            file_base64="",
            filename="",
            message=f"解析Excel失败: {str(e)}",
            column_names=[]
        )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8100)
