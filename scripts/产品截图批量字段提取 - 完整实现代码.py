# ============================================================
# 产品截图批量字段提取工具
# 功能：读取本地图片文件夹中的产品截图，调用视觉大模型API
#       提取产品名称、配料、烹饪方式等结构化字段，导出Excel
# 作者：根据业务需求定制
# ============================================================


# ====================================================================
# 第一部分：导入依赖库
# 说明：这些是程序运行所需的工具包，首次使用前需要安装
# 安装命令（在终端/命令行运行）：
#   pip install openai pandas tqdm openpyxl Pillow
# ====================================================================

import os                     # 操作系统接口：用于读取环境变量、拼接路径
import json                   # JSON处理：用于解析大模型返回的结构化数据
import base64                 # Base64编码：将图片二进制数据转成文本，API传图用
import time                   # 时间控制：用于设置请求间隔，防止触发API限速
import re                     # 正则表达式：用于清洗模型返回文本中的多余字符
from pathlib import Path      # 路径处理：比os.path更现代，自动兼容Win/Mac/Linux

import pandas as pd           # 数据处理：将提取结果整理成表格并导出Excel/CSV
from tqdm import tqdm         # 进度条：批量处理时显示当前进度，防止以为程序卡住
from PIL import Image         # 图片处理：用于压缩图片尺寸，降低API调用费用
import io                     # 字节流处理：配合Pillow在内存中操作图片，不产生临时文件

# 注意：openai库是通用客户端，大多数国内模型也兼容其接口格式
from openai import OpenAI     # OpenAI官方客户端（通义千问/智谱等也用这个库）


# ====================================================================
# 第二部分：配置区域 ⚠️ 这里是最需要根据自己情况修改的地方
# ====================================================================

# --------------------------------------------------------------------
# 2.1 模型API配置
# 说明：下面列出了4种常见方案，取消对应注释（删掉#号）即可切换
# --------------------------------------------------------------------

# ✅ 方案A：OpenAI GPT-4o（最强，需要境外网络）
#    申请地址：https://platform.openai.com/
API_KEY    = "sk-xxxxxxxxxxxxxxxxxxxxxxxx"      # ⚠️ 替换成你自己的 API Key
BASE_URL   = "https://api.openai.com/v1"        # OpenAI 官方接口地址（固定）
MODEL_NAME = "gpt-4o"                           # 使用 GPT-4o 视觉模型

# ✅ 方案B：阿里云通义千问VL（国内推荐，有免费额度）
#    申请地址：https://dashscope.aliyun.com/
# API_KEY    = "sk-xxxxxxxxxxxxxxxxxxxxxxxx"     # ⚠️ 替换成你的 DashScope Key
# BASE_URL   = "https://dashscope.aliyuncs.com/compatible-mode/v1"  # 阿里接口地址
# MODEL_NAME = "qwen-vl-plus"                    # 可选：qwen-vl-max（更准但更贵）

# ✅ 方案C：字节跳动火山引擎豆包视觉模型
#    申请地址：https://www.volcengine.com/product/ark
# API_KEY    = "xxxxxxxxxxxxxxxxxxxxxxxx"        # ⚠️ 替换成你的火山引擎 Key
# BASE_URL   = "https://ark.cn-beijing.volces.com/api/v3"
# MODEL_NAME = "doubao-vision-pro-32k"           # 豆包视觉Pro版

# ✅ 方案D：智谱AI GLM-4V（有免费版）
#    申请地址：https://open.bigmodel.cn/
# API_KEY    = "xxxxxxxxxxxxxxxxxxxxxxxx"        # ⚠️ 替换成你的智谱 Key
# BASE_URL   = "https://open.bigmodel.cn/api/paas/v4/"
# MODEL_NAME = "glm-4v-flash"                    # flash版免费，plus版更准确


# --------------------------------------------------------------------
# 2.2 图片目录配置
# ⚠️ 修改为你存放产品截图的文件夹路径
# --------------------------------------------------------------------

IMAGE_FOLDER = "./product_images"               # 图片文件夹（相对当前脚本的路径）
# IMAGE_FOLDER = "C:/Users/你的用户名/Desktop/产品图片"  # Windows 绝对路径示例
# IMAGE_FOLDER = "/Users/yourname/Downloads/产品图片"    # macOS 绝对路径示例

# ⚠️ 支持处理的图片格式，不需要的格式可以删除
IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".webp", ".bmp"]


# --------------------------------------------------------------------
# 2.3 输出文件配置
# ⚠️ 修改结果文件的保存位置和格式
# --------------------------------------------------------------------

OUTPUT_FILE = "./output_products.xlsx"          # Excel格式输出路径
# OUTPUT_FILE = "./output_products.csv"          # 如需CSV格式，改用这行


# --------------------------------------------------------------------
# 2.4 请求控制参数
# ⚠️ 如果遇到"请求太频繁"报错，可以增大 REQUEST_INTERVAL
# --------------------------------------------------------------------

REQUEST_INTERVAL = 1      # 两次API请求之间的间隔秒数（免费版建议设2~3秒）
MAX_RETRY        = 3      # 单张图片失败后的最大重试次数
IMAGE_MAX_SIZE   = 1024   # 图片最长边压缩到该像素数（压缩可显著减少费用）


# --------------------------------------------------------------------
# 2.5 提取字段配置
# ⚠️ 根据业务需要增删字段
# 格式："英文字段名（Excel列名）": "中文描述（告诉模型提取什么）"
# --------------------------------------------------------------------

EXTRACT_FIELDS = {
    "product_name"      : "产品名称",
    "brand"             : "品牌名称",
    "net_weight"        : "净含量/规格（含单位，如180g）",
    "ingredients"       : "配料表（完整列出所有配料）",
    "cooking_method"    : "烹饪方式（如煎、炒、煮、烤等，多种用顿号分隔）",
    "storage_condition" : "储存条件（如冷冻、冷藏、常温）",
    "shelf_life"        : "保质期（含单位，如12个月）",
    "origin"            : "产地",
    "manufacturer"      : "生产单位/厂商名称",
    "price"             : "产品价格（如有，含货币符号）",
}


# ====================================================================
# 第三部分：Prompt模板构建
# 说明：Prompt是发给大模型的"指令"，质量直接影响提取准确率
# ====================================================================

def build_prompt(fields: dict) -> str:
    """
    根据字段配置动态生成提取Prompt。
    这样做的好处是：修改 EXTRACT_FIELDS 后，Prompt 会自动更新，无需手动改两处。

    参数：
        fields: 字段字典，key为英文字段名，value为中文描述
    返回：
        完整的Prompt字符串
    """
    # 将字段字典转换成 "- 英文名: 中文描述" 的格式，方便模型理解
    fields_text = "\n".join([f"- {k}: {v}" for k, v in fields.items()])

    # f-string 多行字符串，{fields_text} 会被替换成上面生成的字段列表
    prompt = f"""请仔细分析这张产品图片，提取以下字段的信息，并严格以JSON格式输出：

{fields_text}

输出要求：
1. 只输出纯JSON，不要任何额外说明文字或markdown代码块标记
2. 图片中找不到的字段，值填写 null
3. 配料表请完整列出，不要省略
4. 所有字段值使用字符串类型

输出示例格式：
{{"product_name": "嫩滑牛肉片", "brand": "随滋", "net_weight": "180g", ...}}"""

    return prompt


# ====================================================================
# 第四部分：图片预处理
# 说明：将图片压缩并编码为Base64字符串，用于API传输
# ====================================================================

def preprocess_image(image_path: str, max_size: int = IMAGE_MAX_SIZE) -> str:
    """
    读取图片文件，压缩尺寸，转换为Base64编码字符串。

    为什么要压缩？
        - 大模型API按token收费，图片越大消耗token越多
        - 1024px 对于文字识别已经足够清晰
        - 压缩后速度也更快

    参数：
        image_path: 图片文件的完整路径
        max_size:   图片最长边的最大像素数
    返回：
        Base64编码的字符串
    """
    # 用Pillow打开图片文件
    img = Image.open(image_path)

    # 将图片统一转换为RGB模式（PNG可能是RGBA，API通常只接受RGB/JPEG）
    if img.mode != "RGB":
        img = img.convert("RGB")

    # 获取原始宽高
    width, height = img.size

    # 如果图片尺寸超过限制，按比例缩小（保持长宽比不变）
    if max(width, height) > max_size:
        # 计算缩放比例：用最长边除以最大允许值
        scale = max_size / max(width, height)
        # 计算新的宽高，int()取整
        new_width  = int(width  * scale)
        new_height = int(height * scale)
        # LANCZOS是高质量的缩放算法，适合文字图片
        img = img.resize((new_width, new_height), Image.LANCZOS)

    # 创建一个内存字节流（相当于内存中的"虚拟文件"，避免写入磁盘）
    buffer = io.BytesIO()
    # 将图片以JPEG格式保存到内存字节流中，quality=90平衡质量和体积
    img.save(buffer, format="JPEG", quality=90)
    # 将字节流的读取位置移回开头（否则读出来是空的）
    buffer.seek(0)

    # 将字节数据编码为Base64字符串（API要求的格式）
    # .decode("utf-8") 将字节型转换为普通字符串
    b64_string = base64.b64encode(buffer.read()).decode("utf-8")

    return b64_string


# ====================================================================
# 第五部分：调用大模型API提取字段
# ====================================================================

def extract_fields_from_image(
    client: OpenAI,
    image_path: str,
    prompt: str,
    model: str = MODEL_NAME
) -> dict:
    """
    将单张图片发送给大模型，获取结构化字段数据。

    参数：
        client:     OpenAI客户端实例（已配置好API Key和Base URL）
        image_path: 图片文件路径
        prompt:     提取指令（由build_prompt生成）
        model:      使用的模型名称
    返回：
        包含提取字段的字典；失败时返回含错误信息的字典
    """
    # 第一步：对图片进行预处理，获得Base64编码
    b64_image = preprocess_image(image_path)

    # 第二步：构建API请求的消息体
    # 视觉模型支持在同一条消息中混合文字和图片
    messages = [
        {
            "role": "user",            # role固定填"user"，表示这是用户发的消息
            "content": [
                {
                    "type": "text",    # 第一部分：文字指令（Prompt）
                    "text": prompt
                },
                {
                    "type": "image_url",   # 第二部分：图片数据
                    "image_url": {
                        # data:image/jpeg;base64, 是固定前缀，告诉API这是Base64图片
                        "url": f"data:image/jpeg;base64,{b64_image}",
                        "detail": "high"   # ⚠️ 可改为"low"降低费用，但识别精度会下降
                    }
                }
            ]
        }
    ]

    # 第三步：发起API请求
    response = client.chat.completions.create(
        model=model,        # 使用配置的模型名称
        messages=messages,  # 发送构建好的消息
        max_tokens=2000,    # ⚠️ 最大返回token数，配料表长的产品可适当增大
        temperature=0,      # 温度=0表示输出最确定的结果，适合信息提取任务
        # ⚠️ 部分模型支持 response_format={"type": "json_object"} 强制JSON输出
        # 如果模型支持，取消下面这行的注释可以让输出更稳定：
        # response_format={"type": "json_object"},
    )

    # 第四步：从响应中提取文本内容
    # response.choices[0] 取第一个回答
    # .message.content 取消息的文字内容
    raw_text = response.choices[0].message.content.strip()

    # 第五步：解析JSON文本
    # 模型有时会在JSON外面包裹 ```json ... ``` 代码块标记，需要先清洗
    # 使用正则表达式提取花括号 {} 之间的内容
    json_match = re.search(r'\{.*\}', raw_text, re.DOTALL)
    if json_match:
        # 找到了JSON内容，提取出来
        json_text = json_match.group()
    else:
        # 没找到JSON格式的内容，直接使用原始文本（后续会报错并被捕获）
        json_text = raw_text

    # 将JSON字符串转换为Python字典
    result = json.loads(json_text)

    return result


# ====================================================================
# 第六部分：带重试机制的安全调用封装
# 说明：网络抖动或API偶尔报错时，自动重试，避免整个任务中断
# ====================================================================

def safe_extract(
    client: OpenAI,
    image_path: str,
    prompt: str,
    max_retry: int = MAX_RETRY
) -> dict:
    """
    带重试机制的字段提取函数。
    失败时会等待后重试，最终失败会返回包含错误信息的字典（而非抛出异常）。

    参数：
        client:     OpenAI客户端实例
        image_path: 图片路径
        prompt:     提取指令
        max_retry:  最大重试次数
    返回：
        成功：字段字典；失败：含 _error 键的字典
    """
    last_error = None  # 记录最后一次错误信息

    # 循环尝试，最多尝试 max_retry 次
    for attempt in range(1, max_retry + 1):
        try:
            # 尝试提取字段
            result = extract_fields_from_image(client, image_path, prompt)
            # 成功则直接返回结果
            return result

        except json.JSONDecodeError as e:
            # JSON解析失败：模型没有按要求返回JSON格式
            last_error = f"JSON解析失败（第{attempt}次）：{str(e)}"
            print(f"  ⚠️  {last_error}，等待重试...")

        except Exception as e:
            # 其他错误：网络超时、API限速、Key无效等
            last_error = f"API调用失败（第{attempt}次）：{str(e)}"
            print(f"  ❌  {last_error}，等待重试...")

        # 重试前等待，等待时间随重试次数增加（指数退避策略）
        # 第1次失败等2秒，第2次等4秒，第3次等8秒
        wait_time = 2 ** attempt
        time.sleep(wait_time)

    # 所有重试都失败，返回包含错误信息的字典
    # 这样失败的图片也会出现在结果表格中，方便后续排查
    return {"_error": last_error}


# ====================================================================
# 第七部分：批量处理主函数
# 说明：扫描图片目录，逐一处理每张图片，汇总结果并保存
# ====================================================================

def batch_process(
    image_folder: str = IMAGE_FOLDER,
    output_file:  str = OUTPUT_FILE
):
    """
    批量处理图片文件夹中的所有产品截图。

    参数：
        image_folder: 图片文件夹路径
        output_file:  结果文件保存路径
    """

    # ---------- 7.1 初始化 ----------

    # 创建OpenAI客户端，注入API Key和Base URL
    # 注意：即使用的是国内模型（如通义千问），也是用这个客户端，因为它们兼容OpenAI接口
    client = OpenAI(
        api_key=API_KEY,
        base_url=BASE_URL
    )

    # 根据字段配置动态构建Prompt
    prompt = build_prompt(EXTRACT_FIELDS)

    # ---------- 7.2 扫描图片文件 ----------

    # 将字符串路径转为Path对象，方便后续操作
    folder = Path(image_folder)

    # 检查目录是否存在
    if not folder.exists():
        print(f"❌ 错误：找不到图片目录 {image_folder}")
        print("请检查路径是否正确，或先创建该目录并放入图片")
        return

    # 扫描目录下所有符合扩展名的图片文件
    # folder.rglob("*") 会递归搜索所有子目录，如只想搜索当前层用 folder.glob("*")
    image_files = [
        f for f in folder.rglob("*")           # 递归获取所有文件
        if f.suffix.lower() in IMAGE_EXTENSIONS # 只保留指定格式的图片
    ]

    # 按文件名排序，方便追踪
    image_files.sort()

    if not image_files:
        print(f"⚠️  在 {image_folder} 中没有找到任何图片文件")
        print(f"支持的格式：{IMAGE_EXTENSIONS}")
        return

    print(f"\n📂 找到 {len(image_files)} 张图片，开始处理...\n")

    # ---------- 7.3 逐张处理图片 ----------

    results = []       # 存储所有图片的提取结果
    success_count = 0  # 成功计数
    fail_count    = 0  # 失败计数

    # tqdm 会将 image_files 包装成带进度条的迭代器
    # desc= 是进度条左侧的标签文字
    for image_path in tqdm(image_files, desc="处理进度"):

        # 调用带重试的提取函数
        extracted = safe_extract(client, str(image_path), prompt)

        # 无论成功失败，都在结果中记录源文件名，方便追溯
        extracted["_source_file"] = image_path.name

        # 判断是否提取成功（成功的结果中没有 _error 键）
        if "_error" not in extracted:
            success_count += 1
        else:
            fail_count += 1
            # 打印失败信息（tqdm.write 不会破坏进度条显示）
            tqdm.write(f"  ❌ 失败：{image_path.name} → {extracted['_error']}")

        # 将结果追加到列表
        results.append(extracted)

        # 每次请求后等待设定的间隔时间，防止触发API限速
        time.sleep(REQUEST_INTERVAL)

    # ---------- 7.4 保存结果 ----------

    print(f"\n\n📊 处理完成：成功 {success_count} 张，失败 {fail_count} 张")
    print(f"💾 正在保存结果到：{output_file}")

    # 将结果列表转换为DataFrame（表格）
    df = pd.DataFrame(results)

    # 将 _source_file 列移到最前面，方便查看
    cols = ["_source_file"] + [c for c in df.columns if c != "_source_file"]
    df = df[cols]

    # 根据输出文件扩展名选择保存格式
    if output_file.endswith(".xlsx"):
        # 保存为Excel，index=False 表示不保存行号
        df.to_excel(output_file, index=False, engine="openpyxl")
    elif output_file.endswith(".csv"):
        # 保存为CSV，encoding确保中文不乱码
        df.to_csv(output_file, index=False, encoding="utf-8-sig")
    else:
        # 默认保存为CSV
        df.to_csv(output_file, index=False, encoding="utf-8-sig")

    print(f"✅ 已成功保存！共 {len(df)} 条记录")

    # ---------- 7.5 打印预览 ----------
    print("\n📋 结果预览（前3条）：")
    # 只显示主要字段列，避免打印太多列
    preview_cols = [c for c in ["_source_file", "product_name", "brand",
                                "cooking_method", "price"] if c in df.columns]
    print(df[preview_cols].head(3).to_string(index=False))


# ====================================================================
# 第八部分：程序入口
# 说明：if __name__ == "__main__" 确保直接运行脚本时才执行
#       被其他脚本import时不会自动执行，是Python的最佳实践
# ====================================================================

if __name__ == "__main__":
    batch_process()