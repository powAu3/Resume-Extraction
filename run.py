# -*- coding: utf-8 -*-
import requests
import json
import os
import sys
from typing import List, Dict, Any, Optional


# ========== 配置加载 ==========

def load_config():
    """从 config.json 加载配置"""
    config_path = os.path.join(os.path.dirname(__file__), "config.json")
    
    if not os.path.exists(config_path):
        print("❌ 配置文件不存在，请复制 config.example.json 为 config.json 并填写配置")
        sys.exit(1)
    
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

# 加载配置
CONFIG = load_config()
TOKEN = CONFIG["token"]
WORKFLOW_ID = CONFIG["workflow_id"]
LOCAL_FILES = CONFIG.get("local_files", [])

# API 请求头
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}


# ========== 数据解析 ==========

class ResumeParser:
    
    #数据解析
   
    @staticmethod
    def smart_json_parse(value: Any) -> Any:
        """
        递归解析嵌套的 JSON 字符串
        
        Coze 返回的数据可能被多次 JSON 序列化，
        这个方法会一直解析直到得到最终的数据结构。
        """
        while isinstance(value, str):
            try:
                value = json.loads(value)
            except json.JSONDecodeError:
                break
        return value
    
    @staticmethod
    def parse_api_response(response_data: Dict) -> List[Dict]:
        """
        解析 API 响应数据
        
        Coze 返回的数据结构较为复杂，需要逐层解析：
        response -> data(str) -> data(list) -> output(str) -> 实际数据
        """
        resumes = []
        
        try:
            # 获取外层 data 字段（字符串形式）
            data_str = response_data.get("data", "{}")
            data = ResumeParser.smart_json_parse(data_str)
            
            # 获取内层 data 数组
            if isinstance(data, dict):
                data_array = data.get("data", [])
            else:
                data_array = data if isinstance(data, list) else []
            
            data_array = ResumeParser.smart_json_parse(data_array)
            
            # 遍历每份简历数据
            if isinstance(data_array, list):
                for item in data_array:
                    item = ResumeParser.smart_json_parse(item)
                    if isinstance(item, dict) and "output" in item:
                        output = ResumeParser.smart_json_parse(item["output"])
                        if isinstance(output, dict):
                            resumes.append(output)
                    elif isinstance(item, dict):
                        resumes.append(item)
                        
        except Exception as e:
            print(f"解析数据时出错: {e}")
            
        return resumes
    
    @staticmethod
    def format_resume(resume: Dict) -> Dict:
        """
        格式化简历数据，统一字段名称并设置默认值
        
        AI 解析的字段名可能不一致（如"年纪"和"年龄"），
        这里做统一处理，确保后续使用时不会出错。
        """
        return {
            "姓名": resume.get("姓名", "未知"),
            "性别": resume.get("性别", "未知"),
            "年龄": resume.get("年纪") or resume.get("年龄", "未知"),
            "出生日期": resume.get("出生日期", "未知"),
            "最高学历": resume.get("最高学历", "未知"),
            "婚配情况": resume.get("婚配情况", "未知"),
            "就读院校": resume.get("就读院校", []),
            "发表论文情况": resume.get("发表论文情况", []),
            "获批项目情况": resume.get("获批项目情况", []),
            "获奖情况": resume.get("获奖情况", []),
            "其他成果": resume.get("其他成果", []),
            "著作情况": resume.get("著作情况", [])
        }


# ========== 文件上传 ==========

def upload_file(file_path: str) -> str:
    """
    上传文件到 Coze 平台
    
    Args:
        file_path: 本地文件路径
        
    Returns:
        上传成功后返回的 file_id
        
    Raises:
        FileNotFoundError: 文件不存在
        Exception: 上传失败
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    print(f"📤 上传文件: {os.path.basename(file_path)}")
    
    with open(file_path, "rb") as f:
        files = {"file": (os.path.basename(file_path), f)}
        resp = requests.post(
            "https://api.coze.cn/v1/files/upload",
            headers={"Authorization": f"Bearer {TOKEN}"},
            files=files
        )
    
    if resp.status_code != 200:
        raise Exception(f"上传失败: {resp.status_code} - {resp.text}")
    
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"Coze 返回错误: {data}")
    
    file_id = data["data"]["id"]
    print(f"获取 file_id: {file_id}")
    return file_id


def build_file_param(file_id: str) -> str:
    """构建 Coze 工作流需要的文件参数格式"""
    return json.dumps({"file_id": file_id}, ensure_ascii=False)


# ========== 工作流调用 ==========

def run_workflow_sync(file_ids: List[str]) -> Optional[List[Dict]]:
    """
    调用 Coze 工作流解析简历
    
    使用同步模式调用，等待工作流执行完成后返回结果。
    通常需要 几分钟完成解析。
    
    Args:
        file_ids: 已上传文件的 ID 列表
        
    Returns:
        解析后的简历数据列表，失败返回 None
    """
    # 构建请求参数
    jianli_params = [build_file_param(fid) for fid in file_ids]
    payload = {
        "workflow_id": WORKFLOW_ID,
        "parameters": {"jianli": jianli_params}
    }

    print("\n⏳ 正在调用 Coze 工作流（同步模式，可能需要1-2分钟）...")
    
    try:
        resp = requests.post(
            "https://api.coze.cn/v1/workflow/run",
            headers=HEADERS,
            json=payload,
            timeout=600  # 10分钟超时
        )
        resp.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return None

    result = resp.json()
    
    # 保存原始响应便于调试
    with open("response_debug.json", "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    print("📝 原始响应已保存至: response_debug.json")
    
    if result.get("code") != 0:
        print(f"工作流返回错误: {result.get('msg', '未知错误')}")
        return None

    # 解析并格式化数据
    parser = ResumeParser()
    resumes = parser.parse_api_response(result)
    
    if not resumes:
        print("未能解析出简历数据")
        return None
    
    formatted_resumes = [parser.format_resume(r) for r in resumes]
    
    print(f"\n成功解析 {len(formatted_resumes)} 份简历")
    
    # 保存格式化后的数据
    with open("formatted_resumes.json", "w", encoding="utf-8") as f:
        json.dump(formatted_resumes, f, indent=2, ensure_ascii=False)
    print(f"格式化数据已保存至: {os.path.abspath('formatted_resumes.json')}")
    
    return formatted_resumes


def parse_from_response_file(file_path: str) -> Optional[List[Dict]]:
    """
    从本地文件解析数据（测试用）
    
    当已经有保存的 API 响应时，可以用这个方法直接解析，
    避免重复调用 API。
    """
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 尝试解析 JSON
        try:
            result = json.loads(content)
        except json.JSONDecodeError:
            # 兼容 Python dict 格式
            result = eval(content)
        
        parser = ResumeParser()
        resumes = parser.parse_api_response(result)
        
        if resumes:
            return [parser.format_resume(r) for r in resumes]
            
    except Exception as e:
        print(f"解析文件失败: {e}")
    
    return None


# ========== 辅助函数 ==========

def print_resume_summary(resumes: List[Dict]):
    """打印简历摘要信息"""
    for i, resume in enumerate(resumes, 1):
        print(f"\n--- 第 {i} 份简历 ---")
        print(f"姓名: {resume['姓名']}")
        print(f"性别: {resume['性别']}")
        print(f"年龄: {resume['年龄']}")
        print(f"最高学历: {resume['最高学历']}")
        print(f"论文数量: {len(resume['发表论文情况'])} 类")
        print(f"项目数量: {len(resume['获批项目情况'])} 项")
        print(f"获奖数量: {len(resume['获奖情况'])} 项")


def generate_ppt(resumes: List[Dict]):
    """调用 PPT 渲染器生成演示文稿"""
    try:
        from ppt_renderer import PPTRenderer
        renderer = PPTRenderer()
        output_path = renderer.render_all(resumes)
        print(f"\n🎨 PPT 已生成: {output_path}")
    except ImportError:
        print("\n💡 提示: 确保 ppt_renderer.py 在同目录下")


# ========== 主程序入口 ==========

def main():
    """程序入口"""
    
    # 模式1: 从本地文件解析（测试用）
    if len(sys.argv) > 1 and sys.argv[1] == "--from-file":
        response_file = sys.argv[2] if len(sys.argv) > 2 else "response.txt"
        print(f"📂 从文件解析模式: {response_file}")
        
        resumes = parse_from_response_file(response_file)
        
        if resumes:
            print(f"\n成功解析 {len(resumes)} 份简历!")
            print_resume_summary(resumes)
            
            # 保存数据
            with open("formatted_resumes.json", "w", encoding="utf-8") as f:
                json.dump(resumes, f, indent=2, ensure_ascii=False)
            print(f"\n数据已保存至: formatted_resumes.json")
            
            # 生成 PPT
            generate_ppt(resumes)
        else:
            print("解析失败")
        return
    
    # 模式2: 完整流程（上传文件 + 调用API + 生成PPT）
    for fp in LOCAL_FILES:
        if not os.path.exists(fp):
            raise FileNotFoundError(f"请先设置正确的本地文件路径: {fp}")

    try:
        # 上传文件
        file_ids = [upload_file(fp) for fp in LOCAL_FILES]

        # 调用工作流
        resumes = run_workflow_sync(file_ids)

        if resumes:
            print("\n任务成功完成！")
            print_resume_summary(resumes)
            generate_ppt(resumes)
        else:
            print("\n任务未成功完成，请检查错误信息。")

    except Exception as e:
        print(f"\n程序异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
