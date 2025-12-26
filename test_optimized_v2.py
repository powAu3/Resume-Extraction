# -*- coding: utf-8 -*-
"""
测试优化模板渲染器V2
"""
import json
import os
from optimized_renderer_v2 import OptimizedTemplateRendererV2


def main():
    """主测试函数"""
    # 设置路径
    template_path = os.path.join(os.path.dirname(__file__), "副本人才引进ppt.pptx")
    data_path = os.path.join(os.path.dirname(__file__), "formatted_resumes.json")
    
    # 加载数据
    print(f"📂 加载数据文件: {data_path}")
    with open(data_path, 'r', encoding='utf-8') as f:
        resumes = json.load(f)
    print(f"✅ 加载了 {len(resumes)} 份简历数据\n")
    
    # 创建渲染器
    print(f"📄 使用模板: {template_path}\n")
    renderer = OptimizedTemplateRendererV2(template_path)
    
    try:
        # 渲染
        output_path = renderer.render_all(resumes)
        print(f"\n🎉 测试完成！输出文件: {output_path}")
        
    except Exception as e:
        print(f"\n❌ 渲染失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

