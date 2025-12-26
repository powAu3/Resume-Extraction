# -*- coding: utf-8 -*-
"""
极限测试数据生成器
生成各种极端情况的数据来测试PPT渲染的健壮性
"""
import json
import random

def generate_long_text(length=200):
    """生成指定长度的文本"""
    return "A" * length

def generate_extreme_resume(name: str, index: int) -> dict:
    """生成极端测试简历数据"""
    
    # 根据索引生成不同极端情况
    scenarios = [
        "many_papers",      # 大量论文
        "many_projects",    # 大量项目
        "many_other",       # 大量其他成果
        "long_text",        # 超长文本
        "many_education",   # 大量教育经历
        "mixed_extreme",    # 混合极端
    ]
    scenario = scenarios[index % len(scenarios)]
    
    resume = {
        "姓名": name,
        "性别": random.choice(["男", "女"]),
        "年龄": random.randint(25, 60),
        "出生日期": random.randint(1970, 2000),
        "最高学历": random.choice(["博士", "硕士学位", "博士学位", "博士后"]),
        "婚配情况": random.choice(["已婚", "未婚", None]),
        "就读院校": [],
        "发表论文情况": [],
        "获批项目情况": [],
        "获奖情况": [],
        "其他成果": [],
        "著作情况": None
    }
    
    # 根据场景生成数据
    if scenario == "many_papers" or scenario == "mixed_extreme":
        # 生成大量论文（50-100篇）
        paper_count = 80 if scenario == "many_papers" else 60
        journals = [
            "《IEEE Transactions on Neural Networks and Learning Systems》",
            "《Proceedings of International Conference on Machine Learning》",
            "《Nature》",
            "《Science》",
            "《Proceedings of the Conference on Neural Information Processing Systems》",
            "《ACM Multimedia》",
            "《Proceedings of the AAAI Conference on Artificial Intelligence》",
        ]
        categories = ["SCI 1区", "CCF A会议", "SCI 2区", "CCF B会议", "SCI 3区"]
        
        for i in range(paper_count):
            journal = random.choice(journals)
            category = random.choice(categories)
            year = random.randint(2015, 2025)
            
            resume["发表论文情况"].append({
                "年份": str(year),
                "期刊名称": journal,
                "篇数": "1篇",
                "类别": category,
                "论文题目列表": [
                    f"论文标题{i+1}: {generate_long_text(100)}"  # 超长标题
                ]
            })
    
    if scenario == "many_projects" or scenario == "mixed_extreme":
        # 生成大量项目（15-25个）
        project_count = 20 if scenario == "many_projects" else 15
        project_types = [
            "国家自然科学基金",
            "国家重点研发计划",
            "973项目",
            "863项目",
            "中科院先导B",
            "科技部重点研发",
            "中科院青年交叉团队",
            "中科院引进海外人才计划",
            "基金委优秀青年基金",
        ]
        
        for i in range(project_count):
            resume["获批项目情况"].append({
                "项目类别": random.choice(project_types),
                "项目名称列表": [
                    f"项目名称{i+1}: {generate_long_text(80)}",  # 超长项目名
                    f"子项目{i+1}-1: {generate_long_text(60)}"
                ],
                "项数": "1项",
                "年份": f"{random.randint(2015, 2025)}-{random.randint(2020, 2030)}",
                "备注": f"{random.randint(50, 500)}万元"
            })
    
    if scenario == "many_other" or scenario == "mixed_extreme":
        # 生成大量其他成果（10-20项）
        other_count = 15 if scenario == "many_other" else 10
        other_types = [
            "发明专利",
            "实用新型专利",
            "软件著作权",
            "标准制定",
            "咨询报告",
            "独著书籍",
            "主编",
            "报纸发表",
        ]
        
        for i in range(other_count):
            resume["其他成果"].append({
                "类别": random.choice(other_types),
                "名称列表": [
                    f"成果名称{i+1}: {generate_long_text(100)}"
                ],
                "项数": "1项",
                "年份": str(random.randint(2015, 2025)),
                "备注": f"备注信息{i+1}: {generate_long_text(50)}"
            })
    
    if scenario == "many_education" or scenario == "mixed_extreme":
        # 生成大量教育经历（5-10条）
        edu_count = 8 if scenario == "many_education" else 5
        universities = [
            "清华大学", "北京大学", "复旦大学", "上海交通大学",
            "浙江大学", "南京大学", "中山大学", "华中科技大学",
            "西安电子科技大学", "东北大学", "吉林大学"
        ]
        majors = [
            "计算机科学与技术", "软件工程", "人工智能",
            "模式识别与智能系统", "数据科学与大数据技术",
            "网络工程", "信息安全", "物联网工程"
        ]
        
        for i in range(edu_count):
            start_year = 2000 + i * 2
            resume["就读院校"].append({
                "时间区间": f"{start_year}.09 - {start_year + 4}.06",
                "院校": random.choice(universities),
                "专业": random.choice(majors),
                "学位": random.choice(["学士", "硕士", "博士", None])
            })
    
    # 生成获奖情况（5-10项）
    award_count = random.randint(5, 10)
    award_names = [
        "优秀青年基金获得者",
        "杰出青年基金获得者",
        "优秀博士学位论文奖",
        "科技进步奖",
        "自然科学奖",
        "技术发明奖",
        "优秀教师奖",
        "优秀科研工作者",
    ]
    
    for i in range(award_count):
        resume["获奖情况"].append({
            "奖项名称": f"{random.choice(award_names)}（{generate_long_text(30)}）",
            "年份": random.randint(2015, 2025),
            "类型": "科研获奖"
        })
    
    # 确保至少有一些基本数据
    if not resume["发表论文情况"]:
        resume["发表论文情况"] = [{
            "年份": "2024",
            "期刊名称": "《Test Journal》",
            "篇数": "1篇",
            "类别": "SCI 1区",
            "论文题目列表": ["Test Paper"]
        }]
    
    if not resume["获批项目情况"]:
        resume["获批项目情况"] = [{
            "项目类别": "测试项目",
            "项目名称列表": ["Test Project"],
            "项数": "1项",
            "年份": "2024-2025",
            "备注": "100万元"
        }]
    
    return resume

def generate_test_data(num_people: int = 5):
    """生成测试数据"""
    names = [
        "张三", "李四", "王五", "赵六", "钱七",
        "孙八", "周九", "吴十", "郑十一", "王十二",
        "超长姓名测试人员ABCDEFGHIJKLMNOPQRSTUVWXYZ",  # 超长姓名
    ]
    
    resumes = []
    for i in range(num_people):
        name = names[i % len(names)]
        if i == num_people - 1:
            name = f"{name}_{i+1}"  # 最后一个用超长姓名
        
        resume = generate_extreme_resume(name, i)
        resumes.append(resume)
    
    return resumes

if __name__ == "__main__":
    # 生成5人的极限测试数据
    print("🎯 生成极限测试数据...")
    test_data = generate_test_data(5)
    
    output_file = "extreme_test_resumes.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(test_data, f, indent=2, ensure_ascii=False)
    
    print(f"✅ 已生成 {len(test_data)} 份极限测试简历")
    print(f"📁 保存至: {output_file}")
    
    # 打印统计信息
    print("\n📊 数据统计：")
    for i, resume in enumerate(test_data, 1):
        print(f"\n第{i}人: {resume['姓名']}")
        print(f"  论文: {len(resume['发表论文情况'])} 组")
        total_papers = sum(len(p.get('论文题目列表', [])) for p in resume['发表论文情况'])
        print(f"  论文总数: {total_papers} 篇")
        print(f"  项目: {len(resume['获批项目情况'])} 个")
        print(f"  其他成果: {len(resume['其他成果'])} 项")
        print(f"  教育经历: {len(resume['就读院校'])} 条")
        print(f"  获奖: {len(resume['获奖情况'])} 项")

