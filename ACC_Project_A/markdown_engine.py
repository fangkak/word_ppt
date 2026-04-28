from datetime import datetime
import os


def build_markdown(config, state_img=None, seq_img=None):
    """
    构建 Markdown 文档
    
    Args:
        config: ProjectConfig 对象
        state_img: 状态机图片路径
        seq_img: 序列图片路径
    
    Returns:
        Markdown 文档内容字符串
    """
    project_info = config.get_project_info()
    functions = config.get_functions()
    safety = config.get_safety()
    state_machine = config.get_state_machine()
    signals = config.get_signals()
    
    md = []
    
    # ==================== 标题页 ====================
    md.append(f"# {project_info.get('project_name', '项目名称')}")
    md.append("")
    md.append(f"## 项目信息")
    md.append("")
    md.append(f"- **项目名称**: {project_info.get('project_name', 'N/A')}")
    md.append(f"- **车型平台**: {project_info.get('vehicle_platform', 'N/A')}")
    md.append(f"- **客户**: {project_info.get('customer', 'N/A')}")
    md.append(f"- **作者**: {project_info.get('author', 'N/A')}")
    md.append(f"- **版本**: {project_info.get('version', 'N/A')}")
    md.append(f"- **日期**: {project_info.get('date', 'N/A')}")
    md.append("")
    md.append(f"**项目描述**: {project_info.get('description', 'N/A')}")
    md.append("")
    md.append("---")
    md.append("")
    
    # ==================== 目录 ====================
    md.append("## 目录")
    md.append("")
    md.append("1. [项目信息](#项目信息)")
    md.append("2. [功能需求](#功能需求)")
    md.append("3. [安全需求](#安全需求)")
    md.append("4. [状态机](#状态机)")
    md.append("5. [信号定义](#信号定义)")
    if state_img:
        md.append("6. [状态机图](#状态机图)")
    if seq_img:
        md.append("7. [序列图](#序列图)")
    md.append("")
    md.append("---")
    md.append("")
    
    # ==================== 功能需求 ====================
    md.append("## 功能需求")
    md.append("")
    if functions:
        for func in functions:
            md.append(f"### {func.get('id', 'N/A')}: {func.get('name', 'N/A')}")
            md.append("")
            md.append(f"**描述**: {func.get('description', 'N/A')}")
            md.append("")
            
            related_states = func.get('related_states', [])
            if related_states:
                md.append(f"**相关状态**: {', '.join(related_states)}")
                md.append("")
            
            related_signals = func.get('related_signals', [])
            if related_signals:
                md.append(f"**相关信号**: {', '.join(related_signals)}")
                md.append("")
    else:
        md.append("暂无功能需求定义")
        md.append("")
    
    md.append("---")
    md.append("")
    
    # ==================== 安全需求 ====================
    md.append("## 安全需求")
    md.append("")
    
    if safety:
        # 安全目标
        safety_goals = safety.get('safety_goals', [])
        if safety_goals:
            md.append("### 安全目标")
            md.append("")
            for goal in safety_goals:
                md.append(f"- **{goal.get('id', 'N/A')}** (ASIL: {goal.get('asil', 'N/A')})")
                md.append(f"  - 危害: {goal.get('hazard', 'N/A')}")
                md.append(f"  - 安全目标: {goal.get('safety_goal', 'N/A')}")
            md.append("")
        
        # 技术需求
        technical_requirements = safety.get('technical_requirements', [])
        if technical_requirements:
            md.append("### 技术需求")
            md.append("")
            for req in technical_requirements:
                md.append(f"- **{req.get('id', 'N/A')}**")
                md.append(f"  - 源于: {req.get('from_goal', 'N/A')}")
                md.append(f"  - 描述: {req.get('description', 'N/A')}")
            md.append("")
    else:
        md.append("暂无安全需求定义")
        md.append("")
    
    md.append("---")
    md.append("")
    
    # ==================== 状态机 ====================
    md.append("## 状态机")
    md.append("")
    
    if state_machine is not None and not state_machine.empty:
        md.append("### 状态转移表")
        md.append("")
        
        # 转换 DataFrame 为 Markdown 表格
        md.append("|" + "|".join(state_machine.columns) + "|")
        md.append("|" + "|".join(["---"] * len(state_machine.columns)) + "|")
        
        for _, row in state_machine.iterrows():
            md.append("|" + "|".join(str(cell) for cell in row) + "|")
        
        md.append("")
    else:
        md.append("暂无状态机定义")
        md.append("")
    
    md.append("---")
    md.append("")
    
    # ==================== 信号定义 ====================
    md.append("## 信号定义")
    md.append("")
    
    if signals is not None and not signals.empty:
        md.append("### 信号表")
        md.append("")
        
        # 转换 DataFrame 为 Markdown 表格
        md.append("|" + "|".join(signals.columns) + "|")
        md.append("|" + "|".join(["---"] * len(signals.columns)) + "|")
        
        for _, row in signals.iterrows():
            md.append("|" + "|".join(str(cell) for cell in row) + "|")
        
        md.append("")
    else:
        md.append("暂无信号定义")
        md.append("")
    
    md.append("---")
    md.append("")
    
    # ==================== 图表 ====================
    if state_img:
        md.append("## 状态机图")
        md.append("")
        md.append(f"![状态机图]({state_img})")
        md.append("")
        md.append("---")
        md.append("")
    
    if seq_img:
        md.append("## 序列图")
        md.append("")
        md.append(f"![序列图]({seq_img})")
        md.append("")
        md.append("---")
        md.append("")
    
    # ==================== 页脚 ====================
    md.append("## 文档信息")
    md.append("")
    md.append(f"- **生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"- **文档版本**: {project_info.get('version', 'N/A')}")
    md.append("")
    
    return "\n".join(md)
