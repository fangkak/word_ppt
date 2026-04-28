import os
import sys

# 添加当前目录到 Python 路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from engine import ProjectConfig
from diagram_engine import generate_state_machine, generate_sequence
from markdown_engine import build_markdown
from word_engine import md_to_word
from ppt_engine import md_to_ppt

PROJECT = os.path.join(current_dir)  # 使用当前目录作为项目路径
OUT = os.path.join(current_dir, "output")
DIAG = os.path.join(OUT, "diagrams")
DOC = os.path.join(OUT, "docs")

def main():
    try:
        print("=" * 60)
        print("开始生成 Word 和 PowerPoint 文档...")
        print("=" * 60)
        print()
        
        # 创建输出目录
        os.makedirs(DIAG, exist_ok=True)
        os.makedirs(DOC, exist_ok=True)
        print(f"✓ 输出目录已创建: {OUT}")
        print()

        # 加载项目配置
        print("📋 加载项目配置...")
        cfg = ProjectConfig(PROJECT)
        print(f"✓ 项目配置已加载")
        print(f"  项目名称: {cfg.get_project_info().get('project_name')}")
        print()

        # 生成图表
        print("🎨 生成图表...")
        state_img = generate_state_machine(cfg.get_state_machine(), DIAG)
        seq_img = generate_sequence(cfg.get_signals(), DIAG)
        print()

        # 生成 Markdown
        print("📝 生成 Markdown 文档...")
        md = build_markdown(cfg, state_img, seq_img)
        md_file = os.path.join(DOC, "ACC_req.md")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"✓ Markdown 文档已生成: {md_file}")
        print()

        # 生成 Word 文档
        print("📄 生成 Word 文档...")
        docx_file = os.path.join(DOC, "ACC_req.docx")
        md_to_word(md_file, docx_file)
        print()

        # 生成 PowerPoint 演示文稿
        print("🎯 生成 PowerPoint 演示文稿...")
        ppt_file = os.path.join(DOC, "ACC_req.pptx")
        md_to_ppt(md_file, ppt_file, cfg.get_project_info())
        print()

        print("=" * 60)
        print("✅ 所有文档生成完成！")
        print("=" * 60)
        print(f"📁 输出文件位置: {DOC}")
        print(f"  - Markdown: {md_file}")
        print(f"  - Word: {docx_file}")
        print(f"  - PowerPoint: {ppt_file}")
        print()
        
    except Exception as e:
        print()
        print("=" * 60)
        print(f"❌ 错误: {str(e)}")
        print("=" * 60)
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
