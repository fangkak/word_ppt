import os
from engine.config_loader import ProjectConfig
from engine.diagram_engine import generate_state_machine, generate_sequence
from engine.markdown_engine import build_markdown
from engine.word_engine import md_to_word

PROJECT = "projects/ACC_Project_A"
OUT = "output"
DIAG = os.path.join(OUT, "diagrams")
DOC = os.path.join(OUT, "docs")

def main():
    os.makedirs(DIAG, exist_ok=True)
    os.makedirs(DOC, exist_ok=True)

    cfg = ProjectConfig(PROJECT)

    state_img = generate_state_machine(cfg.state_machine, DIAG)
    seq_img = generate_sequence(cfg.signals, DIAG)

    md = build_markdown(cfg, state_img, seq_img)
    md_file = os.path.join(DOC, "ACC_req.md")
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(md)

    docx_file = os.path.join(DOC, "ACC_req.docx")
    md_to_word(md_file, docx_file)

    print("生成完成：", docx_file)

if __name__ == "__main__":
    main()
