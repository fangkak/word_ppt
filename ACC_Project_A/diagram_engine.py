import os
import matplotlib.pyplot as plt
import pandas as pd
from datetime import datetime


def generate_state_machine(state_machine_df, output_dir):
    """
    生成状态机图
    
    Args:
        state_machine_df: 状态机 DataFrame
        output_dir: 输出目录
    
    Returns:
        生成的图片文件路径
    """
    try:
        if state_machine_df is None or state_machine_df.empty:
            print("⚠️  状态机数据为空，跳过图表生成")
            return None
        
        # 创建状态机图
        fig, ax = plt.subplots(figsize=(12, 8))
        
        # 获取状态列表
        states = state_machine_df.iloc[:, 0].unique()
        
        # 绘制状态节点
        for i, state in enumerate(states):
            circle = plt.Circle((i * 2, 0), 0.3, color='lightblue', ec='navy', linewidth=2)
            ax.add_patch(circle)
            ax.text(i * 2, 0, state, ha='center', va='center', fontsize=10, weight='bold')
        
        ax.set_xlim(-1, len(states) * 2)
        ax.set_ylim(-1, 1)
        ax.axis('off')
        ax.set_aspect('equal')
        
        # 保存图片
        output_file = os.path.join(output_dir, 'state_machine.png')
        plt.savefig(output_file, dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"✓ 状态机图已生成: {output_file}")
        return output_file
        
    except Exception as e:
        print(f"⚠️  状态机图生成失败: {str(e)}")
        return None


def generate_sequence(signals_df, output_dir):
    """
    生成序列图
    
    Args:
        signals_df: 信号 DataFrame
        output_dir: 输出目录
    
    Returns:
        生成的图片文件路径
    """
    try:
        if signals_df is None or signals_df.empty:
            print("⚠️  信号数据为空，跳过序列图生成")
            return None
        
        # 创建序列图
        fig, ax = plt.subplots(figsize=(12, 8))
        
        # 获取信号列表
        signals = signals_df.columns.tolist() if hasattr(signals_df, 'columns') else []
        
        if not signals:
            print("⚠️  没有找到信号数据")
            return None
        
        # 绘制信号线
        for i, signal in enumerate(signals):
            ax.plot([0, len(signals_df)], [i, i], 'b-', linewidth=2)
            ax.text(-0.5, i, signal, ha='right', va='center', fontsize=9)
        
        ax.set_xlabel('Time', fontsize=10)
        ax.set_ylabel('Signals', fontsize=10)
        ax.set_title('Signal Sequence Diagram', fontsize=12, weight='bold')
        ax.grid(True, alpha=0.3)
        
        # 保存图片
        output_file = os.path.join(output_dir, 'sequence.png')
        plt.savefig(output_file, dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"✓ 序列图已生成: {output_file}")
        return output_file
        
    except Exception as e:
        print(f"⚠️  序列图生成失败: {str(e)}")
        return None
