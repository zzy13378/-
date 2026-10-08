"""真实模式评测与审查模块。

只处理真实数据：真实论文语料、真实检索、真实大模型，以及第一个项目真实跑出来的产出。
不包含任何模拟用例、模拟模型或离线演示。
"""

__all__ = ["alignment", "realcorpus", "question_gen", "judge", "rag_agent", "p4_adapter", "runtime"]
