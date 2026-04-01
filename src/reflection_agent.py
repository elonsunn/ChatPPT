# reflection_agent.py
#
# 基于 LangGraph 的反思（Reflection）机制：
#   generate → reflect → generate → ... → END
#
# 终止条件（满足任一即停止）：
#   1. 当前已完成轮数 >= max_iterations（强制上限，默认 7）
#   2. 反思者输出 DONE_SIGNAL 且已完成轮数 >= MIN_ITERATIONS（质量达标提前退出）

import operator
from typing import Annotated, TypedDict

from langgraph.graph import END, StateGraph
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI

from logger import LOG

# 反思者输出此信号表示内容已达标，可提前结束循环
DONE_SIGNAL = "[已达标]"
MIN_ITERATIONS = 3   # 无论质量如何，至少完成的轮数
DEFAULT_MAX_ITERATIONS = 7  # 最多轮数上限


class ReflectionState(TypedDict):
    user_input: str                                  # 用户原始需求
    current_draft: str                               # 当前最新草稿
    critiques: Annotated[list[str], operator.add]    # 历轮反思意见（追加）
    iteration: int                                   # 已完成的生成轮数
    max_iterations: int                              # 本次调用的最大轮数


def build_reflection_graph(generator_prompt: str, reflector_prompt: str):
    """
    构建并返回已编译的 LangGraph 反思图。

    参数:
        generator_prompt: ChatBot 系统提示（用于生成 PPT 草稿）
        reflector_prompt: 反思者系统提示（用于审查并给出改进意见）
    """
    model = ChatOpenAI(model="gpt-4o-mini", temperature=0.5, max_tokens=4096)

    # ── 生成节点 ─────────────────────────────────────────────────────────────
    def generate(state: ReflectionState) -> dict:
        iteration = state["iteration"]

        if iteration == 0:
            # 第 1 轮：直接根据用户需求生成初始草稿
            LOG.info(f"[Reflection] 第 1 轮：初始生成")
            messages = [
                SystemMessage(content=generator_prompt),
                HumanMessage(content=state["user_input"]),
            ]
        else:
            # 后续轮次：将草稿 + 最新反思意见一并提供给生成器
            last_critique = state["critiques"][-1]
            LOG.info(f"[Reflection] 第 {iteration + 1} 轮：基于反思意见修订")
            revision_input = (
                f"用户需求：\n{state['user_input']}\n\n"
                f"当前草稿：\n{state['current_draft']}\n\n"
                f"反思意见：\n{last_critique}\n\n"
                "请根据以上反思意见对草稿进行优化，保持原有 Markdown 格式要求不变。"
            )
            messages = [
                SystemMessage(content=generator_prompt),
                HumanMessage(content=revision_input),
            ]

        response = model.invoke(messages)
        LOG.debug(f"[Reflection] 生成完成（第 {iteration + 1} 轮），内容长度: {len(response.content)}")
        return {
            "current_draft": response.content,
            "iteration": iteration + 1,
        }

    # ── 反思节点 ─────────────────────────────────────────────────────────────
    def reflect(state: ReflectionState) -> dict:
        iteration = state["iteration"]
        LOG.info(f"[Reflection] 第 {iteration} 轮反思中...")

        # 当已满足最小轮数时，允许反思者提前宣告达标
        allow_done_hint = (
            f"若内容已足够优质，直接回复"{DONE_SIGNAL}"。否则，" if iteration >= MIN_ITERATIONS else ""
        )
        review_input = (
            f"请审查以下第 {iteration} 轮 PPT 内容草稿，{allow_done_hint}"
            f"给出具体改进建议：\n\n{state['current_draft']}"
        )
        messages = [
            SystemMessage(content=reflector_prompt),
            HumanMessage(content=review_input),
        ]

        response = model.invoke(messages)
        LOG.debug(f"[Reflection] 反思完成（第 {iteration} 轮）: {response.content[:120]}...")
        return {"critiques": [response.content]}

    # ── 路由函数：generate 完成后决定继续反思还是结束 ─────────────────────────
    def should_continue(state: ReflectionState) -> str:
        iteration = state["iteration"]
        max_iter = state["max_iterations"]

        # 条件 1：达到最大轮数上限，强制结束
        if iteration >= max_iter:
            LOG.info(f"[Reflection] 已达最大轮数 {max_iter}，输出最终版本")
            return END

        # 条件 2：上一轮反思认为已达标，且满足最小轮数要求
        last_critique = state["critiques"][-1] if state["critiques"] else ""
        if DONE_SIGNAL in last_critique and iteration >= MIN_ITERATIONS:
            LOG.info(f"[Reflection] 第 {iteration} 轮，反思者认为内容已达标，提前结束")
            return END

        return "reflect"

    # ── 构建图 ────────────────────────────────────────────────────────────────
    graph = StateGraph(ReflectionState)
    graph.add_node("generate", generate)
    graph.add_node("reflect", reflect)

    graph.set_entry_point("generate")
    graph.add_conditional_edges(
        "generate",
        should_continue,
        {END: END, "reflect": "reflect"},
    )
    graph.add_edge("reflect", "generate")

    return graph.compile()
