# chatbot.py

from abc import ABC

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.messages import HumanMessage, AIMessage
from langchain_core.runnables.history import RunnableWithMessageHistory

from logger import LOG
from chat_history import get_session_history
from reflection_agent import build_reflection_graph, DEFAULT_MAX_ITERATIONS


class ChatBot(ABC):
    """
    聊天机器人基类，提供两种对话模式：
      - chat_with_history:    单轮直接调用，历史自动追加
      - chat_with_reflection: LangGraph 反思循环（3-7 轮），仅将最终版本写入历史
    """

    def __init__(
        self,
        prompt_file="./prompts/chatbot.txt",
        reflector_prompt_file="./prompts/chatbot_reflector.txt",
        session_id=None,
    ):
        self.prompt_file = prompt_file
        self.reflector_prompt_file = reflector_prompt_file
        self.session_id = session_id if session_id else "default_session_id"

        self.prompt = self._load_file(self.prompt_file)
        self.reflector_prompt = self._load_file(self.reflector_prompt_file)

        self._create_chatbot()
        self._create_reflection_graph()

    # ── 内部初始化 ────────────────────────────────────────────────────────────

    def _load_file(self, path: str) -> str:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except FileNotFoundError:
            raise FileNotFoundError(f"找不到提示文件 {path}!")

    def _create_chatbot(self):
        """初始化带历史记录的普通对话链。"""
        system_prompt = ChatPromptTemplate.from_messages([
            ("system", self.prompt),
            MessagesPlaceholder(variable_name="messages"),
        ])
        self.chatbot = system_prompt | ChatOpenAI(
            model="gpt-4o-mini",
            temperature=0.5,
            max_tokens=4096,
        )
        self.chatbot_with_history = RunnableWithMessageHistory(
            self.chatbot, get_session_history
        )

    def _create_reflection_graph(self):
        """构建 LangGraph 反思图（使用相同的生成器 prompt + 反思者 prompt）。"""
        self.reflection_graph = build_reflection_graph(
            generator_prompt=self.prompt,
            reflector_prompt=self.reflector_prompt,
        )

    # ── 公开接口 ──────────────────────────────────────────────────────────────

    def chat_with_history(self, user_input, session_id=None):
        """
        单轮直接调用，回复自动追加到会话历史。

        参数:
            user_input (str): 用户输入
            session_id (str): 会话 ID（可选）
        返回:
            str: AI 回复
        """
        if session_id is None:
            session_id = self.session_id

        response = self.chatbot_with_history.invoke(
            [HumanMessage(content=user_input)],
            {"configurable": {"session_id": session_id}},
        )
        LOG.debug(f"[ChatBot] {response.content}")
        return response.content

    def chat_with_reflection(
        self,
        user_input: str,
        session_id: str | None = None,
        max_iterations: int = DEFAULT_MAX_ITERATIONS,
    ) -> str:
        """
        通过 LangGraph 反思循环（3-7 轮）提升内容质量，
        **仅将最终生成版本写入 ChatHistory**，中间轮次不污染历史记录。

        参数:
            user_input (str):       用户原始需求
            session_id (str):       会话 ID（可选）
            max_iterations (int):   最大反思轮数（默认 7，最少保证 3 轮）
        返回:
            str: 经多轮反思优化后的最终 PPT 内容
        """
        if session_id is None:
            session_id = self.session_id

        LOG.info(f"[Reflection] 启动反思循环，最大轮数: {max_iterations}")

        # 运行反思图（独立于 RunnableWithMessageHistory，不自动写入历史）
        result = self.reflection_graph.invoke({
            "user_input": user_input,
            "current_draft": "",
            "critiques": [],
            "iteration": 0,
            "max_iterations": max_iterations,
        })

        final_content: str = result["current_draft"]
        iterations_done: int = result["iteration"]
        LOG.info(
            f"[Reflection] 完成，共 {iterations_done} 轮，"
            f"最终内容长度: {len(final_content)} 字符"
        )

        # 手动将「用户消息 + 最终 AI 回复」写入历史，跳过所有中间轮次
        history = get_session_history(session_id)
        history.add_message(HumanMessage(content=user_input))
        history.add_message(AIMessage(content=final_content))

        return final_content
