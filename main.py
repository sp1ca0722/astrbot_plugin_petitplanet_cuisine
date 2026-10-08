import asyncio
import os
import re
import unicodedata
from difflib import SequenceMatcher

import pandas as pd

from astrbot.api.event import filter, AstrMessageEvent, MessageChain
from astrbot.api.message_components import Plain
from astrbot.api.star import Context, Star, register
from astrbot.api import AstrBotConfig, logger

CUTOFF = 0.3
# 删掉品质
NAME_TAIL = re.compile(r"\s*\([^()]*\)\s*$")
# 从消息里剥掉 "/recipe"、"/菜谱" 这类指令前缀，兼容带不带斜杠、带不带空格
# 指令名要与 @filter.command("recipe", alias={'菜谱', '食谱'}) 保持一致（cuisine 是旧命令名，留着兼容）
# 用后瞻 (?![A-Za-z0-9_]) 代替 \b：\b 对中文不生效（中文算 word 字符），会导致 “/菜谱和煦花果茶” 剥不掉
CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:recipe|cuisine|食谱|菜谱)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)
# 数据表跟 main.py 放一起，避免受 AstrBot 启动目录影响
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cuisine.xlsx")

# 平台模式（_conf_schema.json 里的 platform_mode）：只有 QQ 官方机器人渲染 Markdown
MODE_QQ_OFFICIAL = "qq_official"
MODE_ONEBOT = "onebot"
DEFAULT_PLATFORM_MODE = MODE_ONEBOT

# 降级纯文本时要去掉的 Markdown 标记，都是「保留内容、只剥标记」
MD_IMAGE = re.compile(r"!\[([^\]]*)\]\([^)\s]*\)")
MD_LINK = re.compile(r"\[([^\]]*)\]\([^)\s]*\)")
MD_FENCE = re.compile(r"^[ \t]*```.*$\n?", re.MULTILINE)
MD_HR = re.compile(r"^[ \t]{0,3}(?:[-*_][ \t]*){3,}$\n?", re.MULTILINE)
MD_HEADING = re.compile(r"^[ \t]{0,3}#{1,6}[ \t]*", re.MULTILINE)
MD_QUOTE = re.compile(r"^[ \t]{0,3}>[ \t]?", re.MULTILINE)
MD_LIST = re.compile(r"^([ \t]*)[-*+][ \t]+", re.MULTILINE)
MD_CODE = re.compile(r"`([^`\n]+)`")
MD_BOLD = re.compile(r"\*\*(.+?)\*\*")
MD_ITALIC = re.compile(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)")


def base_name(text: str) -> str:
    """全角转半角 → 去首尾空白 → 剥掉尾部的括号标注。"""
    s = unicodedata.normalize("NFKC", str(text)).strip()
    return NAME_TAIL.sub("", s).strip() or s


def strip_markdown(text: str) -> str:
    """把 Markdown 标记去掉、只留内容，给不渲染 Markdown 的平台用。

    顺序有讲究：图片要先于链接（否则 `![x](y)` 会剩下一个 `!`），
    行内代码要先于加粗/斜体（避免代码里的 `*` 被当成强调标记）。
    """
    for pattern, repl in (
        (MD_IMAGE, r"\1"),
        (MD_LINK, r"\1"),
        (MD_FENCE, ""),
        (MD_HR, ""),
        (MD_HEADING, ""),
        (MD_QUOTE, ""),
        (MD_LIST, r"\1"),  # 去掉列表的 "- " 标记，保留缩进
        (MD_CODE, r"\1"),
        (MD_BOLD, r"\1"),
        (MD_ITALIC, r"\1"),
    ):
        text = pattern.sub(repl, text)
    return text


@register("astrbot_plugin_petitplanet_cuisine", "spica", "星布谷地菜谱查询", "1.0.3")
class MyPlugin(Star):
    def __init__(self, context: Context, config: AstrBotConfig | None = None):
        super().__init__(context)
        # 老版本 AstrBot 不传 config，此时按默认模式（默认 onebot = 去 Markdown）处理
        self.config = config or {}
        self._rows = []
        self._names = []
        self._base_names = []
        self._source = ""  # 表头 L1、M1：数据来源与更新日期

    async def initialize(self):
        """插件实例化后自动调用：把菜谱表读进内存，之后每次查询不再碰磁盘。"""
        try:
            df = await asyncio.to_thread(pd.read_excel, EXCEL_PATH, header=None)
        except Exception as e:  # 文件缺失 / 没装 openpyxl
            logger.error(f"读取 cuisine.xlsx 失败（{EXCEL_PATH}）：{e}")
            return

        # 用列表下标定位，避免再拿标签回查 DataFrame（下标是 int，类型干净）
        self._rows = df.iloc[1:].values.tolist()
        self._names = [str(row[1]).strip() for row in self._rows]  # 菜品名在第 1 列
        self._base_names = [base_name(name) for name in self._names]  # 剥掉标注后的主体名
        # 表头首行的 L、M 两格是「数据来源」「更新日期」，供 /菜谱 源 查询
        header = df.iloc[0].tolist()
        self._source = "\n".join(
            str(cell).strip() for cell in header[11:13] if not pd.isna(cell) and str(cell).strip()
        )
        logger.info(f"菜谱加载完成，共 {len(self._rows)} 条；平台模式：{self._platform_mode()}")

    @filter.command("recipe", alias={'菜谱', '食谱'})
    async def cuisine(self, event: AstrMessageEvent):
        """查询菜谱：/recipe 菜名"""
        # 输入内容来自 message_str；剥掉指令前缀（含别名「菜谱」），兼容不剥的情况
        target = CMD_PREFIX.sub("", event.message_str).strip()
        text = self._render(await self._build_reply(target))

        # 这里手动发送、不走 `yield event.plain_result(text)`：
        # AstrBot 的「引用回复」（platform_settings.reply_with_quote）会往消息链头部插入 Reply 段，
        # 该段序列化后带一个空的 chain 字段，部分 NapCat 版本会直接拒绝整条消息
        # （ActionFailed retcode=1400, 'message segment "reply" field "chain" must be a scalar value'），
        # 结果消息完全发不出去。自己发送就不含 Reply 段，保证送达（代价是没有引用效果）。
        # 若日后 AstrBot / NapCat 修好，可换回 yield event.plain_result(text)。
        try:
            await event.send(MessageChain([Plain(text)]))
        except Exception as e:
            logger.error(f"发送菜谱结果失败：{e}")
        # 已经自己发过了，终止事件传播，避免 AstrBot 再把结果拿去（带引用地）发送一遍
        event.stop_event()
        yield

    @filter.command("helpme", alias={'帮助'})
    async def help(self, event: AstrMessageEvent):
        """查询菜谱帮助：/helpme"""
        text = (
            "### 使用帮助\n\n"
            "`/食谱 菜名`(可模糊搜索)\n"
            "- 返回这道菜的食材和做法\n\n"
#            "`/菜详情 菜名`(可模糊搜索)\n"
#            "- 返回这道菜具体信息\n"
            "`/许愿 许愿内容`\n"
            "- 可以进行对bot功能的许愿（）"
        )
        yield event.plain_result(self._render(text))

    def _platform_mode(self) -> str:
        """读配置里的平台模式，兼容大小写和空格；取值不认识时回落到默认值。"""
        mode = str(self.config.get("platform_mode") or DEFAULT_PLATFORM_MODE).strip().lower()
        return mode if mode in (MODE_QQ_OFFICIAL, MODE_ONEBOT) else DEFAULT_PLATFORM_MODE

    def _render(self, text: str) -> str:
        """按平台模式整理输出：OneBot 不会渲染 Markdown，得先把标记去掉。"""
        return strip_markdown(text) if self._platform_mode() == MODE_ONEBOT else text

    async def _build_reply(self, target: str) -> str:
        """把查询结果整理成要回复的纯文本。"""
        if not self._rows:
            return "菜谱数据没加载成功，请检查 cuisine.xlsx 是否放在插件目录下。"
        if not target:
            return "请输入菜名，例如：/食谱 和煦花果茶"
        # 「源」是固定关键词：直接回表头的 L1、M1（数据来源、更新日期），不走模糊匹配
        if target == "源":
            return self._source or "表格里没有填写数据来源与更新日期。"

        # 用主体名算相似度：输入“梦幻番茄汤汁面”和表里的“梦幻番茄汤汁面(金)”就是完全一致(1.0)
        target_base = base_name(target)
        scored = [
            (SequenceMatcher(None, target_base, base).ratio(), pos)
            for pos, base in enumerate(self._base_names)
        ]
        # key 只按分数排，稳定性保证同分时仍是表里的先后顺序
        matched = sorted(
            (item for item in scored if item[0] >= CUTOFF),
            key=lambda item: item[0],
            reverse=True,
        )

        if not matched:
            return await self._ask_ai(target)

        best_score, best_pos = matched[0]
        # 有特殊效果才打印；注意判断和打印要用同一个下标，否则会打印出空行
        row = ["" if pd.isna(cell) else cell for cell in self._rows[best_pos]]

        lines = []
        if best_score < 1:
            lines.append("猜你想搜：")
        lines.append(f"**[{row[0]}] {row[1]}**")
        # 空的食材格子直接跳过，有内容的各自包成行内代码
        ingredients = [str(cell).strip() for cell in row[2:6]]
        ingredients = [f"`{name}`" for name in ingredients if name]
        if ingredients:
            lines.append("食材：" + " ".join(ingredients))
        lines.append(f"厨具：{row[6]}")
        if row[10] != "":
            lines.append(str(row[10]))
        # 只有一项匹配时没有「其他的菜」可列，不输出这一行
        if best_score < 1 and len(matched) > 1:
            lines.append("> **或是其他的菜？**")
            lines.append("> " + " ".join(self._names[pos] for _, pos in matched[1:]))
        return "\n".join(lines)

    async def _ask_ai(self, target: str) -> str:
        """表里查不到时交给大模型兜底。"""
        # openai 的 import 要 1.2s 左右，只有真需要问 AI 时才导入；命中菜品时可以完全跳过
        from openai import AsyncOpenAI

        api_key = os.environ.get("DEEPSEEK_API_KEY")
        if not api_key:
            logger.error("缺少环境变量 DEEPSEEK_API_KEY，无法调用 AI 兜底")
            return "没有这道菜哦，也不知道它要什么食材……（请管理员配置 DEEPSEEK_API_KEY）"

        client = AsyncOpenAI(api_key=api_key, base_url="https://api.deepseek.com")

        try:
            response = await client.chat.completions.create(
                model="deepseek-flash",
                messages=[
                    {
                        "role": "system",
                        "content": "你是一个厨艺大师，擅长根据食材和厨具推荐菜谱。",
                    },
                    {
                        "role": "user",
                        "content": f"现在用户询问这道菜{target}，但现有的菜单里没有这个菜，请用以下说法来回答：没有这道菜哦，但基本上就是【食材】【食材】（或者更多，最多4个），然后使用【厨具】。语言请一定要简洁。",
                    },
                ],
                stream=False,
                # thinking 关闭 = 非思考模式；reasoning_effort="high" 会把思考模式重新打开，两者冲突，所以不传
                extra_body={"thinking": {"type": "disabled"}},
            )
        except Exception as e:
            logger.error(f"调用 DeepSeek 失败：{e}")
            return "没有这道菜哦，查询服务暂时有点问题，稍后再试试。"

        return response.choices[0].message.content or "没有这道菜哦。"

    async def terminate(self):
        """可选择实现异步的插件销毁方法，当插件被卸载/停用时会调用。"""
