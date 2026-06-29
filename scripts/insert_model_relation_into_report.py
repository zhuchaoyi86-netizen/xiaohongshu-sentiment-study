from pathlib import Path

from docx import Document


SRC = Path("/Users/xinxinhuashe/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_ah3md0gkwwsx22_bc68/temp/drag/基于文本挖掘的小红书食品种草笔记情感对点赞热度的影响研究.docx")
OUT = Path("/Users/xinxinhuashe/Documents/1/python课/小组共享材料/基于文本挖掘的小红书食品种草笔记情感对点赞热度的影响研究_补充模型关系版.docx")


HEADING = "鲁棒性检验"


def main():
    doc = Document(str(SRC))

    target = None
    for para in doc.paragraphs:
        if para.text.strip() == HEADING:
            target = para
            break

    if target is None:
        raise ValueError("未找到“鲁棒性检验”标题，无法插入新内容。")

    # 将原“鲁棒性检验”降为第四部分下的二级标题
    target.style = doc.styles["Heading 2"]

    inserts = [
        ("基准线性回归模型结果分析", "Heading 2"),
        (
            "在实证分析部分，本文首先构建基准线性回归模型，以缩尾并取对数后的点赞量作为被解释变量，以情绪密度、情感得分和混合评价等变量作为核心解释变量，同时控制文本长度、配图数量、话题标签数量、标点语气特征、营销话术标记以及发布时间时段等因素。之所以采用这一设定，是因为食品种草笔记的点赞热度并不是由文本情绪单独决定的，而是情感表达、图文呈现与发布时间共同作用的结果。通过在同一模型中纳入这些变量，本文能够更清晰地识别情感因素的独立影响。",
            "Normal",
        ),
        (
            "基准回归结果显示，模型样本量为2899条，调整后的R方约为0.16，说明文本情感、图文呈现与发布时间能够对点赞热度差异提供一定解释。从核心变量看，情绪密度的回归系数为0.0949，在普通标准误下显著为正，表明在控制其他因素后，情绪表达越集中、越鲜明的食品种草笔记，整体上越容易获得更高点赞。相比之下，混合评价变量的回归系数为负0.3389，且在基准模型中显著，这意味着一篇帖子如果同时呈现推荐与保留意见，用户点赞行为会受到明显抑制。与此同时，配图数量、感叹号数量以及非凌晨发布时间均表现出正向作用，说明平台中的互动热度不仅受到情感内容影响，也受到内容呈现形式和发布时机的共同塑造。",
            "Normal",
        ),
        ("变量之间的模型关系", "Heading 2"),
        (
            "进一步从变量之间的关系看，本文的结果并不支持“文本越正向，点赞就一定越高”的简单线性判断，而是呈现出更有层次的作用机制。首先，情绪密度与点赞热度总体呈正相关，这说明相比于单纯堆叠正向词汇，用户更容易对情绪表达集中、态度鲜明的内容作出互动反应。其次，情感得分在与情绪密度同时进入模型后并未表现出稳定的正向作用，说明单纯增加情绪词数量并不一定直接提升点赞，真正更重要的是情绪在文本中的组织方式和表达密度。",
            "Normal",
        ),
        (
            "此外，混合评价变量的显著负向影响进一步揭示了食品种草笔记中的结构性特征：当文本中同时包含“推荐/种草”和“避雷/保留意见”等正负并存的信息时，用户更容易形成谨慎判断，从而降低点赞意愿。配图数量和发布时间变量则说明，小红书平台中的点赞热度具有明显的平台情境特征，视觉呈现和用户活跃时段会放大或削弱文本情绪的传播效果。因此，本文认为，食品种草笔记的点赞热度本质上是“情感表达强弱与结构 + 图文呈现 + 发布时间”共同作用的结果，其中情绪密度和混合评价是最值得重点关注的两个文本变量。",
            "Normal",
        ),
    ]

    current = target
    for text, style in reversed(inserts):
        current = current.insert_paragraph_before(text, style=style)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT))
    print(f"已生成：{OUT}")


if __name__ == "__main__":
    main()
