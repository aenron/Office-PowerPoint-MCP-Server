import unittest

import ppt_mcp_server


def call_tool(name, **kwargs):
    return ppt_mcp_server.app._tool_manager._tools[name].fn(**kwargs)


class PartyThemeLayoutTests(unittest.TestCase):
    def test_party_theme_and_layout_are_listed(self):
        options = call_tool("list_presentation_options")

        theme_ids = {theme["theme_id"] for theme in options["themes"]}
        layout_ids = {layout["layout_id"] for layout in options["layouts"]}

        self.assertIn("party_red", theme_ids)
        self.assertIn("party_work_summary", layout_ids)

    def test_party_work_summary_can_be_rendered(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_party_work_summary",
            title="党建工作总结汇报",
            subtitle="2026年度重点工作",
            theme="party_red",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "party_work_summary",
                    "title": "党建工作推进情况",
                    "statement": "围绕组织建设、思想引领、作风提升和服务群众形成闭环推进。",
                    "items": [
                        {"title": "组织建设", "points": ["规范支部制度", "强化党员管理"]},
                        {"title": "理论学习", "points": ["落实第一议题", "开展专题研讨"]},
                        {"title": "作风提升", "points": ["问题清单整改", "优化服务流程"]},
                    ],
                }
            ],
        )

        self.assertEqual(result["theme"], "party_red")
        self.assertEqual(result["rendered_slide_types"], ["party_work_summary"])
        self.assertEqual(result["slide_count"], 1)


class ThemeBackgroundTests(unittest.TestCase):
    def test_generated_theme_background_has_no_edge_bars(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_theme_background_without_bars",
            title="主题背景测试",
            theme="academic_burgundy",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "research_questions",
                    "title": "研究问题",
                    "questions": ["问题一", "问题二"],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_theme_background_without_bars"]
        slide = presentation.slides[0]
        edge_bars = [
            shape for shape in slide.shapes
            if (
                abs(shape.left.inches - 0) < 0.01
                and abs(shape.top.inches - 0) < 0.01
                and (
                    (shape.width.inches < 0.3 and shape.height.inches > 7.0)
                    or (shape.width.inches > 12.5 and shape.height.inches < 0.2)
                )
            )
        ]

        self.assertEqual(edge_bars, [])


if __name__ == "__main__":
    unittest.main()
