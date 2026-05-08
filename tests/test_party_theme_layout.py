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
        self.assertIn("expert_text_panel", layout_ids)

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

    def test_expert_text_panel_can_be_rendered(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_expert_text_panel",
            title="模块解析页",
            theme="academic_default",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_text_panel",
                    "title": "C3k2_MCA：增强小目标缺陷感知能力",
                    "statement": "这是页头摘要。",
                    "panel_title": "核心观点",
                    "panel_points": ["说明一", "说明二", "说明三"],
                    "body_title": "详细说明",
                    "body_paragraphs": ["第一段正文。", "第二段正文。"],
                }
            ],
        )

        self.assertEqual(result["rendered_slide_types"], ["expert_text_panel"])
        presentation = ppt_mcp_server.presentations["test_expert_text_panel"]
        slide = presentation.slides[0]
        texts = [shape.text.strip() for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        self.assertIn("核心观点", texts)
        self.assertIn("详细说明", texts)
        self.assertTrue(any("第一段正文" in text for text in texts))


class ThemeBackgroundTests(unittest.TestCase):
    def assert_no_edge_bars(self, slide):
        edge_bars = [
            shape for shape in slide.shapes
            if (
                (
                    abs(shape.left.inches - 0) < 0.01
                    and abs(shape.top.inches - 0) < 0.01
                    and (
                        (shape.width.inches < 0.3 and shape.height.inches > 7.0)
                        or (shape.width.inches > 12.5 and shape.height.inches < 0.2)
                    )
                )
                or (
                    abs(shape.left.inches - 0) < 0.01
                    and shape.width.inches > 12.5
                    and shape.height.inches < 0.5
                )
            )
        ]
        self.assertEqual(edge_bars, [])

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
        self.assert_no_edge_bars(presentation.slides[0])

    def test_expert_layouts_have_no_edge_bars(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_expert_layout_without_bars",
            title="专题页测试",
            theme="academic_default",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_title_content",
                    "title": "HWD：降低分辨率同时保留细节信息",
                    "statement": "常规下采样方法会损失边缘和纹理等关键信息。",
                    "items": [
                        {"title": "核心观点", "points": ["通过小波变换保留低频与高频信息"]},
                    ],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_expert_layout_without_bars"]
        self.assert_no_edge_bars(presentation.slides[0])

    def test_expert_title_content_header_is_left_aligned(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_expert_title_content_left_aligned",
            title="专题页测试",
            theme="academic_default",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_title_content",
                    "title": "C3k2_MCA：增强小目标缺陷感知能力",
                    "statement": "这是用于验证标题与装饰线位置的说明文本。",
                    "items": [
                        {"title": "核心观点", "points": ["说明一", "说明二"]},
                    ],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_expert_title_content_left_aligned"]
        slide = presentation.slides[0]

        title_shapes = [
            shape for shape in slide.shapes
            if getattr(shape, "has_text_frame", False)
            and shape.text.strip() == "C3k2_MCA：增强小目标缺陷感知能力"
        ]
        self.assertEqual(len(title_shapes), 1)
        self.assertLess(title_shapes[0].left.inches, 1.0)

        centered_bars = [
            shape for shape in slide.shapes
            if (
                5.5 < shape.left.inches < 6.2
                and 1.0 < shape.top.inches < 1.4
                and 1.0 < shape.width.inches < 1.7
                and shape.height.inches < 0.1
            )
        ]
        self.assertEqual(centered_bars, [])

    def test_expert_title_content_has_expanded_text_areas(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_expert_title_content_expanded_areas",
            title="专题页测试",
            theme="academic_default",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_title_content",
                    "title": "HWD：降低分辨率同时保留细节信息",
                    "statement": "常规下采样方法虽然可以降低特征图分辨率、扩大感受野，但也会丢失边缘、纹理等关键信息。",
                    "items": [
                        {"title": "核心观点", "points": ["说明一", "说明二", "说明三", "说明四"]},
                    ],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_expert_title_content_expanded_areas"]
        slide = presentation.slides[0]

        statement_shapes = [
            shape for shape in slide.shapes
            if getattr(shape, "has_text_frame", False)
            and "常规下采样方法虽然可以降低特征图分辨率" in shape.text
        ]
        self.assertEqual(len(statement_shapes), 1)
        self.assertGreaterEqual(statement_shapes[0].height.inches, 0.8)

        card_shapes = [
            shape for shape in slide.shapes
            if abs(shape.left.inches - 1.0) < 0.05
            and abs(shape.top.inches - 2.55) < 0.05
            and shape.width.inches > 5.0
        ]
        self.assertEqual(len(card_shapes), 1)
        self.assertGreaterEqual(card_shapes[0].height.inches, 1.9)

    def test_party_layout_has_no_edge_bars(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_party_layout_without_bars",
            title="党建工作总结",
            theme="party_red",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "party_work_summary",
                    "title": "党建工作推进情况",
                    "items": [{"title": "组织建设", "points": ["规范支部制度"]}],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_party_layout_without_bars"]
        self.assert_no_edge_bars(presentation.slides[0])


if __name__ == "__main__":
    unittest.main()
