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
        self.assertIn("expert_forum_blue", theme_ids)
        self.assertNotIn("government_blue", theme_ids)
        self.assertIn("tech_dark", theme_ids)
        self.assertNotIn("education_warm", theme_ids)
        self.assertNotIn("academic_default", theme_ids)
        self.assertIn("party_summary_panel", layout_ids)
        self.assertNotIn("expert_body_panel", layout_ids)
        self.assertNotIn("expert_card_overview", layout_ids)
        self.assertNotIn("expert_image_text", layout_ids)
        self.assertNotIn("expert_process_path", layout_ids)
        self.assertNotIn("expert_scope_list", layout_ids)
        self.assertIn("agenda", layout_ids)
        self.assertIn("problem_solution", layout_ids)
        self.assertIn("case_study", layout_ids)
        self.assertIn("section_agenda", layout_ids)
        self.assertNotIn("before_after", layout_ids)
        self.assertNotIn("swot", layout_ids)
        self.assertIn("risk_matrix", layout_ids)
        self.assertIn("roadmap", layout_ids)
        self.assertNotIn("team_roles", layout_ids)
        self.assertIn("image_showcase", layout_ids)
        self.assertNotIn("architecture", layout_ids)
        self.assertNotIn("party_work_summary", layout_ids)
        self.assertNotIn("expert_text_panel", layout_ids)
        self.assertNotIn("template_names", options)
        self.assertNotIn("compatible_content_fields", options)
        self.assertNotIn("compatible_source_fields", options)
        self.assertNotIn("discovery_tool", options["template_support"])

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
                    "type": "party_summary_panel",
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
        self.assertEqual(result["rendered_slide_types"], ["party_summary_panel"])
        self.assertEqual(result["slide_count"], 1)

    def test_priority_layouts_can_be_rendered(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_priority_layouts",
            title="常用版式测试",
            theme="business_blue",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "agenda",
                    "title": "汇报提纲",
                    "agenda": [
                        {"title": "研究背景", "points": ["说明问题来源"]},
                        {"title": "解决方案", "points": ["说明推进路径"]},
                    ],
                },
                {
                    "type": "problem_solution",
                    "title": "问题与方案",
                    "problem": ["流程分散", "信息重复录入"],
                    "causes": ["系统割裂", "标准不统一"],
                    "solution": ["统一入口", "建立数据规范"],
                    "result": "提升协同效率。",
                },
                {
                    "type": "case_study",
                    "title": "案例分析",
                    "case_name": "试点项目",
                    "case_background": "围绕重点场景开展试点。",
                    "case_actions": ["梳理流程", "上线工具"],
                    "case_results": ["效率提升", "体验改善"],
                    "case_insights": ["标准化后再推广"],
                },
            ],
        )

        self.assertEqual(result["rendered_slide_types"], ["agenda", "problem_solution", "case_study"])
        self.assertEqual(result["slide_count"], 3)

    def test_additional_priority_layouts_can_be_rendered(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_additional_priority_layouts",
            title="扩展版式测试",
            theme="government_blue",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "section_agenda",
                    "title": "第一部分 工作基础",
                    "section_no": "01",
                    "agenda": [
                        {"title": "政策依据", "points": ["说明背景"]},
                        {"title": "实施范围", "points": ["说明边界"]},
                    ],
                },
                {
                    "type": "before_after",
                    "title": "改造前后对比",
                    "before": ["流程分散", "口径不一"],
                    "after": ["统一入口", "统一标准"],
                    "metrics": [{"label": "效率", "value": "提升30%"}],
                },
                {
                    "type": "swot",
                    "title": "形势分析",
                    "strengths": ["基础较好"],
                    "weaknesses": ["协同不足"],
                    "opportunities": ["政策支持"],
                    "threats": ["外部不确定"],
                },
                {
                    "type": "risk_matrix",
                    "title": "风险清单",
                    "risks": [
                        {"risk": "进度延迟", "level": "高", "impact": "影响上线", "mitigation": "周例会跟踪"},
                    ],
                },
                {
                    "type": "roadmap",
                    "title": "推进路线图",
                    "roadmap": [
                        {"title": "启动", "points": ["完成调研"]},
                        {"title": "试点", "points": ["小范围验证"]},
                    ],
                },
                {
                    "type": "team_roles",
                    "title": "职责分工",
                    "roles": [
                        {"title": "项目组", "points": ["统筹推进", "风险协调"]},
                        {"title": "业务组", "points": ["需求确认"]},
                    ],
                },
                {
                    "type": "image_showcase",
                    "title": "现场图片",
                    "image_path": "missing-image.png",
                    "image_caption": "示意说明",
                    "notes": ["用于验证无图片占位也可渲染"],
                },
            ],
        )

        self.assertEqual(
            result["rendered_slide_types"],
            ["section_agenda", "before_after", "swot", "risk_matrix", "roadmap", "team_roles", "image_showcase"],
        )
        self.assertEqual(result["theme"], "business_blue")
        self.assertEqual(result["slide_count"], 7)

    def test_hidden_layouts_remain_renderable_for_compatibility(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_hidden_layouts_compatibility",
            title="隐藏版式兼容",
            theme="expert_forum_blue",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_body_panel",
                    "title": "专家正文页",
                    "panel_points": ["要点一"],
                    "body_paragraphs": ["正文说明。"],
                },
                {
                    "type": "team_roles",
                    "title": "职责分工",
                    "roles": [{"title": "项目组", "points": ["统筹推进"]}],
                },
            ],
        )

        self.assertEqual(result["rendered_slide_types"], ["expert_body_panel", "team_roles"])

    def test_expert_body_panel_can_be_rendered(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_expert_text_panel",
            title="模块解析页",
            theme="expert_forum_blue",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_body_panel",
                    "title": "C3k2_MCA：增强小目标缺陷感知能力",
                    "statement": "这是页头摘要。",
                    "panel_title": "核心观点",
                    "panel_points": ["说明一", "说明二", "说明三"],
                    "body_title": "详细说明",
                    "body_paragraphs": ["第一段正文。", "第二段正文。"],
                }
            ],
        )

        self.assertEqual(result["rendered_slide_types"], ["expert_body_panel"])
        presentation = ppt_mcp_server.presentations["test_expert_text_panel"]
        slide = presentation.slides[0]
        texts = [shape.text.strip() for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        self.assertIn("核心观点", texts)
        self.assertIn("详细说明", texts)
        self.assertTrue(any("第一段正文" in text for text in texts))

    def test_expert_body_panel_does_not_invent_panel_title(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_expert_body_panel_without_default_core_view",
            title="模块解析页",
            theme="expert_forum_blue",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_body_panel",
                    "title": "SBA：增强边界特征与语义特征融合",
                    "statement": "这是页头摘要。",
                    "panel_points": ["说明一", "说明二"],
                    "body_paragraphs": ["第一段正文。", "第二段正文。"],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_expert_body_panel_without_default_core_view"]
        slide = presentation.slides[0]
        texts = [shape.text.strip() for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        self.assertNotIn("核心观点", texts)
        self.assertTrue(any("说明一" in text for text in texts))

    def test_expert_card_overview_does_not_invent_section_title(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_expert_card_overview_without_default_core_view",
            title="专题页测试",
            theme="expert_forum_blue",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_card_overview",
                    "title": "SBA：增强边界特征与语义特征融合",
                    "items": ["边界增强", "语义融合"],
                }
            ],
        )

        presentation = ppt_mcp_server.presentations["test_expert_card_overview_without_default_core_view"]
        slide = presentation.slides[0]
        texts = [shape.text.strip() for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        self.assertNotIn("核心观点", texts)
        self.assertTrue(any("边界增强" in text for text in texts))

    def test_layouts_do_not_invent_content_labels(self):
        unwanted_labels = {
            "核心观点",
            "核心内容",
            "核心结论",
            "对比项",
            "优化后",
            "现状/痛点",
            "目标/方案",
            "核心概念",
            "关系假设 / 分析命题",
            "总体框架",
            "重点工作",
            "数据来源",
            "样本范围",
            "变量设计",
            "分析方法",
            "研究贡献",
            "局限与边界",
            "后续研究方向",
            "条目",
            "图示区域",
            "Left",
            "Right",
        }
        cases = [
            (
                "summary",
                {
                    "type": "summary",
                    "title": "汇总页",
                    "statement": "摘要内容",
                    "text": "补充说明",
                },
            ),
            (
                "comparison",
                {
                    "type": "comparison",
                    "title": "对比页",
                    "comparisons": [{"before": "改造前"}, {"after": "改造后"}],
                },
            ),
            (
                "quote",
                {
                    "type": "quote",
                    "statement": "结论正文",
                },
            ),
            (
                "expert_card_overview",
                {
                    "type": "expert_card_overview",
                    "title": "卡片页",
                    "text": "卡片正文",
                },
            ),
            (
                "theoretical_framework",
                {
                    "type": "theoretical_framework",
                    "title": "框架页",
                    "framework": "框架说明",
                },
            ),
            (
                "party_summary_panel",
                {
                    "type": "party_summary_panel",
                    "title": "党建页",
                    "text": "党建正文",
                },
            ),
            (
                "literature_matrix",
                {
                    "type": "literature_matrix",
                    "title": "文献页",
                    "items": ["文献正文"],
                },
            ),
            (
                "expert_image_text",
                {
                    "type": "expert_image_text",
                    "title": "图文页",
                    "items": ["说明正文"],
                },
            ),
            (
                "method_design",
                {
                    "type": "method_design",
                    "title": "方法页",
                    "content": "方法正文",
                },
            ),
            (
                "contribution_limitations",
                {
                    "type": "contribution_limitations",
                    "title": "贡献页",
                    "contributions": ["贡献内容"],
                    "limitations": ["局限内容"],
                    "implications": ["后续内容"],
                },
            ),
        ]

        for layout_id, slide_spec in cases:
            with self.subTest(layout_id=layout_id):
                presentation_id = f"test_no_invented_labels_{layout_id}"
                call_tool(
                    "generate_presentation",
                    presentation_id=presentation_id,
                    title="默认文案检查",
                    theme="expert_forum_blue",
                    auto_cover=False,
                    show_footer=False,
                    show_page_number=False,
                    slides=[slide_spec],
                )

                presentation = ppt_mcp_server.presentations[presentation_id]
                texts = {
                    shape.text.strip()
                    for shape in presentation.slides[0].shapes
                    if getattr(shape, "has_text_frame", False) and shape.text.strip()
                }
                self.assertFalse(unwanted_labels.intersection(texts))

    def test_empty_slide_list_does_not_create_core_view_title(self):
        call_tool(
            "generate_presentation",
            presentation_id="test_empty_slide_list_without_core_view",
            title="空内容检查",
            subtitle="只有副标题内容",
            theme="expert_forum_blue",
            auto_cover=False,
            auto_closing=False,
            show_footer=False,
            show_page_number=False,
            slides=[],
        )

        presentation = ppt_mcp_server.presentations["test_empty_slide_list_without_core_view"]
        texts = {
            shape.text.strip()
            for shape in presentation.slides[0].shapes
            if getattr(shape, "has_text_frame", False) and shape.text.strip()
        }
        self.assertNotIn("核心观点", texts)

    def test_legacy_ids_still_work(self):
        result = call_tool(
            "generate_presentation",
            presentation_id="test_legacy_ids",
            title="兼容性测试",
            theme="academic_default",
            auto_cover=False,
            show_footer=False,
            show_page_number=False,
            slides=[
                {
                    "type": "expert_text_panel",
                    "title": "旧版式兼容",
                    "panel_points": ["说明一"],
                    "body_paragraphs": ["第一段正文。"],
                },
                {
                    "type": "party_work_summary",
                    "title": "旧党建页兼容",
                    "items": [{"title": "组织建设", "points": ["说明一"]}],
                },
            ],
        )

        self.assertEqual(result["theme"], "expert_forum_blue")
        self.assertEqual(result["rendered_slide_types"], ["expert_body_panel", "party_summary_panel"])


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
