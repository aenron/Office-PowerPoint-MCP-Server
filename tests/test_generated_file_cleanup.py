import os
import tempfile
import unittest
from datetime import date, datetime, time, timedelta

import ppt_mcp_server
import utils as ppt_utils


def call_tool(name, **kwargs):
    return ppt_mcp_server.app._tool_manager._tools[name].fn(**kwargs)


class GeneratedFileCleanupTests(unittest.TestCase):
    def temporary_directory(self):
        os.makedirs(r"C:\tmp", exist_ok=True)
        return tempfile.TemporaryDirectory(dir=r"C:\tmp")

    def set_file_date(self, path, file_date):
        timestamp = datetime.combine(file_date, time(12, 0)).timestamp()
        os.utime(path, (timestamp, timestamp))

    def test_cleanup_stale_generated_files_keeps_only_today_pptx(self):
        today = date(2026, 5, 9)
        yesterday = today - timedelta(days=1)

        with self.temporary_directory() as directory:
            today_pptx = os.path.join(directory, "today.pptx")
            stale_pptx = os.path.join(directory, "stale.pptx")
            stale_text = os.path.join(directory, "stale.txt")

            for path in [today_pptx, stale_pptx, stale_text]:
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write("x")

            self.set_file_date(today_pptx, today)
            self.set_file_date(stale_pptx, yesterday)
            self.set_file_date(stale_text, yesterday)

            result = ppt_utils.cleanup_stale_generated_files(directory, keep_date=today)

            self.assertTrue(os.path.exists(today_pptx))
            self.assertFalse(os.path.exists(stale_pptx))
            self.assertTrue(os.path.exists(stale_text))
            self.assertEqual(result["deleted_files"], [stale_pptx])
            self.assertEqual(result["failed_files"], [])

    def test_export_presentation_does_not_clean_custom_output_directory(self):
        today = date.today()
        yesterday = today - timedelta(days=1)

        with self.temporary_directory() as directory:
            stale_pptx = os.path.join(directory, "old-export.pptx")
            with open(stale_pptx, "w", encoding="utf-8") as handle:
                handle.write("old")
            self.set_file_date(stale_pptx, yesterday)

            call_tool(
                "generate_presentation",
                presentation_id="cleanup_export_test",
                title="清理测试",
                theme="business_blue",
                auto_cover=False,
                show_footer=False,
                show_page_number=False,
                slides=[
                    {
                        "type": "summary",
                        "title": "今日文件",
                        "items": [{"title": "内容", "points": ["说明"]}],
                    }
                ],
            )
            result = call_tool(
                "export_presentation",
                presentation_id="cleanup_export_test",
                file_name="new-export.pptx",
                output_directory=directory,
            )

            self.assertNotIn("error", result)
            self.assertTrue(os.path.exists(stale_pptx))
            self.assertTrue(os.path.exists(os.path.join(directory, "new-export.pptx")))
            self.assertEqual(result["cleanup"]["deleted_files"], [])
            self.assertIn("skipped", result["cleanup"])


if __name__ == "__main__":
    unittest.main()
