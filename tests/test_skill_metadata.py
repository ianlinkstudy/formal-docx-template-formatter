from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


class SkillMetadataTests(unittest.TestCase):
    def test_skill_identity_and_brand(self) -> None:
        skill_text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        frontmatter = skill_text.split("---", 2)[1]
        self.assertIn("name: formal-docx-template-formatter", frontmatter)
        self.assertIn("short-description:", frontmatter)
        self.assertIn("compatibility:", frontmatter)
        self.assertIn("鼓翼成长星球", skill_text)
        self.assertIn(
            "鼓翼成长星球",
            (ROOT / "README.md").read_text(encoding="utf-8"),
        )

    def test_release_version_is_consistent(self) -> None:
        version = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
        self.assertRegex(version, r"^\d+\.\d+\.\d+$")
        changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
        self.assertIn(version, changelog)
        self.assertTrue((ROOT / "docs" / f"release-notes-v{version}.md").is_file())

    def test_publishable_skill_contains_runtime_assets(self) -> None:
        for relative in (
            "agents/openai.yaml",
            "profiles/gb9704.json",
            "requirements.txt",
            "scripts/format_docx.py",
            "scripts/build_release.py",
            "scripts/template_cli.py",
            "scripts/verify_docx.py",
            "references/template-workflow.md",
            "references/template-schema.md",
            "references/verification-and-objects.md",
        ):
            self.assertTrue((ROOT / relative).is_file(), relative)
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
        self.assertRegex(requirements, re.compile(r"python-docx"))
        self.assertRegex(requirements, re.compile(r"docxcompose"))


if __name__ == "__main__":
    unittest.main()
