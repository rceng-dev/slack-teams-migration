import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from slack_migrate import inspect_exports
from slack_to_teams.slack_to_teams import Config, _unique_message_times, slack_to_html


class MigrationTests(unittest.TestCase):
    def test_slack_markup_conversion(self):
        users = {"U1": {"profile": {"real_name": "Ada Lovelace"}}}
        rendered = slack_to_html("Hi <@U1> in <#C1|general> - *done*", users, {})
        self.assertIn("<b>@Ada Lovelace</b>", rendered)
        self.assertIn("<b>#general</b>", rendered)
        self.assertIn("<b>done</b>", rendered)

    def test_duplicate_millisecond_timestamps_are_made_unique(self):
        with tempfile.TemporaryDirectory() as tmp:
            channel = Path(tmp)
            (channel / "2026-01-01.json").write_text(
                json.dumps([{"ts": "1.0001"}, {"ts": "1.0002"}]), encoding="utf-8"
            )
            times = list(_unique_message_times(channel).values())
            self.assertEqual(2, len(set(times)))

    def test_config_uses_environment_secret_and_relative_paths(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "tenantId": "tenant",
                        "clientId": "client",
                        "clientSecret": "",
                        "ownerUpn": "owner@example.com",
                        "exportPath": "export",
                        "stateDir": "state",
                        "teamName": "Archive",
                    }
                ),
                encoding="utf-8",
            )
            with patch.dict(os.environ, {"SLACK_MIGRATION_CLIENT_SECRET": "secret"}):
                config = Config(config_path)
            self.assertEqual("secret", config.client_secret)
            self.assertEqual(root / "export", config.export_path)
            self.assertEqual(root / "state", config.state_dir)

    def test_inspection_reads_both_export_shapes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            channel = root / "channels"
            dm = root / "dms"
            channel.mkdir()
            dm.joinpath("one-to-one").mkdir(parents=True)
            channel.joinpath("users.json").write_text('[{"id":"U1"}]', encoding="utf-8")
            channel.joinpath("channels.json").write_text('[{"name":"general"}]', encoding="utf-8")
            channel.joinpath("general").mkdir()
            channel.joinpath("general", "2026-01-01.json").write_text("[]", encoding="utf-8")
            dm.joinpath("one-to-one", "messages.json").write_text(
                '{"users":[],"messages":[{"ts":"1"}]}', encoding="utf-8"
            )
            self.assertEqual(0, inspect_exports(channel, dm))


if __name__ == "__main__":
    unittest.main()
