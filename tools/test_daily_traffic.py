"""Execute the workflow's embedded scripts against isolated traffic and Git fixtures."""
from __future__ import annotations

import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from test_check_workflows import WORKFLOWS

JOBS = WORKFLOWS["daily-traffic.yml"]["jobs"]


def script(job: str, name: str) -> str:
    return str(next(step["run"] for step in JOBS[job]["steps"] if step["name"] == name))


def embedded(job: str, name: str) -> str:
    match = re.search(r"<< 'PY'\n(.*?)\nPY", script(job, name), re.S)
    assert match is not None
    return match[1]


class TrafficTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.work = self.root / "work"
        self.work.mkdir()
        self.env = {**os.environ, "PUSH_TOKEN": "fixture", "TRAFFIC_TOKEN": "fixture"}
        self.build = self.root / "build-traffic.py"
        self.build.write_text(embedded("traffic", "Build CSV history").replace("/tmp/", f"{self.root}/"))
        self.snapshot("2026-11-07", clones=10)

    def run_script(self, source: str, *, python: bool = True) -> subprocess.CompletedProcess[str]:
        return subprocess.run([sys.executable, "-c", source] if python else ["bash", "-c", source],
                              cwd=self.work, env=self.env, capture_output=True, text=True)

    def build_data(self) -> None:
        result = self.run_script(self.build.read_text())
        self.assertEqual(result.returncode, 0, result.stderr)

    def snapshot(self, day: str, *, clones: int = 10, views: int = 50,
                 stars: int = 3, forks: int = 2, referrer: str = "") -> None:
        for metric, count in (("views", views), ("clones", clones)):
            (self.root / f"{metric}.json").write_text(json.dumps({"count": count, "uniques": 2,
                metric: [{"timestamp": day + "T00:00:00Z", "count": count, "uniques": 2}]}))
        (self.root / "repo.json").write_text(json.dumps({"stargazers_count": stars, "forks_count": forks}))
        (self.root / "referrers.json").write_text(json.dumps(
            [{"referrer": referrer, "count": 5, "uniques": 2}] if referrer else []))
        (self.root / "paths.json").write_text("[]")
        (self.root / "traffic-snapshot-time").write_text(day + "T12:00:00+00:00")

    def rows(self, name: str) -> list[dict[str, str]]:
        with (self.work / "data" / name).open(newline="") as stream:
            return list(csv.DictReader(stream))

    def alert(self) -> list[list[str]]:
        # Capture gh invocations without accessing GitHub or needing credentials.
        prefix = """import subprocess, json
calls = []
def fake_run(args, **kwargs):
    calls.append(args)
    return subprocess.CompletedProcess(args, 0, '', '')
subprocess.run = fake_run
import atexit
atexit.register(lambda: print('CALLS=' + json.dumps(calls)))
"""
        result = self.run_script(prefix + embedded("alert", "Create alert issue on spike or new referrer"))
        self.assertEqual(result.returncode, 0, result.stderr)
        calls: list[list[str]] = json.loads(result.stdout.split("CALLS=")[1])
        return calls

    def test_no_change_and_clone_spike_do_not_alert_but_keep_clones(self) -> None:
        self.build_data()
        self.snapshot("2026-11-08")
        self.build_data()
        self.assertEqual(self.alert(), [])
        self.snapshot("2026-11-08", clones=100)
        self.build_data()
        self.assertEqual(self.alert(), [])
        self.assertEqual(self.rows("traffic.csv")[-1]["clones_14d"], "100")
        self.assertEqual(self.rows("traffic_daily.csv")[-1]["clones"], "100")
        badge = json.loads((self.work / "data/badges/clones.json").read_text())
        self.assertEqual(badge["message"], "110")

    def test_each_remaining_trigger_creates_an_alert(self) -> None:
        self.build_data()
        cases = ({"views": 150}, {"stars": 4}, {"forks": 3}, {"referrer": "example.org"})
        for changes in cases:
            with self.subTest(changes=changes):
                self.snapshot("2026-11-08", **changes)
                self.build_data()
                calls = self.alert()
                self.assertEqual([call[2] for call in calls], ["list", "create"])
                self.assertIn("- Clones: 10 vs 10", calls[1][-1])
                self.assertNotIn("Clones spike", calls[1][-1])

    def test_bad_response_leaves_history_unchanged(self) -> None:
        self.build_data()
        before = {p: p.read_bytes() for p in (self.work / "data").rglob("*") if p.is_file()}
        self.snapshot("2026-11-08")
        (self.root / "paths.json").write_text('{"message": "failure"}')
        result = self.run_script(self.build.read_text())
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(before, {p: p.read_bytes() for p in before})

    def test_older_snapshot_cannot_replace_newer_remote_data(self) -> None:
        self.snapshot("2026-11-09", clones=500)
        self.build_data()
        before = (self.work / "data/traffic.csv").read_bytes()
        self.snapshot("2026-11-08", clones=20)
        self.build_data()
        self.assertEqual((self.work / "data/traffic.csv").read_bytes(), before)

    def git(self, *args: str, cwd: Path | None = None) -> str:
        result = subprocess.run(["git", *args], cwd=cwd or self.work, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.strip()

    @unittest.skipUnless(shutil.which("bash") and shutil.which("git"), "requires Bash and Git")
    def test_rejected_push_rebuilds_on_concurrent_remote_history(self) -> None:
        remote = self.root / "remote.git"
        self.git("init", "--bare", str(remote))
        self.git("init", "-b", "traffic-history")
        self.git("config", "user.email", "fixture@example.invalid")
        self.git("config", "user.name", "Traffic fixture")
        self.build_data()
        self.git("add", ".")
        self.git("commit", "-m", "Initial snapshot")
        self.git("remote", "add", "origin", str(remote))
        self.git("push", "origin", "HEAD:traffic-history")
        # A second writer publishes after this checkout was prepared.
        other = self.root / "other"
        self.git("clone", "-b", "traffic-history", str(remote), str(other))
        (other / "concurrent.txt").write_text("Keep this concurrent update\n")
        with (other / "data/traffic_daily.csv").open("a") as stream:
            stream.write("2026-11-01,7,3,4,2\n")
        self.git("add", ".", cwd=other)
        self.git("-c", "user.name=Other", "-c", "user.email=other@example.invalid",
                 "commit", "-m", "Concurrent history", cwd=other)
        concurrent = self.git("rev-parse", "HEAD", cwd=other)
        self.git("push", "origin", "HEAD", cwd=other)
        self.snapshot("2026-11-08", clones=100)
        self.build_data()
        source = script("traffic", "Commit traffic history").replace("/tmp/", f"{self.root}/")
        result = self.run_script(source, python=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("rejected", result.stderr)
        self.git("merge-base", "--is-ancestor", concurrent, "HEAD")
        self.assertTrue((self.work / "concurrent.txt").exists())
        self.assertEqual({r["date"] for r in self.rows("traffic_daily.csv")},
                         {"2026-11-01", "2026-11-07", "2026-11-08"})
        self.assertEqual(self.git("rev-parse", "HEAD"),
                         self.git("rev-parse", "traffic-history", cwd=remote))

        # Persistent rejection stops at three attempts and publishes nothing.
        hook = remote / "hooks/pre-receive"
        hook.write_text("#!/bin/sh\nexit 1\n")
        hook.chmod(0o755)
        accepted = self.git("rev-parse", "traffic-history", cwd=remote)
        self.snapshot("2026-11-09", clones=200)
        self.build_data()
        result = self.run_script(source, python=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(result.stderr.count("[remote rejected]"), 3)
        self.assertIn("failed after three attempts", result.stdout)
        self.assertEqual(self.git("rev-parse", "traffic-history", cwd=remote), accepted)

    @unittest.skipUnless(shutil.which("bash") and shutil.which("curl"), "requires Bash and curl")
    def test_http_retry_and_persistent_failure(self) -> None:
        statuses = [503, 200]
        requests: list[str] = []

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:
                requests.append(self.path)
                self.send_response(statuses.pop(0) if statuses else 200)
                self.end_headers()
                self.wfile.write(b"{}")

            def log_message(self, format: str, *args: object) -> None:
                pass

        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.addCleanup(server.server_close)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.shutdown)
        source = next(str(step["run"]) for step in JOBS["traffic"]["steps"]
                      if "TRAFFIC_TOKEN" in step.get("env", {}))
        source = source.replace("${{ github.repository }}", "fixture/repo")
        source = source.replace("https://api.github.com/repos/$REPO",
                                f"http://127.0.0.1:{server.server_port}")
        source = source.replace("/tmp/", f"{self.root}/")
        for bound in ("--retry 5", "--retry-max-time 300", "--max-time 60", "--connect-timeout 15"):
            self.assertIn(bound, source)
        self.build_data()
        before = (self.work / "data/traffic.csv").read_bytes()
        result = self.run_script(source, python=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(requests), 6)  # five endpoints, first attempted twice
        requests.clear()
        statuses[:] = [401]
        result = self.run_script(source, python=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(len(requests), 1)
        self.assertEqual((self.work / "data/traffic.csv").read_bytes(), before)

    def test_job_credentials_and_permissions_are_separated(self) -> None:
        self.assertEqual(JOBS["traffic"]["permissions"], {"contents": "write"})
        self.assertEqual(JOBS["alert"]["permissions"], {"issues": "write"})
        self.assertEqual(JOBS["alert"]["needs"], "traffic")
        self.assertNotIn("environment", JOBS["alert"])
        for step in JOBS["alert"]["steps"]:
            self.assertNotIn("actions/checkout", step.get("uses", ""))
            self.assertNotIn("secrets.", repr(step))
            self.assertNotIn("git push", step.get("run", ""))


if __name__ == "__main__":
    unittest.main()
