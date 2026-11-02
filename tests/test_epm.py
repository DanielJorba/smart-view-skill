"""Offline tests for epm.py: guard, parsing, rendering, credentials and HTTP against a local fake pod."""
import base64
import contextlib
import importlib.util
import io
import json
import os
import sys
import tempfile
import threading
import unittest
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "skills" / "smart-view" / "scripts" / "epm.py"
spec = importlib.util.spec_from_file_location("epm", SCRIPT)
epm = importlib.util.module_from_spec(spec)
spec.loader.exec_module(epm)

USER, PASSWORD = "test.user", "not-a-real-password"

SLICE_RESPONSE = {
    "pov": ["BaseData", "FY15", "Plan", "Working", "410", "P_160"],
    "columns": [["Jan", "Feb", "Mar", "Q1"]],
    "rows": [
        {"headers": ["Project Number"], "data": ["1", "2", "3", " "]},
        {"headers": ["Project Type"], "data": ["Other", "IT", "Construction", ""]},
        {"headers": ["Project Investment"], "data": ["100000", "110000", "200000", "410000"]},
    ],
}

DIMENSION = {
    "name": "Entity", "path": "/Entity", "children": [
        {"name": "Total Entity", "path": "/Entity/Total Entity", "dataStorage": "Never Share", "level": 2, "children": [
            {"name": "TD", "alias": "Total Department", "path": "/Entity/Total Entity/TD", "dataStorage": "Never Share",
             "level": 1, "children": [{"name": "410", "path": "/Entity/Total Entity/TD/410", "level": 0}]},
            {"name": "Enterprise Global", "path": "/Entity/Total Entity/Enterprise Global", "level": 0},
        ]},
        {"name": "Management Rollup", "path": "/Entity/Management Rollup", "dataStorage": "Dynamic Calc", "level": 1,
         "children": [{"name": "410", "path": "/Entity/Management Rollup/410", "dataStorage": "Shared", "level": 0}]},
    ],
}


class FakePod(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body):
        data = json.dumps(body).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _authorized(self):
        expected = "Basic " + base64.b64encode(f"{USER}:{PASSWORD}".encode()).decode()
        if self.headers.get("Authorization") != expected:
            self._send(401, {"detail": "Invalid credentials"})
            return False
        return True

    def do_GET(self):
        if not self._authorized():
            return
        if self.path == "/HyperionPlanning/rest/v3/applications":
            self._send(200, {"items": [{"name": "Vision", "type": "EPBCS"}]})
        elif self.path.startswith("/HyperionPlanning/rest/v3/applications/Vision/plantypes/Plan1/dimensions/Entity"):
            self._send(200, DIMENSION)
        elif self.path == "/interop/rest/11.1.2.3.600/applicationsnapshots/outbox%2Flogs%2F7.log/contents":
            data = b"Pipeline Name  : 001\nSEND_TO : someone@example.com\n"
            self.send_response(200)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif self.path == "/HyperionPlanning/rest/v3/applications/Vision/jobdefinitions":
            self._send(200, {"items": [
                {"jobType": "RULES", "jobName": "Calc_Revenue", "links": []},
                {"jobType": "RULESET", "jobName": "Close_All"},
                {"jobType": "IMPORT_DATA", "jobName": "Load_GL"},
                {"jobType": "RULES", "jobName": "Agg_Revenue"},
            ]})
        elif self.path == "/aif/rest/V1/jobs/7":
            self._send(200, {"status": "0", "jobId": "7", "jobStatus": "SUCCESS", "processType": "PIPELINE_EXECUTE",
                             "logFileName": "outbox/logs/7.log", "executedBy": "someone@example.com"})
        else:
            self._send(404, {"detail": f"No resource {self.path}"})

    def do_POST(self):
        if not self._authorized():
            return
        length = int(self.headers.get("Content-Length", 0))
        FakePod.last_body = json.loads(self.rfile.read(length))
        self._send(200, SLICE_RESPONSE)


def run(argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            code = epm.main(argv)
        except SystemExit as e:
            code = e.code
    return code, out.getvalue(), err.getvalue()


class PureFunctions(unittest.TestCase):
    def test_migrates_the_former_epm_cloud_folder_once(self):
        with tempfile.TemporaryDirectory() as tmp:
            old, new = Path(tmp) / "epm-cloud-skill", Path(tmp) / "smart-view-skill"
            (old / "cache" / "p").mkdir(parents=True)
            (old / "pods.json").write_text('{"p": {}}', encoding="utf-8")
            (old / "cache" / "p" / "Entity.json").write_text("{}", encoding="utf-8")
            new.mkdir()
            (new / "connections.json").write_text("[]", encoding="utf-8")
            epm._migrate(old, new)
            self.assertEqual((new / "pods.json").read_text(encoding="utf-8"), '{"p": {}}')
            self.assertTrue((new / "cache" / "p" / "Entity.json").exists())
            self.assertTrue((old / "pods.json").exists())
            (new / "pods.json").write_text('{"q": {}}', encoding="utf-8")
            epm._migrate(old, new)
            self.assertEqual((new / "pods.json").read_text(encoding="utf-8"), '{"q": {}}')

    def test_guard_allows_reads_only(self):
        epm._guard("GET", "/HyperionPlanning/rest/v3/applications")
        epm._guard("GET", "/interop/rest/11.1.2.3.600/applicationsnapshots")
        epm._guard("POST", "/HyperionPlanning/rest/v3/applications/A/plantypes/C/exportdataslice")
        blocked = [
            ("POST", "/HyperionPlanning/rest/v3/applications/A/jobs"),
            ("POST", "/HyperionPlanning/rest/v3/applications/A/plantypes/C/importdataslice"),
            ("PUT", "/HyperionPlanning/rest/v3/applications/A/plantypes/C/exportdataslice"),
            ("DELETE", "/interop/rest/11.1.2.3.600/applicationsnapshots/x"),
            ("GET", "/epm/rest/v1/anything"),
        ]
        for method, path in blocked:
            with self.assertRaises(epm.EpmError, msg=f"{method} {path}"):
                epm._guard(method, path)

    def test_raw_get_deny_ignores_query(self):
        self.assertIsNone(epm.RAW_GET_DENY.search("/HyperionPlanning/rest/v3/applications/A/jobs/12"))
        self.assertIsNotNone(epm.RAW_GET_DENY.search("/interop/rest/v1/services/recreate"))

    def test_split_members(self):
        self.assertEqual(epm.split_members('A, B ,ILvl0Descendants(X),"Name, with comma"'),
                         ["A", "B", "ILvl0Descendants(X)", "Name, with comma"])

    def test_build_payload(self):
        p = epm.build_slice_payload("Year=FY26;Scenario=Actual", ["Entity=410,420", "Account=ILvl0Descendants(PL)"],
                                    ["Period=Jan,Feb"])
        g = p["gridDefinition"]
        self.assertEqual(g["pov"], {"dimensions": ["Year", "Scenario"], "members": [["FY26"], ["Actual"]]})
        self.assertEqual(g["rows"][0]["dimensions"], ["Entity", "Account"])
        self.assertEqual(g["rows"][0]["members"], [["410", "420"], ["ILvl0Descendants(PL)"]])
        self.assertTrue(g["suppressMissingRows"])
        with self.assertRaises(epm.EpmError):
            epm.build_slice_payload("Period=Jan", ["Account=A"], ["Period=Feb"])
        with self.assertRaises(epm.EpmError):
            epm.build_slice_payload("", ["Account=A"], [])

    def test_slice_table_and_mask(self):
        pov, labels, rows = epm.slice_table(SLICE_RESPONSE)
        self.assertEqual(labels, ["Jan", "Feb", "Mar", "Q1"])
        self.assertEqual(rows[2], (["Project Investment"], ["100000", "110000", "200000", "410000"]))
        self.assertEqual([epm.mask(v) for v in rows[1][1]], ["Other", "IT", "Construction", ""])
        self.assertEqual(epm.mask("-1.5E3"), "<n>")
        self.assertEqual(epm.mask("12,5"), "<n>")

    def test_tree(self):
        lines = epm.render_tree(DIMENSION, max_depth=1, max_lines=50)
        self.assertEqual(lines[0], "Entity")
        self.assertIn("  Total Entity  [Never Share, L2]  +2 children", lines)
        deep = epm.render_tree(DIMENSION, max_depth=5, max_lines=50)
        self.assertIn("      TD (Total Department)  [Never Share, L1]".strip(), [l.strip() for l in deep])
        self.assertEqual(len(epm.find_nodes(DIMENSION, "410")), 2)
        self.assertTrue(epm.render_tree(DIMENSION, 5, 2)[-1].startswith("..."))

    def test_unmangle_git_bash_path(self):
        self.assertEqual(epm.unmangle_path("C:/Program Files/Git/HyperionPlanning/rest/v3/applications"),
                         "/HyperionPlanning/rest/v3/applications")
        self.assertEqual(epm.unmangle_path("interop/rest/11.1.2.3.600/applicationsnapshots"),
                         "/interop/rest/11.1.2.3.600/applicationsnapshots")
        self.assertEqual(epm.unmangle_path("/aif/rest/V1/jobs/1"), "/aif/rest/V1/jobs/1")

    def test_redact(self):
        self.assertEqual(epm.redact("SEND_TO : a.b@example.com,c@x.es"), "SEND_TO : <user>,<user>")
        self.assertEqual(epm.redact("jdoe_example.com_ExportedData_Year.zip"), "<user>_ExportedData_Year.zip")
        self.assertEqual(epm.redact("data/snapshots/data/AIF_PROCESSES.dat"), "data/snapshots/data/AIF_PROCESSES.dat")

    def test_normalize_url(self):
        self.assertEqual(epm.normalize_url("https://epm-x.example.com/HyperionPlanning/SmartView"),
                         "https://epm-x.example.com")
        self.assertEqual(epm.normalize_url("http://127.0.0.1:8080/x"), "http://127.0.0.1:8080")
        with self.assertRaises(epm.EpmError):
            epm.normalize_url("http://epm-x.example.com")


@unittest.skipUnless(os.name == "nt", "Windows Credential Manager")
class WithFakePod(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        os.environ["SMARTVIEW_SKILL_HOME"] = cls.tmp.name
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), FakePod)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        cls.pod = f"unittest-{uuid.uuid4().hex[:8]}"
        code, out, err = run(["pod-add", cls.pod, "--url", f"http://127.0.0.1:{cls.server.server_port}",
                              "--app", "Vision", "--cube", "Plan1", "--env", "test"])
        assert code == 0, err

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        epm.cred_delete(epm.CRED_PREFIX + cls.pod)
        epm.cred_delete(epm.LEGACY_CRED_PREFIX + cls.pod)
        os.environ.pop("SMARTVIEW_SKILL_HOME", None)
        os.environ.pop(epm.FIGURES_ENV, None)
        cls.tmp.cleanup()

    def setUp(self):
        epm.cred_write(epm.CRED_PREFIX + self.pod, USER, PASSWORD)

    def test_credential_roundtrip(self):
        self.assertEqual(epm.cred_read(epm.CRED_PREFIX + self.pod), (USER, PASSWORD))
        self.assertIsNone(epm.cred_read(epm.CRED_PREFIX + "does-not-exist-" + uuid.uuid4().hex))

    def test_cred_set_refuses_without_terminal(self):
        sys.stdin, saved = io.StringIO("secret\n"), sys.stdin
        try:
            code, out, err = run(["cred-set", self.pod])
        finally:
            sys.stdin = saved
        self.assertEqual(code, 2)
        self.assertIn("interactive", err)

    def test_ping_and_wrong_password(self):
        code, out, err = run(["ping", self.pod])
        self.assertEqual(code, 0, err)
        self.assertIn("Applications: Vision", out)
        epm.cred_write(epm.CRED_PREFIX + self.pod, USER, "wrong")
        code, out, err = run(["ping", self.pod])
        self.assertEqual(code, 1)
        self.assertIn("HTTP 401", err)
        log = (Path(self.tmp.name) / "logs" / "calls.log").read_text(encoding="utf-8")
        self.assertNotIn(PASSWORD, log)
        self.assertNotIn("wrong", log)

    def test_tree_uses_cache(self):
        code, out, err = run(["tree", self.pod, "Entity", "410", "--depth", "0"])
        self.assertEqual(code, 0, err)
        self.assertIn("appears 2 times", out)
        code, out, err = run(["find", self.pod, "Entity", "department"])
        self.assertIn("1 matches", out)
        log = (Path(self.tmp.name) / "logs" / "calls.log").read_text(encoding="utf-8")
        self.assertEqual(log.count("/dimensions/Entity"), 1)

    def test_slice_masks_unless_allowed(self):
        os.environ[epm.FIGURES_ENV] = "0"
        code, out, err = run(["slice", self.pod, "--pov", "Year=FY15;Scenario=Plan", "--rows", "Account=A,B",
                              "--cols", "Period=IDescendants(Q1)"])
        self.assertEqual(code, 0, err)
        self.assertIn("Project Investment\t<n>\t<n>\t<n>\t<n>", out)
        self.assertNotIn("410000", out)
        self.assertEqual(FakePod.last_body["gridDefinition"]["columns"][0]["members"], [["IDescendants(Q1)"]])
        os.environ[epm.FIGURES_ENV] = "1"
        code, out, err = run(["slice", self.pod, "--pov", "Year=FY15", "--rows", "Account=A", "--cols", "Period=Q1"])
        self.assertIn("410000", out)
        code, out, err = run(["slice", self.pod, "--pov", "Year=FY15", "--rows", "Account=A", "--cols", "Period=Q1",
                              "--hide-figures"])
        self.assertNotIn("410000", out)

    def test_slice_digest(self):
        os.environ[epm.FIGURES_ENV] = "0"
        args = ["slice", self.pod, "--pov", "Year=FY15", "--rows", "Account=A", "--cols", "Period=Q1", "--digest"]
        code, out, err = run(args)
        self.assertEqual(code, 0, err)
        self.assertNotIn("410000", out)
        self.assertNotIn("Project Investment\t<n>", out)
        lines = dict((l.split("\t")[1], l.split("\t")[2]) for l in out.splitlines() if l.startswith(("row\t", "col\t")))
        self.assertEqual(len(lines["Jan"]), 12)
        self.assertNotEqual(lines["Jan"], lines["Feb"])
        self.assertEqual(run(args)[1], out)
        key = epm.digest_key()
        self.assertEqual(epm.digest(key, ["100000", ""]), epm.digest(key, ["100000.0", ""]))

    def test_slice_to_file_needs_the_switch(self):
        os.environ[epm.FIGURES_ENV] = "0"
        target = Path(self.tmp.name) / "out" / "slice.xlsx"
        args = ["slice", self.pod, "--pov", "Year=FY15", "--rows", "Account=A", "--cols", "Period=Q1",
                "--out", str(target)]
        code, out, err = run(args)
        self.assertEqual(code, 2)
        self.assertIn(epm.FIGURES_ENV, err)
        self.assertFalse(target.exists() or target.with_suffix(".csv").exists())
        os.environ[epm.FIGURES_ENV] = "1"
        code, out, err = run(args)
        self.assertEqual(code, 0, err)
        written = target if target.exists() else target.with_suffix(".csv")
        self.assertTrue(written.exists())
        if written.suffix == ".xlsx":
            from openpyxl import load_workbook
            ws = load_workbook(written)["Data"]
            self.assertEqual(ws.cell(row=4, column=5).value, 410000)

    def test_rules_are_cached_and_searched(self):
        code, out, err = run(["rules", self.pod, "--refresh"])
        self.assertEqual(code, 0, err)
        self.assertIn("3 rules and rulesets", out)
        self.assertNotIn("Load_GL", out)
        code, out, err = run(["rules", self.pod, "revenue"])
        self.assertIn("2 rules and rulesets matching 'revenue'", out)
        code, out, err = run(["jobdefs", self.pod, "--type", "import_data"])
        self.assertIn("Load_GL", out)
        log = (Path(self.tmp.name) / "logs" / "calls.log").read_text(encoding="utf-8")
        self.assertEqual(log.count("/jobdefinitions"), 1)

    def test_sv_connection(self):
        code, out, err = run(["sv", self.pod])
        self.assertEqual(code, 0, err)
        self.assertIn("Inline connection: https://127.0.0.1/HyperionPlanning/SmartView|127.0.0.1|Vision|Plan1", out)
        self.assertIn(f"pod:{self.pod}", out)

    def test_legacy_credential_still_works(self):
        epm.cred_delete(epm.CRED_PREFIX + self.pod)
        epm.cred_write(epm.LEGACY_CRED_PREFIX + self.pod, USER, PASSWORD)
        try:
            code, out, err = run(["ping", self.pod])
            self.assertEqual(code, 0, err)
        finally:
            epm.cred_delete(epm.LEGACY_CRED_PREFIX + self.pod)

    def test_users_are_never_printed(self):
        code, out, err = run(["status"])
        self.assertEqual(code, 0, err)
        self.assertIn("credential set", out)
        self.assertNotIn(USER, out)
        code, out, err = run(["cred-show", self.pod])
        self.assertNotIn(USER, out)
        code, out, err = run(["cat", self.pod, "outbox/logs/7.log"])
        self.assertEqual(code, 0, err)
        self.assertIn("Pipeline Name  : 001", out)
        self.assertIn("SEND_TO : <user>", out)
        code, out, err = run(["dijob", self.pod, "7"])
        self.assertEqual(code, 0, err)
        self.assertIn("Process 7 PIPELINE_EXECUTE: success (SUCCESS)", out)
        self.assertNotIn("someone", out)

    def test_build_permission(self):
        from unittest import mock
        test_pod, prod_pod = f"{self.pod}-t", f"{self.pod}-p"
        run(["pod-add", test_pod, "--url", "https://epm1-test-example.epm.example.com/HyperionPlanning", "--app", "A"])
        run(["pod-add", prod_pod, "--url", "https://epm1-example.epm.example.com", "--app", "A"])
        self.assertEqual(run(["build-check", test_pod])[0], 3)
        self.assertEqual(run(["build-check", self.pod])[0], 3)
        self.assertEqual(run(["buildlog", test_pod, "created", "X"])[0], 3)
        code, out, err = run(["allow-build", test_pod])
        self.assertEqual(code, 2)
        self.assertIn("interactive", err)
        with mock.patch.object(epm, "_interactive", return_value=True), \
                mock.patch("builtins.input", return_value=prod_pod):
            code, out, err = run(["allow-build", prod_pod])
        self.assertEqual(code, 2)
        self.assertIn("not a test pod", err)
        with mock.patch.object(epm, "_interactive", return_value=True), \
                mock.patch("builtins.input", return_value="wrong"):
            self.assertEqual(run(["allow-build", test_pod])[0], 2)
        with mock.patch.object(epm, "_interactive", return_value=True), \
                mock.patch("builtins.input", return_value=test_pod):
            self.assertEqual(run(["allow-build", test_pod])[0], 0)
        code, out, err = run(["build-check", test_pod])
        self.assertEqual(code, 0, err)
        self.assertIn("https://epm1-test-example.epm.example.com/", out)
        self.assertIn("BUILD enabled", run(["status"])[1])
        self.assertEqual(run(["buildlog", test_pod, "created", "integration", "Meta_X"])[0], 0)
        self.assertIn("created integration Meta_X", run(["buildlog", test_pod, "--show"])[1])
        run(["pod-add", test_pod, "--url", "https://epm2-test-example.epm.example.com", "--app", "A"])
        self.assertEqual(run(["build-check", test_pod])[0], 3)
        with mock.patch.object(epm, "_interactive", return_value=True), \
                mock.patch("builtins.input", return_value=test_pod):
            run(["allow-build", test_pod])
        run(["deny-build", test_pod])
        self.assertEqual(run(["build-check", test_pod])[0], 3)

    def test_blocked_raw_get(self):
        code, out, err = run(["get", self.pod, "/interop/rest/v1/services/recreate"])
        self.assertEqual(code, 2)
        self.assertIn("Blocked", err)


if __name__ == "__main__":
    unittest.main(verbosity=2)
