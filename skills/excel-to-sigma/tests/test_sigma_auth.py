"""Sigma auth resolver: resolution order, CLI profile, refresh, no-leak, vendored hashes.

Everything is mocked: no network, no `sigma` binary, no keychain, no browser.
"""
from __future__ import annotations

import contextlib
import hashlib
import io
import json
import os
import re
import sys
import tempfile
import time
import types
import unittest
import urllib.error
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))

import get_token  # noqa: E402
import sigma_auth  # noqa: E402

from test_draft_mode import FLAT, FORECAST, Fixtures, run  # noqa: E402

BASE = "https://aws-api.sigmacomputing.com"
OTHER_BASE = "https://api.us-a.aws.sigmacomputing.com"
ENV_TOKEN = "env-token-AAAAAAAA"
FILE_TOKEN = "file-token-BBBBBBBB"
CLI_TOKEN = "cli-token-CCCCCCCC"
CLI_TOKEN_2 = "cli-token-DDDDDDDD"
PROVIDER_TOKEN = "provider-token-EEEEEEEE"
ALL_TOKENS = (ENV_TOKEN, FILE_TOKEN, CLI_TOKEN, CLI_TOKEN_2, PROVIDER_TOKEN)

SECRET_ENV_KEYS = ("SIGMA_API_TOKEN", "SIGMA_BASE_URL", "SIGMA_WORKDIR", "SIGMA_CLI_PROFILE",
                   "SIGMA_TOKEN_MINTED_AT", "SIGMA_AUTH_METHOD", "SIGMA_AUTH_MODE",
                   "SIGMA_CLIENT_ID", "SIGMA_CLIENT_SECRET")


def iso(epoch):
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(epoch))


def provider_result(token=PROVIDER_TOKEN, base=BASE, minted=None, method="browser"):
    return get_token.TokenResult(base, token, minted or iso(time.time()), method)


class FakeCli:
    """Stand-in for subprocess.run + shutil.which for the `sigma` binary."""

    def __init__(self, tokens=(CLI_TOKEN,), base=BASE, token_rc=0, status_rc=0, status_text=None):
        self.tokens = list(tokens)
        self.token_calls = 0
        self.calls = []
        self.token_rc = token_rc
        self.status_rc = status_rc
        self.status_text = status_text if status_text is not None else (
            f"Profile:    work\nBase URL:   {base}\nAuth: OK (token acquired)\n")

    def run(self, argv, **kwargs):
        self.calls.append(list(argv))
        assert kwargs.get("capture_output") and not kwargs.get("shell")
        if argv[1:3] == ["auth", "token"]:
            token = self.tokens[min(self.token_calls, len(self.tokens) - 1)]
            self.token_calls += 1
            return types.SimpleNamespace(returncode=self.token_rc,
                                         stdout=(token + "\n") if self.token_rc == 0 else "",
                                         stderr="")
        if argv[1:3] == ["auth", "status"]:
            return types.SimpleNamespace(returncode=self.status_rc, stdout=self.status_text, stderr="")
        raise AssertionError(argv)


class AuthTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        self.cwd = self.dir / "cwd"
        self.cwd.mkdir()
        previous = os.getcwd()
        os.chdir(self.cwd)
        self.addCleanup(os.chdir, previous)
        env = mock.patch.dict(os.environ, {}, clear=False)
        env.start()
        self.addCleanup(env.stop)
        for key in SECRET_ENV_KEYS:
            os.environ.pop(key, None)
        # Never read the real ~/.sigma-migration/env.
        neutral = mock.patch.object(get_token, "_NEUTRAL_ENV", str(self.dir / "no-such-env"))
        neutral.start()
        self.addCleanup(neutral.stop)
        sigma_auth.reset()
        self.addCleanup(sigma_auth.reset)
        self.mint = mock.patch.object(get_token, "mint_token")
        self.mint_mock = self.mint.start()
        self.addCleanup(self.mint.stop)
        self.mint_mock.side_effect = AssertionError("provider reached")

    def use_cli(self, cli, found=True):
        which = mock.patch.object(sigma_auth.shutil, "which",
                                  return_value="/usr/bin/sigma" if found else None)
        runner = mock.patch.object(sigma_auth.subprocess, "run", side_effect=cli.run)
        which.start(); runner.start()
        self.addCleanup(which.stop); self.addCleanup(runner.stop)

    def write_auth_json(self, directory, token=FILE_TOKEN, base=BASE, minted=None, **extra):
        directory.mkdir(parents=True, exist_ok=True)
        payload = {"SIGMA_API_TOKEN": token, "SIGMA_BASE_URL": base,
                   "SIGMA_TOKEN_MINTED_AT": minted or iso(time.time()),
                   "SIGMA_AUTH_METHOD": "browser", **extra}
        path = directory / "auth.json"
        path.write_text(json.dumps(payload), encoding="utf-8")
        return path


class ResolutionOrderTests(AuthTestCase):
    def test_env_token_wins_over_everything(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE,
                           "SIGMA_WORKDIR": str(self.dir / "w")})
        self.write_auth_json(self.dir / "w")
        self.write_auth_json(self.cwd)
        cli = FakeCli()
        self.use_cli(cli)
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        stderr = io.StringIO()
        with contextlib.redirect_stderr(stderr):
            base, token, method = sigma_auth.get_session()
        self.assertEqual((base, token), (BASE, ENV_TOKEN))
        self.assertEqual(method, "env")
        self.assertEqual(cli.calls, [])
        self.mint_mock.assert_not_called()
        self.assertIn("takes precedence over the Sigma CLI profile", stderr.getvalue())
        self.assertNotIn(ENV_TOKEN, stderr.getvalue())

    def test_env_token_without_base_url_is_an_error(self):
        os.environ["SIGMA_API_TOKEN"] = ENV_TOKEN
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "SIGMA_BASE_URL is not;"):
            sigma_auth.get_session()

    def test_env_base_url_must_be_a_published_sigma_host(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": "https://evil.example.com"})
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "non-Sigma host"):
            sigma_auth.get_session()
        os.environ["SIGMA_BASE_URL"] = "https://nope.sigmacomputing.com"
        sigma_auth.reset()
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "published Sigma API host"):
            sigma_auth.get_session()

    def test_auth_json_in_workdir_beats_cwd(self):
        os.environ["SIGMA_WORKDIR"] = str(self.dir / "w")
        self.write_auth_json(self.dir / "w", token=FILE_TOKEN)
        self.write_auth_json(self.cwd, token=CLI_TOKEN)
        base, token, method = sigma_auth.get_session()
        self.assertEqual((base, token, method), (BASE, FILE_TOKEN, "browser"))
        self.mint_mock.assert_not_called()

    def test_auth_json_in_cwd_is_the_next_fallback(self):
        os.environ["SIGMA_WORKDIR"] = str(self.dir / "empty")
        self.write_auth_json(self.cwd, token=FILE_TOKEN)
        self.assertEqual(sigma_auth.get_session()[1], FILE_TOKEN)

    def test_auth_json_beats_cli_profile_and_provider(self):
        self.write_auth_json(self.cwd)
        cli = FakeCli()
        self.use_cli(cli)
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        self.assertEqual(sigma_auth.get_session()[1], FILE_TOKEN)
        self.assertEqual(cli.calls, [])

    def test_auth_json_base_url_conflicting_with_env_fails_loudly(self):
        self.write_auth_json(self.cwd, base=BASE)
        os.environ["SIGMA_BASE_URL"] = OTHER_BASE
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "differs from the base URL"):
            sigma_auth.get_session()

    def test_broken_auth_json_fails_loudly_without_echoing_contents(self):
        (self.cwd / "auth.json").write_text("{not json " + FILE_TOKEN, encoding="utf-8")
        with self.assertRaises(sigma_auth.SigmaAuthError) as raised:
            sigma_auth.get_session()
        self.assertNotIn(FILE_TOKEN, str(raised.exception))
        (self.cwd / "auth.json").write_text(json.dumps({"SIGMA_BASE_URL": BASE}), encoding="utf-8")
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "no SIGMA_API_TOKEN"):
            sigma_auth.get_session()

    def test_cli_profile_is_used_when_no_token_is_present(self):
        cli = FakeCli()
        self.use_cli(cli)
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        base, token, method = sigma_auth.get_session()
        self.assertEqual((base, token, method), (BASE, CLI_TOKEN, "sigma-cli"))
        self.assertEqual(cli.calls, [["/usr/bin/sigma", "auth", "token", "-p", "work"],
                                     ["/usr/bin/sigma", "auth", "status", "-p", "work"]])
        self.mint_mock.assert_not_called()

    def test_provider_is_the_last_resort(self):
        self.mint_mock.side_effect = None
        self.mint_mock.return_value = provider_result()
        base, token, method = sigma_auth.get_session()
        self.assertEqual((base, token, method), (BASE, PROVIDER_TOKEN, "browser"))
        self.mint_mock.assert_called_once_with()

    def test_provider_failure_names_every_way_to_sign_in(self):
        self.mint_mock.side_effect = get_token.TokenProviderError("no Sigma authentication is available")
        with self.assertRaises(sigma_auth.SigmaAuthError) as raised:
            sigma_auth.get_session()
        text = str(raised.exception)
        self.assertIn("browser-login.sh", text)
        self.assertIn("--sigma-profile", text)
        self.assertIn("SIGMA_CLIENT_ID", text)

    def test_session_is_resolved_once(self):
        self.mint_mock.side_effect = None
        self.mint_mock.return_value = provider_result()
        sigma_auth.get_session()
        sigma_auth.get_session()
        self.assertEqual(self.mint_mock.call_count, 1)

    def test_session_repr_never_contains_the_token(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        sigma_auth.get_session()
        self.assertNotIn(ENV_TOKEN, repr(sigma_auth._SESSION))


class CliProfileTests(AuthTestCase):
    def setUp(self):
        super().setUp()
        os.environ["SIGMA_CLI_PROFILE"] = "work"

    def test_base_url_is_parsed_from_the_status_text_line(self):
        cli = FakeCli(base=OTHER_BASE,
                      status_text=("Profile: work\nUser:       a@b.c\n"
                                   f"Base URL:   {OTHER_BASE}\nAuth: OK (token acquired)\n"))
        self.use_cli(cli)
        self.assertEqual(sigma_auth.get_session()[0], OTHER_BASE)

    def test_missing_sigma_binary_message(self):
        self.use_cli(FakeCli(), found=False)
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, r"Sigma CLI \('sigma'\) is not on PATH"):
            sigma_auth.get_session()

    def test_not_signed_in_message_names_the_login_command(self):
        self.use_cli(FakeCli(token_rc=1))
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError,
                                    r"not signed in.*sigma auth login -p work"):
            sigma_auth.get_session()

    def test_empty_token_output_is_not_signed_in(self):
        self.use_cli(FakeCli(tokens=("",)))
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "not signed in"):
            sigma_auth.get_session()

    def test_base_url_mismatch_with_env_fails_loudly(self):
        self.use_cli(FakeCli(base=BASE))
        os.environ["SIGMA_BASE_URL"] = OTHER_BASE
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "differs from Sigma CLI profile"):
            sigma_auth.get_session()

    def test_matching_env_base_url_is_fine(self):
        self.use_cli(FakeCli(base=BASE))
        os.environ["SIGMA_BASE_URL"] = BASE + "/"  # normalized to the origin
        self.assertEqual(sigma_auth.get_session()[0], BASE)

    def test_env_base_url_stands_in_when_status_has_no_base_url_line(self):
        self.use_cli(FakeCli(status_text="Auth: OK (token acquired)\n"))
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "could not read the base URL"):
            sigma_auth.get_session()
        os.environ["SIGMA_BASE_URL"] = BASE
        sigma_auth.reset()
        self.assertEqual(sigma_auth.get_session()[0], BASE)

    def test_cli_base_url_must_be_a_published_sigma_host(self):
        self.use_cli(FakeCli(base="https://evil.example.com"))
        with self.assertRaisesRegex(sigma_auth.SigmaAuthError, "non-Sigma host"):
            sigma_auth.get_session()

    def test_profile_names_cannot_smuggle_flags(self):
        for bad in ("-x", "a b", "", "a;b", "../x"):
            with self.assertRaises(sigma_auth.SigmaAuthError, msg=bad):
                sigma_auth.use_cli_profile(bad)

    def test_use_cli_profile_sets_the_environment_variable(self):
        sigma_auth.use_cli_profile("other")
        self.assertEqual(os.environ["SIGMA_CLI_PROFILE"], "other")

    def test_token_with_odd_characters_is_refused(self):
        self.use_cli(FakeCli(tokens=("bad token\nwith newline",)))
        with self.assertRaises(sigma_auth.SigmaAuthError) as raised:
            sigma_auth.get_session()
        self.assertNotIn("bad token", str(raised.exception))


def http_error(code, body=b'{"message":"nope"}'):
    return urllib.error.HTTPError(BASE + "/v2/x", code, "err", {}, io.BytesIO(body))


class FakeResponse:
    def __init__(self, payload):
        self.raw = json.dumps(payload).encode() if payload is not None else b""

    def read(self):
        return self.raw

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class FakeHttp:
    """Scripted replacement for sigma_auth._urlopen; records the bearer used per call."""

    def __init__(self, *outcomes):
        self.outcomes = list(outcomes)
        self.requests = []

    def __call__(self, request):
        self.requests.append((request.get_method(), request.full_url,
                              request.get_header("Authorization")))
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return FakeResponse(outcome)


class ApiTests(AuthTestCase):
    def patch_http(self, *outcomes):
        http = FakeHttp(*outcomes)
        patcher = mock.patch.object(sigma_auth, "_urlopen", side_effect=http)
        patcher.start()
        self.addCleanup(patcher.stop)
        return http

    def call(self, *args, **kwargs):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                result = sigma_auth.api(*args, **kwargs)
                error = None
            except Exception as exc:  # noqa: BLE001 - the test inspects it
                result, error = None, exc
        return result, error, out.getvalue(), err.getvalue()

    def assertNoTokens(self, *texts):
        for text in texts:
            for token in ALL_TOKENS:
                self.assertNotIn(token, text)

    def test_success_returns_json_and_sends_the_bearer(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        http = self.patch_http({"workbookId": "wb-1"})
        result, error, out, err = self.call("POST", "/v2/workbooks", {"name": "x"})
        self.assertIsNone(error)
        self.assertEqual(result, {"workbookId": "wb-1"})
        self.assertEqual(http.requests, [("POST", BASE + "/v2/workbooks", f"Bearer {ENV_TOKEN}")])
        self.assertNoTokens(out, err)

    def test_empty_body_is_none(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        self.patch_http(None)
        self.assertIsNone(self.call("GET", "/v2/x")[0])

    def test_401_refreshes_once_through_the_cli_and_retries(self):
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        cli = FakeCli(tokens=(CLI_TOKEN, CLI_TOKEN_2))
        self.use_cli(cli)
        http = self.patch_http(http_error(401), {"ok": True})
        result, error, out, err = self.call("GET", "/v2/whoami")
        self.assertIsNone(error)
        self.assertEqual(result, {"ok": True})
        self.assertEqual([r[2] for r in http.requests],
                         [f"Bearer {CLI_TOKEN}", f"Bearer {CLI_TOKEN_2}"])
        self.assertEqual(cli.token_calls, 2)
        self.assertEqual(out, "")  # the first 401 is not an error to the caller
        self.assertNoTokens(out, err)

    def test_401_refreshes_once_through_the_provider_and_retries(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        self.mint_mock.side_effect = None
        self.mint_mock.return_value = provider_result()
        http = self.patch_http(http_error(401), {"ok": 1})
        result, error, _, _ = self.call("GET", "/v2/whoami")
        self.assertIsNone(error)
        self.assertEqual([r[2] for r in http.requests],
                         [f"Bearer {ENV_TOKEN}", f"Bearer {PROVIDER_TOKEN}"])
        self.mint_mock.assert_called_once_with()

    def test_second_401_is_an_error_after_exactly_one_refresh(self):
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        cli = FakeCli(tokens=(CLI_TOKEN, CLI_TOKEN_2))
        self.use_cli(cli)
        http = self.patch_http(http_error(401), http_error(401, b'{"message":"still no"}'))
        result, error, out, err = self.call("GET", "/v2/whoami")
        self.assertIsInstance(error, urllib.error.HTTPError)
        self.assertEqual(error.code, 401)
        self.assertEqual(len(http.requests), 2)
        self.assertEqual(cli.token_calls, 2)
        self.assertIn("HTTP 401", out)
        self.assertIn("still no", out)
        self.assertNoTokens(out, err)

    def test_refresh_failure_on_401_raises_a_clear_error(self):
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        cli = FakeCli(tokens=(CLI_TOKEN,))
        self.use_cli(cli)
        # first fetch fine, the refresh fetch fails (profile signed out meanwhile)
        original = cli.run

        def flaky(argv, **kwargs):
            if argv[1:3] == ["auth", "token"] and cli.token_calls >= 1:
                cli.token_rc = 1
            return original(argv, **kwargs)
        patcher = mock.patch.object(sigma_auth.subprocess, "run", side_effect=flaky)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.patch_http(http_error(401))
        result, error, out, err = self.call("GET", "/v2/whoami")
        self.assertIsInstance(error, sigma_auth.SigmaAuthError)
        self.assertIn("could not be refreshed", str(error))
        self.assertIn("sigma auth login -p work", str(error))
        self.assertNoTokens(out, err, str(error))

    def test_non_401_errors_print_status_and_first_1500_chars_then_raise(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        self.mint_mock.side_effect = AssertionError("must not refresh on a 500")
        http = self.patch_http(http_error(500, b"x" * 3000))
        result, error, out, _ = self.call("POST", "/v2/workbooks", {"a": 1})
        self.assertIsInstance(error, urllib.error.HTTPError)
        self.assertEqual(len(http.requests), 1)
        self.assertTrue(out.startswith("HTTP 500 "))
        self.assertEqual(out.strip().split(" ", 2)[2], "x" * 1500)

    def test_proactive_refresh_after_fifty_minutes(self):
        old = time.time() - 51 * 60
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE,
                           "SIGMA_TOKEN_MINTED_AT": iso(old)})
        self.mint_mock.side_effect = None
        self.mint_mock.return_value = provider_result()
        http = self.patch_http({"ok": 1})
        _, error, _, _ = self.call("GET", "/v2/whoami")
        self.assertIsNone(error)
        self.assertEqual(http.requests[0][2], f"Bearer {PROVIDER_TOKEN}")
        self.mint_mock.assert_called_once_with()

    def test_no_refresh_before_fifty_minutes(self):
        young = time.time() - 49 * 60
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE,
                           "SIGMA_TOKEN_MINTED_AT": iso(young)})
        http = self.patch_http({"ok": 1})
        self.call("GET", "/v2/whoami")
        self.assertEqual(http.requests[0][2], f"Bearer {ENV_TOKEN}")
        self.mint_mock.assert_not_called()

    def test_no_proactive_refresh_when_mint_time_is_unknown(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        http = self.patch_http({"ok": 1}, {"ok": 2})
        with mock.patch.object(sigma_auth, "_now", return_value=time.time() + 10 * 3600):
            self.call("GET", "/v2/a")
            self.call("GET", "/v2/b")
        self.mint_mock.assert_not_called()
        self.assertEqual(len(http.requests), 2)

    def test_proactive_refresh_applies_to_a_long_running_session(self):
        now = {"t": 1_800_000_000.0}
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE,
                           "SIGMA_TOKEN_MINTED_AT": iso(now["t"])})
        self.mint_mock.side_effect = None
        self.mint_mock.return_value = provider_result(minted=iso(now["t"] + 51 * 60))
        http = self.patch_http({"ok": 1}, {"ok": 2}, {"ok": 3})
        with mock.patch.object(sigma_auth, "_now", side_effect=lambda: now["t"]):
            self.call("GET", "/v2/a")
            now["t"] += 51 * 60
            self.call("GET", "/v2/b")      # crosses 50 minutes: refresh
            now["t"] += 10 * 60
            self.call("GET", "/v2/c")      # fresh token is 10 minutes old: no refresh
        self.assertEqual([r[2] for r in http.requests],
                         [f"Bearer {ENV_TOKEN}", f"Bearer {PROVIDER_TOKEN}", f"Bearer {PROVIDER_TOKEN}"])
        self.assertEqual(self.mint_mock.call_count, 1)

    def test_failed_proactive_refresh_warns_and_keeps_going(self):
        old = time.time() - 55 * 60
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE,
                           "SIGMA_TOKEN_MINTED_AT": iso(old)})
        self.mint_mock.side_effect = get_token.TokenProviderError("no auth available")
        http = self.patch_http({"ok": 1})
        result, error, out, err = self.call("GET", "/v2/whoami")
        self.assertIsNone(error)
        self.assertEqual(http.requests[0][2], f"Bearer {ENV_TOKEN}")
        self.assertIn("proactive Sigma token refresh failed", err)
        self.assertNoTokens(out, err)

    def test_cli_tokens_are_not_proactively_refreshed(self):
        os.environ["SIGMA_CLI_PROFILE"] = "work"
        cli = FakeCli()
        self.use_cli(cli)
        self.patch_http({"ok": 1})
        with mock.patch.object(sigma_auth, "_now", return_value=time.time() + 5 * 3600):
            self.call("GET", "/v2/a")
        self.assertEqual(cli.token_calls, 1)

    def test_base_url_argument_must_match_the_authenticated_host(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        self.patch_http({"ok": 1})
        _, error, _, _ = self.call("GET", "/v2/x", base_url=OTHER_BASE)
        self.assertIsInstance(error, sigma_auth.SigmaAuthError)
        self.assertIn("does not match", str(error))
        _, error, _, _ = self.call("GET", "/v2/x", base_url=BASE)
        self.assertIsNone(error)

    def test_path_must_be_absolute(self):
        os.environ.update({"SIGMA_API_TOKEN": ENV_TOKEN, "SIGMA_BASE_URL": BASE})
        _, error, _, _ = self.call("GET", "v2/x")
        self.assertIsInstance(error, sigma_auth.SigmaAuthError)

    def test_redirects_are_never_followed_with_the_bearer(self):
        handler = sigma_auth._NoRedirect()
        self.assertIsNone(handler.redirect_request(mock.Mock(), None, 302, "Found", {},
                                                   "https://evil.example.com/steal"))


class FlagTests(AuthTestCase):
    def test_add_and_apply_cli_args(self):
        import argparse
        parser = argparse.ArgumentParser()
        sigma_auth.add_cli_args(parser)
        sigma_auth.apply_cli_args(parser.parse_args([]))
        self.assertNotIn("SIGMA_CLI_PROFILE", os.environ)
        sigma_auth.apply_cli_args(parser.parse_args(["--sigma-profile", "work"]))
        self.assertEqual(os.environ["SIGMA_CLI_PROFILE"], "work")


class BuilderWiringTests(Fixtures):
    """The builders authenticate only through sigma_auth and never print a token."""

    BUILDERS = ("build-forecast-model.py", "build-research-model.py", "build-table-report.py",
                "build-input-table-wb.py", "build-dm-on-view.py", "build-research-workbook.py",
                "batch-convert.py")

    def setUp(self):
        super().setUp()
        env = mock.patch.dict(os.environ, {}, clear=False)
        env.start()
        self.addCleanup(env.stop)
        for key in SECRET_ENV_KEYS:
            os.environ.pop(key, None)
        sigma_auth.reset()
        self.addCleanup(sigma_auth.reset)

    def test_every_post_capable_builder_offers_sigma_profile_and_no_inline_client_credentials(self):
        for name in self.BUILDERS:
            text = (SCRIPTS / name).read_text(encoding="utf-8")
            self.assertIn("sigma_auth", text, name)
            self.assertNotIn("/v2/auth/token", text, name)
            self.assertNotIn("client_credentials", text, name)
            self.assertNotIn("SIGMA_CLIENT_SECRET\"]", text, name)

    def test_forecast_env_token_api_are_thin_wrappers_over_sigma_auth(self):
        with mock.patch.object(sigma_auth, "get_session", return_value=(BASE, ENV_TOKEN, "env")) as session, \
                mock.patch.object(sigma_auth, "api", return_value={"ok": 1}) as api:
            e = FORECAST.env()
            self.assertEqual(FORECAST.token(e), ENV_TOKEN)
            self.assertEqual(e["SIGMA_BASE_URL"], BASE)
            self.assertEqual(FORECAST.api(BASE, ENV_TOKEN, "POST", "/v2/x", {"a": 1}), {"ok": 1})
        session.assert_called_once_with()
        api.assert_called_once_with("POST", "/v2/x", {"a": 1}, base_url=BASE)

    def test_sigma_profile_flag_reaches_the_resolver_on_post(self):
        source = self.source()
        plan_path, plan = self.plan(source, issues=FLAT)
        review = self.dir / "review.json"
        review.write_text(json.dumps({"sourceSha256": plan["sourceSha256"], "issues": [
            {"code": "FLAT_CARRY_POLICY", "owner": "Finance", "disposition": "confirmed policy",
             "verified": True}]}), encoding="utf-8")
        with mock.patch.object(FORECAST, "env", side_effect=RuntimeError("auth reached")):
            with self.assertRaisesRegex(RuntimeError, "auth reached"):
                run(FORECAST.main, [str(plan_path), "--post", "--sigma-profile", "work",
                                    "--review-dispositions", str(review)])
        self.assertEqual(os.environ["SIGMA_CLI_PROFILE"], "work")

    def test_dry_run_never_touches_auth_even_with_a_profile_flag(self):
        source = self.source()
        plan_path, _ = self.plan(source, issues=FLAT)
        with mock.patch.object(FORECAST, "env", side_effect=AssertionError("auth reached")):
            code, _, _ = run(FORECAST.main, [str(plan_path), "--sigma-profile", "work",
                                             "--out-dir", str(self.dir / "out")])
        self.assertEqual(code, 0)
        self.assertNotIn("SIGMA_CLI_PROFILE", os.environ)

    def test_forecast_post_end_to_end_prints_no_token(self):
        source = self.source()
        plan_path, plan = self.plan(source, issues=FLAT)
        review = self.dir / "review.json"
        review.write_text(json.dumps({"sourceSha256": plan["sourceSha256"], "issues": [
            {"code": "FLAT_CARRY_POLICY", "owner": "Finance", "disposition": "confirmed policy",
             "verified": True}]}), encoding="utf-8")
        cli = FakeCli(tokens=(CLI_TOKEN, CLI_TOKEN_2))
        posted = {}

        def fake_urlopen(request):
            method, url = request.get_method(), request.full_url
            path = url[len(BASE):]
            if request.get_header("Authorization") == f"Bearer {CLI_TOKEN}" and not posted.get("401"):
                posted["401"] = True
                raise http_error(401)
            if (method, path) == ("POST", "/v2/dataModels/spec"):
                posted["dm"] = json.loads(request.data)
                return FakeResponse({"dataModelId": "dm-1"})
            if (method, path) == ("GET", "/v2/dataModels/dm-1/spec"):
                return FakeResponse(posted["dm"])
            if method == "POST" and path == "/v2/workbooks":
                body = json.loads(request.data)
                return FakeResponse({"dryRun": True, "warnings": []} if body.get("dryRun")
                                    else {"workbookId": "wb-1", "url": BASE + "/wb-1"})
            raise AssertionError((method, path))

        which = mock.patch.object(sigma_auth.shutil, "which", return_value="/usr/bin/sigma")
        runner = mock.patch.object(sigma_auth.subprocess, "run", side_effect=cli.run)
        http = mock.patch.object(sigma_auth, "_urlopen", side_effect=fake_urlopen)
        with which, runner, http:
            code, out, err = run(FORECAST.main, [str(plan_path), "--post", "--sigma-profile", "work",
                                                 "--review-dispositions", str(review),
                                                 "--out-dir", str(self.dir / "out")])
        self.assertEqual(code, 0, err)
        self.assertIn("workbookId: wb-1", out)
        self.assertEqual(cli.token_calls, 2)  # one 401 -> exactly one refresh
        for token in ALL_TOKENS:
            self.assertNotIn(token, out + err)


class VendoredRuntimeTests(unittest.TestCase):
    def test_vendored_files_match_the_recorded_sha256(self):
        text = (SCRIPTS / "AUTH_RUNTIME.md").read_text(encoding="utf-8")
        rows = re.findall(r"^\| ((?:scripts|refs)/\S+) \| ([0-9a-f]{64}) \|$", text, re.MULTILINE)
        self.assertEqual(sorted(path for path, _ in rows), sorted([
            "scripts/get_token.py", "scripts/get-token.sh", "scripts/browser-login.sh",
            "scripts/refresh-token.sh", "scripts/lib/browser-login-platform.sh",
            "refs/browser-oauth-login.md"]))
        for path, expected in rows:
            actual = hashlib.sha256((ROOT / path).read_bytes()).hexdigest()
            self.assertEqual(actual, expected,
                             f"{path} drifted from the vendored runtime; re-vendor deliberately "
                             "and update scripts/AUTH_RUNTIME.md")

    def test_the_record_names_the_source_commit(self):
        text = (SCRIPTS / "AUTH_RUNTIME.md").read_text(encoding="utf-8")
        self.assertIn("twells89/sigma-skills", text)
        self.assertIn("95e67b7", text)

    def test_vendored_shell_scripts_stay_executable(self):
        for name in ("get-token.sh", "browser-login.sh", "refresh-token.sh", "get_token.py"):
            self.assertTrue(os.access(SCRIPTS / name, os.X_OK), name)


if __name__ == "__main__":
    unittest.main()
