"""One Sigma auth + REST resolver for every excel-to-sigma script that calls Sigma.

Resolution order (first hit wins):

  1. ``SIGMA_API_TOKEN`` in the process environment (+ ``SIGMA_BASE_URL``), as-is.
  2. ``<SIGMA_WORKDIR>/auth.json``, then ``./auth.json`` (written by
     ``python scripts/get_token.py --workdir <dir>``).
  3. A Sigma CLI profile, when ``--sigma-profile NAME`` / ``SIGMA_CLI_PROFILE`` is set:
     token from ``sigma auth token -p NAME``, base URL from ``sigma auth status -p NAME``.
  4. The vendored provider ``get_token.mint_token()`` (auto mode: browser keychain
     session/refresh, then client credentials from env or ``~/.sigma-migration/env``).

``api()`` refreshes proactively once a token with a known mint time is 50 minutes old,
and refreshes exactly once on an HTTP 401 through the same source before giving up.
Tokens are never printed, logged or put in exception text.
"""
import calendar
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))
import get_token as _gt  # noqa: E402  (vendored provider; stdlib only)

REFRESH_AFTER_SECONDS = 50 * 60
CLI_BINARY = "sigma"
CLI_TIMEOUT_SECONDS = 60
HTTP_TIMEOUT_SECONDS = 300
_PROFILE_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
_BASE_URL_LINE_RE = re.compile(r"^\s*Base URL:\s*(\S+)\s*$", re.MULTILINE)
# OAuth profiles report the bare API host instead of a Base URL line (sigma 1.2.3).
_API_HOST_LINE_RE = re.compile(r"^\s*API host:\s*([A-Za-z0-9.-]+)\s*$", re.MULTILINE)
# The CLI lets these override a profile's own sign-in; strip them so the profile's
# OAuth session is what `auth token` / `auth status` use.
_CLI_ENV_OVERRIDES = ("SIGMA_CLIENT_ID", "SIGMA_CLIENT_SECRET", "SIGMA_API_TOKEN", "SIGMA_BASE_URL")


class SigmaAuthError(RuntimeError):
    """Sigma authentication could not be resolved or refreshed."""


class Session:
    __slots__ = ("base_url", "token", "method", "minted_at", "source", "profile")

    def __init__(self, base_url, token, method, minted_at, source, profile=None):
        self.base_url = base_url
        self.token = token
        self.method = method
        self.minted_at = minted_at  # epoch seconds, or None when unknown
        self.source = source        # env | auth.json | cli | provider
        self.profile = profile

    def __repr__(self):  # never expose the token
        return f"Session(base_url={self.base_url!r}, method={self.method!r}, source={self.source!r})"


_SESSION = None


def reset():
    """Drop the cached session (tests; or to force a fresh resolution)."""
    global _SESSION
    _SESSION = None


# ------------------------------------------------------------------ CLI flag glue
def add_cli_args(parser):
    parser.add_argument("--sigma-profile", metavar="NAME",
                        help="use this Sigma CLI profile (`sigma auth login -p NAME`) for the "
                             "Sigma API token instead of browser OAuth / client credentials "
                             "(same as SIGMA_CLI_PROFILE)")


def apply_cli_args(args):
    profile = getattr(args, "sigma_profile", None)
    if profile:
        use_cli_profile(profile)


def use_cli_profile(name):
    _validate_profile(name)
    os.environ["SIGMA_CLI_PROFILE"] = name
    reset()


# ------------------------------------------------------------------ helpers
def _validate_profile(name):
    if not isinstance(name, str) or not _PROFILE_RE.fullmatch(name):
        raise SigmaAuthError(
            "invalid Sigma CLI profile name (letters, digits, '.', '_' and '-' only)")
    return name


def _norm_base(url, label):
    try:
        return _gt._assert_sigma_url(url, label, base_url=True)
    except _gt.TokenProviderError as exc:
        raise SigmaAuthError(str(exc)) from exc


def _check_token(token, label):
    try:
        return _gt._validate_access_token(token)
    except _gt.TokenProviderError as exc:
        raise SigmaAuthError(f"{label}: {exc}") from exc


def _parse_minted_at(value):
    if isinstance(value, str) and _gt._MINTED_AT_RE.fullmatch(value):
        return calendar.timegm(time.strptime(value, "%Y-%m-%dT%H:%M:%SZ"))
    return None


def _now():
    return time.time()


def _env_base_url(load_neutral):
    if load_neutral and not os.environ.get("SIGMA_BASE_URL"):
        try:
            _gt._load_neutral_env()
        except _gt.TokenProviderError as exc:
            raise SigmaAuthError(str(exc)) from exc
    value = os.environ.get("SIGMA_BASE_URL")
    return _norm_base(value, "SIGMA_BASE_URL") if value else None


def _cli_profile():
    name = os.environ.get("SIGMA_CLI_PROFILE")
    return _validate_profile(name) if name else None


def _run_cli(args):
    binary = shutil.which(CLI_BINARY)
    if not binary:
        raise SigmaAuthError(
            f"the Sigma CLI ('{CLI_BINARY}') is not on PATH; install it, or unset "
            "SIGMA_CLI_PROFILE and use `eval \"$(bash scripts/browser-login.sh)\"` or "
            "SIGMA_CLIENT_ID/SIGMA_CLIENT_SECRET instead")
    try:
        env = {k: v for k, v in os.environ.items() if k not in _CLI_ENV_OVERRIDES}
        return subprocess.run([binary, *args], capture_output=True, text=True,
                              timeout=CLI_TIMEOUT_SECONDS, check=False, env=env)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise SigmaAuthError(f"could not run `{CLI_BINARY} {' '.join(args[:2])}`: "
                             f"{type(exc).__name__}") from exc


def _cli_token(profile):
    done = _run_cli(["auth", "token", "-p", profile])
    token = (done.stdout or "").strip()
    if done.returncode != 0 or not token:
        raise SigmaAuthError(
            f"Sigma CLI profile '{profile}' is not signed in (or has no usable token); "
            f"run `sigma auth login -p {profile}` and retry")
    return _check_token(token, f"Sigma CLI profile '{profile}' token")


def _cli_base_url(profile):
    done = _run_cli(["auth", "status", "-p", profile])
    if done.returncode != 0:
        return None
    # sigma 1.2.3 writes the status report to stderr, not stdout; read both.
    text = (done.stdout or "") + "\n" + (done.stderr or "")
    match = _BASE_URL_LINE_RE.search(text)
    if match:
        return _norm_base(match.group(1), f"Sigma CLI profile '{profile}' base URL")
    host = _API_HOST_LINE_RE.search(text)
    if host:
        return _norm_base("https://" + host.group(1), f"Sigma CLI profile '{profile}' API host")
    return None


# ------------------------------------------------------------------ sources
def _from_env():
    token = os.environ.get("SIGMA_API_TOKEN")
    if not token:
        return None
    base = _env_base_url(load_neutral=True)
    if not base:
        raise SigmaAuthError("SIGMA_API_TOKEN is set but SIGMA_BASE_URL is not; set it to "
                             "your Sigma cloud API host")
    if os.environ.get("SIGMA_CLI_PROFILE"):
        print("note: SIGMA_API_TOKEN is set in the environment and takes precedence over "
              "the Sigma CLI profile; unset it to use the profile.", file=sys.stderr)
    return Session(base, _check_token(token, "SIGMA_API_TOKEN"),
                   os.environ.get("SIGMA_AUTH_METHOD") or "env",
                   _parse_minted_at(os.environ.get("SIGMA_TOKEN_MINTED_AT")), "env")


def _auth_json_paths():
    paths = []
    workdir = os.environ.get("SIGMA_WORKDIR")
    if workdir:
        paths.append(Path(workdir) / "auth.json")
    paths.append(Path.cwd() / "auth.json")
    return paths


def _from_auth_json():
    for path in _auth_json_paths():
        if not path.is_file():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise SigmaAuthError(f"{path} is not valid JSON ({type(exc).__name__}); "
                                 "re-run `python scripts/get_token.py --workdir <dir>`") from exc
        token = data.get("SIGMA_API_TOKEN") if isinstance(data, dict) else None
        if not token:
            raise SigmaAuthError(f"{path} has no SIGMA_API_TOKEN; re-run "
                                 "`python scripts/get_token.py --workdir <dir>`")
        file_base = data.get("SIGMA_BASE_URL")
        env_base = _env_base_url(load_neutral=False)
        file_base = _norm_base(file_base, f"{path.name} SIGMA_BASE_URL") if file_base else None
        if file_base and env_base and file_base != env_base:
            raise SigmaAuthError(f"SIGMA_BASE_URL ({env_base}) differs from the base URL in "
                                 f"{path} ({file_base}); refusing to guess")
        base = file_base or env_base
        if not base:
            raise SigmaAuthError(f"{path} has no SIGMA_BASE_URL and none is set in the "
                                 "environment")
        return Session(base, _check_token(token, f"{path.name} SIGMA_API_TOKEN"),
                       data.get("SIGMA_AUTH_METHOD") or "auth.json",
                       _parse_minted_at(data.get("SIGMA_TOKEN_MINTED_AT")), "auth.json")
    return None


def _from_cli():
    profile = _cli_profile()
    if not profile:
        return None
    token = _cli_token(profile)
    env_base = _env_base_url(load_neutral=False)
    cli_base = _cli_base_url(profile)
    if env_base and cli_base and env_base != cli_base:
        raise SigmaAuthError(
            f"SIGMA_BASE_URL ({env_base}) differs from Sigma CLI profile '{profile}' "
            f"({cli_base}); unset SIGMA_BASE_URL or pick the matching profile")
    base = cli_base or env_base
    if not base:
        raise SigmaAuthError(
            f"could not read the base URL for Sigma CLI profile '{profile}' from "
            f"`sigma auth status -p {profile}`; set SIGMA_BASE_URL")
    return Session(base, token, "sigma-cli", None, "cli", profile)


def _provider_session(expected_base=None):
    if expected_base and not os.environ.get("SIGMA_BASE_URL"):
        os.environ["SIGMA_BASE_URL"] = expected_base
    try:
        result = _gt.mint_token()
    except _gt.TokenProviderError as exc:
        raise SigmaAuthError(
            f"{exc}. Sign in with `eval \"$(bash scripts/browser-login.sh)\"`, run "
            "`sigma auth login` and pass --sigma-profile NAME, or set SIGMA_CLIENT_ID and "
            "SIGMA_CLIENT_SECRET") from exc
    base = _norm_base(result.base_url, "SIGMA_BASE_URL")
    if expected_base and base != expected_base:
        raise SigmaAuthError(f"refreshed token is for {base}, not {expected_base}")
    return Session(base, _check_token(result.token, "minted token"), result.auth_method,
                   _parse_minted_at(result.minted_at), "provider")


def _resolve():
    return _from_env() or _from_auth_json() or _from_cli() or _provider_session()


def _refresh(session):
    """Replace the session's token through the source it came from."""
    if session.source == "cli":
        session.token = _cli_token(session.profile)
        return session
    fresh = _provider_session(expected_base=session.base_url)
    session.token, session.method = fresh.token, fresh.method
    session.minted_at = fresh.minted_at
    return session


def get_session():
    """Return (base_url, token, method), resolving auth on first use."""
    global _SESSION
    if _SESSION is None:
        _SESSION = _resolve()
    _maybe_refresh(_SESSION)
    return _SESSION.base_url, _SESSION.token, _SESSION.method


def _maybe_refresh(session):
    if session.minted_at is None or _now() - session.minted_at < REFRESH_AFTER_SECONDS:
        return
    try:
        _refresh(session)
    except SigmaAuthError as exc:
        # Still within the token's lifetime: keep going; a real 401 re-raises below.
        session.minted_at = None
        print(f"warning: proactive Sigma token refresh failed ({exc}); continuing with "
              "the current token", file=sys.stderr)


# ------------------------------------------------------------------ REST
class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """Never forward the bearer token to a redirect target."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_OPENER = urllib.request.build_opener(_NoRedirect)


def _urlopen(request):
    return _OPENER.open(request, timeout=HTTP_TIMEOUT_SECONDS)


def _send(session, method, path, body):
    request = urllib.request.Request(
        session.base_url + path,
        data=(json.dumps(body).encode() if body else None), method=method,
        headers={"Authorization": f"Bearer {session.token}",
                 "Content-Type": "application/json", "Accept": "application/json"})
    with _urlopen(request) as response:
        raw = response.read()
    return json.loads(raw) if raw and raw.strip() else None


def _error_body(exc):
    try:
        return exc.read().decode(errors="replace")[:1500]
    finally:
        exc.close()


def api(method, path, body=None, base_url=None):
    """Call the Sigma REST API; one refresh+retry on 401, then print+raise like before."""
    if not path.startswith("/"):
        raise SigmaAuthError(f"API path must start with '/': {path!r}")
    get_session()
    session = _SESSION
    if base_url and _norm_base(base_url, "base URL") != session.base_url:
        raise SigmaAuthError(f"base URL {base_url} does not match the authenticated "
                             f"Sigma host {session.base_url}")
    retried = False
    while True:
        try:
            return _send(session, method, path, body)
        except urllib.error.HTTPError as exc:
            if exc.code == 401 and not retried:
                retried = True
                try:
                    _refresh(session)
                except SigmaAuthError as refresh_error:
                    print("HTTP", exc.code, _error_body(exc))
                    raise SigmaAuthError(
                        f"Sigma returned 401 and the token could not be refreshed: "
                        f"{refresh_error}") from exc
                exc.close()
                continue
            print("HTTP", exc.code, _error_body(exc))
            raise
