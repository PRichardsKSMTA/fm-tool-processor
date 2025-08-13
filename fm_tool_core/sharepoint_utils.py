from __future__ import annotations

from pathlib import Path

try:  # pragma: no cover - optional dependency
    import requests
except Exception:  # pragma: no cover - requests may be absent
    requests = None  # type: ignore

try:  # pragma: no cover - optional dependency
    from office365.runtime.client_request_exception import ClientRequestException
except Exception:  # pragma: no cover - library may be absent

    class ClientRequestException(Exception):  # type: ignore
        pass


from .constants import ROOT_SP_SITE, SP_PASS, SP_USERNAME
from .exceptions import FlowError


def sp_ctx(site_url: str | None = None):
    if not (SP_USERNAME and SP_PASS):
        raise FlowError("SharePoint credentials missing", work_completed=False)

    base = site_url.rstrip("/") if site_url else ROOT_SP_SITE
    try:
        from office365.runtime.auth.user_credential import UserCredential
        from office365.sharepoint.client_context import ClientContext
    except Exception as exc:  # pragma: no cover - import failure
        raise FlowError(f"SharePoint SDK missing: {exc}", work_completed=False)

    return ClientContext(base).with_credentials(UserCredential(SP_USERNAME, SP_PASS))


def sp_exists(ctx, rel_url: str) -> bool:
    try:
        ctx.web.get_file_by_server_relative_url(rel_url).get().execute_query()
        return True
    except Exception:
        return False


def sp_upload(ctx, folder: str, fname: str, local: Path):
    tgt = ctx.web.get_folder_by_server_relative_url(folder)
    with local.open("rb") as f:
        content = f.read()
    try:
        tgt.upload_file(fname, content).execute_query()
    except ClientRequestException as exc:
        raise FlowError(
            f"SharePoint upload failed: {exc}", work_completed=False
        ) from exc
    except Exception as exc:  # pragma: no cover - optional
        if requests is not None and isinstance(exc, requests.HTTPError):
            raise FlowError(
                f"SharePoint upload failed: {exc}", work_completed=False
            ) from exc
        raise


# Backwards compatible aliases used by tests
sharepoint_upload = sp_upload
sharepoint_file_exists = sp_exists

__all__ = [
    "sp_ctx",
    "sp_exists",
    "sp_upload",
    "sharepoint_upload",
    "sharepoint_file_exists",
]
