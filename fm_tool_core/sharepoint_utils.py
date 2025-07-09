from __future__ import annotations

from pathlib import Path  # Work with filesystem paths

from .constants import ROOT_SP_SITE, SP_PASSWORD, SP_USERNAME
from .exceptions import FlowError


def sp_ctx(site_url: str | None = None):
    # Build a SharePoint client context
    if not (SP_USERNAME and SP_PASSWORD):
        raise FlowError("SharePoint credentials missing", work_completed=False)

    base = site_url.rstrip("/") if site_url else ROOT_SP_SITE
    try:
        from office365.runtime.auth.user_credential import UserCredential
        from office365.sharepoint.client_context import ClientContext
    except Exception as exc:  # pragma: no cover - import failure
        raise FlowError(f"SharePoint SDK missing: {exc}", work_completed=False)

    return ClientContext(base).with_credentials(
        UserCredential(SP_USERNAME, SP_PASSWORD)
    )


def sp_exists(ctx, rel_url: str) -> bool:
    # Check if a file already exists on SharePoint
    try:
        ctx.web.get_file_by_server_relative_url(rel_url).get().execute_query()
        return True
    except Exception:
        return False


def sp_upload(ctx, folder: str, fname: str, local: Path):
    # Upload a local file to a SharePoint folder
    tgt = ctx.web.get_folder_by_server_relative_url(folder)
    with local.open("rb") as f:
        content = f.read()
    tgt.upload_file(fname, content).execute_query()


# Backwards compatible aliases used by tests
sharepoint_upload = sp_upload
sharepoint_file_exists = sp_exists

# Exported symbols when importing * from this module
__all__ = [
    "sp_ctx",
    "sp_exists",
    "sp_upload",
    "sharepoint_upload",
    "sharepoint_file_exists",
]
