import pytest
from unittest.mock import MagicMock

from fm_tool_core.exceptions import FlowError
from fm_tool_core import sharepoint_utils


def _ctx_with_exc(exc):
    ctx = MagicMock()
    tgt = MagicMock()
    ctx.web.get_folder_by_server_relative_url.return_value = tgt
    upload = tgt.upload_file.return_value
    upload.execute_query.side_effect = exc
    return ctx


def test_sp_upload_client_request_exception(tmp_path):
    local = tmp_path / "f.txt"
    local.write_text("data")
    ctx = _ctx_with_exc(sharepoint_utils.ClientRequestException("msg", 500, "err"))
    with pytest.raises(FlowError) as exc:
        sharepoint_utils.sp_upload(ctx, "/folder", "f.txt", local)
    assert not exc.value.work_completed
    assert "SharePoint upload failed" in str(exc.value)
