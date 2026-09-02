"""Phase 0 smoke tests.

Real parser/aggregate/combiner tests with fixture PDFs arrive in Phase 1
(see documents/UPGRADE_PLAN.md).
"""
import os
import uuid

import app as app_module
from combinepdf import classify


def test_pages_render():
    client = app_module.app.test_client()
    assert client.get("/").status_code == 200
    assert client.get("/combine").status_code == 200


def test_unknown_job_id_returns_404():
    client = app_module.app.test_client()
    resp = client.get("/count/status/does-not-exist")
    assert resp.status_code == 404


def test_count_rejects_empty_upload():
    client = app_module.app.test_client()
    resp = client.post("/count", data={})
    assert resp.status_code == 400


def test_download_deletes_file_after_complete_download():
    report_id = uuid.uuid4().hex
    report_path = os.path.join(app_module.UPLOAD_DIR, f"{report_id}.xlsx")
    with open(report_path, "wb") as fh:
        fh.write(b"fake-xlsx-bytes")

    client = app_module.app.test_client()
    resp = client.get(f"/download/{report_id}")
    assert resp.status_code == 200
    assert resp.data == b"fake-xlsx-bytes"           # whole body consumed
    assert not os.path.exists(report_path)           # deleted on completion

    # Second attempt: file is gone -> 404
    assert client.get(f"/download/{report_id}").status_code == 404


def test_download_keeps_file_when_missing():
    client = app_module.app.test_client()
    assert client.get("/download/nope-not-real").status_code == 404
    assert client.get("/download_combine/nope-not-real").status_code == 404


def test_combiner_classify_filenames():
    assert classify("INV651195.pdf") == ("651195", "INV")
    assert classify("PL651195.pdf") == ("651195", "PL")
    assert classify("651195.pdf") == ("651195", "FREIGHT")
    assert classify("FRT_651195.pdf") == ("651195", "FREIGHT")
    assert classify("random_notes.pdf") == (None, None)
