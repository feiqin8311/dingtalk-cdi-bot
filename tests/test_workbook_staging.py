from io import BytesIO
from pathlib import Path
import subprocess
import sys
from time import perf_counter
from unittest.mock import patch

import pytest

from app.core import workbook_staging as staging


def test_staging_subprocess_writes_only_the_temporary_file(tmp_path):
    target = tmp_path / '开票表.xlsx'
    target.write_bytes(b'old workbook')
    staged = staging.stage_customs_workbook(target, BytesIO(b'new workbook'))
    assert staged.parent == target.parent
    assert staged.name.startswith('.ezarc-register-')
    assert staged.read_bytes() == b'new workbook'
    assert target.read_bytes() == b'old workbook'


def test_hung_worker_times_out_and_does_not_publish(tmp_path):
    target = tmp_path / '开票表.xlsx'
    target.write_bytes(b'old workbook')
    spawn = subprocess.Popen
    child = None

    def spawn_hung_worker(_command, **kwargs):
        nonlocal child
        child = spawn([sys.executable, '-I', '-c', 'import time; time.sleep(30)'], **kwargs)
        return child

    try:
        start = perf_counter()
        with patch.object(staging, 'STAGING_TIMEOUT_SECONDS', 0.2), patch.object(
            staging.subprocess, 'Popen', side_effect=spawn_hung_worker,
        ):
            with pytest.raises(staging.WorkbookStagingTimeout, match='刷盘超过 0.2 秒.*原表未更新'):
                staging.stage_customs_workbook(target, BytesIO(b'new workbook'))
        assert perf_counter() - start < 5
        assert child.wait(timeout=5) != 0
    finally:
        if child is not None and child.poll() is None:
            child.kill()
            child.wait(timeout=5)
    assert target.read_bytes() == b'old workbook'


class SimulatedWorker:
    def __init__(self, completion_time):
        self.now = 0
        self.completion_time = completion_time
        self.killed = False

    def wait(self, timeout):
        self.now += timeout
        if self.now < self.completion_time:
            raise subprocess.TimeoutExpired('fake staging worker', timeout)
        return 0

    def kill(self):
        self.killed = True


def test_slow_worker_can_finish_after_the_warning(tmp_path):
    worker = SimulatedWorker(completion_time=120)
    messages = []
    with patch.object(staging.subprocess, 'Popen', return_value=worker) as spawn, \
            patch.object(staging, 'monotonic', side_effect=lambda: worker.now):
        staging.stage_customs_workbook(tmp_path / '开票表.xlsx', BytesIO(b'data'), progress=messages.append)
    assert worker.now == 120
    assert not worker.killed
    spawn.assert_called_once()
    assert messages[0] == '报关开票：共享盘刷盘较慢，已等待 1 分 30 秒'
    assert messages[-1].endswith('1 分 59 秒')


@pytest.mark.parametrize('stop_at, error, expected_seconds', [
    (110, staging.WorkbookStagingCancelled, 110),
    (None, staging.WorkbookStagingTimeout, 600),
])
def test_waiting_worker_can_be_stopped_and_has_a_ten_minute_limit(tmp_path, stop_at, error, expected_seconds):
    worker = SimulatedWorker(completion_time=900)
    path = tmp_path / '开票表.xlsx'
    path.write_bytes(b'old workbook')
    with patch.object(staging.subprocess, 'Popen', return_value=worker), \
            patch.object(staging, 'monotonic', side_effect=lambda: worker.now):
        with pytest.raises(error, match='原表未更新'):
            staging.stage_customs_workbook(path, BytesIO(b'data'),
                                           cancelled=lambda: stop_at is not None and worker.now >= stop_at)
    assert worker.now == expected_seconds
    assert worker.killed
    assert path.read_bytes() == b'old workbook'


def test_app_process_can_run_worker_without_starting_qt(tmp_path):
    source = tmp_path / 'source.xlsx'
    source.write_bytes(b'new workbook')
    stage = tmp_path / '.ezarc-register-ui.tmp'
    entry = Path(__file__).resolve().parents[1] / 'Ui_PyQt.py'
    result = subprocess.run([sys.executable, str(entry), staging.STAGING_FLAG, str(source), str(stage)],
                            check=True, capture_output=True, timeout=10)
    assert result.returncode == 0
    assert stage.read_bytes() == b'new workbook'
