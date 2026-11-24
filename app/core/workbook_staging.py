"""Write and flush a staging file in a subprocess that cannot publish a workbook."""

from io import BytesIO
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from time import monotonic
from uuid import uuid4


STAGING_FLAG = '--stage-customs-workbook'
STAGING_WARNING_SECONDS = 90
STAGING_TIMEOUT_SECONDS = 10 * 60
STAGING_CANCELLED_MESSAGE = '已停止本次易达报关开票，原表未更新。'


class WorkbookStagingAborted(Exception):
    """The pending workbook will not be published by this attempt."""


class WorkbookStagingTimeout(WorkbookStagingAborted):
    pass


class WorkbookStagingCancelled(WorkbookStagingAborted):
    pass


def run_staging_worker(args):
    """Worker entry point: write a uniquely named .tmp, never replace the target."""
    if len(args) != 2:
        return 2
    source, staged = map(Path, args)
    if not staged.name.startswith('.ezarc-register-') or staged.suffix != '.tmp':
        return 2
    created = False
    try:
        # Close the local file before starting any potentially blocking network write.
        data = source.read_bytes()
        with staged.open('xb') as dest:
            created = True
            dest.write(data)
            dest.flush()
            os.fsync(dest.fileno())
    except OSError:
        if created:
            try:
                staged.unlink(missing_ok=True)
            except OSError:
                pass
        return 1
    return 0


def _wait_for_staging(worker, staged, progress, cancelled):
    started = monotonic()
    warned_at = None
    residue = f'共享目录可能残留临时文件：{staged.name}'
    while True:
        if cancelled is not None and cancelled():
            raise WorkbookStagingCancelled(f'{STAGING_CANCELLED_MESSAGE}{residue}')
        elapsed = monotonic() - started
        remaining = STAGING_TIMEOUT_SECONDS - elapsed
        if remaining <= 0:
            raise WorkbookStagingTimeout(
                f'易达报关开票临时文件写入或刷盘超过 {STAGING_TIMEOUT_SECONDS:g} 秒，原表未更新。{residue}'
            )
        if progress is not None and elapsed >= STAGING_WARNING_SECONDS and int(elapsed) != warned_at:
            warned_at = int(elapsed)
            minutes, seconds = divmod(warned_at, 60)
            progress(f'报关开票：共享盘刷盘较慢，已等待 {minutes} 分 {seconds:02d} 秒')
        try:
            exit_code = worker.wait(timeout=min(1, remaining))
        except subprocess.TimeoutExpired:
            continue
        if cancelled is not None and cancelled():
            raise WorkbookStagingCancelled(f'{STAGING_CANCELLED_MESSAGE}{residue}')
        return exit_code


def stage_customs_workbook(target: Path, updated: BytesIO, *, progress=None, cancelled=None) -> Path:
    staged = target.with_name(f'.ezarc-register-{uuid4().hex}.tmp')
    with TemporaryDirectory(prefix='ezarc-stage-', ignore_cleanup_errors=True) as local:
        source = Path(local) / 'updated.xlsx'
        with updated.getbuffer() as data:
            source.write_bytes(data)
        if getattr(sys, 'frozen', False):
            command = [sys.executable]
        else:
            command = [sys.executable, '-I', str(Path(__file__).resolve())]
        command += [STAGING_FLAG, str(source), str(staged.absolute())]
        if cancelled is not None and cancelled():
            raise WorkbookStagingCancelled(STAGING_CANCELLED_MESSAGE)
        worker = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                  stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        try:
            exit_code = _wait_for_staging(worker, staged, progress, cancelled)
        except WorkbookStagingAborted:
            try:
                worker.kill()
            except OSError:
                # Even if Windows cannot immediately cancel the I/O, this worker can only write its .tmp.
                pass
            # Do not wait for termination or touch the share here: cleanup can block on the same SMB wait.
            raise
        if exit_code != 0:
            raise OSError(f'易达报关开票临时文件写入或刷盘失败，原表未更新：{target}')
    return staged


if __name__ == '__main__':
    if sys.argv[1:2] != [STAGING_FLAG]:
        sys.exit(2)
    sys.exit(run_staging_worker(sys.argv[2:]))
