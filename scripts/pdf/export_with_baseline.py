#!/usr/bin/env python3
"""Bounded, isolated native/compatible export; publish nothing on conversion failure."""
import argparse
import hashlib
import os
import pathlib
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import time


class ExportTimeout(Exception):
    pass


def digest(path):
    sha = hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            sha.update(chunk)
    return sha.hexdigest()


def terminate(process, managed_group=False):
    # Each child starts its own session. Reap the whole group, even if the
    # soffice launcher exits before soffice.bin.
    try:
        if managed_group:
            process.terminate()
        else:
            os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    try:
        process.wait(timeout=1)
    except subprocess.TimeoutExpired:
        pass
    try:
        if managed_group:
            process.kill()
        else:
            os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    process.wait(timeout=2)


class Processes:
    def __init__(self, deadline, log, managed_group=False):
        self.deadline, self.log = deadline, log
        self.managed_group = managed_group
        self.children, self.readers = [], []
        self.lock = threading.Lock()
        self.logged = 0

    def start(self, command):
        child = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                 start_new_session=not self.managed_group)
        self.children.append(child)

        def read():
            with child.stdout:
                while chunk := child.stdout.read(4096):
                    with self.lock:
                        remaining = 65536 - self.logged
                        if remaining > 0:
                            self.log.write(chunk[:remaining])
                            self.log.flush()
                            self.logged += min(len(chunk), remaining)
        reader = threading.Thread(target=read, daemon=True)
        reader.start()
        self.readers.append(reader)
        return child

    def wait(self, child):
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ExportTimeout('PDF conversion deadline exceeded')
        try:
            code = child.wait(timeout=remaining)
        except subprocess.TimeoutExpired as e:
            raise ExportTimeout('PDF conversion deadline exceeded') from e
        if code:
            if code == 69:
                raise FileNotFoundError('A required converter dependency is unavailable; see conversion log')
            raise RuntimeError(f'Converter process failed with exit {code}')

    def close(self):
        for child in reversed(self.children):
            terminate(child, self.managed_group)
        for reader in self.readers:
            reader.join(timeout=1)


def main():
    p = argparse.ArgumentParser()
    p.add_argument('docx', type=pathlib.Path)
    p.add_argument('output_directory', type=pathlib.Path)
    p.add_argument('--uno-jar', required=True, type=pathlib.Path)
    p.add_argument('--soffice', default='soffice')
    p.add_argument('--java', default='java')
    p.add_argument('--timeout', type=int, default=120)
    p.add_argument('--cjk-font', default='Noto Serif CJK SC',
                   help='Installed fallback for missing CJK fonts in compatible output; empty disables')
    p.add_argument('--managed-process-group', action='store_true',
                   help='Service-only: parent owns and reaps this entire setsid group')
    a = p.parse_args()
    if sys.platform != 'linux':
        raise RuntimeError('This process-isolated helper currently requires Linux')
    if a.managed_process_group and os.getpid() != os.getpgrp():
        raise RuntimeError('Managed conversion must be launched as its own setsid group leader')
    if not 1 <= a.timeout <= 600:
        raise ValueError('--timeout must be between 1 and 600 seconds')
    source, out, jar = a.docx.resolve(), a.output_directory.resolve(), a.uno_jar.resolve()
    java_source = pathlib.Path(__file__).with_name('LibreOfficeBaselineExporter.java')
    if not all(path.is_file() for path in (source, jar, java_source)):
        raise FileNotFoundError('Input DOCX, public UNO jar and helper Java source must exist')
    out.mkdir(parents=True, exist_ok=True)
    stem = source.stem
    names = [stem+'-native.pdf', stem+'-baseline-compatible.pdf', stem+'-baseline-audit.tsv', stem+'-source.sha256']
    if any((out/name).exists() for name in names):
        raise FileExistsError('Refusing to overwrite existing outputs; choose a new directory')
    before = digest(source)
    deadline = time.monotonic() + a.timeout
    published = []
    try:
        with tempfile.TemporaryDirectory(prefix='baseline-export-', dir=out) as temporary:
            work = pathlib.Path(temporary)
            classes, profile = work/'classes', work/'profile'
            classes.mkdir(); profile.mkdir()
            logpath = out/(stem+'-libreoffice.log')
            # Exclusive creation prevents two jobs sharing a directory.
            with logpath.open('xb') as log:
                processes = Processes(deadline, log, a.managed_process_group)
                try:
                    processes.wait(processes.start([a.java, 'com.sun.tools.javac.Main', '-proc:none', '-cp', str(jar), '-d', str(classes), str(java_source)]))
                    with socket.socket() as s:
                        s.bind(('127.0.0.1', 0)); port = s.getsockname()[1]
                    processes.start([a.soffice, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--invisible', '--norestore', '--nodefault', '--nofirststartwizard',
                                     f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager'])
                    processes.wait(processes.start([a.java, '-cp', str(classes)+os.pathsep+str(jar), 'LibreOfficeBaselineExporter', str(port), str(source),
                                                    str(work/names[0]), str(work/names[1]), str(work/names[2]), a.cjk_font]))
                finally:
                    processes.close()
            if digest(source) != before:
                raise RuntimeError('Source DOCX changed unexpectedly')
            for name in names[:2]:
                with (work/name).open('rb') as f:
                    if f.read(5) != b'%PDF-':
                        raise RuntimeError('Converter output is not a PDF')
            (work/names[3]).write_text(before+'  '+source.name+'\n')
            # No-overwrite publication; all assertions run first.
            for name in names:
                os.link(work/name, out/name)
                published.append(out/name)
        print('Unchanged source SHA256:', before)
        print('Native:', out/names[0])
        print('Baseline-compatible:', out/names[1])
        print('Audit:', out/names[2])
        with (out/(stem+'-libreoffice.log')).open('rb') as log:
            for line in log.read(65536).decode('utf-8', errors='replace').splitlines():
                if line.startswith('CJK_FONT_FALLBACK '):
                    print(line)
    except BaseException:
        for path in published:
            path.unlink(missing_ok=True)
        # The service deletes its private job directory after failure. Keep bounded
        # diagnostics in its server log rather than losing the converter's reason.
        logpath = out/(stem+'-libreoffice.log')
        if logpath.is_file():
            with logpath.open('rb') as log:
                sys.stderr.write(log.read(65536).decode('utf-8', errors='replace'))
        raise


if __name__ == '__main__':
    def interrupted(signum, frame):
        raise ExportTimeout('PDF export cancelled')
    signal.signal(signal.SIGTERM, interrupted)
    try:
        main()
    except ExportTimeout as failure:
        print(str(failure), file=sys.stderr); sys.exit(124)
    except FileNotFoundError as failure:
        print(str(failure), file=sys.stderr); sys.exit(69)
    except Exception as failure:
        print(str(failure), file=sys.stderr); sys.exit(1)
