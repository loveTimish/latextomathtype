#!/usr/bin/env python3
"""Bounded warmed DOCX comparison of already-built classes, with no downloads."""
import argparse
import hashlib
import json
import os
import pathlib
import shutil
import signal
import statistics
import subprocess
import time

ROOT = pathlib.Path(__file__).resolve().parents[2]

def hashes(path):
    return {str(p.relative_to(path)): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted(path.rglob('*')) if p.is_file()}

def summarize(results, rounds):
    summary = []
    for case in ('repeat40', 'unique40'):
        for window, minimum in (('all_measured', 0), ('last_half', rounds // 2)):
            row = {'case': case, 'window': window}
            for metric in ('ms', 'thread_cpu_ms', 'process_cpu_ms', 'allocated_bytes', 'compilation_ms', 'gc_ms'):
                values = {}
                for variant in ('before', 'after'):
                    forks = [[s[metric] for s in r['samples'] if s['case'] == case and s['round'] >= minimum]
                             for r in results if r['variant'] == variant]
                    medians = [statistics.median(f) for f in forks]
                    values[variant] = statistics.median(medians)
                    row[f'{variant}_{metric}_fork_medians'] = medians
                    row[f'{variant}_{metric}_median'] = values[variant]
                row[f'{metric}_change_percent'] = (values['after'] / values['before'] - 1) * 100 if values['before'] else None
            summary.append(row)
    return summary

def run(args, variant, fork, before, after, probe):
    cp = (ROOT / 'target/linux-review/classpath.txt').read_text().strip().split(':')
    cp = [str(probe)] + [str(before if variant == 'before' else after) if p == str(ROOT / 'target/classes') else p
                       for p in cp if p and p != str(ROOT / 'target/test-classes')]
    destination = args.output / f'{variant}-{fork}'
    destination.mkdir()
    command = ['java', '-Xms256m', '-Xmx1024m',
               f'-Dpaperword.mathjax.node.command={ROOT}/node_modules/node/bin/node',
               '-Dpaperword.render.cache.enabled=false', '-cp', ':'.join(cp),
               'com.lz.paperword.core.render.WarmDocxProbe', str(destination),
               str(args.warmups), str(args.rounds), str(args.profile).lower()]
    started = time.monotonic()
    with (args.output / f'{variant}-{fork}.log').open('w') as log:
        process = subprocess.Popen(command, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT, start_new_session=True)
        try:
            process.wait(timeout=180)
            if process.returncode:
                raise RuntimeError(f'{variant}-{fork} failed; see log')
        finally:
            try: os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError: pass
            process.wait()
    result = {'variant': variant, 'fork': fork, 'wall_s': time.monotonic() - started,
              'samples': [], 'warmup': [], 'command': command}
    for line in (args.output / f'{variant}-{fork}.log').read_text().splitlines():
        for prefix, key in (('SAMPLE ', 'samples'), ('WARMUP ', 'warmup')):
            if line.startswith(prefix): result[key].append(json.loads(line[len(prefix):]))
        for prefix, key in (('STATE_BEFORE ', 'state_before'), ('STATE_AFTER ', 'state_after'), ('VALIDATED ', 'validation')):
            if line.startswith(prefix): result[key] = json.loads(line[len(prefix):])
    if len(result['samples']) != args.rounds * 2 or 'validation' not in result:
        raise RuntimeError(f'{variant}-{fork} incomplete')
    print(json.dumps({k: result[k] for k in ('variant', 'fork', 'wall_s', 'validation')}), flush=True)
    return result

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--before-classes', type=pathlib.Path, required=True)
    parser.add_argument('--output', type=pathlib.Path, required=True)
    parser.add_argument('--forks', type=int, default=3)
    parser.add_argument('--warmups', type=int, default=10)
    parser.add_argument('--rounds', type=int, default=20)
    parser.add_argument('--profile', action='store_true')
    parser.add_argument('--prepare-only', action='store_true')
    args = parser.parse_args()
    args.output = args.output.resolve()
    args.output.mkdir(parents=True, exist_ok=True)
    before = args.before_classes.resolve()
    after = args.output / 'after-classes'
    probe = args.output / 'probe-classes'
    if not after.exists(): shutil.copytree(ROOT / 'target/classes', after)
    probe.mkdir(exist_ok=True)
    cp = (ROOT / 'target/linux-review/classpath.txt').read_text().strip()
    subprocess.run(['java', 'com.sun.tools.javac.Main', '-proc:none', '-cp', cp, '-d', str(probe),
                    'scripts/linux-review/LinuxBenchmark.java', 'scripts/linux-review/WarmDocxProbe.java'], cwd=ROOT, check=True)
    manifest = {'synthetic_input': True, 'warmups_per_case': args.warmups, 'rounds_per_case': args.rounds,
                'forks_per_variant': args.forks, 'profile': args.profile, 'before_classes': str(before),
                'before_sha256': hashes(before), 'after_sha256': hashes(after)}
    (args.output / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    if args.prepare_only: raise SystemExit(0)
    results = []
    reference = None
    for fork in range(args.forks):
        for variant in (('before', 'after') if fork % 2 == 0 else ('after', 'before')):
            results.append(run(args, variant, fork, before, after, probe))
            validation = json.loads((args.output / f'{variant}-{fork}/validation.json').read_text())
            if reference is None: reference = validation
            if validation != reference: raise RuntimeError('MTEF or preview differs between variants/forks')
            (args.output / 'raw.json').write_text(json.dumps(results, indent=2) + '\n')
    summary = summarize(results, args.rounds)
    (args.output / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n')
    print(json.dumps(summary, indent=2))
