#!/usr/bin/env python3
"""Compare already-built 432e4dd classes against current classes; no downloads."""
import argparse
import importlib.util
import json
import os
import pathlib
import signal
import statistics
import subprocess
import time

ROOT=pathlib.Path(__file__).resolve().parents[2]
spec=importlib.util.spec_from_file_location('baseline_runner',pathlib.Path(__file__).with_name('run_benchmarks.py'))
helpers=importlib.util.module_from_spec(spec);spec.loader.exec_module(helpers)

def run(args, variant, mode, iteration):
    cp=(ROOT/'target/linux-review/classpath.txt').read_text().strip()
    if variant=='before':cp=cp.replace(str(ROOT/'target/classes'),str(args.before_classes.resolve()))
    destination=args.output/f'{variant}-{mode}-{iteration}';destination.mkdir(parents=True)
    log=args.output/f'{variant}-{mode}-{iteration}.log'
    command=['java','-Xms256m','-Xmx1024m',f'-Dpaperword.mathjax.node.command={ROOT}/node_modules/node/bin/node',
             '-Dpaperword.render.cache.enabled=false','-cp',cp,'com.lz.paperword.core.render.LinuxBenchmark',mode,str(destination)]
    started=time.perf_counter();peak=family=0
    with log.open('w') as output:
        process=subprocess.Popen(command,cwd=ROOT,stdout=output,stderr=subprocess.STDOUT,start_new_session=True)
        try:
            while process.poll() is None:
                a,b=helpers.rss_tree(process.pid);peak=max(peak,a);family=max(family,b)
                if time.perf_counter()-started>120:raise TimeoutError('benchmark exceeded120s')
                time.sleep(.025)
            if process.returncode:raise RuntimeError(f'failed: {log}')
        finally:
            try:os.killpg(process.pid,signal.SIGKILL)
            except ProcessLookupError:pass
            process.wait()
    samples=[json.loads(line[6:]) for line in log.read_text().splitlines() if line.startswith('BENCH ')]
    result={'variant':variant,'mode':mode,'fork':iteration,'wall_s':time.perf_counter()-started,
            'peak_jvm_rss_mib':peak/1024,'peak_family_rss_mib':family/1024,'samples':samples}
    if any(s['errors'] for s in samples):raise RuntimeError(f'correctness failure: {log}')
    print(json.dumps(result),flush=True)
    return result

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--before-classes',type=pathlib.Path,required=True)
    parser.add_argument('--output',type=pathlib.Path,required=True);parser.add_argument('--forks',type=int,default=3)
    parser.add_argument('--modes',nargs='+',default=['performance','distinct','paper40']);args=parser.parse_args()
    args.output.mkdir(parents=True,exist_ok=True);results=[]
    for mode in args.modes:
        for i in range(args.forks):
            for variant in (['before','after'] if i%2==0 else ['after','before']):
                results.append(run(args,variant,mode,i))
                (args.output/'raw.json').write_text(json.dumps(results,indent=2)+'\n')
    groups={}
    for r in results:
        for s in r['samples']:groups.setdefault((r['variant'],s['case']),[]).append(s['ms'])
    summary=[]
    for case in sorted({k[1] for k in groups}):
        before=statistics.median(groups['before',case]);after=statistics.median(groups['after',case]);
        summary.append({'case':case,'before_median_ms':before,'after_median_ms':after,'change_percent':(after/before-1)*100,
                        'before_raw_ms':groups['before',case],'after_raw_ms':groups['after',case]})
    (args.output/'summary.json').write_text(json.dumps(summary,indent=2)+'\n')
