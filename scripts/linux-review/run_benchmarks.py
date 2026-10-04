#!/usr/bin/env python3
"""Run already compiled baseline and optimized classes sequentially; no downloads."""
import argparse,json,os,pathlib,subprocess,time
ROOT=pathlib.Path(__file__).resolve().parents[2]
OUT=ROOT/'target/linux-review'

def rss_tree(pid):
    proc={}
    for d in pathlib.Path('/proc').iterdir():
        if not d.name.isdigit():continue
        try:
            v={line.split(':',1)[0]:line.split(':',1)[1].strip() for line in (d/'status').read_text().splitlines() if ':' in line}
            proc[int(d.name)]=(int(v.get('PPid',0)),int(v.get('VmRSS','0 kB').split()[0]))
        except (OSError,ValueError):pass
    own=proc.get(pid,(0,0))[1];family={pid}
    while True:
        expanded=family|{i for i,(parent,_) in proc.items() if parent in family}
        if expanded==family:break
        family=expanded
    return own,sum(proc.get(i,(0,0))[1] for i in family)

def run(variant,mode,runid):
    cp=(OUT/'classpath.txt').read_text().strip()
    if variant=='baseline':cp=cp.replace(str(ROOT/'target/classes'),str(OUT/'baseline-classes'))
    destination=OUT/f'{variant}-{mode}-{runid}';destination.mkdir(exist_ok=True)
    log=OUT/f'{variant}-{mode}-{runid}.log'
    cmd=['java','-Xms256m','-Xmx1024m',f'-Dpaperword.mathjax.node.command={ROOT}/node_modules/node/bin/node','-Dpaperword.render.cache.enabled=false','-cp',cp,'com.lz.paperword.core.render.LinuxBenchmark',mode,str(destination)]
    start=time.perf_counter();peak=family=0
    with log.open('w') as output:
        p=subprocess.Popen(cmd,cwd=ROOT,stdout=output,stderr=subprocess.STDOUT)
        while p.poll() is None:
            a,b=rss_tree(p.pid);peak=max(peak,a);family=max(family,b)
            if time.perf_counter()-start>120:p.kill();raise TimeoutError(f'{variant} {mode} exceeded 120s')
            time.sleep(.025)
        code=p.wait()
    data={'variant':variant,'mode':mode,'run':runid,'exit':code,'wall_s':time.perf_counter()-start,'peak_jvm_rss_mib':peak/1024,'peak_family_rss_mib':family/1024,'samples':[]}
    for line in log.read_text().splitlines():
        if line.startswith('BENCH '):data['samples'].append(json.loads(line[6:]))
    print(json.dumps(data),flush=True)
    if code:raise RuntimeError(f'benchmark failed: {log}')
    return data

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--mode',default='performance');ap.add_argument('--forks',type=int,default=3);a=ap.parse_args();results=[]
    for i in range(a.forks):
        order=['baseline','optimized'] if i%2==0 else ['optimized','baseline']
        for variant in order:results.append(run(variant,a.mode,i))
    (OUT/f'comparison-{a.mode}.json').write_text(json.dumps(results,indent=2)+'\n')
