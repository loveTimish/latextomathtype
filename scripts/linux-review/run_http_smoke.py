#!/usr/bin/env python3
"""Cold JVM health/export checks on loopback, synthetic payloads only."""
import argparse,io,json,pathlib,socket,struct,subprocess,time,urllib.request,zipfile,zlib
ROOT=pathlib.Path(__file__).resolve().parents[2];OUT=ROOT/'target/linux-review'
parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--variant',action='append',choices=['baseline','optimized'])
parser.add_argument('--forks',type=int,default=3)
parser.add_argument('--baseline-script',type=pathlib.Path)
args=parser.parse_args();variants=args.variant or ['optimized']
if not 1<=args.forks<=10:parser.error('--forks must be between 1 and 10')
if 'baseline' in variants and (args.baseline_script is None or not args.baseline_script.is_file() or not (OUT/'baseline-classes').is_dir()):
 parser.error('baseline requires baseline-classes and --baseline-script from its original checkout; do not mix renderer bundles')
assets=OUT/'http-assets';assets.mkdir(parents=True,exist_ok=True)
def png_chunk(kind,data):
 return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)
png=assets/'synthetic-image.png'
png.write_bytes(b'\x89PNG\r\n\x1a\n'+png_chunk(b'IHDR',struct.pack('>IIBBBBB',1,1,8,2,0,0,0))+png_chunk(b'IDAT',zlib.compress(b'\x00\x22\x66\xaa'))+png_chunk(b'IEND',b''))
payload=json.dumps({'paper':{'name':'Synthetic Linux HTTP smoke'},'sections':[{'headline':'Test','images':[str(png)],'questions':[{'serialNumber':1,'questionType':6,'content':'$x^2+1$ and $\\frac{1}{2}$'}]}]}).encode()
records=[]
for variant in variants:
 for i in range(args.forks):
  cp=(OUT/'classpath.txt').read_text().strip()
  if variant=='baseline':cp=cp.replace(str(ROOT/'target/classes'),str(OUT/'baseline-classes'))
  with socket.socket() as s:s.bind(('127.0.0.1',0));port=s.getsockname()[1]
  cmd=['java','-Xms256m','-Xmx1024m','-Dserver.address=127.0.0.1',f'-Dserver.port={port}','-Dspring.main.banner-mode=off','-Dlogging.level.root=WARN','-Dlogging.level.com.lz.paperword=WARN','-Dpaperword.render.cache.enabled=false',f'-Dpaperword.assets.root={assets}',f'-Dpaperword.mathjax.node.command={ROOT}/node_modules/node/bin/node','-cp',cp,'com.lz.paperword.PaperToWordApplication']
  if variant=='baseline':cmd.insert(1,f'-Dpaperword.mathjax.script={args.baseline_script.resolve()}')
  with (OUT/f'http-{variant}-{i}.log').open('w') as log:
   start=time.perf_counter();p=subprocess.Popen(cmd,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
   try:
    while True:
     try:
      with urllib.request.urlopen(f'http://127.0.0.1:{port}/api/export/health',timeout=.25) as response:
       assert response.status==200;break
     except Exception:
      if p.poll() is not None or time.perf_counter()-start>30:raise RuntimeError('health startup failed')
      time.sleep(.025)
    health_ms=(time.perf_counter()-start)*1000;exports=[]
    for j in range(2):
     t=time.perf_counter();req=urllib.request.Request(f'http://127.0.0.1:{port}/api/export/word',data=payload,headers={'Content-Type':'application/json'},method='POST')
     with urllib.request.urlopen(req,timeout=30) as response:docx=response.read();assert response.status==200
     elapsed=(time.perf_counter()-t)*1000
     with zipfile.ZipFile(io.BytesIO(docx)) as z:
      names=z.namelist();oles=[x for x in names if x.startswith('word/embeddings/')];wmfs=[x for x in names if x.endswith('.wmf')];pngs=[x for x in names if x.startswith('word/media/') and x.endswith('.png')]
      assert len(oles)==2 and len(wmfs)==2 and len(pngs)==1
      assert z.read(pngs[0])==png.read_bytes()
     (OUT/f'http-{variant}-{i}-{j}.docx').write_bytes(docx);exports.append({'ms':elapsed,'bytes':len(docx),'ole':len(oles),'wmf':len(wmfs),'png':len(pngs)})
    record={'variant':variant,'fork':i,'health_ms':health_ms,'exports':exports};records.append(record);print(json.dumps(record),flush=True)
   finally:
    p.terminate()
    try:p.wait(timeout=10)
    except subprocess.TimeoutExpired:p.kill();p.wait()
(OUT/'http-smoke.json').write_text(json.dumps(records,indent=2)+'\n')
