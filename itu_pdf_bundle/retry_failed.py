#!/usr/bin/env python3
from __future__ import annotations
import base64, csv, hashlib, html, http.cookiejar, json, os, random, re, ssl, time
import urllib.error, urllib.parse, urllib.request, zlib
from pathlib import Path
from typing import Any

MANIFEST_B64='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' 
OUT_ROOT=Path("out")/"ITU-R建议书1200-1400MHz汇总（第二轮补充）"
LOG_DIR=OUT_ROOT/"清单与日志"
GITHUB_TOKEN=os.environ.get("GITHUB_TOKEN","")
UAS=[
 "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36",
 "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/139.0.0.0 Safari/537.36",
 "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/138.0.0.0 Safari/537.36",
]
CTX=ssl.create_default_context()
JAR=http.cookiejar.CookieJar()
OPENER=urllib.request.build_opener(urllib.request.HTTPCookieProcessor(JAR))


def manifest():
 return json.loads(zlib.decompress(base64.b64decode(MANIFEST_B64)).decode("utf-8"))


def uniq(xs):
 out=[]; seen=set()
 for x in xs:
  key=x[1] if isinstance(x,tuple) else x
  if key and key not in seen: seen.add(key); out.append(x)
 return out


def get(url, *, binary=True, attempts=5, referer=None, timeout=150):
 last=None
 for n in range(attempts):
  hdr={"User-Agent":random.choice(UAS),"Accept":"application/pdf,text/html,*/*;q=0.8","Accept-Language":"zh-CN,zh;q=0.9,en;q=0.8","Cache-Control":"no-cache","Connection":"close"}
  if referer: hdr["Referer"]=referer
  if GITHUB_TOKEN and "api.github.com" in url: hdr.update({"Authorization":f"Bearer {GITHUB_TOKEN}","Accept":"application/vnd.github+json","X-GitHub-Api-Version":"2022-11-28"})
  try:
   req=urllib.request.Request(url,headers=hdr)
   with OPENER.open(req,timeout=timeout,context=CTX) if False else OPENER.open(req,timeout=timeout) as r:
    data=r.read()
   if b"Request Rejected" in data[:4096] or b"The requested URL was rejected" in data[:4096]:
    last=RuntimeError("ITU WAF Request Rejected")
    try:
     warm=urllib.request.Request("https://www.itu.int/",headers={"User-Agent":random.choice(UAS),"Accept":"text/html,*/*"})
     with OPENER.open(warm,timeout=30) as rr: rr.read(1024)
    except Exception: pass
    time.sleep(5+4*n+random.random()*3)
    continue
   return data
  except urllib.error.HTTPError as e:
   last=e
   if e.code in (404,410): break
   time.sleep(3+3*n+random.random()*2)
  except Exception as e:
   last=e; time.sleep(3+3*n+random.random()*2)
 raise RuntimeError(f"download failed {url}: {last!r}")


def text(url,**kw):
 b=get(url,**kw)
 for enc in ("utf-8","utf-16","latin-1"):
  try:return b.decode(enc)
  except UnicodeDecodeError:pass
 return b.decode("utf-8","replace")


def metadata(version):
 slug="itu-r-"+version.lower().replace(".","-")+".yaml"
 errs=[]
 for branch in ("v2","master"):
  u=f"https://raw.githubusercontent.com/relaton/relaton-data-itu-r/{branch}/data/{slug}"
  try:
   t=text(u,attempts=3,timeout=60)
   urls=re.findall(r"https://www\.itu\.int/dms_pubrec/[^\s'\"]+?PDF-E\.pdf",t)
   if urls:return uniq([html.unescape(x.rstrip(".,;")) for x in urls]),u,None
   errs.append(u+": no PDF URL")
  except Exception as e: errs.append(u+": "+repr(e))
 return [],None," | ".join(errs)


def scrape(item):
 version=item["version"]; rec=item.get("rec") or version.rsplit("-",1)[0]
 pages=uniq([item.get("en_page"),item.get("zh_page"),f"https://www.itu.int/rec/R-REC-{version}/en",f"https://www.itu.int/rec/R-REC-{version}/zh",f"https://www.itu.int/rec/R-REC-{rec}/en",f"https://www.itu.int/rec/R-REC-{rec}/zh"])
 out=[]
 for p in pages:
  try:t=html.unescape(text(p,attempts=3,timeout=90).replace("\\/","/"))
  except Exception:continue
  for u in re.findall(r"(?:https?://www\.itu\.int)?/dms_pubrec/[^\s'\"<>]+?PDF-[A-Z]\.pdf",t,re.I):
   if u.startswith("/"):u="https://www.itu.int"+u
   out.append(u)
 return uniq(out)


def valid_pdf(b):
 return b.startswith(b"%PDF-") and len(b)>=5000


def try_urls(item,res,cands):
 for lang,u,kind in uniq(cands):
  a={"language":lang,"url":u,"source_kind":kind}
  try:
   b=get(u,attempts=5,timeout=180,referer=item.get("en_page"))
   if not valid_pdf(b):
    a.update(status="invalid",bytes=len(b),header=repr(b[:80]));res["attempts"].append(a);continue
   dest=OUT_ROOT/item["target"];dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(b)
   sha=hashlib.sha256(b).hexdigest();a.update(status="ok",bytes=len(b));res["attempts"].append(a)
   m=re.search(r"R-REC-([A-Z]+\.[0-9]+-[0-9]+)-",u,re.I)
   res.update(status="ok",language=lang,source_kind=kind,source_url=u,bytes=len(b),sha256=sha,saved_path=str(dest),resolved_version=m.group(1) if m else item["version"])
   return True
  except Exception as e:a.update(status="error",error=repr(e));res["attempts"].append(a)
 return False


def one(item):
 res={"row":item["row"],"rec":item["rec"],"version":item["version"],"target":item["target"],"status":"failed","attempts":[]}
 meta,ms,me=metadata(item["version"]);res.update(metadata_source=ms,metadata_error=me,metadata_pdf_urls=meta)
 c=[]
 # Explicit correction for the stale workbook version: current in-force release is RA.611-4 (03/2006).
 if item["version"]=="RA.611-1":
  e="https://www.itu.int/dms_pubrec/itu-r/rec/ra/R-REC-RA.611-4-200603-I!!PDF-E.pdf"
  c += [("C",e.replace("PDF-E.pdf","PDF-C.pdf"),"ITU-corrected"),("E",e,"ITU-corrected")]
 for e in meta:
  if "PDF-E.pdf" in e:c.append(("C",e.replace("PDF-E.pdf","PDF-C.pdf"),"ITU-metadata"))
  c.append(("E",e,"ITU-metadata"))
 if item.get("cn_url"):c.append(("C",item["cn_url"],"ITU-workbook"))
 if item.get("en_url"):c.append(("E",item["en_url"],"ITU-workbook"))
 if try_urls(item,res,c):return res
 sc=[]
 for u in scrape(item):
  up=u.upper();lang="C" if "PDF-C.PDF" in up else ("E" if "PDF-E.PDF" in up else "O")
  sc.append((lang,u,"ITU-page"))
 try_urls(item,res,sc)
 return res


def main():
 items=manifest();LOG_DIR.mkdir(parents=True,exist_ok=True);results=[]
 for i,item in enumerate(items,1):
  r=one(item);results.append(r);print(f"[{i:02d}/{len(items)}] {item['version']} {r['status']} {r.get('language','')} {r.get('bytes',0)}",flush=True)
  time.sleep(1.3+random.random()*1.2)
 ok=sum(r["status"]=="ok" for r in results);summary={"requested":len(items),"downloaded":ok,"failed":len(items)-ok,"chinese":sum(r.get("language")=="C" and r["status"]=="ok" for r in results),"english":sum(r.get("language")=="E" and r["status"]=="ok" for r in results),"total_bytes":sum(r.get("bytes",0) for r in results)}
 (LOG_DIR/"第二轮下载结果.json").write_text(json.dumps({"summary":summary,"results":results},ensure_ascii=False,indent=2),encoding="utf-8")
 with (LOG_DIR/"第二轮下载结果.csv").open("w",encoding="utf-8-sig",newline="") as f:
  fields=["row","rec","version","resolved_version","status","language","source_kind","bytes","sha256","target","source_url","metadata_source","metadata_error"]
  w=csv.DictWriter(f,fieldnames=fields,extrasaction="ignore");w.writeheader();w.writerows(results)
 print(json.dumps(summary,ensure_ascii=False,indent=2),flush=True)
 return 0

if __name__=="__main__":raise SystemExit(main())
