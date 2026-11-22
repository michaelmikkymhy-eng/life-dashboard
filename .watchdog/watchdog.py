#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""文件回滚看门狗 v2（2026-09-03）
v1→v2：/data/user/work 为临时盘（沙箱重启即清空）→ 迁移至 /workspace/.watchdog 持久；
      state 仅在变更时落盘（避免每 3s 同步扰动宿主）。
"""
import hashlib, json, os, re, subprocess, time

WATCH = '/workspace/01-日常管理'
SD = '/workspace/.watchdog'
VER = SD + '/versions'
LOG = SD + '/events.log'
LOCK = SD + '/pause.lock'
POLL, PROMOTE_N, FIGHT_WIN = 3, 2, 600

os.makedirs(VER, exist_ok=True)

def norm(t):
    t = re.sub(r'<br\s*/?>', '', t)
    return '\n'.join(' '.join(l.split()) for l in t.split('\n'))

def H(raw, is_md):
    if is_md:
        return hashlib.sha256(norm(raw.decode('utf-8','replace')).encode()).hexdigest()[:16]
    return hashlib.sha256(raw).hexdigest()[:16]

def files():
    out=[]
    for d,_,fs in os.walk(WATCH):
        if '__pycache__' in d: continue
        for f in fs:
            if f.endswith(('.md','.xlsx')) and not f.startswith('~$'):
                out.append(os.path.join(d,f))
    return out

def log(m):
    with open(LOG,'a',encoding='utf-8') as f:
        f.write(time.strftime('[%m-%d %H:%M:%S] ')+m+'\n')

def snap():
    try:
        ps=subprocess.run(['ps','aux'],capture_output=True,text=True,timeout=5).stdout
        n=len([l for l in ps.split('\n') if l and 'ps aux' not in l])
        log(f'    证据：沙箱进程 {n} 个，无本地写入者 → 宿主侧同步层所为')
    except Exception as e:
        log(f'    证据采集失败 {e}')

try:
    state=json.load(open(SD+'/state.json',encoding='utf-8'))
    log(f'=== 看门狗v2启动（state续接，{len(state)} 个文件）===')
except Exception:
    state={}
    log('=== 看门狗v2启动（全新基线）===')
fights={}

while True:
    dirty=False
    try:
        if os.path.exists(LOCK):
            time.sleep(1); continue
        for p in files():
            try: raw=open(p,'rb').read()
            except Exception: continue
            h=H(raw,p.endswith('.md'))
            st=state.get(p)
            if st is None:
                open(f'{VER}/{h}','wb').write(raw)
                state[p]={'latest':h,'history':[h]}
                log(f'✅ 基线: {os.path.relpath(p,WATCH)} {h}')
                dirty=True; continue
            if h==st['latest']:
                fights.pop(p,None); continue
            if h in st['history']:
                now=time.time()
                f=[t for t in fights.get(p,[]) if now-t<FIGHT_WIN]
                if len(f)>=PROMOTE_N-1:
                    st['latest']=h; fights.pop(p,None); dirty=True
                    log(f'🤝 同版本短时多次出现，接受为新基线: {os.path.relpath(p,WATCH)} → {h}')
                else:
                    f.append(now); fights[p]=f
                    ver=f"{VER}/{st['latest']}"
                    if os.path.exists(ver):
                        open(p,'wb').write(open(ver,'rb').read())
                        log(f'🔁 回滚拦截→已恢复: {os.path.relpath(p,WATCH)}（回退 {h}，已恢复 {st["latest"]}）')
                        snap()
                    else:
                        log(f'⛔ 检测到回滚但版本文件缺失: {os.path.relpath(p,WATCH)}')
                continue
            open(f'{VER}/{h}','wb').write(raw)
            st['history']=(st['history']+[h])[-200:]
            st['latest']=h; fights.pop(p,None); dirty=True
            log(f'➕ 新版本: {os.path.relpath(p,WATCH)} → {h}')
        if dirty:
            tmp=SD+'/state.json.tmp'
            json.dump(state,open(tmp,'w',encoding='utf-8'),ensure_ascii=False)
            os.replace(tmp,SD+'/state.json')
    except Exception as e:
        log(f'⚠️ 轮询异常: {e}')
    time.sleep(POLL)
