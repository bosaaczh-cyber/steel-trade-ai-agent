# -*- coding: utf-8 -*-
"""
GMT.AI 钢贸通 — Web 管理后台 v3.1
登录系统 + 公司配置 + 全自动外贸流水线
"""
import os, json, re, csv, time, hashlib, threading
from functools import wraps
from datetime import datetime, date, timedelta
from pathlib import Path
from flask import (Flask, render_template, request, jsonify,
                   send_file, redirect, url_for, session)
from dotenv import load_dotenv
load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "change-this-secret")

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
USERS_FILE = DATA_DIR / "users.json"

# ═══ 用户系统 ═══
def _hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

def load_users():
    if not USERS_FILE.exists():
        default = {"admin": {"password": _hash_pw("gmt2026"), "role": "admin", "created": str(date.today()),
            "company": {"name":"我的钢管公司","name_en":"My Steel Co","location":"河北沧州",
                "capacity":"年产能12万吨","certs":"API 5L / ISO 9001","delivery":"30天交货",
                "email":"sales@company.com","whatsapp":"+86-138-0000-0000","website":"","advantages":""}}}
        save_users(default); return default
    with open(USERS_FILE, "r", encoding="utf-8") as f: return json.load(f)

def save_users(u):
    with open(USERS_FILE, "w", encoding="utf-8") as f: json.dump(u, f, ensure_ascii=False, indent=2)

def get_company():
    u = load_users().get(session.get("username",""))
    return u.get("company",{}) if u else {}

def login_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if "username" not in session: return redirect("/login")
        return f(*a, **kw)
    return dec

def user_data_dir(username=None):
    d = DATA_DIR / (username or session.get("username","default")); d.mkdir(exist_ok=True); return d

def load_send_log(username=None):
    f = user_data_dir(username) / "send_log.json"
    if not f.exists(): return []
    try:
        with open(f,"r",encoding="utf-8") as fh: return json.load(fh)
    except: return []

def save_send_log(log, username=None):
    with open(user_data_dir(username)/"send_log.json","w",encoding="utf-8") as f: json.dump(log,f,ensure_ascii=False,indent=2)

def load_leads_db(username=None):
    f = user_data_dir(username) / "leads_db.json"
    if not f.exists(): return {"leads":[],"followups":[]}
    try:
        with open(f,"r",encoding="utf-8") as fh: return json.load(fh)
    except: return {"leads":[],"followups":[]}

def save_leads_db(db, username=None):
    with open(user_data_dir(username)/"leads_db.json","w",encoding="utf-8") as f: json.dump(db,f,ensure_ascii=False,indent=2)

# ═══ 任务状态 ═══
task_status = {}; task_lock = threading.Lock()
def get_task(name):
    k = f"{session.get('username','x')}:{name}"
    with task_lock:
        if k not in task_status: task_status[k]={"running":False,"progress":0,"message":"","result":None}
        return task_status[k]
def update_task(name, username=None, **kw):
    if not username: username=session.get("username","x")
    k=f"{username}:{name}"
    with task_lock:
        if k not in task_status: task_status[k]={"running":False,"progress":0,"message":"","result":None}
        task_status[k].update(kw)

# ═══ 登录 / 注册 / 登出 ═══
@app.route("/login", methods=["GET","POST"])
def login_page():
    if request.method=="GET":
        return render_template("login.html") if "username" not in session else redirect("/")
    u=request.form.get("username","").strip().lower(); p=request.form.get("password","")
    if not u or not p: return render_template("login.html",error="请填写用户名和密码")
    users=load_users(); user=users.get(u)
    if not user or user["password"]!=_hash_pw(p): return render_template("login.html",error="用户名或密码错误")
    session["username"]=u; session.permanent=True; return redirect("/")

@app.route("/register", methods=["GET","POST"])
def register_page():
    if request.method=="GET": return render_template("register.html")
    u=request.form.get("username","").strip().lower(); p=request.form.get("password","")
    c=request.form.get("confirm",""); cn=request.form.get("company","").strip()
    if not u or not p: return render_template("register.html",error="请填写用户名和密码")
    if len(p)<6: return render_template("register.html",error="密码至少6位")
    if p!=c: return render_template("register.html",error="两次密码不一致")
    users=load_users()
    if u in users: return render_template("register.html",error="用户名已存在")
    users[u]={"password":_hash_pw(p),"role":"user","created":str(date.today()),
        "company":{"name":cn or "我的公司","name_en":"","location":"","capacity":"","certs":"",
            "delivery":"","email":"","whatsapp":"","website":"","advantages":""}}
    save_users(users); session["username"]=u; session.permanent=True; return redirect("/settings")

@app.route("/logout")
def logout(): session.clear(); return redirect("/login")

# ═══ 公司设置 ═══
@app.route("/settings")
@login_required
def settings_page(): return render_template("settings.html",company=get_company())

@app.route("/api/settings", methods=["POST"])
@login_required
def api_save_settings():
    data=request.json or {}; users=load_users(); u=session["username"]
    if u not in users: return jsonify({"error":"用户不存在"}),400
    users[u]["company"]={k:data.get(k,"") for k in ["name","name_en","location","capacity","certs","delivery","email","whatsapp","website","advantages"]}
    save_users(users); return jsonify({"success":True})

# ═══ 首页 ═══
@app.route("/")
@login_required
def dashboard():
    log=load_send_log(); db=load_leads_db()
    ts=len(log); tc=sum(1 for l in log if l.get("status")=="success"); tf=ts-tc
    sr=round(tc/ts*100,1) if ts else 0; tl=len(db.get("leads",[]))
    ds={}
    for e in log:
        d=e.get("date","")
        if d not in ds: ds[d]={"success":0,"failed":0}
        ds[d]["success" if e.get("status")=="success" else "failed"]+=1
    today=date.today(); cd,cs,cf=[],[],[]
    for i in range(6,-1,-1):
        d=str(today-timedelta(days=i)); cd.append(d[-5:])
        s=ds.get(d,{"success":0,"failed":0}); cs.append(s["success"]); cf.append(s["failed"])
    rl=sorted(log,key=lambda x:f"{x.get('date','')}{x.get('time','')}",reverse=True)[:15]
    return render_template("dashboard.html",total_sent=ts,total_success=tc,total_failed=tf,
        success_rate=sr,total_leads=tl,chart_dates=json.dumps(cd),chart_success=json.dumps(cs),
        chart_failed=json.dumps(cf),recent_log=rl,company=get_company())

# ═══ 找客户 ═══
@app.route("/leads")
@login_required
def leads_page():
    db=load_leads_db(); return render_template("leads.html",leads=db.get("leads",[]),task=get_task("find_leads"))

@app.route("/api/find-leads", methods=["POST"])
@login_required
def api_find_leads():
    if get_task("find_leads")["running"]: return jsonify({"error":"任务正在运行中"}),400
    data=request.json or {}; r=data.get("region","middle_east"); ind=data.get("industry","oil_gas")
    kw=data.get("keywords",""); u=session["username"]
    def _run():
        update_task("find_leads",username=u,running=True,progress=10,message="正在搜索...")
        try:
            from steel_master_web import run_find_leads
            res=run_find_leads(r,ind,kw,progress_cb=lambda p,m:update_task("find_leads",username=u,progress=p,message=m))
            db=load_leads_db(u); ex={l.get("company","") for l in db.get("leads",[])}; added=0
            for l in res:
                if l.get("company") not in ex: db["leads"].append(l); added+=1
            save_leads_db(db,u)
            update_task("find_leads",username=u,running=False,progress=100,message=f"完成！找到{len(res)}家，新增{added}家",result=res)
        except Exception as e:
            update_task("find_leads",username=u,running=False,progress=0,message=f"出错：{e}",result=None)
    threading.Thread(target=_run,daemon=True).start(); return jsonify({"status":"started"})

@app.route("/api/leads/<int:idx>", methods=["DELETE"])
@login_required
def api_delete_lead(idx):
    db=load_leads_db(); leads=db.get("leads",[])
    if 0<=idx<len(leads): r=leads.pop(idx); save_leads_db(db); return jsonify({"removed":r.get("company","")})
    return jsonify({"error":"索引越界"}),400

@app.route("/api/leads/add", methods=["POST"])
@login_required
def api_add_lead():
    data=request.json or {}; db=load_leads_db()
    db["leads"].append({k:data.get(k,"") for k in ["company","country","contact","title","email","need","grade"]} | {"status":"new","added_at":str(date.today())})
    save_leads_db(db); return jsonify({"success":True})

# ═══ 邮箱挖掘 ═══
@app.route("/api/mine-email", methods=["POST"])
@login_required
def api_mine_email():
    try:
        from steel_email_finder import find_email
        data=request.json or {}; idx=data.get("idx",-1); co=data.get("company","")
        if not co: return jsonify({"error":"公司名不能为空"}),400
        res=find_email(co,data.get("country",""),data.get("title",""),verbose=False)
        if res.get("email") and idx>=0:
            db=load_leads_db(); leads=db.get("leads",[])
            if 0<=idx<len(leads):
                leads[idx]["email"]=res["email"]; leads[idx]["email_confidence"]=res.get("confidence",0)
                save_leads_db(db)
        return jsonify({"email":res.get("email",""),"confidence":res.get("confidence",0),"source":res.get("source",""),"all_emails":res.get("all_emails",[]),"company":co})
    except ImportError: return jsonify({"error":"steel_email_finder.py 未找到"}),500
    except Exception as e: return jsonify({"error":str(e)}),500

# ═══ 发邮件 ═══
@app.route("/send")
@login_required
def send_page():
    db=load_leads_db(); all_leads=db.get("leads",[])
    leads=[]
    for i,l in enumerate(all_leads):
        if l.get("email") and "@" in l.get("email",""):
            l["_db_idx"]=i; leads.append(l)
    log=load_send_log(); sent={l.get("email","") for l in log if l.get("status")=="success"}
    for l in leads: l["_sent"]=l.get("email","") in sent
    return render_template("send.html",leads=leads,task=get_task("send_email"),company=get_company())

@app.route("/api/send-test", methods=["POST"])
@login_required
def api_send_test():
    try:
        from Steel_sender import send_single_email
        data=request.json or {}; co=get_company(); to=data.get("to","")
        if not to: return jsonify({"error":"请填写收件邮箱"}),400
        cn=co.get("name_en") or co.get("name","GMT.AI")
        return jsonify(send_single_email(to,f"【测试】{cn} 邮件系统测试",f"Hi,\n\nTest from {cn}.\nTime: {datetime.now()}\n\nBest,\n{cn}"))
    except ImportError: return jsonify({"error":"Steel_sender.py 未找到"}),500

@app.route("/api/send-batch", methods=["POST"])
@login_required
def api_send_batch():
    if get_task("send_email")["running"]: return jsonify({"error":"正在发送中"}),400
    data=request.json or {}; targets=data.get("targets",[]); u=session["username"]
    if not targets: return jsonify({"error":"没有目标"}),400
    def _run():
        update_task("send_email",username=u,running=True,progress=0,message="准备发送...")
        try:
            from Steel_sender import send_single_email
            udir=DATA_DIR/u; udir.mkdir(exist_ok=True)
            log_file=udir/"send_log.json"
            sc=fc=0; total=len(targets); td=str(date.today())
            for i,t in enumerate(targets):
                update_task("send_email",username=u,progress=int((i+1)/total*100),message=f"发送中 {i+1}/{total}")
                r=send_single_email(t["email"],t.get("subject",""),t.get("body",""),t.get("contact",""))
                try:
                    with open(log_file,"r",encoding="utf-8") as fh: log=json.load(fh)
                except: log=[]
                log.append({"date":td,"time":datetime.now().strftime("%H:%M:%S"),"company":t.get("company",""),
                    "contact":t.get("contact",""),"email":t["email"],"subject":t.get("subject",""),
                    "status":"success" if r.get("success") else "failed","error":r.get("error",""),"request_id":r.get("request_id","")})
                with open(log_file,"w",encoding="utf-8") as fh: json.dump(log,fh,ensure_ascii=False,indent=2)
                if r.get("success"): sc+=1
                else: fc+=1
                if i<total-1: time.sleep(0.8)
            update_task("send_email",username=u,running=False,progress=100,message=f"完成！成功{sc}封，失败{fc}封")
        except Exception as e: update_task("send_email",username=u,running=False,progress=0,message=f"出错：{e}")
    threading.Thread(target=_run,daemon=True).start(); return jsonify({"status":"started","count":len(targets)})

@app.route("/api/generate-email", methods=["POST"])
@login_required
def api_generate_email():
    try:
        from crewai import Agent, Task, Crew, LLM
        data=request.json or {}; co=get_company()
        mn=co.get("name_en") or co.get("name","Our Company")
        my_llm=LLM(model="openai/claude-sonnet-4-6",api_key=os.environ.get("CLAUDE_API_KEY",""),base_url=os.environ.get("CLAUDE_API_URL",""),temperature=0.4)
        writer=Agent(role="专业外贸开发信撰写专家",goal="撰写高转化率开发信",backstory=f"你是{mn}的资深开发信专家。",llm=my_llm,verbose=False,allow_delegation=False)
        ln="用西班牙语（usted格式）" if data.get("lang")=="es" else "用英语"
        ct=data.get("contact","Procurement Manager")
        task=Task(description=(
            f"为以下客户写开发信（150-200词，{ln}）：\n\n"
            f"收件人：{ct}\n公司：{data.get('company','')},{data.get('country','')}\n需求：{data.get('need','')}\n\n"
            f"发件人：{mn}，{co.get('location','')}\n"
            f"核心优势：{co.get('capacity','')}, {co.get('certs','')}, {co.get('delivery','')}\n"
            f"联系：{co.get('email','')} | WhatsApp: {co.get('whatsapp','')}\n\n"
            f"格式：\nSubject: [...]\nDear {ct},\n[正文]\nBest regards,\n[签名用{mn}]"),
            expected_output="完整开发信",agent=writer)
        result=str(Crew(agents=[writer],tasks=[task],verbose=False).kickoff())
        subj=""; body=[]
        for line in result.split("\n"):
            line=line.strip().replace("**","")
            if line.lower().startswith("subject:"): subj=line.split(":",1)[-1].strip()
            elif line: body.append(line)
        return jsonify({"subject":subj,"body":"\n".join(body)})
    except Exception as e: return jsonify({"error":str(e)}),500

# ═══ MTC ═══
@app.route("/mtc")
@login_required
def mtc_page(): return render_template("mtc.html")

@app.route("/api/mtc/generate", methods=["POST"])
@login_required
def api_mtc_generate():
    try:
        from steel_mtc import fill_mtc
        r=fill_mtc(request.json or {})
        if r: return jsonify({"success":True,"file":r})
        return jsonify({"error":"生成失败"}), 500
    except Exception as e: return jsonify({"error":str(e)}),500

@app.route("/api/mtc/template")
@login_required
def api_mtc_template():
    from steel_mtc import DEMO_ORDER; return jsonify(DEMO_ORDER)

@app.route("/api/task-status/<name>")
@login_required
def api_task_status(n): return jsonify(get_task(n))

@app.route("/download/<path:fn>")
@login_required
def download_file(fn):
    fp=BASE_DIR/fn; return send_file(fp,as_attachment=True) if fp.exists() else ("文件不存在",404)

# ═══ 客户列表 API（无刷新更新）═══
@app.route("/api/leads/list")
@login_required
def api_leads_list():
    db=load_leads_db(); return jsonify({"leads":db.get("leads",[])})

# ═══ 批量 AI 生成 + 群发 ═══
@app.route("/api/batch-generate-send", methods=["POST"])
@login_required
def api_batch_generate_send():
    if get_task("batch_gen_send")["running"]: return jsonify({"error":"批量任务正在运行"}),400
    data=request.json or {}; indices=data.get("indices",[]); lang=data.get("lang","en")
    u=session["username"]; company=get_company()
    if not indices: return jsonify({"error":"没有选中客户"}),400

    def _run():
        update_task("batch_gen_send",username=u,running=True,progress=0,message="准备中...")
        try:
            from crewai import Agent, Task, Crew, LLM
            from Steel_sender import send_single_email
            my_llm=LLM(model="openai/claude-sonnet-4-6",api_key=os.environ.get("CLAUDE_API_KEY",""),
                base_url=os.environ.get("CLAUDE_API_URL",""),temperature=0.4)
            mn=company.get("name_en") or company.get("name","Our Company")
            writer=Agent(role="专业外贸开发信撰写专家",goal="撰写高转化率开发信",backstory=f"你是{mn}的资深开发信专家。",
                llm=my_llm,verbose=False,allow_delegation=False)

            udir=DATA_DIR/u; udir.mkdir(exist_ok=True)
            db_file=udir/"leads_db.json"
            with open(db_file,"r",encoding="utf-8") as fh: db=json.load(fh)
            leads=db.get("leads",[])

            total=len(indices); sc=fc=0; td=str(date.today())
            for i,idx in enumerate(indices):
                if idx<0 or idx>=len(leads): continue
                lead=leads[idx]
                email_addr=lead.get("email","")
                if not email_addr or "@" not in email_addr: fc+=1; continue

                update_task("batch_gen_send",username=u,progress=int((i+0.3)/total*100),
                    message=f"[{i+1}/{total}] AI写信: {lead.get('company','')[:20]}...")
                ct=lead.get("contact","") or "Procurement Manager"
                ln="用西班牙语（usted格式）" if lang=="es" else "用英语"
                task=Task(description=(
                    f"为以下客户写开发信（150-200词，{ln}）：\n\n"
                    f"收件人：{ct}\n公司：{lead.get('company','')},{lead.get('country','')}\n需求：{lead.get('need','')}\n\n"
                    f"发件人：{mn}，{company.get('location','')}\n"
                    f"核心优势：{company.get('capacity','')}, {company.get('certs','')}, {company.get('delivery','')}\n"
                    f"联系：{company.get('email','')} | WhatsApp: {company.get('whatsapp','')}\n\n"
                    f"格式：\nSubject: [...]\nDear {ct},\n[正文]\nBest regards,\n[签名用{mn}]"),
                    expected_output="完整开发信",agent=writer)
                result=str(Crew(agents=[writer],tasks=[task],verbose=False).kickoff())
                subj=""; body=[]
                for line in result.split("\n"):
                    line=line.strip().replace("**","")
                    if line.lower().startswith("subject:"): subj=line.split(":",1)[-1].strip()
                    elif line: body.append(line)
                body_text="\n".join(body)

                update_task("batch_gen_send",username=u,progress=int((i+0.7)/total*100),
                    message=f"[{i+1}/{total}] 发送: {email_addr[:30]}...")
                r=send_single_email(email_addr,subj,body_text,ct)

                log_file=udir/"send_log.json"
                try:
                    with open(log_file,"r",encoding="utf-8") as fh2: log=json.load(fh2)
                except: log=[]
                log.append({"date":td,"time":datetime.now().strftime("%H:%M:%S"),"company":lead.get("company",""),
                    "contact":ct,"email":email_addr,"subject":subj,
                    "status":"success" if r.get("success") else "failed","error":r.get("error",""),"request_id":r.get("request_id","")})
                with open(log_file,"w",encoding="utf-8") as fh2: json.dump(log,fh2,ensure_ascii=False,indent=2)

                if r.get("success"): sc+=1
                else: fc+=1
                if i<total-1: time.sleep(0.8)

            update_task("batch_gen_send",username=u,running=False,progress=100,
                message=f"完成！AI生成{total}封，成功{sc}封，失败{fc}封")
        except Exception as e:
            update_task("batch_gen_send",username=u,running=False,progress=0,message=f"出错：{e}")
    threading.Thread(target=_run,daemon=True).start()
    return jsonify({"status":"started","count":len(indices)})

# ═══ 群开配置 API（国家+行业列表）═══
@app.route("/api/bulk-config")
@login_required
def api_bulk_config():
    countries={
        "middle_east":{"label":"中东","countries":["Saudi Arabia","UAE","Kuwait","Qatar","Oman","Bahrain","Iraq","Iran","Jordan","Lebanon"]},
        "south_america":{"label":"南美洲","countries":["Brazil","Argentina","Chile","Colombia","Peru","Venezuela","Ecuador","Bolivia","Uruguay","Paraguay"]},
        "central_america":{"label":"中美洲/加勒比","countries":["Mexico","Panama","Costa Rica","Guatemala","Honduras","Dominican Republic","Cuba","El Salvador","Nicaragua"]},
        "southeast_asia":{"label":"东南亚","countries":["Vietnam","Indonesia","Thailand","Philippines","Malaysia","Singapore","Myanmar","Cambodia","Laos"]},
        "south_asia":{"label":"南亚","countries":["India","Pakistan","Bangladesh","Sri Lanka","Nepal"]},
        "east_asia":{"label":"东亚","countries":["South Korea","Japan","Taiwan","Mongolia"]},
        "central_asia":{"label":"中亚","countries":["Kazakhstan","Uzbekistan","Turkmenistan","Tajikistan","Kyrgyzstan"]},
        "africa_west":{"label":"西非","countries":["Nigeria","Ghana","Senegal","Ivory Coast","Cameroon","Mali","Burkina Faso"]},
        "africa_east":{"label":"东非","countries":["Kenya","Tanzania","Ethiopia","Uganda","Rwanda","Mozambique","Madagascar"]},
        "africa_north":{"label":"北非","countries":["Egypt","Algeria","Morocco","Tunisia","Libya","Sudan"]},
        "africa_south":{"label":"南部非洲","countries":["South Africa","Angola","Zambia","Zimbabwe","Botswana","Namibia","DRC"]},
        "europe_west":{"label":"西欧","countries":["United Kingdom","Germany","France","Italy","Spain","Netherlands","Belgium","Portugal","Switzerland","Austria"]},
        "europe_east":{"label":"东欧","countries":["Turkey","Poland","Romania","Czech Republic","Hungary","Bulgaria","Serbia","Croatia","Ukraine","Greece"]},
        "europe_north":{"label":"北欧","countries":["Norway","Sweden","Finland","Denmark","Iceland","Estonia","Latvia","Lithuania"]},
        "russia_cis":{"label":"俄罗斯/独联体","countries":["Russia","Belarus","Georgia","Armenia","Azerbaijan","Moldova"]},
        "north_america":{"label":"北美","countries":["United States","Canada"]},
        "oceania":{"label":"大洋洲","countries":["Australia","New Zealand","Papua New Guinea","Fiji"]},
    }
    industries={
        "oil_gas":{"name":"石油天然气","en":"Oil & Gas"},
        "mining":{"name":"矿业","en":"Mining"},
        "construction":{"name":"建筑/EPC承包商","en":"Construction & EPC"},
        "petrochemical":{"name":"石化化工","en":"Petrochemical"},
        "trader":{"name":"钢管贸易商","en":"Steel Pipe Trader"},
        "water":{"name":"水务/水处理","en":"Water & Wastewater"},
        "power":{"name":"电力/能源","en":"Power & Energy"},
        "shipbuilding":{"name":"造船/海工","en":"Shipbuilding & Offshore"},
        "automotive":{"name":"汽车制造","en":"Automotive"},
        "hvac":{"name":"暖通空调","en":"HVAC"},
        "agriculture":{"name":"农业/灌溉","en":"Agriculture & Irrigation"},
        "infrastructure":{"name":"基础设施/桥梁","en":"Infrastructure"},
        "pipeline":{"name":"管道工程","en":"Pipeline"},
        "steel_structure":{"name":"钢结构","en":"Steel Structure"},
        "fire_protection":{"name":"消防","en":"Fire Protection"},
        "furniture":{"name":"家具/装饰管","en":"Furniture & Decoration"},
    }
    return jsonify({"countries":countries,"industries":industries})

# ═══ 群开系统 ═══
@app.route("/api/bulk-outreach", methods=["POST"])
@login_required
def api_bulk_outreach():
    if get_task("bulk")["running"]: return jsonify({"error":"群开任务正在运行"}),400
    data=request.json or {}; region=data.get("region","middle_east")
    industries=data.get("industries",["oil_gas"]); u=session["username"]
    company=get_company()

    def _run():
        update_task("bulk",username=u,running=True,progress=5,message="群开启动...")
        try:
            from steel_bulk_web import run_bulk_outreach
            result=run_bulk_outreach(region,industries,company,
                progress_cb=lambda p,m:update_task("bulk",username=u,progress=p,message=m))
            # 保存结果到用户目录
            with open(user_data_dir(u)/"bulk_result.json","w",encoding="utf-8") as f:
                json.dump(result,f,ensure_ascii=False,indent=2)
            update_task("bulk",username=u,running=False,progress=100,
                message=f"群开完成！{result['total_companies']}家公司",result=result)
        except Exception as e:
            update_task("bulk",username=u,running=False,progress=0,message=f"出错：{e}")
    threading.Thread(target=_run,daemon=True).start()
    return jsonify({"status":"started"})

@app.route("/api/bulk-results")
@login_required
def api_bulk_results():
    f=user_data_dir()/"bulk_result.json"
    if not f.exists(): return jsonify({"error":"还没有群开结果"}),404
    with open(f,"r",encoding="utf-8") as fh: return jsonify(json.load(fh))

@app.route("/api/bulk-import", methods=["POST"])
@login_required
def api_bulk_import():
    """把群开结果导入客户库"""
    data=request.json or {}; items=data.get("companies",[])
    db=load_leads_db(); existing={l.get("company","") for l in db.get("leads",[])}
    added=0
    for c in items:
        if c.get("company") and c["company"] not in existing:
            db["leads"].append({"company":c["company"],"country":c.get("country",""),
                "contact":"","title":"","email":"","website":c.get("website",""),
                "need":c.get("industry",""),"grade":"B","status":"new","added_at":str(date.today())})
            added+=1
    save_leads_db(db)
    return jsonify({"success":True,"added":added})

# ═══ 客户详情 ═══
@app.route("/leads/<int:idx>")
@login_required
def lead_detail(idx):
    db=load_leads_db(); leads=db.get("leads",[])
    if idx<0 or idx>=len(leads): return redirect("/leads")
    lead=leads[idx]; log=load_send_log()
    history=[l for l in log if l.get("email")==lead.get("email") or l.get("company")==lead.get("company")]
    history.sort(key=lambda x:f"{x.get('date','')}{x.get('time','')}",reverse=True)
    return render_template("detail.html",lead=lead,idx=idx,history=history)

@app.route("/api/leads/<int:idx>/update", methods=["POST"])
@login_required
def api_update_lead(idx):
    data=request.json or {}; db=load_leads_db(); leads=db.get("leads",[])
    if idx<0 or idx>=len(leads): return jsonify({"error":"索引越界"}),400
    for k in ["company","country","contact","title","email","need","grade","status","note"]:
        if k in data: leads[idx][k]=data[k]
    save_leads_db(db); return jsonify({"success":True})

# ═══ CSV 导入导出 ═══
@app.route("/api/leads/export-csv")
@login_required
def api_export_csv():
    import io
    db=load_leads_db(); leads=db.get("leads",[])
    output=io.StringIO()
    writer=csv.writer(output)
    writer.writerow(["公司","国家","联系人","职位","邮箱","评级","需求","状态","备注"])
    for l in leads:
        writer.writerow([l.get("company",""),l.get("country",""),l.get("contact",""),
            l.get("title",""),l.get("email",""),l.get("grade",""),l.get("need",""),
            l.get("status",""),l.get("note","")])
    output.seek(0)
    fn=f"leads_export_{date.today()}.csv"
    fp=user_data_dir()/fn
    with open(fp,"w",encoding="utf-8-sig",newline="") as f: f.write(output.getvalue())
    return send_file(fp,as_attachment=True,download_name=fn)

@app.route("/api/leads/import-csv", methods=["POST"])
@login_required
def api_import_csv():
    import io
    if "file" not in request.files: return jsonify({"error":"没有上传文件"}),400
    f=request.files["file"]
    content=f.read().decode("utf-8-sig")
    reader=csv.DictReader(io.StringIO(content))
    db=load_leads_db(); existing={l.get("company","") for l in db.get("leads",[])}; added=0
    for row in reader:
        co=row.get("公司") or row.get("company") or row.get("Company","")
        if co and co not in existing:
            db["leads"].append({"company":co,"country":row.get("国家") or row.get("country",""),
                "contact":row.get("联系人") or row.get("contact",""),
                "title":row.get("职位") or row.get("title",""),
                "email":row.get("邮箱") or row.get("email",""),
                "grade":row.get("评级") or row.get("grade","B"),
                "need":row.get("需求") or row.get("need",""),
                "status":"new","added_at":str(date.today())})
            added+=1
    save_leads_db(db)
    return jsonify({"success":True,"added":added,"total":len(db["leads"])})

@app.context_processor
def inject_globals():
    return {"current_user":session.get("username",""),"company_name":get_company().get("name","") if "username" in session else ""}

if __name__=="__main__":
    print("GMT.AI 钢贸通 — http://127.0.0.1:5000"); app.run(debug=True,host="0.0.0.0",port=5000)
