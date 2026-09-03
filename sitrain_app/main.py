from __future__ import annotations

import json
import os
import threading
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from . import __version__, db
from .config import ROOT, ensure_data_dirs
from .documents import render_preview
from .glossary_excel import export_global_terms, parse_global_terms
from .pipeline import add_version, prepare_file, publish, refresh_machine_version, run_translation, start_automatic_translation
from .providers import ProviderError, TranslationRequest, translate
from .storage import checksum, file_dir, save_upload
from .sync import apply_sync, create_mappings
from .terminology import normalize

STATIC = Path(__file__).parent/"static"
SOURCE_LANGUAGES = {"auto", "Chinese", "English", "German", "French", "Italian", "Spanish"}
TARGET_LANGUAGES = {"Chinese", "English"}


@asynccontextmanager
async def lifespan(_app: FastAPI):
    ensure_data_dirs(); db.init_db()
    for task in db.rows("SELECT id FROM tasks WHERE status='queued' AND task_type='translate'"):
        threading.Thread(target=run_translation,args=(task["id"],),daemon=True).start()
    for file in db.rows("SELECT id FROM files WHERE status='preflight'"):
        threading.Thread(target=prepare_file,args=(file["id"],),daemon=True).start()
    for file in db.rows("SELECT id FROM files WHERE status='ready'"):
        threading.Thread(target=start_automatic_translation,args=(file["id"],),daemon=True).start()
    yield


app = FastAPI(title="SITRAIN 教材智能翻译与校对平台",version=__version__,lifespan=lifespan)


class ProjectIn(BaseModel):
    name: str = Field(min_length=1,max_length=120)
    description: str = Field(default="",max_length=1000)


class ChapterIn(BaseModel):
    name: str = Field(min_length=1,max_length=120)


class UnitPatch(BaseModel):
    human_text: str | None = None
    status: str | None = None
    revision: int


class ProviderIn(BaseModel):
    name: str
    protocol: str
    base_url: str = ""
    api_key: str | None = None
    model: str = ""
    enabled: bool = True
    is_default: bool = False
    is_fallback: bool = False
    config: dict = {}


class ProviderCreate(BaseModel):
    name: str = Field(min_length=1, max_length=80)
    base_url: str = Field(min_length=8, max_length=500)
    model: str = Field(min_length=1, max_length=160)
    api_key: str | None = None


class TermIn(BaseModel):
    project_id: str | None = None
    source: str
    target: str
    aliases: list[str] = []
    chinese_aliases: list[str] = []
    mode: str = "required"
    case_sensitive: bool = False


class MappingPatch(BaseModel):
    slide_number: int = Field(ge=1)
    confirmed: bool = True


@app.get("/api/health")
def health():
    active = db.row("SELECT name,model,api_key,last_test_status FROM providers WHERE enabled=1 AND is_default=1 LIMIT 1")
    if not active:
        llm = {"status":"neutral", "label":"LLM 未选择"}
    elif not active["api_key"]:
        llm = {"status":"neutral", "label":"LLM 未配置"}
    else:
        state = active.get("last_test_status", "untested")
        status = {"success":"online", "failed":"error"}.get(state, "warning")
        text = {"online":"已验证", "error":"连接失败", "warning":"待测试"}[status]
        llm = {"status":status, "label":active["name"]}
    return {"ok":True,"version":__version__,"llm":llm}


@app.get("/api/dashboard")
def dashboard():
    return {"projects":db.rows("SELECT p.*,(SELECT COUNT(*) FROM chapters c WHERE c.project_id=p.id) chapter_count FROM projects p ORDER BY updated_at DESC"),
            "files":db.rows("SELECT * FROM files ORDER BY updated_at DESC LIMIT 100"),
            "tasks":db.rows("SELECT * FROM tasks WHERE status IN ('queued','running') ORDER BY updated_at DESC LIMIT 50"),
            "counts":{"pending_review":db.scalar("SELECT COUNT(*) FROM files WHERE status='review'"),
                      "published":db.scalar("SELECT COUNT(*) FROM files WHERE status='published'")}}


@app.post("/api/projects")
def project_create(body: ProjectIn):
    pid,now = db.new_id(),db.utcnow()
    db.execute("INSERT INTO projects(id,name,description,created_at,updated_at) VALUES(?,?,?,?,?)",(pid,body.name,body.description,now,now))
    db.execute("INSERT INTO project_members(project_id,user_id,role) VALUES(?,?,?)",(pid,"local-admin","owner"))
    return db.row("SELECT * FROM projects WHERE id=?",(pid,))


@app.get("/api/projects/{project_id}")
def project_get(project_id: str):
    project = db.row("SELECT * FROM projects WHERE id=?",(project_id,))
    if not project: raise HTTPException(404,"教材项目不存在")
    project["chapters"] = db.rows("SELECT * FROM chapters WHERE project_id=? ORDER BY position,created_at",(project_id,))
    for chapter in project["chapters"]:
        chapter["files"] = db.rows("SELECT * FROM files WHERE chapter_id=? ORDER BY created_at",(chapter["id"],))
    return project


@app.post("/api/projects/{project_id}/chapters")
def chapter_create(project_id: str,body: ChapterIn):
    if not db.row("SELECT id FROM projects WHERE id=?",(project_id,)): raise HTTPException(404,"教材项目不存在")
    cid,now = db.new_id(),db.utcnow(); pos=db.scalar("SELECT COALESCE(MAX(position),-1)+1 FROM chapters WHERE project_id=?",(project_id,))
    db.execute("INSERT INTO chapters(id,project_id,name,position,created_at,updated_at) VALUES(?,?,?,?,?,?)",(cid,project_id,body.name,pos,now,now))
    return db.row("SELECT * FROM chapters WHERE id=?",(cid,))


@app.post("/api/files/upload")
async def upload(background: BackgroundTasks,upload: UploadFile=File(...),project_id: str|None=Form(None),chapter_id: str|None=Form(None),source_language: str=Form("auto"),target_language: str=Form("Chinese")):
    if source_language not in SOURCE_LANGUAGES:
        raise HTTPException(422, "源语言仅支持自动识别、中文、英文、德文、法文、意大利文或西班牙文")
    if target_language not in TARGET_LANGUAGES:
        raise HTTPException(422, "目标语言仅支持中文或英文")
    if source_language != "auto" and source_language == target_language:
        raise HTTPException(422, "源语言与目标语言不能相同；混合语言文档请使用自动识别")
    if bool(project_id) != bool(chapter_id): raise HTTPException(400,"教材模式必须同时指定项目和章节")
    if chapter_id and not db.row("SELECT id FROM chapters WHERE id=? AND project_id=?",(chapter_id,project_id)): raise HTTPException(404,"章节不存在或不属于该项目")
    fid,now = db.new_id(),db.utcnow(); path=await save_upload(upload,fid); ext=path.suffix.lower()
    kind={".pptx":"powerpoint",".ppt":"powerpoint",".docx":"word",".doc":"word",".xlsx":"excel",".xls":"excel",".pdf":"pdf"}.get(ext,"text")
    db.execute("INSERT INTO files(id,project_id,chapter_id,original_name,extension,kind,mode,status,source_language,target_language,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
               (fid,project_id,chapter_id,path.name,ext,kind,"chapter" if chapter_id else "standalone","preflight",source_language,target_language,now,now))
    add_version(fid,"original",path,{"immutable":True})
    background.add_task(prepare_file,fid)
    return {"id":fid,"status":"preflight","name":path.name}


@app.get("/api/files/{file_id}")
def file_get(file_id: str):
    item=db.row("SELECT * FROM files WHERE id=?",(file_id,))
    if not item: raise HTTPException(404,"文件不存在")
    item["versions"]=db.rows("SELECT id,version_type,sequence,size,checksum,metadata_json,created_at FROM file_versions WHERE file_id=? ORDER BY created_at",(file_id,))
    item["unit_counts"]={r["status"]:r["n"] for r in db.rows("SELECT status,COUNT(*) n FROM translation_units WHERE file_id=? GROUP BY status",(file_id,))}
    return item


@app.get("/api/files/{file_id}/units")
def units(file_id: str,status: str|None=None):
    sql="SELECT * FROM translation_units WHERE file_id=?"; params=[file_id]
    if status: sql+=" AND status=?"; params.append(status)
    return {"units":db.rows(sql+" ORDER BY page,CAST(stable_key AS INTEGER),stable_key",params)}


@app.get("/api/files/{file_id}/preview/{page_number}")
def file_preview(file_id: str,page_number: int):
    file=db.row("SELECT * FROM files WHERE id=?",(file_id,))
    if not file: raise HTTPException(404,"文件不存在")
    version=db.row("SELECT * FROM file_versions WHERE file_id=? AND version_type IN ('published','machine','working','original') ORDER BY CASE version_type WHEN 'published' THEN 1 WHEN 'machine' THEN 2 WHEN 'working' THEN 3 ELSE 4 END,sequence DESC LIMIT 1",(file_id,))
    try: path=render_preview(Path(version["path"]),page_number,file_dir(file_id)/"previews")
    except RuntimeError as exc: raise HTTPException(409,str(exc)) from exc
    return FileResponse(path,media_type="image/png")


@app.patch("/api/units/{unit_id}")
def unit_patch(unit_id: str,body: UnitPatch,background: BackgroundTasks):
    current=db.row("SELECT * FROM translation_units WHERE id=?",(unit_id,))
    if not current: raise HTTPException(404,"翻译单元不存在")
    if current["revision"]!=body.revision: raise HTTPException(409,"该译文已被其他修改更新，请刷新后重试")
    text=body.human_text if body.human_text is not None else current.get("human_text")
    status=body.status or ("modified" if body.human_text is not None else current["status"])
    db.execute("UPDATE translation_units SET human_text=?,status=?,revision=revision+1,updated_at=? WHERE id=?",(text,status,db.utcnow(),unit_id))
    if status=="confirmed" and text:
        file=db.row("SELECT project_id FROM files WHERE id=?",(current["file_id"],))
        db.execute("INSERT INTO translation_memory(id,project_id,source_norm,source,target,origin_unit_id,approved_at) VALUES(?,?,?,?,?,?,?) ON CONFLICT(project_id,source_norm) DO UPDATE SET target=excluded.target,origin_unit_id=excluded.origin_unit_id,approved_at=excluded.approved_at",
                   (db.new_id(),file.get("project_id"),normalize(current["source_text"]),current["source_text"],text,unit_id,db.utcnow()))
    background.add_task(refresh_machine_version, current["file_id"])
    return db.row("SELECT * FROM translation_units WHERE id=?",(unit_id,))


@app.post("/api/files/{file_id}/translate")
def translate_start(file_id: str,background: BackgroundTasks,provider_id: str|None=None):
    file=db.row("SELECT * FROM files WHERE id=?",(file_id,))
    if not file or file["status"] not in {"ready","review","failed"}: raise HTTPException(409,"文件尚未完成预检或当前状态不可翻译")
    key=f"translate:{file_id}:{db.scalar('SELECT COALESCE(MAX(revision),0) FROM translation_units WHERE file_id=?',(file_id,))}:{provider_id or 'default'}"
    existing=db.row("SELECT * FROM tasks WHERE idempotency_key=?",(key,))
    if existing and existing["status"] in {"queued","running","completed"}: return existing
    tid,now=db.new_id(),db.utcnow(); db.execute("INSERT INTO tasks(id,file_id,task_type,status,idempotency_key,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(tid,file_id,"translate","queued",key,now,now))
    background.add_task(run_translation,tid,provider_id)
    return db.row("SELECT * FROM tasks WHERE id=?",(tid,))


@app.post("/api/files/{file_id}/publish")
def file_publish(file_id: str):
    try: vid=publish(file_id)
    except RuntimeError as exc: raise HTTPException(409,str(exc)) from exc
    return {"version_id":vid}


@app.get("/api/versions/{version_id}/download")
def download(version_id: str):
    version=db.row("SELECT * FROM file_versions WHERE id=?",(version_id,))
    if not version: raise HTTPException(404,"版本不存在")
    path=Path(version["path"])
    if not path.exists(): raise HTTPException(410,"版本文件已丢失")
    return FileResponse(path,filename=path.name)


@app.get("/api/providers")
def provider_list():
    records=db.rows("SELECT * FROM providers WHERE id<>'custom' ORDER BY CASE id WHEN 'qwen-mt' THEN 1 WHEN 'deepseek' THEN 2 WHEN 'kimi' THEN 3 ELSE 4 END,name")
    for r in records:
        r["api_key_configured"]=bool(r.pop("api_key", "")); r.pop("api_key",None)
        r["is_active"] = bool(r["enabled"] and r["is_default"])
    return {"providers":records}


@app.put("/api/providers/{provider_id}")
def provider_save(provider_id: str,body: ProviderIn):
    current=db.row("SELECT * FROM providers WHERE id=?",(provider_id,))
    if not current or provider_id == "custom": raise HTTPException(404,"Provider 不存在")
    key=current["api_key"] if body.api_key is None or body.api_key=="" else body.api_key
    activate = bool(body.enabled and body.is_default)
    if activate and not key:
        raise HTTPException(422,"请先填写 API Key，才能设为当前翻译服务")
    with db.connect() as conn:
        if activate:
            conn.execute("UPDATE providers SET enabled=0,is_default=0 WHERE id<>?",(provider_id,))
        conn.execute("UPDATE providers SET name=?,protocol=?,base_url=?,api_key=?,model=?,enabled=?,is_default=?,is_fallback=?,config_json=?,updated_at=? WHERE id=?",
                     (body.name,body.protocol,body.base_url,key,body.model,int(activate),int(activate),int(body.is_fallback),json.dumps(body.config,ensure_ascii=False),db.utcnow(),provider_id))
    return {"ok":True}


@app.post("/api/providers")
def provider_create(body: ProviderCreate):
    provider_id = f"custom-{db.new_id()[:12]}"
    db.execute("INSERT INTO providers(id,name,protocol,base_url,api_key,model,enabled,is_default,is_fallback,config_json,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
               (provider_id,body.name,"openai_chat",body.base_url.rstrip("/"),body.api_key or "",body.model,0,0,0,
                json.dumps({"temperature":0.1,"batch_size":20,"concurrency":2,"rpm":0,"timeout":120,"retries":3,"backoff":2,"structured_output":True,"headers":{},"prompt":"Translate each item faithfully from {source_lang} into {target_lang}. Return only a JSON object using the same keys."},ensure_ascii=False),db.utcnow()))
    return {"id":provider_id}


@app.delete("/api/providers/{provider_id}")
def provider_delete(provider_id: str):
    provider = db.row("SELECT id,is_default FROM providers WHERE id=?", (provider_id,))
    if not provider or not provider_id.startswith("custom-"):
        raise HTTPException(404, "仅可删除自定义 Provider")
    db.execute("DELETE FROM providers WHERE id=?", (provider_id,))
    return {"ok": True, "was_active": bool(provider["is_default"])}


@app.post("/api/providers/{provider_id}/test")
def provider_test(provider_id: str):
    provider = db.row("SELECT * FROM providers WHERE id=?",(provider_id,))
    if not provider or provider_id == "custom": raise HTTPException(404,"Provider 不存在")
    if not provider.get("api_key"):
        raise HTTPException(422,"尚未配置 API Key")
    try:
        translate(provider,TranslationRequest([("connection-check","Connection test")],"English","Chinese",[],[]))
    except ProviderError as exc:
        db.execute("UPDATE providers SET last_test_status='failed',last_tested_at=?,last_test_message=?,updated_at=? WHERE id=?",
                   (db.utcnow(),str(exc)[:300],db.utcnow(),provider_id))
        raise HTTPException(502,f"连接失败：{exc}") from exc
    db.execute("UPDATE providers SET last_test_status='success',last_tested_at=?,last_test_message='',updated_at=? WHERE id=?",
               (db.utcnow(),db.utcnow(),provider_id))
    return {"ok":True,"message":"连接成功"}


@app.get("/api/glossary")
def glossary(project_id: str|None=None):
    return {"terms":db.rows("SELECT * FROM glossary_terms WHERE project_id IS ? ORDER BY LENGTH(source) DESC",(project_id,))}


@app.post("/api/glossary")
def glossary_add(body: TermIn):
    tid=db.new_id(); db.execute("INSERT INTO glossary_terms(id,project_id,source,target,aliases_json,chinese_aliases_json,mode,case_sensitive,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                               (tid,body.project_id,body.source,body.target,json.dumps(body.aliases,ensure_ascii=False),json.dumps(body.chinese_aliases,ensure_ascii=False),body.mode,int(body.case_sensitive),db.utcnow()))
    return db.row("SELECT * FROM glossary_terms WHERE id=?",(tid,))


@app.get("/api/glossary/export")
def glossary_export():
    terms = db.rows("SELECT * FROM glossary_terms WHERE project_id IS NULL ORDER BY LENGTH(source) DESC")
    return StreamingResponse(iter([export_global_terms(terms)]),
                             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": "attachment; filename=SITRAIN-global-glossary.xlsx"})


@app.post("/api/glossary/import")
async def glossary_import(upload: UploadFile = File(...)):
    if Path(upload.filename or "").suffix.lower() != ".xlsx":
        raise HTTPException(415, "请上传 .xlsx 格式的术语表")
    terms, errors = parse_global_terms(await upload.read())
    existing = {normalize(term["source"]) for term in db.rows("SELECT source FROM glossary_terms WHERE project_id IS NULL")}
    added = skipped = 0
    with db.connect() as conn:
        for term in terms:
            source = str(term["source"])
            key = normalize(source)
            if key in existing:
                skipped += 1
                continue
            conn.execute("INSERT INTO glossary_terms(id,project_id,source,target,aliases_json,chinese_aliases_json,mode,case_sensitive,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                         (db.new_id(), None, source, str(term["target"]), json.dumps(term["aliases"], ensure_ascii=False), json.dumps(term["chinese_aliases"], ensure_ascii=False), "required", 0, db.utcnow()))
            existing.add(key)
            added += 1
    return {"added": added, "skipped": skipped, "invalid": len(errors), "errors": errors[:20]}


@app.post("/api/chapters/{chapter_id}/sync/map")
def sync_map(chapter_id: str):
    try: return {"mappings":create_mappings(chapter_id)}
    except RuntimeError as exc: raise HTTPException(409,str(exc)) from exc


@app.post("/api/chapters/{chapter_id}/sync/apply")
def sync_apply(chapter_id: str):
    try: return {"version_id":apply_sync(chapter_id)}
    except RuntimeError as exc: raise HTTPException(409,str(exc)) from exc


@app.patch("/api/mappings/{mapping_id}")
def sync_mapping_confirm(mapping_id: str, body: MappingPatch):
    mapping=db.row("SELECT * FROM ppt_word_mappings WHERE id=?",(mapping_id,))
    if not mapping: raise HTTPException(404,"页面映射不存在")
    db.execute("UPDATE ppt_word_mappings SET slide_number=?,confirmed=?,confidence=CASE WHEN ? THEN MAX(confidence,0.99) ELSE confidence END,updated_at=? WHERE id=?",
               (body.slide_number,int(body.confirmed),int(body.confirmed),db.utcnow(),mapping_id))
    return db.row("SELECT * FROM ppt_word_mappings WHERE id=?",(mapping_id,))


@app.get("/api/chapters/{chapter_id}/package")
def chapter_package(chapter_id: str):
    chapter=db.row("SELECT * FROM chapters WHERE id=?",(chapter_id,))
    if not chapter: raise HTTPException(404,"章节不存在")
    versions=db.rows("SELECT v.* FROM file_versions v JOIN files f ON f.id=v.file_id WHERE f.chapter_id=? AND v.version_type='published' AND v.sequence=(SELECT MAX(v2.sequence) FROM file_versions v2 WHERE v2.file_id=v.file_id AND v2.version_type='published')",(chapter_id,))
    if not versions: raise HTTPException(409,"章节尚无发布版本")
    target=file_dir(chapter_id)/f"{chapter['name']}_交付包.zip"
    if target.exists(): target=target.with_name(f"{target.stem}_{db.new_id()[:8]}.zip")
    with zipfile.ZipFile(target,"x",zipfile.ZIP_DEFLATED) as archive:
        for v in versions: archive.write(v["path"],Path(v["path"]).name)
        report={"chapter":chapter["name"],"generated_at":db.utcnow(),"files":[{"name":Path(v["path"]).name,"sha256":v["checksum"]} for v in versions]}
        archive.writestr("QA-交付清单.json",json.dumps(report,ensure_ascii=False,indent=2))
    return FileResponse(target,filename=target.name)


app.mount("/",StaticFiles(directory=STATIC,html=True),name="static")
